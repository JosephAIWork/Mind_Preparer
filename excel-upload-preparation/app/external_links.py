"""External links: what a workbook reads from another workbook, and the
values Excel keeps for it.

A formula such as ='[1]Other sheet'!$D$13*H21 reads a cell of another
file. Excel stores a copy of every external cell it has read
(xl/externalLinks/externalLinkN.xml) and shows those copies when the file
cannot be reached -- a link to a file on a colleague's machine can leave
tens of thousands of formulas living on that cache. Mind cannot follow an
external link at all: every such formula is an error there, and the package
part makes the converter report "external links".

`read_external_links` reads the links and their caches; `freeze_formula`
writes the cached value into the formula in place of the reference
(a number, a text, TRUE/FALSE), or NA() where Excel has no copy (that
reference is #REF! in Excel already). Excel's values do not change.
"""
from __future__ import annotations

import posixpath
import re
import urllib.parse
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .formula_utils import mask_strings

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
# '[1]Sheet name'!$D$13  or  [1]Sheet!D13  -- the reference may be a range
EXT_REF_RE = re.compile(r"(?:'\[(\d+)\]((?:[^']|'')+)'|\[(\d+)\]([A-Za-z0-9_.]+))!(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)")


def _target_path(target: str, workbook: Path) -> tuple[str, bool]:
    """The file a link points to, as a path, and whether it can be reached from here."""
    t = urllib.parse.unquote(target)
    if t.lower().startswith("file:///"):
        t = t[8:]
    try:
        p = Path(t)
        if not p.is_absolute():
            p = (workbook.parent / t).resolve()
        return str(p), p.exists()
    except OSError:
        return t, False


def _cell_value(cell: ET.Element) -> Any:
    v = cell.find("m:v", NS)
    text = v.text if v is not None else None
    kind = cell.get("t", "n")
    if kind == "str":
        return text or ""
    if kind == "b":
        return text == "1"
    if kind == "e":
        return text or "#N/A"
    if text is None or text == "":
        return None
    try:
        f = float(text)
    except ValueError:
        return text
    return int(f) if f.is_integer() and "e" not in text.lower() and "." not in text else f


def read_external_links(path: str | Path) -> list[dict[str, Any]]:
    """Every external link of the workbook, in the [1], [2], ... order formulas use:
    {index, part, target, file_name, reachable, sheets, cache: {(sheet, "D13"): value}, cached_cells}."""
    path = Path(path)
    out: list[dict[str, Any]] = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            return out
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")).findall(f"{{{REL_NS}}}Relationship")}
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        refs = wb.find("m:externalReferences", NS)
        if refs is None:
            return out
        for index, ref in enumerate(refs.findall("m:externalReference", NS), start=1):
            rid = ref.get(f"{{{NS['r']}}}id")
            part = rels.get(rid) or ""
            part = posixpath.normpath(posixpath.join("xl", part)) if not part.startswith("/") else part.lstrip("/")
            link: dict[str, Any] = {"index": index, "part": part, "target": None, "file_name": None, "reachable": False, "sheets": [], "cache": {}, "cached_cells": 0}
            if part in names:
                root = ET.fromstring(z.read(part))
                book = root.find("m:externalBook", NS)
                if book is not None:
                    sheets = [s.get("val") for s in book.findall("m:sheetNames/m:sheetName", NS)]
                    link["sheets"] = sheets
                    for data in book.findall("m:sheetDataSet/m:sheetData", NS):
                        sheet_id = int(data.get("sheetId", "0"))
                        sheet = sheets[sheet_id] if sheet_id < len(sheets) else str(sheet_id)
                        for cell in data.iter(f"{{{NS['m']}}}cell"):
                            link["cache"][(sheet, cell.get("r", "").replace("$", "").upper())] = _cell_value(cell)
                    link["cached_cells"] = len(link["cache"])
                    rid2 = book.get(f"{{{NS['r']}}}id")
                    part_rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
                    if part_rels in names:
                        targets = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read(part_rels)).findall(f"{{{REL_NS}}}Relationship")}
                        target = targets.get(rid2)
                        if target:
                            link["target"], link["reachable"] = _target_path(target, path)
                            link["file_name"] = posixpath.basename(urllib.parse.unquote(target).replace("\\", "/"))
            out.append(link)
    return out


def _literal(value: Any) -> str:
    if value is None:
        return "NA()"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(value)
    if isinstance(value, str):
        if value.startswith("#"):
            return "NA()"
        return '"' + value.replace('"', '""') + '"'
    return "NA()"


def external_references(formula: str) -> list[dict[str, Any]]:
    """Every external reference in a formula: {start, end, index, sheet, ref, range}."""
    if not formula or not formula.startswith("="):
        return []
    masked = mask_strings(formula, mask_sheet_quotes=False)
    out = []
    for m in EXT_REF_RE.finditer(masked):
        index = int(m.group(1) or m.group(3))
        sheet = (m.group(2) or m.group(4) or "").replace("''", "'")
        ref = m.group(5)
        out.append({"start": m.start(), "end": m.end(), "index": index, "sheet": sheet, "ref": ref, "range": ":" in ref})
    return out


def freeze_formula(formula: str, links: list[dict[str, Any]]) -> tuple[str | None, list[str], str | None]:
    """The formula with each external reference replaced by Excel's cached value
    -> (rewritten formula or None, what was written, why not)."""
    refs = external_references(formula)
    if not refs:
        return None, [], None
    by_index = {link["index"]: link for link in links}
    notes: list[str] = []
    out = formula
    for r in reversed(refs):
        if r["range"]:
            return None, notes, f"a range of the other workbook ({r['ref']}): no single value to write"
        link = by_index.get(r["index"])
        if link is None:
            return None, notes, f"link [{r['index']}] is not in the workbook"
        key = (r["sheet"], r["ref"].replace("$", "").upper())
        if key in link["cache"]:
            value = link["cache"][key]
            literal = _literal(value)
            notes.append(f"{r['sheet']}!{key[1]} = {literal[:40]}")
        else:
            literal = "NA()"
            notes.append(f"{r['sheet']}!{key[1]}: no copy kept, #REF! in Excel -> NA()")
        out = out[: r["start"]] + literal + out[r["end"] :]
    return out, notes, None
