"""Workbook preparation (1.4.0, extended 1.5.0): turn findings into concrete,
reviewable *operations* and apply the ones the user approves to a fresh copy
through Excel.

Design (SKILL.md non-negotiables #1, #4, #6, #7):
  * `plan_actions()` never touches a file: every action lists exactly which
    cells / sheets it would change, with before/after, and says what it had
    to skip and why.
  * `apply_operations()` copies the source (hash first), opens the copy in
    Excel once, applies the operations, saves, verifies the result opens in
    Excel, and appends the change log sidecar. openpyxl is only a fallback
    for the value/formula operations (it can't keep references intact on
    structural edits, so those refuse without Excel).
  * Nothing here converts formulas to values or removes VBA/links/names/
    formatting. Content is only cleared when it is *moved* (a label that
    becomes a grid title) or when the user explicitly asked the assistant
    for it.

Operation shape: {op, action_id, rule_id, sheet, before, after, note, ...}
  set_value / set_formula / clear_cell   -> cell ("B3" or a range "B3:D3")
  rename_sheet                           -> after = new name
  insert_row / insert_column             -> row / column
  set_sheet_visibility                   -> after = True|False
  unprotect_sheet, explicit_colors (cells)
Cell operations flagged `after_inserts` run after the row/column inserts
(their coordinates are post-insert).
"""
from __future__ import annotations

import difflib
import json
import re
import time
from pathlib import Path
from typing import Any

import openpyxl

from .change_apply import append_change_log, plan_loop_case_fix
from .excel_com import close_quietly, com_available, excel_session, open_for_write, save_in_place, verify_opens_in_excel
from .formula_utils import (
    cell_refs_in_formula,
    col_to_num,
    collapse_na_reference_calls,
    describe_reference_problem,
    mask_strings,
    num_to_col,
    parse_ref,
    ref_text,
    reference_arg_problems,
)
from .grids import all_grids, grid_containing, looks_like_title, parse_flags
from .inventory import cell_value, make_immutable_copy, sha256_of, values_cell
from .theme_colors import explicit_theme_colors_in_package
from .validators.io import EXPORT_COLUMNS
from .validators.kb import documented_flags

STRUCTURAL_OPS = {"rename_sheet", "insert_row", "insert_column", "explicit_colors", "unprotect_sheet", "set_sheet_visibility"}
CELL_OPS = {"set_value", "set_formula", "clear_cell", "set_array_formula"}
HIDE_MARKER = "&&Hide"
EXCEL_SHEET_NAME_MAX = 31
TITLE_NAME_MAX = 60

SPECIAL_HEADERS = {
    "exportsettings": EXPORT_COLUMNS,
    "projectsettings": ["Name", "Value", "Locked", "Hidden"],
    "parameters": ["Label", "Type", "PossibleValues"],
    "inputsettings": ["GridName", "FileNamePattern"],
    "calculationsteps": [],
}


def _findings_by_id(validation_report: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {f["rule_id"]: f for f in validation_report["findings"]}


def _op(op: str, action_id: str, rule_id: str, sheet: str, **kw: Any) -> dict[str, Any]:
    return {"op": op, "action_id": action_id, "rule_id": rule_id, "sheet": sheet, **kw}


def _is_text(v: Any) -> bool:
    return isinstance(v, str) and v.strip() != "" and not v.startswith("=")


# --- individual action planners -------------------------------------------------
def plan_hide_marker(analysis, report) -> tuple[list[dict], list[str]]:
    """STR-007: hidden sheets get the '&&Hide' marker in their name (Excel
    updates every formula reference when a sheet is renamed through COM)."""
    ops, skipped = [], []
    for s in analysis["workbooks"][0]["sheets"]:
        if s["state"] == "visible" or HIDE_MARKER.lower() in s["name"].lower():
            continue
        base = s["name"]
        new_name = f"{base} {HIDE_MARKER}"
        if len(new_name) > EXCEL_SHEET_NAME_MAX:
            new_name = f"{base[: EXCEL_SHEET_NAME_MAX - len(HIDE_MARKER) - 1].rstrip()} {HIDE_MARKER}"
        ops.append(_op("rename_sheet", "hide_marker", "STR-007", s["name"], before=s["name"], after=new_name, note="hidden sheet without the documented marker"))
    return ops, skipped


def plan_loop_name_case(analysis, report) -> tuple[list[dict], list[str]]:
    """LOOP-002 + RES-002 (case-only mismatches): rewrite loop-name string
    arguments to the casing the loop is defined with."""
    ops, skipped = [], []
    renames, edits = plan_loop_case_fix(analysis)
    seen = set()
    for e in edits:
        ops.append(_op("set_formula", "loop_name_case", "LOOP-002", e["sheet"], cell=e["cell"], before=e["before"], after=e["after"], note=f"renames {renames}"))
        seen.add((e["sheet"], e["cell"]))
    res = _findings_by_id(report).get("RES-002", {})
    for site in (res.get("observed") or {}).get("case_mismatch", []):
        key = (site["sheet"], site["cell"])
        if key in seen:
            continue
        formulas = [f for f in analysis["workbooks"][0]["formulas"] if f["sheet"] == site["sheet"] and f["cell"] == site["cell"]]
        if not formulas:
            continue
        before = formulas[0]["formula"]
        after = before.replace(f'"{site["name"]}"', f'"{site["defined_as"]}"')
        if after != before:
            ops.append(_op("set_formula", "loop_name_case", "RES-002", site["sheet"], cell=site["cell"], before=before, after=after, note=f'"{site["name"]}" -> "{site["defined_as"]}"'))
            seen.add(key)
    return ops, skipped


def _title_edit(g: dict[str, Any], new_title: str, action_id: str, rule_id: str, note: str) -> dict[str, Any]:
    return _op("set_value", action_id, rule_id, g["sheet"], cell=g["title_cell"], before=g["title"], after=new_title, note=note)


def plan_reorder_input_flag(analysis, report) -> tuple[list[dict], list[str]]:
    """INP-005: a /Reorder grid must carry /Input."""
    ops, skipped = [], []
    for g in all_grids(analysis):
        if "reorder" in g["flag_names"] and "input" not in g["flag_names"] and g.get("title_cell"):
            ops.append(_title_edit(g, g["title"].rstrip() + " /Input", "reorder_input_flag", "INP-005", "/Reorder requires /Input (KB 'Input manager reorder columns')"))
    return ops, skipped


def plan_flag_spelling(analysis, report) -> tuple[list[dict], list[str]]:
    """FLG-001: an undocumented flag that is a near-miss of a documented one
    is corrected to the documented spelling; anything ambiguous is skipped."""
    ops, skipped = [], []
    docs = documented_flags()
    if not docs:
        return ops, ["references/mind-flags.yaml missing"]
    names = {n: spec["name"] for n, spec in docs.items()}
    for g in all_grids(analysis):
        if not g.get("title_cell"):
            continue
        title = g["title"]
        new_title = title
        for fl in g["flags"]:
            if fl["name"] in names:
                continue
            matches = difflib.get_close_matches(fl["name"], list(names), n=2, cutoff=0.72)
            if len(matches) == 1 or (len(matches) > 1 and difflib.SequenceMatcher(None, fl["name"], matches[0]).ratio() - difflib.SequenceMatcher(None, fl["name"], matches[1]).ratio() > 0.08):
                fixed = fl["raw"].replace("/" + fl["raw_name"], "/" + names[matches[0]], 1)
                new_title = new_title.replace(fl["raw"], fixed, 1)
            else:
                skipped.append(f"{g['sheet']}!{g['title_cell']} {fl['raw']}: no unambiguous documented flag")
        if new_title != title:
            ops.append(_title_edit(g, new_title, "flag_spelling", "FLG-001", "undocumented flag corrected to the documented spelling"))
    return ops, skipped


def plan_special_headers(analysis, report) -> tuple[list[dict], list[str]]:
    """EXP-003 / PRJ-002 / PAR-002 / INP-003: headers of the documented special
    grids that are near-misses of the documented column names."""
    ops, skipped = [], []
    for g in all_grids(analysis):
        for flag, columns in SPECIAL_HEADERS.items():
            if flag not in g["flag_names"] or not columns:
                continue
            lower = {c.lower(): c for c in columns}
            for idx, h in enumerate(g.get("header_values", [])):
                if not isinstance(h, str) or h.strip().lower() in lower or h.startswith("="):
                    continue
                positional = columns[idx] if idx < len(columns) and flag in ("projectsettings", "parameters", "inputsettings") else None
                matches = difflib.get_close_matches(h.strip().lower(), list(lower), n=1, cutoff=0.6)
                target = lower[matches[0]] if matches else positional
                if target is None or (positional and target != positional and flag != "exportsettings"):
                    skipped.append(f"{g['sheet']}!{g['ref']} header '{h}' (col {idx + 1}): no documented column matches")
                    continue
                cell = ref_text(g["first_col"] + idx, g["first_row"])
                rule = {"exportsettings": "EXP-003", "projectsettings": "PRJ-002", "parameters": "PAR-002", "inputsettings": "INP-003"}[flag]
                ops.append(_op("set_value", "special_headers", rule, g["sheet"], cell=cell, before=h, after=target, note=f"/{flag} column header"))
    return ops, skipped


def _row_insert_is_safe(grids: list[dict[str, Any]], sheet: str, row: int, own: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    """Grids on `sheet` that a whole-row insert at `row` would cut in two."""
    return [o for o in grids if o["sheet"] == sheet and o is not own and o["first_row"] < row <= o["last_row"]]


# --- what the model's formulas read (1.6.8) ------------------------------------------
# A title is text written into a cell; a moved label is a cell emptied; a row
# insert shifts every cell below it. None of that may change a computed value.
# The real Shlomo model showed how it can: `=COUNTA(CoverageDetails!A:A)` counted
# a title written into column A and the model picked a different policy, and
# `=Information!A1` on five sheets read 0 after the company name was "moved"
# into a title. So every write site asks this index first.
_REF_INDEX_CACHE: dict[str, dict[str, Any]] = {}

# 1.8.0 -- two ways a formula reads a sheet that Excel does not move along when a row
# is inserted there. Both changed computed values on a real model (PVFP: 2,462 cells,
# 1,781 of them into errors) after an ordinary Prep:
#   * a 3-D reference, =SUM('LoB 1:>>'!K65): it reads the SAME cell on every sheet of the
#     range, by position. A row inserted on one of them shifts that sheet alone.
#   * an address built as text, =INDIRECT("'"&$E$2&"'!H"&n): text is not a reference.
_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')
_THREE_D_QUOTED_RE = re.compile(r"'((?:[^']|'')+)'!")
_THREE_D_BARE_RE = re.compile(r"(?<![A-Za-z0-9_.'])([A-Za-z0-9_.]+):([A-Za-z0-9_.]+)!")
_INDIRECT_RE = re.compile(r"(?<![A-Za-z0-9_.])INDIRECT\s*\(", re.IGNORECASE)


def _sheet_order(wb: dict[str, Any]) -> list[str]:
    """Sheet names in tab order, the sheets left out of the scan included."""
    names = [s["name"] for s in wb.get("sheets", [])]
    for ig in sorted(wb.get("ignored_sheets", []), key=lambda s: s.get("index", 0)):
        names.insert(min(int(ig.get("index", len(names))), len(names)), ig["name"])
    return names


def _three_d_sheets(formula: str, order: list[str]) -> list[tuple[str, list[str]]]:
    """[("First:Last", [every sheet from First to Last in tab order])] for each 3-D reference."""
    text = _STRING_LITERAL_RE.sub('""', formula)
    pairs = []
    for m in _THREE_D_QUOTED_RE.finditer(text):
        inner = m.group(1).replace("''", "'")
        if ":" in inner and "[" not in inner:
            pairs.append(tuple(inner.split(":", 1)))
    pairs += [(m.group(1), m.group(2)) for m in _THREE_D_BARE_RE.finditer(text)]
    lower = [n.lower() for n in order]
    out = []
    for a, b in pairs:
        if a.lower() in lower and b.lower() in lower:
            i, j = sorted((lower.index(a.lower()), lower.index(b.lower())))
            out.append((f"{a}:{b}", order[i:j + 1]))
    return out


def _text_addressed_sheets(analysis: dict[str, Any], formula: str, home: str | None, order: list[str], seen_values: dict[tuple[str, int, int], Any]) -> list[str]:
    """The sheets an INDIRECT in `formula` may build an address on: a sheet named
    in one of its text pieces ("'CoC "&n&"'!F14" names every sheet starting with
    'CoC '), or named by a cell the formula reads ($E$2 holds 'LoB 1'). A cell
    that holds some other text names a sheet that does not exist: that INDIRECT
    points nowhere and no sheet is at stake. When nothing can be read at all:
    the formula's own sheet for an address without a sheet, every sheet
    otherwise -- unknown is not safe."""
    literals = [m.group(0)[1:-1].replace('""', '"') for m in _STRING_LITERAL_RE.finditer(formula)]
    found: list[str] = []
    for lit in literals:
        low = lit.lower()
        piece = low.lstrip("'").split("!")[0]
        for name in order:
            n = name.lower()
            if n in low or (len(piece) >= 3 and "'" not in piece and n.startswith(piece) and low.lstrip().startswith("'")):
                if name not in found:
                    found.append(name)
    names = {n.lower(): n for n in order}
    names_nothing_here = False
    for ref in cell_refs_in_formula(formula):
        if ref.get("cells") != 1:
            continue
        sheet = ref.get("sheet") or home
        if not sheet or ":" in sheet:
            continue
        key = (sheet, ref["r1"], ref["c1"])
        if key not in seen_values:
            seen_values[key] = values_cell(analysis, sheet, ref["r1"], ref["c1"])
        v = seen_values[key]
        if isinstance(v, str) and v.strip():
            if v.strip().lower() not in names:
                names_nothing_here = True
            elif names[v.strip().lower()] not in found:
                found.append(names[v.strip().lower()])
    if found or names_nothing_here:
        return found
    if any("!" in lit for lit in literals):
        return list(order)
    return [home] if home else []


def reference_index(analysis: dict[str, Any]) -> dict[str, Any]:
    """Every cell range the workbook's formulas and defined names read:
    {'rects': {sheet: [(r1, r2, c1, c2, owner)]}, 'whole_cols': {sheet: {col: owner}},
     'whole_rows': {sheet: {row: owner}}} where owner names the reading formula,
    plus (1.8.0) 'positional': {sheet: owner} for the sheets inside a 3-D reference and
    'text_addressed': {sheet: owner} for the sheets an INDIRECT builds addresses on."""
    key = str(analysis.get("analysis_id") or id(analysis)) + ":" + str(analysis["workbooks"][0].get("copy_path"))
    cached = _REF_INDEX_CACHE.get(key)
    if cached is not None:
        return cached
    rects: dict[str, list[tuple[int, int, int, int, str]]] = {}
    whole_cols: dict[str, dict[int, str]] = {}
    whole_rows: dict[str, dict[int, str]] = {}
    positional: dict[str, str] = {}
    text_addressed: dict[str, str] = {}
    wb = analysis["workbooks"][0]
    order = _sheet_order(wb)
    seen_values: dict[tuple[str, int, int], Any] = {}
    sources: list[tuple[str, str | None, str]] = [(f["formula"], f["sheet"], f"{f['sheet']}!{f['cell']}") for f in wb.get("formulas", []) if isinstance(f.get("formula"), str)]
    for d in wb.get("defined_names", []):
        val = str(d.get("value") or "")
        if val and "#REF!" not in val:
            sources.append((val if val.startswith("=") else "=" + val, None, f"name {d.get('name')}"))
    for formula, home, owner in sources:
        spans: dict[str, list[str]] = {}
        if ":" in formula and "!" in formula:
            for label, sheets in _three_d_sheets(formula, order):
                spans[label.lower()] = sheets
                for name in sheets:
                    positional.setdefault(name, f"{owner} ('{label}')")
        if "INDIRECT" in formula.upper() and _INDIRECT_RE.search(_STRING_LITERAL_RE.sub('""', formula)):
            for name in _text_addressed_sheets(analysis, formula, home, order, seen_values):
                text_addressed.setdefault(name, owner)
        for ref in cell_refs_in_formula(formula):
            sheet = ref.get("sheet") or home
            if not sheet:
                continue
            # a 3-D reference reads the same cells on every sheet of its range
            for sheet in spans.get(sheet.lower(), [sheet]):
                if ref.get("whole_column"):
                    cols = whole_cols.setdefault(sheet, {})
                    for c in range(ref["c1"], ref["c2"] + 1):
                        cols.setdefault(c, owner)
                elif ref.get("whole_row"):
                    rows = whole_rows.setdefault(sheet, {})
                    for r in range(ref["r1"], ref["r2"] + 1):
                        rows.setdefault(r, owner)
                else:
                    rects.setdefault(sheet, []).append((ref["r1"], ref["r2"], ref["c1"], ref["c2"], owner))
    index = {"rects": rects, "whole_cols": whole_cols, "whole_rows": whole_rows, "positional": positional, "text_addressed": text_addressed}
    _REF_INDEX_CACHE.clear()  # one workbook at a time is plenty
    _REF_INDEX_CACHE[key] = index
    return index


def referenced_by(index: dict[str, Any], sheet: str, row: int, col: int) -> str | None:
    """The formula (or name) that reads this cell, else None."""
    owner = index["whole_cols"].get(sheet, {}).get(col) or index["whole_rows"].get(sheet, {}).get(row)
    if owner:
        return owner
    for r1, r2, c1, c2, own in index["rects"].get(sheet, []):
        if r1 <= row <= r2 and c1 <= col <= c2:
            return own
    return None


def whole_reference_to(index: dict[str, Any], sheet: str) -> str | None:
    """A formula that reads a whole column or row of `sheet` -- inserting a
    row there shifts what it counts or indexes."""
    cols = index["whole_cols"].get(sheet) or {}
    rows = index["whole_rows"].get(sheet) or {}
    return next(iter(cols.values()), None) or next(iter(rows.values()), None)


def insert_blocker(index: dict[str, Any], sheet: str) -> str | None:
    """Why a row must not be inserted on `sheet`, in words -- or None when
    nothing reads it in a way an insert would shift."""
    whole = whole_reference_to(index, sheet)
    if whole:
        return f"inserting a row would shift what {whole} counts or indexes over whole columns of this sheet"
    three_d = index.get("positional", {}).get(sheet)
    if three_d:
        return f"inserting a row would shift what the 3-D reference in {three_d} reads here: it finds its cells by position, on every sheet of its range"
    text = index.get("text_addressed", {}).get(sheet)
    if text:
        return f"inserting a row would shift what {text} reads here through an address built as text (INDIRECT)"
    return None


def plan_separate_merged_grids(analysis, report) -> tuple[list[dict], list[str]]:
    """STR-001: a '#Title' trapped inside a grid gets an empty row inserted
    above it -- only when no other grid on the sheet spans that row (a
    whole-row insert would split it too)."""
    ops, skipped = [], []
    grids = all_grids(analysis)
    index = reference_index(analysis)
    for g in grids:
        for cell in g.get("inner_title_cells", []):
            ref = parse_ref(cell)
            if not ref:
                continue
            row = ref["r1"]
            if row <= g["first_row"]:
                # 1.7.2: a '#Title' on the grid's own first row (next to a label or
                # a header) is not separated by a row above it -- the whole block
                # just moves down and the next scan finds the same trapped title.
                # This was an endless Prep round; it is a manual job.
                skipped.append(f"{g['sheet']}!{cell}: the title shares the grid's first row with other content -- an empty row above would not separate it; move the title to its own row by hand")
                continue
            others = _row_insert_is_safe(grids, g["sheet"], row, own=g)
            if others:
                skipped.append(f"{g['sheet']}!{cell}: inserting a row would also split {', '.join(o['display_name'] for o in others[:3])} -- separate manually")
                continue
            blocker = insert_blocker(index, g["sheet"])
            if blocker:
                skipped.append(f"{g['sheet']}!{cell}: {blocker} -- separate manually")
                continue
            ops.append(_op("insert_row", "separate_merged_grids", "STR-001", g["sheet"], row=row, before=f"'{cell_value(analysis, g['sheet'], row, ref['c1'])}' directly below grid {g['display_name']}", after=f"empty row {row} inserted; title moves to row {row + 1}", note="Excel shifts every reference automatically"))
    return ops, skipped


# --- context-aware grid titles (1.5.0) ---------------------------------------------
def _clean_name(text: Any) -> str:
    s = str(text).strip()
    s = s.lstrip("#").strip()
    s = re.sub(r"\s*/[A-Za-z][A-Za-z0-9_.()]*", " ", s)  # a '/Flag' token would be read as a flag
    s = s.replace("&&", " ")
    s = re.sub(r"[\r\n\t]+", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip(" :-–—")
    return s[:TITLE_NAME_MAX].strip()


def _unique_name(name: str, used: set[str]) -> str:
    base = name or "Grid"
    candidate = base
    n = 2
    while candidate.lower() in used:
        candidate = f"{base} ({n})"
        n += 1
    used.add(candidate.lower())
    return candidate


def _section_header_name(analysis: dict[str, Any], g: dict[str, Any]) -> tuple[str, str] | None:
    """The section heading written in the row *directly* above the grid.

    Real models label a block with a heading row that is itself part of a
    small grid (a section number next to a section title, or a row label
    beside the header row), so it is neither blank nor a 'standalone' text
    cell and the label searches below never see it. Preference:
      1. the cell just left of the grid on that row  ('Mortality' beside
         the 'Age | MNS | MS' header row),
      2. the leftmost text cell inside the grid's own columns
         ('Demographic Assumptions' above the block it heads).
    Formulas, numbers and existing '#Titles' are never used.
    """
    r0, c0, c1 = g["first_row"], g["first_col"], g["last_col"]
    if r0 <= 1:
        return None
    row = r0 - 1
    candidates = [c0 - 1] if c0 > 1 else []
    candidates += list(range(c0, min(c1, c0 + 3) + 1))
    for c in candidates:
        v = cell_value(analysis, g["sheet"], row, c)
        if _is_text(v) and not looks_like_title(v):
            name = _clean_name(v)
            if name:
                return name, f"section heading above at {ref_text(c, row)}"
    return None


def _context_name(analysis: dict[str, Any], g: dict[str, Any], labels: list[dict[str, Any]], consumed: set[tuple[str, str]]) -> tuple[str, str, dict[str, Any] | None]:
    """(name, source, consumed_label) for an untitled grid, from the cells around it:
    the section heading directly above it, a standalone label above it (a blank
    row or more between), a standalone label to its left, its text header row,
    else sheet + anchor."""
    r0, c0, c1 = g["first_row"], g["first_col"], g["last_col"]
    sheet = g["sheet"]
    heading = _section_header_name(analysis, g)
    if heading:
        return heading[0], heading[1], None
    above = [
        lab for lab in labels
        if lab["sheet"] == sheet and (lab["sheet"], lab["cell"]) not in consumed
        and r0 - 4 <= lab["row"] <= r0 - 2 and c0 - 1 <= lab["col"] <= c1
    ]
    if above:
        lab = max(above, key=lambda x: (x["row"], -abs(x["col"] - c0)))
        name = _clean_name(lab["text"])
        if name:
            return name, f"label above at {lab['cell']}", lab
    left = [
        lab for lab in labels
        if lab["sheet"] == sheet and (lab["sheet"], lab["cell"]) not in consumed
        and r0 <= lab["row"] <= r0 + 1 and c0 - 4 <= lab["col"] <= c0 - 2
    ]
    if left:
        lab = max(left, key=lambda x: (x["col"], -x["row"]))
        name = _clean_name(lab["text"])
        if name:
            return name, f"label to the left at {lab['cell']}", lab
    headers = [h for h in g.get("header_values", []) if _is_text(h)]
    if g.get("header_is_all_text") and headers:
        name = _clean_name(" ".join(str(h).strip() for h in headers[:3]) + (" ..." if len(headers) > 3 else ""))
        if name:
            return name, "header row", None
    if g["n_cols"] == 1 and _is_text(g.get("header_values", [None])[0]) and g["n_rows"] > 1:
        name = _clean_name(g["header_values"][0])
        if name:
            return name, "first cell", None
    return _clean_name(f"{sheet} {g['anchor']}"), "sheet name + position", None


def standalone_labels(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    """Loose text cells that are not themselves '#Titles' -- the pool the
    naming heuristics draw a grid's name from."""
    labels: list[dict[str, Any]] = []
    for s in analysis["workbooks"][0]["sheets"]:
        for c in s.get("standalone_text_cells", []):
            ref = parse_ref(c["cell"])
            if ref and not c.get("looks_like_title"):
                labels.append({"sheet": s["name"], "cell": c["cell"], "row": ref["r1"], "col": ref["c1"], "text": c["text"]})
    return labels


def deterministic_name(analysis: dict[str, Any], grid: dict[str, Any], labels: list[dict[str, Any]] | None = None) -> tuple[str, str]:
    """(name, source) the heuristics alone would give a grid -- what the
    assistant's proposal is compared against."""
    try:
        name, source, _ = _context_name(analysis, grid, labels if labels is not None else standalone_labels(analysis), set())
        return name, source
    except Exception as exc:
        return "", f"unavailable ({exc})"


def is_weak_name(name: str | None, sheet: str) -> bool:
    """True when a grid name carries no meaning: absent, the '<Sheet> <Anchor>'
    fallback, or a scrap of a header row ('I', 'I I ...'). These are the names
    worth sending to the assistant for a context-aware alternative."""
    text = (name or "").strip()
    if not text:
        return True
    m = re.fullmatch(r"(.+?) ([A-Z]{1,3}\d+)", text)
    if m and m.group(1) == sheet:
        return True
    if len(text) <= 2:
        return True
    return bool(re.fullmatch(r"[I|\s.]+", text))


def plan_create_grid_titles(
    analysis,
    report,
    names: dict[str, str] | None = None,
    exclude: set[str] | None = None,
    reserved: set[str] | list[str] | None = None,
    pre_titled: set[tuple[str, str]] | None = None,
    pre_inserted: dict[str, set[int]] | None = None,
) -> tuple[list[dict], list[str]]:
    """STR-004 / STR-002 / STR-001: every untitled grid gets exactly one
    '#Name' title with a name taken from its surroundings.

    Cases, in order:
      A. caption above a table -- a text cell whose right neighbour is empty
         while the row below starts a block that is wider: Mind (and this
         app) read the caption as a one-column grid that swallows the table's
         first column, and the rest of the table as a second grid ("two
         headers for one range"). Prepending '#' turns the caption into the
         title and the table below becomes one grid.
      B. the cell above the grid is empty -> write '#<name>' there; the name
         comes from a nearby standalone label (which is then moved into the
         title), else the header row, else sheet + position.
      C. the grid starts on row 1 -> insert a row first (only when that
         cuts no other grid), then write the title.
    Names are unique across the workbook; a grid that already has a title
    or contains a trapped title (STR-001) is never given another one.

    `names` (optional) maps "Sheet!Ref" to a name proposed by the assistant
    (app/grid_naming.py). It is only ever consulted when the deterministic
    name would be a weak one, and it goes through the same cleaning and
    uniqueness rules, so it can change *which* name is written but never
    where it is written or whether the operation is safe.

    `exclude` / `reserved` / `pre_titled` / `pre_inserted` come from the Grid
    Namer (plan_named_areas): grids the user handled by hand are left alone,
    their chosen names are already taken, and their write sites and row
    inserts count as used, so one combined apply never collides."""
    ops: list[dict[str, Any]] = []
    skipped: list[str] = []
    exclude = exclude or set()
    grids = all_grids(analysis)
    index = reference_index(analysis)
    used = {g["name"].strip().lower() for g in grids if g.get("name")}
    used |= {str(n).strip().lower() for n in (reserved or ()) if str(n).strip()}
    labels = standalone_labels(analysis)
    consumed: set[tuple[str, str]] = set()
    titled_cells: set[tuple[str, str]] = set(pre_titled or ())
    inserted_rows: dict[str, set[int]] = {k: set(v) for k, v in (pre_inserted or {}).items()}
    by_sheet: dict[str, list[dict[str, Any]]] = {}
    for g in grids:
        by_sheet.setdefault(g["sheet"], []).append(g)

    # A. captions that split a table
    caption_handled: set[int] = set()
    for g in grids:
        if g.get("name") or g["n_cols"] != 1 or not _is_text(g.get("header_values", [None])[0]):
            continue
        r0, c0 = g["first_row"], g["first_col"]
        right = [h for h in by_sheet.get(g["sheet"], []) if h is not g and h["first_row"] == r0 + 1 and h["first_col"] == c0 + 1]
        if not right:
            continue
        if f"{g['sheet']}!{g['ref']}" in exclude or f"{right[0]['sheet']}!{right[0]['ref']}" in exclude:
            continue  # the user named this pair (or one of it) in the Grid Namer
        caption = str(g["header_values"][0])
        reader = referenced_by(index, g["sheet"], r0, c0)
        if reader:
            skipped.append(f"{g['sheet']}!{g['anchor']}: the caption is read by {reader} -- turning it into a title would change that result")
            caption_handled.add(id(g))
            caption_handled.add(id(right[0]))
            continue
        name = _unique_name(_clean_name(caption), used)
        ops.append(_op("set_value", "create_grid_titles", "STR-001", g["sheet"], cell=g["anchor"], before=caption, after=f"#{name}", note=f"caption above the table {right[0]['ref']}: becomes its title so the table is one grid"))
        titled_cells.add((g["sheet"], g["anchor"]))
        caption_handled.add(id(g))
        caption_handled.add(id(right[0]))

    # B / C. every other untitled grid
    for g in grids:
        if g.get("name") or id(g) in caption_handled or f"{g['sheet']}!{g['ref']}" in exclude:
            continue
        if g.get("inner_title_cells"):
            skipped.append(f"{g['sheet']}!{g['ref']}: contains a trapped '#Title' (STR-001) -- separate the grids first")
            continue
        r0, c0 = g["first_row"], g["first_col"]
        name, source, label = _context_name(analysis, g, labels, consumed)
        proposed = _clean_name((names or {}).get(f"{g['sheet']}!{g['ref']}", ""))
        if proposed and is_weak_name(name, g["sheet"]):
            name, source = proposed, "the assistant, from the surrounding cells"
        name = _unique_name(name, used)
        title = f"#{name}"
        above_cell = ref_text(c0, r0 - 1) if r0 > 1 else None
        above_val = cell_value(analysis, g["sheet"], r0 - 1, c0) if r0 > 1 else None
        above_is_free = (
            r0 > 1
            and (g["sheet"], above_cell) not in titled_cells
            and (above_val is None or (isinstance(above_val, str) and above_val.strip() == ""))
            and grid_containing(by_sheet.get(g["sheet"], []), r0 - 1, c0) is None
        )
        if above_is_free:
            reader = referenced_by(index, g["sheet"], r0 - 1, c0)
            if reader:
                skipped.append(f"{g['sheet']}!{g['ref']}: the cell above ({above_cell}) is read by {reader} -- a title there would change that result")
                continue
            ops.append(_op("set_value", "create_grid_titles", "STR-004", g["sheet"], cell=above_cell, before=None, after=title, note=f"name from {source}"))
            titled_cells.add((g["sheet"], above_cell))
        else:
            # The cell above is taken (a section heading, another grid, another
            # grid's title) or the grid starts on row 1: make room with a whole
            # row insert -- but only when no other grid spans that row and no
            # formula counts or indexes whole columns of this sheet.
            reason = "starts on row 1" if r0 == 1 else f"the cell above ({above_cell}) is not free"
            already = r0 in inserted_rows.get(g["sheet"], set())
            if not already:
                cutting = _row_insert_is_safe(grids, g["sheet"], r0, own=g)
                if cutting:
                    skipped.append(f"{g['sheet']}!{g['ref']}: {reason} and inserting a row would cut {cutting[0]['display_name']}")
                    continue
                blocker = insert_blocker(index, g["sheet"])
                if blocker:
                    skipped.append(f"{g['sheet']}!{g['ref']}: {reason} and {blocker}")
                    continue
            reader = referenced_by(index, g["sheet"], r0, c0)
            if reader:
                skipped.append(f"{g['sheet']}!{g['ref']}: {reason} and the title row would fall inside the range {reader} reads")
                continue
            if not already:
                before = f"grid {g['ref']} starts on row 1" if r0 == 1 else f"'{above_val}' sits directly above grid {g['ref']}"
                ops.append(_op("insert_row", "create_grid_titles", "STR-004", g["sheet"], row=r0, before=before, after=f"empty row {r0} inserted (the grid moves down one row)", note="makes room for the title; Excel shifts every reference automatically"))
                inserted_rows.setdefault(g["sheet"], set()).add(r0)
            ops.append(_op("set_value", "create_grid_titles", "STR-004", g["sheet"], cell=ref_text(c0, r0), before=None, after=title, note=f"name from {source}; written after the row insert", after_inserts=True, insert_at=r0))
            titled_cells.add((g["sheet"], ref_text(c0, r0)))
        if label is not None:
            lref = parse_ref(label["cell"])
            reader = referenced_by(index, label["sheet"], lref["r1"], lref["c1"]) if lref else None
            if reader:
                # the label's text still names the grid, but the cell stays: something reads it
                skipped.append(f"{label['sheet']}!{label['cell']}: label kept in place (its text became the title of {g['sheet']}!{g['ref']}) because {reader} reads it")
            else:
                consumed.add((label["sheet"], label["cell"]))
                ops.append(_op("clear_cell", "create_grid_titles", "STR-002", label["sheet"], cell=label["cell"], before=label["text"], after=None, note=f"label moved into the title of {g['sheet']}!{g['ref']}"))

    for lab in labels:
        if (lab["sheet"], lab["cell"]) not in consumed and not (lab["sheet"], lab["cell"]) in titled_cells:
            skipped.append(f"standalone text '{str(lab['text'])[:30]}' at {lab['sheet']}!{lab['cell']} is not next to a grid -- Mind ignores it; move it or delete it by hand")
    return ops, skipped


def _canonical_flags(flags: Any, docs: dict[str, dict[str, Any]]) -> tuple[list[str], list[str]]:
    """User-typed flags -> ('Input', 'Resize.A', ...) with the documented casing
    for the name part; anything that does not parse as exactly one '/Flag'
    token comes back in the second list."""
    out: list[str] = []
    bad: list[str] = []
    seen: set[str] = set()
    for f in flags or []:
        text = str(f).strip().lstrip("/").strip()
        if not text:
            continue
        parsed = parse_flags("/" + text)
        if len(parsed) != 1 or parsed[0]["raw"] != "/" + text:
            bad.append(str(f))
            continue
        spec = docs.get(parsed[0]["name"])
        if spec:
            text = spec["name"] + text[len(parsed[0]["raw_name"]):]
        if text.lower() in seen:
            continue
        seen.add(text.lower())
        out.append(text)
    return out, bad


def plan_named_areas(analysis: dict[str, Any], areas: list[dict[str, Any]]) -> dict[str, Any]:
    """The Grid Namer screen (1.7.0): the user selected an area of a sheet,
    chose a name and flags, and submitted. Naming cannot redraw Mind's grid
    boundaries -- only label what its detection already sees -- so each area
    is resolved to the detected grid it touches (exactly one, or the
    documented caption-above-a-table pair) and the '#Name /Flags' title is
    written with the same reference-safety rules as create_grid_titles: a
    write site some formula reads is refused, with the reader named.

    areas: [{"sheet", "ref", "name", "flags": ["Input", ...]}], ref being the
    user's selection (a cell or a range) in that sheet.

    Returns {"ops", "skipped", "named" (grid keys "Sheet!Ref" a title was
    planned for), "exclude" (grid keys the automatic pass must leave alone,
    including refused ones -- retrying them would only repeat the refusal),
    "reserved" (names now taken), "titled_cells", "inserted_rows"} -- the last
    four feed plan_create_grid_titles so one combined apply never collides."""
    ops: list[dict[str, Any]] = []
    skipped: list[str] = []
    named: list[str] = []
    exclude: set[str] = set()
    reserved: list[str] = []
    titled_cells: set[tuple[str, str]] = set()
    inserted_rows: dict[str, set[int]] = {}
    grids = all_grids(analysis)
    index = reference_index(analysis)
    docs = documented_flags()
    by_sheet: dict[str, list[dict[str, Any]]] = {}
    for g in grids:
        by_sheet.setdefault(g["sheet"], []).append(g)

    # Resolve every area first: a grid the user is renaming releases its old name.
    resolved: list[dict[str, Any]] = []
    for area in areas:
        sheet = str(area.get("sheet") or "")
        label = f"{sheet}!{area.get('ref')}"
        r = parse_ref(str(area.get("ref") or "").split("!")[-1])
        if r is None:
            skipped.append(f"{label}: not a cell or range reference")
            continue
        if sheet not in by_sheet:
            skipped.append(f"{label}: no grids detected on this sheet")
            continue
        name = _clean_name(area.get("name") or "")
        if not name:
            skipped.append(f"{label}: the name is empty once cleaned up -- give the area a real name")
            continue
        flags, bad = _canonical_flags(area.get("flags"), docs)
        if bad:
            skipped.append(f"{label}: unrecognisable flag(s) {', '.join(bad)} -- nothing was written for this area")
            continue
        r1, r2 = min(r["r1"], r["r2"]), max(r["r1"], r["r2"])
        c1, c2 = min(r["c1"], r["c2"]), max(r["c1"], r["c2"])
        hits = []
        for g in by_sheet[sheet]:
            t = parse_ref(g["title_cell"]) if g.get("title_cell") else None
            in_rect = not (g["last_row"] < r1 or g["first_row"] > r2 or g["last_col"] < c1 or g["first_col"] > c2)
            on_title = t is not None and r1 <= t["r1"] <= r2 and c1 <= t["c1"] <= c2
            if in_rect or on_title:
                hits.append(g)
        if len(hits) == 1:
            # A hit that is one half of an untitled caption-above-a-table pair is
            # extended to the pair: titling only one half would leave the other to
            # the automatic pass, whose separate title lands adjacent and makes the
            # re-detected grids merge wrongly (seen live on the caption layout).
            g = hits[0]
            cap = table = None
            if not g.get("name"):
                if g["n_cols"] == 1 and _is_text(g.get("header_values", [None])[0]):
                    t = [h for h in by_sheet[sheet] if h is not g and not h.get("name") and h["first_row"] == g["first_row"] + 1 and h["first_col"] == g["first_col"] + 1]
                    if t:
                        cap, table = g, t[0]
                if cap is None:
                    c = [
                        h for h in by_sheet[sheet]
                        if h is not g and not h.get("name") and h["n_cols"] == 1 and _is_text(h.get("header_values", [None])[0])
                        and g["first_row"] == h["first_row"] + 1 and g["first_col"] == h["first_col"] + 1
                    ]
                    if c:
                        cap, table = c[0], g
            resolved.append({"grid": cap or g, "name": name, "flags": flags, "table": table})
            continue
        # the documented caption-above-a-table split reads as two grids; the
        # user naturally selects both -- turning the caption into the title
        # merges them into one named grid
        if len(hits) == 2:
            for cap, table in ((hits[0], hits[1]), (hits[1], hits[0])):
                if (
                    not cap.get("name") and cap["n_cols"] == 1 and _is_text(cap.get("header_values", [None])[0])
                    and table["first_row"] == cap["first_row"] + 1 and table["first_col"] == cap["first_col"] + 1
                ):
                    resolved.append({"grid": cap, "name": name, "flags": flags, "table": table})
                    break
            else:
                skipped.append(f"{label}: the selection touches 2 grids ({hits[0]['ref']}, {hits[1]['ref']}) -- Mind reads them separately; select and name one at a time")
            continue
        if not hits:
            skipped.append(f"{label}: Mind detects no grid there (an alone text cell or empty space is ignored) -- nothing to name")
        else:
            skipped.append(f"{label}: the selection touches {len(hits)} grids ({', '.join(h['ref'] for h in hits[:6])}) -- Mind reads them separately; select and name one at a time")

    retitled = {id(item["grid"]) for item in resolved}
    used = {g["name"].strip().lower() for g in grids if g.get("name") and id(g) not in retitled}

    for item in resolved:
        g, table = item["grid"], item["table"]
        gkey = f"{g['sheet']}!{g['ref']}"
        exclude.add(gkey)
        if table is not None:
            exclude.add(f"{table['sheet']}!{table['ref']}")
        unique = _unique_name(item["name"], used)
        suffix = "" if unique == item["name"] else f" (made unique: '{item['name']}' is already a grid name)"
        title = "#" + unique + "".join(f" /{f}" for f in item["flags"])

        if table is not None:  # caption case: the caption cell becomes the title of the table below
            reader = referenced_by(index, g["sheet"], g["first_row"], g["first_col"])
            if reader:
                skipped.append(f"{gkey}: the caption is read by {reader} -- turning it into a title would change that result")
                continue
            ops.append(_op("set_value", "named_areas", "STR-004", g["sheet"], cell=g["anchor"], before=g["header_values"][0], after=title, note=f"user-chosen name; the caption becomes the title of the table {table['ref']}{suffix}"))
            titled_cells.add((g["sheet"], g["anchor"]))
            named.append(gkey)
            reserved.append(unique)
            continue

        if g.get("title_cell"):  # already titled: rewrite the title cell
            if (g.get("title") or "").strip() == title:
                skipped.append(f"{gkey}: already titled exactly '{title}'")
                reserved.append(unique)
                continue
            t = parse_ref(g["title_cell"])
            reader = referenced_by(index, g["sheet"], t["r1"], t["c1"]) if t else None
            if reader:
                skipped.append(f"{gkey}: the title cell {g['title_cell']} is read by {reader} -- rewriting it would change that result")
                continue
            ops.append(_op("set_value", "named_areas", "STR-004", g["sheet"], cell=g["title_cell"], before=g["title"], after=title, note=f"user-chosen name and flags{suffix}"))
            titled_cells.add((g["sheet"], g["title_cell"]))
            named.append(gkey)
            reserved.append(unique)
            continue

        # untitled: same write sites and safety checks as create_grid_titles
        r0, c0 = g["first_row"], g["first_col"]
        above_cell = ref_text(c0, r0 - 1) if r0 > 1 else None
        above_val = cell_value(analysis, g["sheet"], r0 - 1, c0) if r0 > 1 else None
        above_is_free = (
            r0 > 1
            and (g["sheet"], above_cell) not in titled_cells
            and (above_val is None or (isinstance(above_val, str) and above_val.strip() == ""))
            and grid_containing(by_sheet.get(g["sheet"], []), r0 - 1, c0) is None
        )
        if above_is_free:
            reader = referenced_by(index, g["sheet"], r0 - 1, c0)
            if reader:
                skipped.append(f"{gkey}: the cell above ({above_cell}) is read by {reader} -- a title there would change that result")
                continue
            ops.append(_op("set_value", "named_areas", "STR-004", g["sheet"], cell=above_cell, before=None, after=title, note=f"user-chosen name and flags{suffix}"))
            titled_cells.add((g["sheet"], above_cell))
        else:
            reason = "starts on row 1" if r0 == 1 else f"the cell above ({above_cell}) is not free"
            already = r0 in inserted_rows.get(g["sheet"], set())
            if not already:
                cutting = _row_insert_is_safe(grids, g["sheet"], r0, own=g)
                if cutting:
                    skipped.append(f"{gkey}: {reason} and inserting a row would cut {cutting[0]['display_name']}")
                    continue
                blocker = insert_blocker(index, g["sheet"])
                if blocker:
                    skipped.append(f"{gkey}: {reason} and {blocker}")
                    continue
            reader = referenced_by(index, g["sheet"], r0, c0)
            if reader:
                skipped.append(f"{gkey}: {reason} and the title row would fall inside the range {reader} reads")
                continue
            if not already:
                before = f"grid {g['ref']} starts on row 1" if r0 == 1 else f"'{above_val}' sits directly above grid {g['ref']}"
                ops.append(_op("insert_row", "named_areas", "STR-004", g["sheet"], row=r0, before=before, after=f"empty row {r0} inserted (the grid moves down one row)", note="makes room for the user-chosen title; Excel shifts every reference automatically"))
                inserted_rows.setdefault(g["sheet"], set()).add(r0)
            ops.append(_op("set_value", "named_areas", "STR-004", g["sheet"], cell=ref_text(c0, r0), before=None, after=title, note=f"user-chosen name and flags{suffix}; written after the row insert", after_inserts=True, insert_at=r0))
            titled_cells.add((g["sheet"], ref_text(c0, r0)))
        named.append(gkey)
        reserved.append(unique)

    return {"ops": ops, "skipped": skipped, "named": named, "exclude": exclude, "reserved": reserved, "titled_cells": titled_cells, "inserted_rows": inserted_rows}


def plan_explicit_colors(analysis, report) -> tuple[list[dict], list[str]]:
    """FMT-002: theme colours become explicit RGB (same look, no theme dependency)."""
    ops, skipped = [], []
    for s in analysis["workbooks"][0]["sheets"]:
        refs = s.get("style_stats", {}).get("theme_color_cell_refs", [])
        total = s.get("style_stats", {}).get("theme_color_cells", 0)
        if refs:
            ops.append(_op("explicit_colors", "explicit_colors", "FMT-002", s["name"], cells=refs, before=f"{len(refs)} cell(s) with theme colours", after="same colours as explicit RGB", note="requires Excel"))
            if total > len(refs):
                skipped.append(f"{s['name']}: only the first {len(refs)} of {total} theme-coloured cells are listed")
    return ops, skipped


def plan_keep_empty_styles(analysis, report) -> tuple[list[dict], list[str]]:
    """FMT-003: KB -- add an apostrophe + space to styled empty cells so Mind
    keeps their style. Changes cell content (COUNTA etc.), so off by default."""
    ops, skipped = [], []
    for s in analysis["workbooks"][0]["sheets"]:
        refs = s.get("style_stats", {}).get("empty_styled_cell_refs", [])
        for cell in refs:
            ops.append(_op("set_value", "keep_empty_styles", "FMT-003", s["name"], cell=cell, before=None, after="' ", note="apostrophe + space keeps the Excel style in Mind"))
        total = s.get("style_stats", {}).get("empty_styled_cells", 0)
        if total > len(refs):
            skipped.append(f"{s['name']}: only the first {len(refs)} of {total} styled empty cells are listed")
    return ops, skipped


def plan_unprotect_sheets(analysis, report) -> tuple[list[dict], list[str]]:
    """FMT-005: sheet protection makes locked cells non-clickable in Mind."""
    ops, skipped = [], []
    for s in analysis["workbooks"][0]["sheets"]:
        if s.get("protected"):
            ops.append(_op("unprotect_sheet", "unprotect_sheets", "FMT-005", s["name"], before="protected", after="unprotected", note="fails if the sheet has a password"))
    return ops, skipped


_SHEET_REF_ERROR_RE = re.compile(r"(?<![:\w'.])(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!#REF!(?![:\w])")


def _na_fix_formula(formula: str, broken_names) -> tuple[str, bool, bool]:
    """Replace broken references with NA(): every standalone reference to a broken
    defined name and every standalone literal #REF!, sheet-qualified or not
    ('BASE Polices'!#REF!, 1.7.2). A call that then holds NA() where Excel needs
    a range (SUMIFS(NA(),...) -- Excel refuses such a formula) becomes NA()
    itself (`collapse_na_reference_calls`; exact, it could only ever produce an
    error). Returns (fixed, has_unfixable, collapsed) where has_unfixable is
    True if a #REF! remains (e.g. a range endpoint A1:#REF! that cannot be
    safely NA()-swapped) or Excel would still refuse the result, and collapsed
    is True if a whole call had to become NA() (the broken reference stood for
    a range)."""
    fixed = formula
    for nm in sorted(broken_names, key=len, reverse=True):
        fixed = re.sub(r"(?<![A-Za-z0-9_.])" + re.escape(nm) + r"(?![A-Za-z0-9_.])", "NA()", fixed)
    fixed = _SHEET_REF_ERROR_RE.sub("NA()", fixed)
    fixed = re.sub(r"(?<![:!\w])#REF!(?![:\w])", "NA()", fixed)  # standalone #REF! only
    swapped = fixed
    fixed = collapse_na_reference_calls(swapped)
    return fixed, ("#REF!" in fixed or bool(reference_arg_problems(fixed))), fixed != swapped


def plan_fix_broken_refs(analysis, report) -> tuple[list[dict], list[str]]:
    """REF-001 fix: rewrite formulas that reference broken (#REF!) defined names or
    contain a literal #REF!, replacing the broken reference with NA(). Conservative:
    a broken reference only ever produced an error and NA() is also an error, so no
    valid result changes (IFERROR/valid branches keep their value). Array formulas are
    rewritten as arrays; a #REF! that cannot be safely swapped (range endpoint) is
    skipped for review."""
    ops, skipped = [], []
    for f, fixed, unfixable, collapsed in _broken_reference_sites(analysis):
        sheet, cell, formula = f["sheet"], f["cell"], f["formula"]
        if unfixable or collapsed:
            skipped.append(f"{sheet}!{cell}: the broken reference stands for a whole range (e.g. a deleted column inside SUMIFS) -- NA() cannot replace a range; see the action 'Replace formulas built on a broken range with =NA()'")
            continue
        note = "replaced broken #REF! / broken-name reference with NA() (error stays an error; valid result preserved)"
        if f.get("array_ref"):
            ops.append(_op("set_array_formula", "fix_broken_refs", "REF-001", sheet, range=str(f["array_ref"]), cell=str(f["array_ref"]), before=formula, after=fixed, note=note))
        else:
            ops.append(_op("set_formula", "fix_broken_refs", "REF-001", sheet, cell=cell, before=formula, after=fixed, note=note))
    return ops, skipped


def _broken_reference_sites(analysis: dict[str, Any]) -> list[tuple[dict[str, Any], str, bool, bool]]:
    """Every formula with a broken reference: (formula record, NA()-rewritten
    formula, unfixable, collapsed) where `unfixable` means a #REF! remains after
    the rewrite -- it stood for a range (A1:#REF!), which NA() cannot replace
    without changing the formula's meaning -- and `collapsed` means a call that
    needs a range (SUMIFS('Sheet'!#REF!,...)) had to become NA() as a whole."""
    wb0 = analysis["workbooks"][0]
    broken = sorted({d["name"] for d in wb0.get("defined_names", []) if "#REF!" in str(d.get("value") or "")})
    name_re = re.compile(r"(?<![A-Za-z0-9_.])(" + "|".join(re.escape(n) for n in broken) + r")(?![A-Za-z0-9_.])") if broken else None
    lit_re = re.compile(r"#REF!")
    out = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        masked = mask_strings(formula)
        if not ((name_re is not None and name_re.search(masked)) or lit_re.search(masked)):
            continue
        fixed, unfixable, collapsed = _na_fix_formula(formula, broken)
        out.append((f, fixed, unfixable or fixed == formula, collapsed))
    return out


def plan_fix_broken_refs_whole(analysis, report) -> tuple[list[dict], list[str]]:
    """REF-001, the range case (1.7.2): a formula whose broken reference stands
    for a whole range cannot be patched piecewise, so the *whole* formula
    becomes =NA(). Conservative for the same reason as the piecewise fix: the
    cell already evaluated to an error, and it still does -- no valid value
    changes. It is a bigger decision (a calculation is replaced, not repaired),
    so it is its own action with its own checkbox and a caution note.
    1.7.4: where only the call that needs the range has to go
    (`collapse_na_reference_calls`), only that call becomes NA() --
    =IFERROR(SUMIF(#REF!,1),0) -> =IFERROR(NA(),0) keeps its 0."""
    ops, skipped = [], []
    for f, fixed, unfixable, collapsed in _broken_reference_sites(analysis):
        if not (unfixable or collapsed):
            continue
        sheet, cell, formula = f["sheet"], f["cell"], f["formula"]
        after = "=NA()" if unfixable else fixed
        note = (
            "the broken reference stood for a range; the whole formula is replaced by =NA() (the cell already returned an error)"
            if after == "=NA()"
            else "the broken reference stood for a range; the call that needed it is replaced by NA() (it could only return an error)"
        )
        if f.get("array_ref"):
            ops.append(_op("set_array_formula", "fix_broken_refs_whole", "REF-001", sheet, range=str(f["array_ref"]), cell=str(f["array_ref"]), before=formula, after=after, note=note))
        else:
            ops.append(_op("set_formula", "fix_broken_refs_whole", "REF-001", sheet, cell=cell, before=formula, after=after, note=note))
    return ops, skipped


# --- FRM-003: spilled-range references (1.7.2) --------------------------------------------
# Excel stores `A1#` (a reference to the range a dynamic-array formula spills
# into) as `_xlfn.ANCHORARRAY(A1)`. Mind rejects every such formula with
# "Unsupported formula: ANCHORARRAY()". The spill's current extent is known:
# Excel writes the anchor as an array formula whose `ref` is the spilled
# range. Replacing the reference by that fixed range keeps today's result and
# gives Mind a formula it compiles; if the source grows later, the range will
# not follow (the caution says so).
SPILL_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?ANCHORARRAY\(\s*((?:'[^']+'|[A-Za-z0-9_.]+)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})\s*\)", re.IGNORECASE)
SPILL_HASH_RE = re.compile(r"(?<![A-Za-z0-9_.#])((?:'[^']+'|[A-Za-z0-9_.]+)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})#", re.IGNORECASE)


def _absolute(ref: str) -> str:
    """'B12' -> '$B$12' (a single A1 reference)."""
    m = re.match(r"^[$]?([A-Za-z]{1,3})[$]?([0-9]{1,7})$", ref)
    return "$" + m.group(1) + "$" + m.group(2) if m else ref


def _spill_anchors(analysis: dict[str, Any]) -> dict[tuple[str, str], str]:
    """(sheet, anchor cell) -> the range its array formula covers, from the inventory."""
    out: dict[tuple[str, str], str] = {}
    for f in analysis["workbooks"][0].get("formulas", []):
        ref = f.get("array_ref")
        if ref:
            out[(f["sheet"], str(f["cell"]).replace("$", "").upper())] = str(ref).replace("$", "").upper()
    return out


def freeze_spill_refs(formula: str, sheet: str, anchors: dict[tuple[str, str], str]) -> tuple[str, list[str]]:
    """Rewrite every `A1#` / ANCHORARRAY(A1) in `formula` (a formula on `sheet`)
    as the fixed range the anchor spills into today. Returns (fixed formula,
    references that could not be resolved -- their anchor is not an array
    formula in the inventory, so the spill extent is unknown)."""
    masked = mask_strings(formula)  # string literals become spaces; positions line up with `formula`
    unresolved: list[str] = []
    frozen: list[str] = []
    pieces: list[tuple[int, int, str]] = []
    for pattern in (SPILL_CALL_RE, SPILL_HASH_RE):
        for m in pattern.finditer(masked):
            prefix = formula[m.start(1):m.end(1)] if m.group(1) else ""
            cell = formula[m.start(2):m.end(2)]
            anchor_sheet = prefix[:-1].strip("'") if prefix else sheet
            rng = anchors.get((anchor_sheet, cell.replace("$", "").upper()))
            if rng is None:
                unresolved.append(f"{prefix}{cell}#")
                continue
            if "$" in cell:  # keep the reference absolute when the original was
                rng = ":".join(_absolute(part) for part in rng.split(":"))
            pieces.append((m.start(), m.end(), f"{prefix}{rng}"))
            frozen.append(f"{prefix}{cell}# -> {prefix}{rng}")
    if not pieces:
        return formula, unresolved
    pieces.sort()
    out, pos = [], 0
    for start, end, text in pieces:
        if start < pos:  # overlapping match (cannot happen with these patterns, but never corrupt a formula)
            continue
        out.append(formula[pos:start])
        out.append(text)
        pos = end
    out.append(formula[pos:])
    FROZEN_NOTES[formula] = frozen
    return "".join(out), unresolved


FROZEN_NOTES: dict[str, list[str]] = {}  # last rewrite's "A1# -> A1:A3" texts, for the operation note


def plan_freeze_spill_refs(analysis, report) -> tuple[list[dict], list[str]]:
    """FRM-003: each spilled-range reference becomes the range it covers today."""
    ops, skipped = [], []
    anchors = _spill_anchors(analysis)
    for f in analysis["workbooks"][0].get("formulas", []):
        formula = f.get("formula") or ""
        masked = mask_strings(formula)
        if not (SPILL_CALL_RE.search(masked) or SPILL_HASH_RE.search(masked)):
            continue
        sheet, cell = f["sheet"], f["cell"]
        fixed, unresolved = freeze_spill_refs(formula, sheet, anchors)
        if unresolved:
            skipped.append(f"{sheet}!{cell}: the spill extent of {', '.join(unresolved)} is unknown (its anchor is not an array formula in the file) -- replace the reference by the real range by hand")
            continue
        what = "; ".join(FROZEN_NOTES.get(formula, []))
        note = f"{what} -- frozen to the range the spill covers today (Mind rejects ANCHORARRAY); a 1-cell range means the source spills one cell in this file"
        if f.get("array_ref") and ":" in str(f["array_ref"]):
            ops.append(_op("set_array_formula", "freeze_spill_refs", "FRM-003", sheet, range=str(f["array_ref"]), cell=str(f["array_ref"]), before=formula, after=fixed, note=note))
        else:
            ops.append(_op("set_formula", "freeze_spill_refs", "FRM-003", sheet, cell=cell, before=formula, after=fixed, note=note))
    return ops, skipped


# 1.7.2: every action carries its `level` -- "blocking" when Mind's converter
# refuses the workbook (or computes wrong) without it, "optional" when Mind reads
# the file as it is and the change only improves it. The level is the priority
# of the rule the action serves (rules/*.yaml: REQUIRED -> blocking); it is not
# a judgment about any particular workbook. Blocking actions are on by default;
# `caution` says why a blocking action still deserves a look before Apply.
BLOCKING, OPTIONAL = "blocking", "optional"
ACTIONS: list[dict[str, Any]] = [
    {"id": "fix_broken_refs", "title": "Replace broken (#REF!) references with NA() so Mind can compile", "rule_ids": ["REF-001"], "default_on": True, "level": BLOCKING,
     "caution": "rewrites formulas: a broken reference becomes NA(); the cell already returned an error, so no valid result changes", "planner": plan_fix_broken_refs},
    {"id": "fix_broken_refs_whole", "title": "Replace formulas built on a broken range with =NA()", "rule_ids": ["REF-001"], "default_on": True, "level": BLOCKING,
     "caution": "replaces the whole formula (its broken reference stood for a range, e.g. a deleted column inside SUMIFS); restore the real range by hand instead if you know it", "planner": plan_fix_broken_refs_whole},
    {"id": "freeze_spill_refs", "title": "Replace spilled-range references (A1#) by the fixed range they cover today", "rule_ids": ["FRM-003"], "default_on": True, "level": BLOCKING,
     "caution": "Mind rejects ANCHORARRAY() (Office 365 dynamic arrays); each 'A1#' becomes the range the spill covers now, so a source that grows later will not be followed -- add MM_RANGE by hand if it must", "planner": plan_freeze_spill_refs},
    {"id": "separate_merged_grids", "title": "Insert an empty row before a '#Title' trapped inside a grid", "rule_ids": ["STR-001"], "default_on": True, "level": BLOCKING,
     "caution": "inserts whole rows through Excel; formulas follow, but VBA code or other workbooks addressing these rows by number will not", "planner": plan_separate_merged_grids},
    {"id": "loop_name_case", "title": "Normalise loop-name capitalisation (MM_LOOP / MM_RESULT / MM_DIMSIZE ...)", "rule_ids": ["LOOP-002", "RES-002"], "default_on": True, "level": BLOCKING, "planner": plan_loop_name_case},
    {"id": "reorder_input_flag", "title": "Add /Input to /Reorder grids", "rule_ids": ["INP-005"], "default_on": True, "level": BLOCKING, "planner": plan_reorder_input_flag},
    {"id": "special_headers", "title": "Correct headers of /ExportSettings, /ProjectSettings, /Parameters, /InputSettings grids", "rule_ids": ["EXP-003", "PRJ-002", "PAR-002", "INP-003"], "default_on": True, "level": BLOCKING, "planner": plan_special_headers},
    {"id": "create_grid_titles", "title": "Give every untitled grid one '#Name' title taken from the cells around it (captions, labels, headers)", "rule_ids": ["STR-004", "STR-002", "STR-001"], "default_on": True, "level": OPTIONAL,
     "level_note": "Mind accepts an untitled grid but lists it as 'Untitled' -- unusable for whoever opens the model", "planner": plan_create_grid_titles},
    {"id": "flag_spelling", "title": "Correct misspelt flags to the documented spelling", "rule_ids": ["FLG-001"], "default_on": True, "level": OPTIONAL, "planner": plan_flag_spelling},
    {"id": "hide_marker", "title": "Add '&&Hide' to hidden sheet names", "rule_ids": ["STR-007"], "default_on": True, "level": OPTIONAL, "planner": plan_hide_marker},
    {"id": "explicit_colors", "title": "Replace theme colours with explicit RGB", "rule_ids": ["FMT-002"], "default_on": True, "level": OPTIONAL, "planner": plan_explicit_colors},
    {"id": "keep_empty_styles", "title": "Put an apostrophe + space in styled empty cells", "rule_ids": ["FMT-003"], "default_on": False, "level": OPTIONAL, "planner": plan_keep_empty_styles},
    {"id": "unprotect_sheets", "title": "Remove sheet protection", "rule_ids": ["FMT-005"], "default_on": False, "level": OPTIONAL, "planner": plan_unprotect_sheets},
]
ACTION_LEVELS: dict[str, str] = {a["id"]: a["level"] for a in ACTIONS}


def merged_ranges(analysis: dict[str, Any]) -> dict[str, list[tuple[int, int, int, int]]]:
    """Per sheet, every merged range as (r1, r2, c1, c2) -- from the inventory."""
    out: dict[str, list[tuple[int, int, int, int]]] = {}
    for s in analysis["workbooks"][0].get("sheets", []):
        rects = []
        for m in s.get("merged_cells", []) or []:
            ref = parse_ref(str(m))
            if ref and not ref.get("whole_column") and not ref.get("whole_row"):
                rects.append((ref["r1"], ref["r2"], ref["c1"], ref["c2"]))
        out[s["name"]] = rects
    return out


def merged_conflict(merged: dict[str, list[tuple[int, int, int, int]]], sheet: str, cell: str) -> str | None:
    """The merged range that contains `cell` without `cell` being its first cell
    (Excel ignores a write there), else None."""
    ref = parse_ref(str(cell))
    if not ref:
        return None
    for r1, r2, c1, c2 in merged.get(sheet, []):
        if r1 <= ref["r1"] <= r2 and c1 <= ref["c1"] <= c2 and (ref["r1"], ref["c1"]) != (r1, c1):
            return f"{ref_text(c1, r1)}:{ref_text(c2, r2)}"
    return None


def _drop_merged_targets(ops: list[dict[str, Any]], skipped: list[str], merged: dict[str, list[tuple[int, int, int, int]]]) -> list[dict[str, Any]]:
    """1.7.2: a cell write aimed at a non-first cell of a merged range is moved
    from the operations to the skip list (Excel would ignore it silently).
    Writes planned for post-insert coordinates are left to the executor's guard."""
    kept = []
    for o in ops:
        if o["op"] in ("set_value", "set_formula", "clear_cell") and not o.get("after_inserts"):
            span = merged_conflict(merged, o["sheet"], str(o.get("cell", "")))
            if span:
                skipped.append(f"{o['sheet']}!{o['cell']}: inside the merged range {span} -- Excel only writes into its first cell; unmerge it or write '{o.get('after')}' there by hand")
                continue
        kept.append(o)
    return kept


def plan_actions(analysis: dict[str, Any], validation_report: dict[str, Any], grid_names: dict[str, str] | None = None) -> list[dict[str, Any]]:
    """Every action with its planned operations and skip notes -- no file is touched.

    `grid_names` (optional, "Sheet!Ref" -> name) comes from the assistant
    (app/grid_naming.suggest_names) and is only used by `create_grid_titles`,
    and there only where the deterministic name would be a weak one."""
    out = []
    merged = merged_ranges(analysis)
    for a in ACTIONS:
        try:
            if a["id"] == "create_grid_titles":
                ops, skipped = a["planner"](analysis, validation_report, grid_names)
            else:
                ops, skipped = a["planner"](analysis, validation_report)
            ops = _drop_merged_targets(list(ops), skipped, merged)
        except Exception as exc:  # a planner bug must never hide the other actions
            ops, skipped = [], [f"planner error: {exc}"]
        out.append({
            "id": a["id"], "title": a["title"], "rule_ids": a["rule_ids"], "default_on": a["default_on"],
            "level": a["level"], "level_note": a.get("level_note"), "caution": a.get("caution"),
            "operations": ops, "count": len(ops), "skipped": skipped,
        })
    return out


def formula_replacement_op(sheet: str, cell: str, before: str, after: str, rule_id: str) -> dict[str, Any]:
    """A user-approved formula replacement (FRM-002 / RSK-004 / FORMULA-002)."""
    after = after.strip()
    if not after.startswith("="):
        raise ValueError("the replacement must be a formula starting with '='")
    return _op("set_formula", "formula_replacement", rule_id, sheet, cell=cell, before=before, after=after, note="user-approved replacement")


# --- executor ------------------------------------------------------------------
def _ordered(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Cell ops on original coordinates, then inserts (bottom-up / right-to-left),
    then cell ops written for post-insert coordinates, then sheet-level ops."""
    phase1 = [o for o in operations if o["op"] in CELL_OPS | {"explicit_colors", "unprotect_sheet"} and not o.get("after_inserts")]
    row_inserts = sorted((o for o in operations if o["op"] == "insert_row"), key=lambda o: (o["sheet"], -int(o["row"])))
    col_inserts = sorted((o for o in operations if o["op"] == "insert_column"), key=lambda o: (o["sheet"], -col_to_num(str(o["column"]))))
    phase2 = [_resolve_insert_at(o, row_inserts) for o in operations if o["op"] in CELL_OPS and o.get("after_inserts")]
    sheet_level = [o for o in operations if o["op"] in ("rename_sheet", "set_sheet_visibility")]
    return phase1 + row_inserts + col_inserts + phase2 + sheet_level


def _resolve_insert_at(op: dict[str, Any], row_inserts: list[dict[str, Any]]) -> dict[str, Any]:
    """A cell written into a row this plan inserts (`insert_at` = the row the
    companion insert targets) lands one row lower for every *earlier* row this
    plan inserts on the same sheet -- including inserts planned by other
    actions the user selected. Recomputed from `insert_at` every time, so the
    same operation list can be ordered repeatedly without drifting."""
    at = op.get("insert_at")
    if at is None:
        return op
    at = int(at)
    shift = sum(1 for o in row_inserts if o["sheet"] == op["sheet"] and int(o["row"]) < at)
    if not shift:
        return op
    ref = parse_ref(str(op["cell"]))
    if not ref:
        return op
    return {**op, "cell": ref_text(ref["c1"], at + shift)}


EXCEL_REFUSED_FORMULA = -2146827284  # 0x800A03EC: Excel would not accept the value/formula
# Excel answers these while it is busy for an instant (a background recalculation, a
# redraw): the call did nothing and succeeds a moment later. One real Apply of 10,404
# formulas lost exactly one of them to RPC_E_CALL_REJECTED.
TRANSIENT_COM_HRESULTS = {-2147418111, -2147417846}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
TRANSIENT_RETRIES = 5


def _is_transient(exc: BaseException) -> bool:
    return getattr(exc, "hresult", None) in TRANSIENT_COM_HRESULTS or (exc.args[:1] and exc.args[0] in TRANSIENT_COM_HRESULTS)


def _com_message(exc: BaseException, op: dict[str, Any] | None = None) -> str:
    """Excel's own message out of a pywin32 com_error, else a plain-English
    one: Excel often gives no text at all (just 0x800A03EC), which used to
    reach the user as a raw tuple."""
    info = getattr(exc, "excepinfo", None)
    if info and len(info) > 2 and info[2]:
        return str(info[2])
    scode = info[5] if info and len(info) > 5 else None
    if scode == EXCEL_REFUSED_FORMULA:
        formula = str((op or {}).get("after") or "")
        problems = reference_arg_problems(formula) if formula.startswith("=") else []
        if problems:
            return f"Excel refused this formula: {describe_reference_problem(problems[0])} (NA(), numbers, text and calculations are not accepted there)"
        if formula.startswith("="):
            return "Excel refused this formula -- it is not a formula Excel can store (check the syntax and the function arguments)"
        return "Excel refused this change (error 0x800A03EC)"
    return str(exc)


def _cells_of(ref_or_range: Any) -> list[str]:
    return [c.upper() for c in _range_cells(str(ref_or_range or "").replace("$", "").upper())]


def _merged_guard(rng: Any, sheet: str, target: str) -> None:
    """Raise when `rng` (a single cell or a range) touches a merged area it does
    not own: Excel ignores a value written into a merged cell other than the
    area's first cell, with no error. A write on the first cell is fine."""
    try:
        merged = rng.MergeCells
    except Exception:
        return
    if not merged:
        return
    area = None
    try:
        area = rng.Cells(1, 1).MergeArea
        first = str(area.Cells(1, 1).Address).replace("$", "")
        span = str(area.Address).replace("$", "")
    finally:
        area = None
    top_left = _cells_of(target)[0]
    if first.upper() != top_left.upper():
        raise ValueError(
            f"{sheet}!{target} is part of the merged range {span} -- Excel only accepts a value in its first cell ({first}); "
            "unmerge it or write the title into that cell by hand"
        )


def _read_back_guard(rng: Any, o: dict[str, Any], sheet: str, target: str) -> None:
    """After a write, confirm the cell is no longer empty (set_value / set_formula)
    or is empty (clear_cell). A silent no-op must never be reported as applied."""
    try:
        first = rng.Cells(1, 1)
        try:
            formula = first.Formula
        finally:
            first = None
    except Exception:
        return  # cannot read back: keep the operation, the verify-opens step still runs
    has_content = formula is not None and str(formula) != ""
    if o["op"] == "clear_cell":
        if has_content:
            raise ValueError(f"{sheet}!{target} still has content after clearing")
    elif o.get("after") not in (None, ""):
        if not has_content:
            raise ValueError(f"{sheet}!{target}: Excel accepted the write but the cell is still empty (protected, merged or otherwise refused)")


def _op_target(o: dict[str, Any]) -> str:
    """'Sheet!Cell' (or row / column / the sheet alone) an operation writes to, for progress messages."""
    where = o.get("range") or o.get("cell") or (f"row {o['row']}" if o.get("row") is not None else None) or (f"column {o['column']}" if o.get("column") else None)
    return f"{o['sheet']}!{where}" if where else str(o["sheet"])


def _apply_with_excel(copy_path: Path, operations: list[dict[str, Any]], progress: Any = None) -> tuple[list[dict], list[dict]]:
    applied: list[dict] = []
    failed: list[dict] = []
    # Every cell any operation writes. A multi-cell array formula (Ctrl+Shift+Enter)
    # can only be changed as a whole -- Excel refuses "part of an array" -- so it is
    # dismantled first when the operations cover all of its cells.
    covered: set[tuple[str, str]] = set()
    for o in operations:
        if o["op"] in CELL_OPS:
            for c in _cells_of(o.get("range") or o.get("cell")):
                covered.add((o["sheet"], c))
    dismantled: set[tuple[str, str]] = set()

    def _array_of(ws: Any, target: str) -> str | None:
        """Address (e.g. 'C7:C9') of the multi-cell array formula containing the
        top-left cell of `target`, else None."""
        rng = ws.Range(target).Cells(1, 1)
        try:
            if not rng.HasArray:
                return None
            arr = rng.CurrentArray
            try:
                addr = str(arr.Address).replace("$", "")
            finally:
                arr = None
            return addr if ":" in addr else None
        finally:
            rng = None

    def _write_cell(ws: Any, o: dict[str, Any]) -> dict[str, Any]:
        sheet, target = o["sheet"], str(o["cell"])
        record = o
        arr = _array_of(ws, target)
        if arr is not None:
            arr_cells = _cells_of(arr)
            if all((sheet, c) in covered for c in arr_cells):
                if (sheet, arr) not in dismantled:
                    ws.Range(arr).ClearContents()
                    dismantled.add((sheet, arr))
                record = {**o, "array": arr}
            elif o["op"] == "set_formula" and _cells_of(target) == [arr_cells[0]]:
                # Only the top-left cell is addressed: replace the whole array formula (Excel's own semantics).
                ws.Range(arr).FormulaArray = o["after"]
                return {**o, "array": arr, "cell": arr, "note": f"{o.get('note') or ''} (whole array formula {arr} replaced)".strip()}
            else:
                raise ValueError(
                    f"{sheet}!{target} is part of the array formula {arr} -- Excel cannot change part of an array; "
                    f"include every cell of {arr} in the fix, or use set_array_formula on {arr}"
                )
        else:
            for sh, a in dismantled:  # a cell of an array dismantled by an earlier operation
                if sh == sheet and set(_cells_of(target)) <= set(_cells_of(a)):
                    record = {**o, "array": a}
                    break
        rng = ws.Range(target)
        try:
            # 1.7.2 guard: Excel silently ignores a write into any cell of a merged
            # range other than its first cell -- no error, nothing written. Prep
            # then reported "applied, verified" and the next scan proposed the
            # same write again, forever. Refuse, and say why.
            _merged_guard(rng, sheet, target)
            try:
                merged = bool(rng.MergeCells)
            except Exception:
                merged = False
            if merged:
                # The first cell of a merged range: Excel refuses ClearContents /
                # value writes on that single cell ("We can't do that to a merged
                # cell") but accepts them on the whole merge area.
                area = rng.Cells(1, 1).MergeArea
                rng = None
                rng = area
            if o["op"] == "set_formula":
                if rng.HasArray:  # single-cell array formula: replace it entirely
                    rng.ClearContents()
                rng.Formula = o["after"]
            elif o["op"] == "set_value":
                if rng.HasArray:
                    rng.ClearContents()
                rng.Value = o["after"]
            else:
                rng.ClearContents()
            _read_back_guard(rng, o, sheet, target)
        finally:
            rng = None
        return record

    def _write_array(ws: Any, o: dict[str, Any]) -> dict[str, Any]:
        target = str(o.get("range") or o["cell"]).replace("$", "").upper()
        existing = _array_of(ws, target)
        if existing is not None and (existing.upper() == target or all((o["sheet"], c) in covered for c in _cells_of(existing))):
            ws.Range(existing).ClearContents()
            dismantled.add((o["sheet"], existing))
        rng = ws.Range(target)
        try:
            try:
                rng.FormulaArray = o["after"]
            except Exception as exc:
                raise ValueError(
                    f"could not write the array formula over {o['sheet']}!{target}: {_com_message(exc, o)} -- if the range overlaps an "
                    "existing array formula, the operations must cover all of that array (or use its exact range)"
                ) from exc
        finally:
            rng = None
        return {**o, "array": target}

    def _run(excel: Any) -> None:
        wb = None
        try:
            ordered = _ordered(operations)
            total = len(ordered)
            if progress is not None:
                progress("apply_write", "Opening the copy in Excel", 0.0, done=0, total=total)
            wb = open_for_write(excel, copy_path)
            # Manual calculation while writing: in automatic mode Excel recalculates every
            # dependent after each write, and 12,678 writes into a 70k-formula model ran for
            # hours. The mode is restored before the save, so the file keeps its own setting
            # and is recalculated once.
            calc_mode = None
            try:
                calc_mode = excel.Calculation
                excel.Calculation = -4135  # xlCalculationManual
            except Exception:
                calc_mode = None
            for i, o in enumerate(ordered, start=1):
                if progress is not None:
                    progress("apply_write", f"Change {i} of {total}: {_op_target(o)}", (i - 1) / total, done=i - 1, total=total, sheet=o["sheet"])
                ws = None
                try:
                    for attempt in range(TRANSIENT_RETRIES + 1):
                        try:
                            ws = wb.Worksheets(o["sheet"])
                            record = _write_one(wb, ws, o, i, total)
                            break
                        except Exception as exc:
                            ws = None
                            if attempt < TRANSIENT_RETRIES and _is_transient(exc):
                                time.sleep(0.5 * (attempt + 1))  # Excel was busy for an instant: the same call again
                                continue
                            raise
                    applied.append(record)
                except Exception as exc:
                    failed.append({**o, "error": _com_message(exc, o)[:300]})
                finally:
                    ws = None
            if progress is not None:
                progress("apply_save", "Excel is recalculating and saving the copy", None, done=total, total=total)
            if calc_mode is not None:
                try:
                    excel.Calculation = calc_mode
                except Exception:
                    pass
            save_in_place(wb, copy_path)
        finally:
            close_quietly(wb)

    def _write_one(wb: Any, ws: Any, o: dict[str, Any], i: int, total: int) -> dict[str, Any]:
        """One operation on its sheet; returns the record to log as applied."""
        record = o
        if o["op"] in ("set_formula", "set_value", "clear_cell"):
            record = _write_cell(ws, o)
        elif o["op"] == "set_array_formula":
            record = _write_array(ws, o)
        elif o["op"] == "unprotect_sheet":
            ws.Unprotect()
        elif o["op"] == "set_sheet_visibility":
            ws.Visible = -1 if o["after"] else 0  # xlSheetVisible / xlSheetHidden
        elif o["op"] == "explicit_colors":
            if not o.get("in_style_table"):  # 1.8.0: already done for every sheet at once (app/theme_colors.py)
                _explicit_colors(ws, o, i, total, progress)
        elif o["op"] == "insert_row":
            ws.Rows(int(o["row"])).Insert()
        elif o["op"] == "insert_column":
            ws.Columns(col_to_num(str(o["column"]))).Insert()
        elif o["op"] == "rename_sheet":
            ws.Name = o["after"]
        else:
            raise ValueError(f"unknown op {o['op']}")
        return record

    with excel_session() as excel:
        _run(excel)
    return applied, failed


XL_COLOR_INDEX_NONE = -4142


def _recolor(rng: Any, single: bool) -> bool:
    """Give a range its own colours back as explicit RGB (font, then fill).
    -> False when the range holds more than one font colour or more than one
    fill: Excel answers Null for a mixed range, and it must be split.
    A single cell is never mixed (a font colour Excel cannot name -- rich text
    with several colours -- is left alone, as before)."""
    try:
        font = rng.Font.Color
    except Exception:
        font = False
    try:
        index = rng.Interior.ColorIndex
    except Exception:
        index = False
    if not single and (font is None or index is None):
        return False
    fill: Any = False
    if index not in (False, None, XL_COLOR_INDEX_NONE):
        try:
            fill = rng.Interior.Color
        except Exception:
            fill = False
        if fill is None:
            if not single:
                return False  # one palette index, several real colours
            fill = False
    if font not in (False, None):
        try:
            rng.Font.Color = font
        except Exception:
            pass
    if fill is not False:
        try:
            rng.Interior.Color = fill
        except Exception:
            pass
    return True


def _explicit_colors(ws: Any, o: dict[str, Any], i: int, total: int, progress: Any = None, blockwise: bool = True) -> None:
    """FMT-002 at the cell level: theme colours become the same colours as
    explicit RGB.

    1.8.0: by blocks. The cells are covered with rectangles and a rectangle
    that has one font colour and one fill is re-coloured in one go; a mixed
    one is split into its rows, a mixed row into its cells -- the cell-by-cell
    assignment this replaces, which took ~10 COM calls a cell (37 minutes for
    the 77,000 cells of one model, 124,000 on another). The result is the same
    cell for cell (`blockwise=False` is the old way, kept for the test that
    says so).

    Since `_colors_in_the_style_table` this is the fallback only (a package
    whose style table or theme cannot be read): Excel keeps the default text
    colour as a theme colour when a cell is assigned the colour it already
    shows, so this way FMT-002 comes back at the next scan."""
    from .autofix import rectangles

    cells = [str(c) for c in o["cells"]]
    n_cells = len(cells)
    done = 0
    last_told = -1

    def tell() -> None:
        nonlocal last_told
        if progress is not None and (done - last_told >= 25 or last_told < 0):
            last_told = done
            progress("apply_write", f"Change {i} of {total}: {o['sheet']}, colours of cell {min(done + 1, n_cells)} of {n_cells}", (i - 1 + done / max(1, n_cells)) / total, done=i - 1, total=total, sheet=o["sheet"], inner_done=done, inner_total=n_cells)

    coords: list[tuple[int, int]] = []
    odd: list[str] = []
    for c in cells:
        ref = parse_ref(c.replace("$", ""))
        if blockwise and ref and not ref["whole_column"] and not ref["whole_row"] and ref["r1"] == ref["r2"] and ref["c1"] == ref["c2"]:
            coords.append((ref["r1"], ref["c1"]))
        else:
            odd.append(c)
    for r1, c1, r2, c2 in rectangles(coords):
        tell()
        block = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        try:
            whole = _recolor(block, (r1, c1) == (r2, c2))
        finally:
            block = None
        if not whole:
            for r in range(r1, r2 + 1):
                row = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
                try:
                    row_done = _recolor(row, c1 == c2)
                finally:
                    row = None
                if not row_done:
                    for c in range(c1, c2 + 1):
                        cell = ws.Cells(r, c)
                        try:
                            _recolor(cell, True)
                        finally:
                            cell = None
                done += c2 - c1 + 1
                tell()
        else:
            done += (r2 - r1 + 1) * (c2 - c1 + 1)
    for c in odd:
        tell()
        cell = ws.Range(c)
        try:
            _recolor(cell, True)
        finally:
            cell = None
        done += 1


def _range_cells(ref_text_: str) -> list[str]:
    ref = parse_ref(ref_text_)
    if not ref or ref["whole_column"] or ref["whole_row"]:
        return [ref_text_]
    return [ref_text(c, r) for r in range(ref["r1"], ref["r2"] + 1) for c in range(ref["c1"], ref["c2"] + 1)]


def _apply_with_openpyxl(copy_path: Path, operations: list[dict[str, Any]], progress: Any = None) -> tuple[list[dict], list[dict]]:
    applied, failed = [], []
    total = len(operations)
    if progress is not None:
        progress("apply_write", "Opening the copy (openpyxl)", 0.0, done=0, total=total)
    wb = openpyxl.load_workbook(copy_path, data_only=False, keep_vba=copy_path.suffix.lower() == ".xlsm")
    for i, o in enumerate(operations, start=1):
        if progress is not None:
            progress("apply_write", f"Change {i} of {total}: {_op_target(o)}", (i - 1) / total, done=i - 1, total=total, sheet=o["sheet"])
        if o["op"] in STRUCTURAL_OPS or o["op"] == "set_array_formula" or o.get("after_inserts"):
            failed.append({**o, "error": "requires Excel (references/styles would not be kept intact)"})
            continue
        try:
            ws = wb[o["sheet"]]
            for cell in _range_cells(o["cell"]):
                ws[cell].value = None if o["op"] == "clear_cell" else o["after"]
            applied.append(o)
        except Exception as exc:
            failed.append({**o, "error": str(exc)[:200]})
    if progress is not None:
        progress("apply_save", "Saving the copy", None, done=total, total=total)
    wb.save(copy_path)
    wb.close()
    return applied, failed


def _colors_in_the_style_table(copy_path: Path, operations: list[dict[str, Any]], warnings: list[str], progress: Any = None) -> list[dict[str, Any]]:
    """FMT-002 (1.8.0): theme colours become RGB in the copy's style table,
    before Excel opens it -- a few hundred fonts and fills instead of hundreds
    of thousands of cells, and the default text colour, which Excel keeps as a
    theme colour when a cell is assigned the colour it already shows, is
    converted too. -> the operations, the colour ones marked as done; unchanged
    when the package cannot be converted (they then go cell by cell, as before)."""
    if not any(o["op"] == "explicit_colors" for o in operations):
        return operations
    if progress is not None:
        progress("apply_write", "Replacing theme colours in the workbook's style table", 0.0, done=0, total=len(operations))
    try:
        done = explicit_theme_colors_in_package(copy_path)
    except Exception as exc:
        warnings.append(f"The style table could not be converted ({str(exc)[:160]}); theme colours were replaced cell by cell.")
        return operations
    if not done or not (done["fonts"] or done["fills"]):
        return operations
    note = f"in the workbook's style table: {done['fonts']} font colour(s) and {done['fills']} fill colour(s), every sheet at once"
    return [{**o, "in_style_table": True, "note": note} if o["op"] == "explicit_colors" else o for o in operations]


def apply_operations(source_path: Path, work_dir: Path, operations: list[dict[str, Any]], prefer_excel: bool = True, progress: Any = None) -> dict[str, Any]:
    """Apply approved operations to a fresh copy of `source_path`.

    `progress(stage, message, fraction, **facts)` (optional) is told every
    stage (app.progress.APPLY_STAGES) and, while writing, every operation
    (`done` / `total`)."""
    if not operations:
        return {"status": "NOT_APPLICABLE", "message": "No operations selected."}
    source_path = Path(source_path).resolve()
    if progress is not None:
        progress("apply_copy", "Copying the workbook (the source is never modified)", None, done=0, total=len(operations))
    copy_path, source_sha256 = make_immutable_copy(source_path, work_dir)
    warnings: list[str] = []
    method = "openpyxl"
    applied: list[dict] = []
    failed: list[dict] = []
    if prefer_excel and com_available():
        operations = _colors_in_the_style_table(copy_path, operations, warnings, progress)
        try:
            applied, failed = _apply_with_excel(copy_path, operations, progress)
            method = "excel_com"
        except Exception as exc:
            warnings.append(f"Excel COM apply failed ({str(exc)[:160]}); fell back to openpyxl.")
    if method == "openpyxl":
        applied, failed = _apply_with_openpyxl(copy_path, operations, progress)
        warnings.append("Written with openpyxl (reduced fidelity): structural operations were refused and the file may not open in Excel.")
    # 1.7.3 guard: changes that did not reach the file were not applied, whatever
    # the writer reported. A copy still byte-identical to its source holds none of them.
    if applied and sha256_of(copy_path) == source_sha256:
        reason = "the changes were made in Excel but the file was not saved: the output is identical to the source"
        failed = [{**o, "error": reason} for o in applied] + failed
        applied = []
        warnings.append(f"Nothing was written: {reason}.")
    if progress is not None:
        progress("apply_verify", "Opening the copy in Excel to check it" if com_available() else "Excel is not available here: the copy is not verified", None, done=len(operations), total=len(operations))
    verification = verify_opens_in_excel(copy_path) if com_available() else {"opens": None}
    entry = {
        "rule_id": sorted({o["rule_id"] for o in operations}),
        "action": "prep_apply_operations",
        "actions": sorted({o["action_id"] for o in operations}),
        "method": method,
        "source_path": str(source_path),
        "source_sha256": source_sha256,
        "output_path": str(copy_path),
        "output_sha256": sha256_of(copy_path),
        "applied": [{k: v for k, v in o.items() if k != "cells"} | ({"cells": len(o["cells"])} if "cells" in o else {}) for o in applied],
        "failed": failed,
        "verified_opens_in_excel": verification.get("opens"),
        "warnings": warnings,
    }
    if progress is not None:
        progress("apply_log", "Writing the change log", None, done=len(operations), total=len(operations))
    append_change_log(copy_path, entry)
    status = "APPLIED" if applied and not failed else ("PARTIAL" if applied else "ERROR")
    return {
        "status": status,
        "method": method,
        "output_path": copy_path,
        "applied": applied,
        "failed": failed,
        "verified_opens_in_excel": verification.get("opens"),
        "warnings": warnings,
        "message": f"{len(applied)} operation(s) applied via {method}" + (f"; {len(failed)} failed" if failed else ""),
        "change_log_entry": entry,
    }


def summarize_plan(actions: list[dict[str, Any]]) -> str:
    return json.dumps([{k: v for k, v in a.items() if k in ("id", "title", "rule_ids", "count", "skipped")} for a in actions], indent=1)


# --- operations proposed by the assistant (1.5.0) --------------------------------------
ASSISTANT_OPS = {"set_value", "set_formula", "clear_cell", "set_array_formula", "rename_sheet", "insert_row", "insert_column", "set_sheet_visibility", "unprotect_sheet"}
BAD_SHEET_CHARS = set('[]:*?/\\')


def validate_proposal(proposal: Any, analysis: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    """Turn the assistant's JSON proposal into executor operations, refusing
    anything malformed. Returns (operations, errors)."""
    errors: list[str] = []
    ops: list[dict[str, Any]] = []
    if not isinstance(proposal, dict) or not isinstance(proposal.get("operations"), list):
        return [], ["proposal must be an object with an 'operations' list"]
    sheets = {s["name"].lower(): s["name"] for s in analysis["workbooks"][0]["sheets"]}
    # sheets the scan skipped still exist in the file: a rename must not collide with them
    ignored_names = {s["name"].lower(): s["name"] for s in analysis["workbooks"][0].get("ignored_sheets", [])}
    summary = str(proposal.get("summary") or "assistant change")[:200]
    for i, raw in enumerate(proposal["operations"], start=1):
        if not isinstance(raw, dict):
            errors.append(f"operation {i}: not an object")
            continue
        op = str(raw.get("op", "")).strip()
        if op not in ASSISTANT_OPS:
            errors.append(f"operation {i}: unknown op '{op}'")
            continue
        sheet_raw = str(raw.get("sheet", "")).strip()
        sheet = sheets.get(sheet_raw.lower())
        if sheet is None:
            errors.append(f"operation {i}: sheet '{sheet_raw}' does not exist")
            continue
        base = {"action_id": "assistant", "rule_id": "CHAT", "sheet": sheet, "note": summary}
        try:
            if op in ("set_value", "set_formula", "clear_cell"):
                cell = str(raw.get("cell", "")).strip().replace("$", "").upper()
                ref = parse_ref(cell)
                if not ref or ref["whole_column"] or ref["whole_row"]:
                    raise ValueError(f"'{raw.get('cell')}' is not a cell or range")
                if ref["sheet"]:
                    raise ValueError("put the sheet in 'sheet', not in 'cell'")
                before = cell_value(analysis, sheet, ref["r1"], ref["c1"]) if ref["cells"] == 1 else f"{ref['cells']} cells"
                if ref["cells"] == 1 and before is None:
                    # openpyxl keeps a multi-cell array formula in its anchor cell only; the other members are not empty in Excel
                    member_of = _array_containing(analysis, sheet, ref["ref"])
                    if member_of:
                        before = f"part of array formula {member_of}"
                if op == "set_value":
                    value = raw.get("value")
                    if isinstance(value, str) and value.startswith("="):
                        raise ValueError("a value starting with '=' must use set_formula")
                    if ref["cells"] == 1 and _same_content(before, value):
                        raise ValueError(f"{sheet}!{ref['ref']} already contains {value!r}; nothing would change -- propose a real change")
                    ops.append({**base, "op": "set_value", "cell": ref["ref"], "before": before, "after": value})
                elif op == "set_formula":
                    formula = str(raw.get("formula", "")).strip()
                    if not formula.startswith("="):
                        raise ValueError("formula must start with '='")
                    _refuse_what_excel_refuses(sheet, ref["ref"], formula)
                    if ref["cells"] == 1 and _same_content(before, formula):
                        raise ValueError(f"{sheet}!{ref['ref']} already contains the formula {formula}; nothing would change -- the formula itself must be different")
                    ops.append({**base, "op": "set_formula", "cell": ref["ref"], "before": before, "after": formula})
                else:
                    if ref["cells"] == 1 and before is None:
                        raise ValueError(f"{sheet}!{ref['ref']} is already empty; nothing would change")
                    ops.append({**base, "op": "clear_cell", "cell": ref["ref"], "before": before, "after": None})
            elif op == "set_array_formula":
                rng_text = str(raw.get("range") or raw.get("cell") or "").strip().replace("$", "").upper()
                ref = parse_ref(rng_text)
                if not ref or ref["whole_column"] or ref["whole_row"]:
                    raise ValueError(f"'{raw.get('range')}' is not a range like C7:C9")
                if ref["sheet"]:
                    raise ValueError("put the sheet in 'sheet', not in 'range'")
                formula = str(raw.get("formula", "")).strip()
                if not formula.startswith("="):
                    raise ValueError("formula must start with '='")
                _refuse_what_excel_refuses(sheet, ref["ref"], formula)
                current = _array_formula_at(analysis, sheet, ref["ref"])
                if current is not None and _same_content(current, formula):
                    raise ValueError(
                        f"{sheet}!{ref['ref']} is already the array formula {formula}; re-entering it changes nothing -- the formula "
                        f"(or the range it covers) must be different: {array_size_hint(ref['ref'], formula) or 'shrink the array to the cells that have source values and clear the rest'}"
                    )
                ops.append({**base, "op": "set_array_formula", "range": ref["ref"], "cell": ref["ref"], "before": current or f"{ref['cells']} cells (array formula)", "after": formula})
            elif op == "rename_sheet":
                new_name = str(raw.get("new_name", "")).strip()
                if not new_name or len(new_name) > EXCEL_SHEET_NAME_MAX or BAD_SHEET_CHARS & set(new_name):
                    raise ValueError("new_name must be 1-31 characters without []:*?/\\")
                if (new_name.lower() in sheets or new_name.lower() in ignored_names) and new_name != sheet:
                    raise ValueError(f"a sheet named '{new_name}' already exists")
                ops.append({**base, "op": "rename_sheet", "before": sheet, "after": new_name})
            elif op == "insert_row":
                row = int(raw.get("row"))
                if row < 1:
                    raise ValueError("row must be >= 1")
                ops.append({**base, "op": "insert_row", "row": row, "before": f"row {row}", "after": f"empty row inserted at {row}"})
            elif op == "insert_column":
                column = str(raw.get("column", "")).strip().upper()
                if not re.fullmatch(r"[A-Z]{1,3}", column):
                    raise ValueError("column must be a letter like 'C'")
                ops.append({**base, "op": "insert_column", "column": column, "before": f"column {column}", "after": f"empty column inserted at {column}"})
            elif op == "set_sheet_visibility":
                visible = raw.get("visible")
                if not isinstance(visible, bool):
                    raise ValueError("visible must be true or false")
                ops.append({**base, "op": "set_sheet_visibility", "before": "current visibility", "after": visible})
            elif op == "unprotect_sheet":
                ops.append({**base, "op": "unprotect_sheet", "before": "protected", "after": "unprotected"})
        except (ValueError, TypeError) as exc:
            errors.append(f"operation {i} ({op}): {exc}")
    ops, array_errors = _array_coverage(ops, analysis)
    return ops, errors + array_errors


def _refuse_what_excel_refuses(sheet: str, cell: str, formula: str) -> None:
    """1.7.2: a formula Excel would refuse to store (a value where a function
    needs a cell range, e.g. SUMIFS(NA(), ...)) is rejected here, with the
    concrete repair, so the assistant fixes it before the user sees it --
    instead of Excel failing it at Apply time."""
    problems = reference_arg_problems(formula)
    if not problems:
        return
    listed = "; ".join(describe_reference_problem(p) for p in problems[:3])
    collapsed = collapse_na_reference_calls(formula)
    if collapsed != formula and not reference_arg_problems(collapsed):
        repair = (
            f"Use {collapsed} instead: a call with a broken range can only ever produce an error, so the whole call "
            "becomes NA() (results, IFERROR branches included, stay the same)"
        )
    else:
        repair = "Point it at a real range (e.g. 'Sheet'!C:C) if you know which one was meant, or replace the whole formula with =NA()"
    raise ValueError(
        f"Excel would refuse the formula for {sheet}!{cell}: {listed}. SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS, "
        f"COUNTBLANK, OFFSET, ROW, COLUMN and SUBTOTAL only accept a cell range there -- not NA(), #N/A, a number, text or a "
        f"calculation. {repair}."
    )


def _same_content(before: Any, after: Any) -> bool:
    """True when writing `after` would leave the cell as it is (formula text
    compared without whitespace/case; other values compared as Excel would)."""
    if before is None or after is None:
        return before is None and after is None
    b, a = str(before).strip(), str(after).strip()
    if a.startswith("=") or b.startswith("="):
        return a.replace(" ", "").upper() == b.replace(" ", "").upper()
    if isinstance(before, bool) or isinstance(after, bool):
        return before is after
    try:
        return float(b) == float(a)
    except ValueError:
        return b == a


def _array_containing(analysis: dict[str, Any], sheet: str, cell: str) -> str | None:
    """The multi-cell array formula range that contains `cell`, if any."""
    cell = cell.replace("$", "").upper()
    for wb in analysis.get("workbooks", []):
        for f in wb.get("formulas", []):
            ref = str(f.get("array_ref") or "").replace("$", "").upper()
            if f.get("sheet") == sheet and ":" in ref and cell in _cells_of(ref):
                return ref
    return None


_ARRAY_REF_RE = re.compile(r"(?:(?:'[^']+'|[A-Za-z0-9_.]+)!)?\$?[A-Za-z]{1,3}\$?\d{1,7}:\$?[A-Za-z]{1,3}\$?\d{1,7}")


def array_size_hint(array_ref: str, formula: str) -> str:
    """How big a multi-cell array formula is versus the ranges it references,
    with the concrete fix when the array overflows its source (the classic
    cause of #N/A in the extra cells of a Ctrl+Shift+Enter array)."""
    arr = parse_ref(str(array_ref).replace("$", "").upper())
    if not arr or arr["whole_column"] or arr["whole_row"] or arr["cells"] < 2:
        return ""
    rows, cols = arr["r2"] - arr["r1"] + 1, arr["c2"] - arr["c1"] + 1
    sizes = []
    for m in _ARRAY_REF_RE.finditer(mask_strings(str(formula or ""))):
        r = parse_ref(m.group(0).replace("$", ""))
        if r and not r["whole_column"] and not r["whole_row"]:
            sizes.append((m.group(0).replace("$", ""), r["r2"] - r["r1"] + 1, r["c2"] - r["c1"] + 1, r["cells"]))
    text = f"the array {arr['ref']} has {rows}x{cols} cells"
    if not sizes:
        return text
    text += " while the formula's ranges are " + ", ".join(f"{s[0]} ({s[1]}x{s[2]})" for s in sizes)
    big = max(sizes, key=lambda s: s[3])
    if cols == 1 and big[2] == 1 and big[1] < rows:
        keep = f"{ref_text(arr['c1'], arr['r1'])}:{ref_text(arr['c1'], arr['r1'] + big[1] - 1)}"
        rest = f"{ref_text(arr['c1'], arr['r1'] + big[1])}:{ref_text(arr['c2'], arr['r2'])}"
        text += (f"; the array is taller than its source, so its last {rows - big[1]} cell(s) can only be #N/A -- fix: set_array_formula on {keep} "
                 f"with the same formula and clear_cell {rest} (or extend the source range to {rows} rows)")
    elif rows == 1 and big[1] == 1 and big[2] < cols:
        keep = f"{ref_text(arr['c1'], arr['r1'])}:{ref_text(arr['c1'] + big[2] - 1, arr['r1'])}"
        rest = f"{ref_text(arr['c1'] + big[2], arr['r1'])}:{ref_text(arr['c2'], arr['r2'])}"
        text += (f"; the array is wider than its source, so its last {cols - big[2]} cell(s) can only be #N/A -- fix: set_array_formula on {keep} "
                 f"with the same formula and clear_cell {rest} (or extend the source range to {cols} columns)")
    return text


def _array_formula_at(analysis: dict[str, Any], sheet: str, ref: str) -> str | None:
    """The array formula currently entered over exactly `ref`, if any."""
    for wb in analysis.get("workbooks", []):
        for f in wb.get("formulas", []):
            if f.get("sheet") == sheet and str(f.get("array_ref") or "").replace("$", "").upper() == ref.upper():
                return str(f.get("formula") or "")
    return None


def _array_coverage(ops: list[dict[str, Any]], analysis: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    """Excel cannot change part of a multi-cell array formula (Ctrl+Shift+Enter).
    Drop -- with a precise message -- every operation that touches an array the
    proposal does not cover completely; a fully covered array is dismantled by
    the executor, and set_array_formula writes one as a whole."""
    arrays: dict[str, list[tuple[str, list[str]]]] = {}
    for wb in analysis.get("workbooks", []):
        for f in wb.get("formulas", []):
            ref = str(f.get("array_ref") or "").replace("$", "").upper()
            if ":" in ref:
                arrays.setdefault(f["sheet"], []).append((ref, _cells_of(ref)))
    if not arrays:
        return ops, []
    written: dict[str, set[str]] = {}
    for o in ops:
        if o["op"] in CELL_OPS:
            written.setdefault(o["sheet"], set()).update(_cells_of(o.get("range") or o.get("cell")))
    kept: list[dict[str, Any]] = []
    errors: list[str] = []
    for o in ops:
        problem = None
        if o["op"] in CELL_OPS:
            target = o.get("range") or o.get("cell")
            cells = set(_cells_of(target))
            for ref, arr_cells in arrays.get(o["sheet"], []):
                if not (cells & set(arr_cells)):
                    continue
                missing = [c for c in arr_cells if c not in written.get(o["sheet"], set())]
                if missing:
                    problem = (
                        f"{o['sheet']}!{target} is part of the array formula {ref} (Ctrl+Shift+Enter) and Excel cannot change part of an array: "
                        f"cover ALL of {ref} -- one set_array_formula on {ref}, or one operation per cell (still uncovered: {', '.join(missing[:10])})"
                    )
                    break
        if problem:
            errors.append(problem)
        else:
            kept.append(o)
    return kept, errors


def describe_operation(o: dict[str, Any]) -> str:
    target = o.get("cell") or (f"row {o['row']}" if "row" in o else f"column {o['column']}" if "column" in o else "")
    return f"{o['op']} {o['sheet']}!{target}".rstrip("!") + f": {str(o.get('before'))[:40]!r} -> {str(o.get('after'))[:60]!r}"
