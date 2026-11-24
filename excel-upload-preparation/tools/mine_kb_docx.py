#!/usr/bin/env python3
"""Deterministic  extraction from  CompleteMindDocn.docx (a page-by-page scrape
of the Milliman Mind knowledge base, kb.milliman-mind.com) into three
reference files consumed by the validators:

  references/mm-function-registry-kb.yaml   every MM_ function the KB documents
                                            (dedicated article: syntax + category)
                                            or merely mentions (changelogs etc.)
  references/supported-excel-functions.yaml the  KB's "Supported Excel formulas"
                                            list (basic-model-design page)
  references/mind-flags.yaml                 documented grid-title / header flags
                                            and sheet/header  markers, each one
                                            checked to actually occur  in the KB text

Like tools/mine_function_registry.py this is  a parser, not an editor: re-run
it against an updated .docx instead  of hand-editing the outputs. It reads the
.docx  with zipfile + regex only (no python-docx dependency).

Usage:
    python tools/mine_kb_docx.py [--docx PATH] [--out-dir references]
"""
from __future__ import annotations

import argparse
import html
import re
import sys
import zipfile
from collections import OrderedDict
from pathlib import Path

import yaml

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DOCX = PACKAGE_ROOT.parent.parent / "CompleteMindDocn.docx"
PAGE_MARKER_RE = re.compile(r"^>>>>>Page>>>> (\d+)$")
FUNC_SLUG_RE = re.compile(r"mm_([a-z0-9_]+)")
MM_TOKEN_RE = re.compile(r"\bMM_[A-Z][A-Z0-9_]*\b")
SYNTAX_RE = re.compile(r"=\s*(MM_[A-Z0-9_]+)\s*\(")
HYPHEN_BREAK_RE = re.compile(r"(MM_[A-Z0-9_]+)- ([A-Z0-9_]+)")  # scrape  artifact: 'MM_HP-  CLOOP'
HOST_PARTS = {"kb.milliman-mind.com", "www.kb.milliman-mind.com"}

# Grid-title  flags (and  a few header/sheet markers) the KB documents. `where`
# says where Mind reads the flag from; `syntax` is the documented shape. Every
# `kb_term` MUST occur (case-insensitively) in the KB text or the script fails,
# so  this table can't silently drift away from the source document.
DOCUMENTED_FLAGS: list[dict] = [
    # ---  inputs /  exports -------------------------------------------------
    {"name": "Input", "where": "title", "kb_term": "/Input", "topic": "input-manager",
     "meaning": "Grid is an assumptions-only input grid, importable through the Input Manager (CSV named like the grid)."},
    {"name": "InputSettings", "where": "title", "kb_term": "/InputSettings", "topic": "input-manager-automatch", "unique": True,
     "meaning": "Unique grid mapping /Input grid names (GridName) to CSV filename patterns (FileNamePattern)."},
    {"name": "Reorder", "where": "title", "kb_term": "/Reorder", "topic": "input-manager-reorder-columns",
     "meaning": "Input columns may arrive in any order; requires /Input, no /NoHeader and unique headers."},
    {"name": "NoHeader", "where": "title", "kb_term": "/NoHeader", "topic": "input-manager-reorder-columns",
     "meaning": "Grid has no header row (first row is data)."},
    {"name": "NoHeaders", "where": "title", "kb_term": "/NoHeaders", "topic": "changelog", "alias_of": "NoHeader",
     "meaning": "Alias spelling of /NoHeader seen in release notes."},
    {"name": "Export", "where": "title", "kb_term": "/Export", "topic": "export",
     "meaning": "Grid is listed in the Export Manager."},
    {"name": "ExportSettings", "where": "title", "kb_term": "/ExportSettings", "topic": "export", "unique": True,
     "meaning": "Unique grid of per-grid export settings (GridName, ExportByInstance, ... AllowScientificFormat)."},
    {"name": "ExportByInstance", "where": "title", "kb_term": "/EXPORTBYINSTANCE", "topic": "export",
     "meaning": "Export one file per instance (overridden by /ExportSettings)."},
    {"name": "ExportByLoop", "where": "title", "kb_term": "/EXPORTBYLOOP", "topic": "export",
     "meaning": "Export one file per loop index (overridden by /ExportSettings)."},
    {"name": "ExportNoInstanceKeys", "where": "title", "kb_term": "/EXPORTNOINSTANCEKEYS", "topic": "export",
     "meaning": "Do not add an InstanceKey column to the export."},
    {"name": "ExportNoLoopKeys", "where": "title", "kb_term": "/EXPORTNOLOOPKEYS", "topic": "export",
     "meaning": "Do not add a loop-index column to the export."},
    {"name": "InputLink", "where": "title", "kb_term": "/InputLink", "topic": "interlinks",
     "meaning": "Target grid of an interlink (receives data from another project's /OutputLink grid of the same name)."},
    {"name": "OutputLink", "where": "title", "kb_term": "/OutputLink", "topic": "interlinks",
     "meaning": "Source grid of an interlink."},
    # --- project / structure ---------------------------------------------
    {"name": "ProjectSettings", "where": "title", "kb_term": "/ProjectSettings", "topic": "projectsettings-grid", "unique": True,
     "meaning": "Unique grid pre-setting project settings: columns Name | Value | Locked | Hidden."},
    {"name": "StepLabels", "where": "title", "kb_term": "/StepLabels", "topic": "step-labels-grid", "unique": True,
     "meaning": "Unique grid defining workflow step labels; must be alone on its own sheet."},
    {"name": "Step", "where": "title", "kb_term": "/Step", "topic": "enable-workflow-mode",
     "meaning": "Assigns the grid to a workflow step."},
    {"name": "Collapse", "where": "title", "kb_term": "/Collapse", "topic": "enable-workflow-mode",
     "meaning": "Grid is collapsed (3 rows shown) in workflow mode."},
    {"name": "Group", "where": "title", "kb_term": "/Group.(Group", "topic": "create-groups-of-grids",
     "syntax": "/Group.(GroupName).x.y", "args": 3,
     "meaning": "Grid belongs to a group: shared group name, then vertical (x) and horizontal (y) position starting at 0."},
    {"name": "Parameters", "where": "title", "kb_term": "/Parameters", "topic": "setup-parameters-editor",
     "meaning": "Parameters-editor grid: columns Label | Type | PossibleValues | Values..."},
    {"name": "Translations", "where": "title", "kb_term": "/Translations", "topic": "translations-of-model", "unique": True,
     "meaning": "Translation grid: headers are ISO 639-1 two-letter language codes, first column the design language."},
    {"name": "ReadOnly", "where": "title", "kb_term": "/ReadOnly", "topic": "changelog",
     "meaning": "Grid (or group) is read-only in Mind."},
    {"name": "NoFormula", "where": "title", "kb_term": "/NoFormula", "topic": "changelog",
     "meaning": "Formulas of the grid are not shown in Mind (except in debug mode)."},
    {"name": "Always", "where": "title", "kb_term": "/Always", "topic": "create-dashboards",
     "meaning": "Dashboard displayed at project launch."},
    {"name": "AlwaysSave", "where": "title", "kb_term": "/AlwaysSave", "topic": "changelog",
     "meaning": "Grid results are always saved (HPC partial-save option)."},
    {"name": "Result", "where": "title", "kb_term": "/Result", "topic": "changelog",
     "meaning": "Result grid flag (release notes only; no dedicated article)."},
    {"name": "HPCInput", "where": "title", "kb_term": "/HPCInput", "topic": "changelog",
     "meaning": "Input grid split across HPC tasks (used with MM_HPCLOOP)."},
    {"name": "ExtSettings", "where": "title", "kb_term": "/ExtSettings", "topic": "changelog", "unique": True,
     "meaning": "Extension settings grid; the release notes say exactly one is expected when used."},
    {"name": "XBRLSettings", "where": "title", "kb_term": "/XBRLSettings", "topic": "changelog",
     "meaning": "XBRL export settings grid."},
    {"name": "ValueExclusionSource", "where": "title", "kb_term": "/valueexclusionsource", "topic": "changelog",
     "meaning": "Source grid of the right-click value-exclusion feature."},
    {"name": "ValueExclusionDest", "where": "title", "kb_term": "/ValueExclusionDest", "topic": "changelog",
     "meaning": "Destination grid of the value-exclusion feature (read-only)."},
    # --- resize ------------------------------------------------------------
    {"name": "Resize", "where": "title", "kb_term": "/Resize", "topic": "dynamic-resizing-of-grids",
     "syntax": "/Resize or /Resize.Name", "max_args": 1,
     "meaning": "Grid gets the same number of rows as the other /Resize(Row) grids (of the same named group)."},
    {"name": "ResizeRow", "where": "title", "kb_term": "/ResizeRow", "topic": "dynamic-resizing-of-grids",
     "syntax": "/ResizeRow or /ResizeRow.Name", "max_args": 1,
     "meaning": "Same as /Resize (rows)."},
    {"name": "ResizeColumn", "where": "title", "kb_term": "/ResizeColumn", "topic": "dynamic-resizing-of-grids",
     "syntax": "/ResizeColumn or /ResizeColumn.Name", "max_args": 1,
     "meaning": "Grid gets the same number of columns as the other /ResizeColumn grids (of the same named group)."},
    # --- instances -----------------------------------------------------------
    {"name": "InstanceKeys", "where": "title", "kb_term": "/InstanceKeys", "topic": "introduction-to-instances", "unique": True,
     "meaning": "InstanceKey table in the main workbook: headers are the workbook names to duplicate."},
    {"name": "InstanceSplit", "where": "title", "kb_term": "/InstanceSplit", "topic": "manage-grid-data-with-instancesplit",
     "meaning": "Grid data is split by the instance keys in its first column."},
    {"name": "InstanceSelect", "where": "header", "kb_term": "/InstanceSelect", "topic": "partial-load-instanceselect",
     "meaning": "Header suffix on the last column of the InstanceKey grid defining instance groups for partial load."},
    # --- calculations / iterations / batches ----------------------------------
    {"name": "CalculationSteps", "where": "title", "kb_term": "/CalculationSteps", "topic": "iterations-design", "unique": True,
     "meaning": "Two-column grid: workbook name | integer calculation step (equal numbers run in parallel)."},
    {"name": "iterationinput", "where": "title", "kb_term": "/iterationinput.(input_name)", "topic": "iterations-design",
     "syntax": "/iterationinput.(name)", "args": 1,
     "meaning": "Iteration input grid; paired by name with an /iterationoutput grid."},
    {"name": "iterationoutput", "where": "title", "kb_term": "/iterationoutput.(input_name)", "topic": "iterations-design",
     "syntax": "/iterationoutput.(name)", "args": 1,
     "meaning": "Iteration output grid feeding the next iteration's /iterationinput of the same name."},
    {"name": "NoKeepResults", "where": "title", "kb_term": "/NoKeepResults", "topic": "iterations-design",
     "meaning": "Keep only the last iteration's results for this grid (saves RAM)."},
    {"name": "BatchResult", "where": "title", "kb_term": "/batchresult", "topic": "mm_batchindex-function",
     "meaning": "Grid keeps results of every batch (BATCH loop visible) when running batches."},
    # --- AOC / AOS ------------------------------------------------------------
    {"name": "AOC", "where": "title", "kb_term": "/AOC", "topic": "analysis-of-sensitivities-aos-and-analysis-of-change-aoc",
     "meaning": "Input grid with a stretched-values version (Analysis of Change)."},
    {"name": "AOCOrder", "where": "title", "kb_term": "/AOCOrder", "topic": "analysis-of-sensitivities-aos-and-analysis-of-change-aoc", "unique": True,
     "meaning": "Unique grid: Grid Name | AOC Order index | Label."},
    {"name": "AOCGroup", "where": "title", "kb_term": "/AOCGROUP.(GroupName)", "topic": "groups-of-aoc-grids-with-aocgroup",
     "syntax": "/AOCGROUP.(GroupName)", "args": 1,
     "meaning": "Puts an /AOC grid into a named AOC group."},
    # --- backup ------------------------------------------------------------------
    {"name": "BackupSource", "where": "title", "kb_term": "/Backupsource.(backup", "topic": "grid-backup-feature",
     "syntax": "/BackupSource.(backup name)", "args": 1,
     "meaning": "Source grid of a named backup (exactly one source per backup name)."},
    {"name": "BackupDest", "where": "title", "kb_term": "/Backupdest.(backup", "topic": "grid-backup-feature",
     "syntax": "/BackupDest.(backup name)", "args": 1,
     "meaning": "Destination grid of a named backup (one or several per backup name)."},
    {"name": "BackupDestAllDimensions", "where": "title", "kb_term": "/Backupdestalldimensions.(backup", "topic": "grid-backup-feature",
     "syntax": "/BackupDestAllDimensions.(backup name)", "args": 1,
     "meaning": "Destination grid receiving the source grid with all dimensions flattened."},
    {"name": "CleanAuditTrailAfterBackup", "where": "title", "kb_term": "/CleanAuditTrailAfterBackup", "topic": "changelog",
     "meaning": "Clears the audit trail after a successful backup."},
    # --- display ----------------------------------------------------------------
    {"name": "HideRows", "where": "header", "kb_term": "/HideRows", "topic": "dynamically-hide-rows-columns-grids",
     "meaning": "Header of an extra right-hand column whose TRUE/1 values hide the matching row."},
]

# Non-flag  markers the  KB documents (used by validators alongside the flags).
DOCUMENTED_MARKERS: list[dict] = [
    {"marker": "#", "where": "title", "kb_term": "preceded by", "meaning": "A cell whose text starts with '#' directly above a grid names it: '#GridName /Flag1 /Flag2'."},
    {"marker": "&&Hide", "where": "sheet_name", "kb_term": "&&Hide", "meaning": "Hides an entire sheet in Mind."},
    {"marker": "&&HideColumn", "where": "header", "kb_term": "&&HideColumn", "meaning": "Hides the column whose header contains it."},
    {"marker": "&hide", "where": "header", "kb_term": "&hide", "meaning": "Hides the whole header line of the grid."},
    {"marker": "#NbSimulations", "where": "title", "kb_term": "#NbSimulations", "meaning": "Grid title locking the number of simulations per run."},
    {"marker": "°Year0 / °Month0 / °Quarter0 / °Half-year0", "where": "field_names_and_formulas", "kb_term": "Year0",
     "meaning": "Date indicators shifted from the project date (°Year-1, °Year1, ...)."},
]


def read_docx_lines(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    lines = []
    for para in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S):
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", para, flags=re.S))
        text = html.unescape(text)
        if text.strip():
            lines.append(text)
    return lines


def split_pages(lines: list[str]) -> list[dict]:
    pages: list[dict] = []
    current: dict | None = None
    for line in lines:
        m = PAGE_MARKER_RE.match(line)
        if m:
            current = {"n": int(m.group(1)), "url": "", "lines": []}
            pages.append(current)
        elif current is not None:
            if not current["url"] and line.startswith("http"):
                current["url"] = line.strip()
            else:
                current["lines"].append(line)
    return pages


def _slug_parts(url: str) -> tuple[str, str]:
    path = re.sub(r"^https?://[^/]+/", "", url)
    path = re.sub(r"^en_US/", "", path)
    parts = [p for p in path.split("/") if p and p not in HOST_PARTS]
    if len(parts) >= 2:
        return parts[-2], parts[-1]
    return "", parts[-1] if parts else ""


def mine_functions(pages: list[dict]) -> list[dict]:
    by_name: "OrderedDict[str, dict]" = OrderedDict()

    # Pass  1: dedicated  function articles (url slug contains mm_<name>).
    for page in pages:
        category, slug = _slug_parts(page["url"])
        if "mm_" not in slug:
            continue
        names = ["MM_" + m.upper() for m in FUNC_SLUG_RE.findall(slug)]
        body = page["lines"]
        description = ""
        syntax_lines: list[str] = []
        for i, line in enumerate(body):
            if line.strip() == "Description" and i + 1 < len(body) and not description:
                description = body[i + 1].strip()
            if line.strip() == "Syntax":
                for cand in body[i + 1 : i + 8]:
                    if "MM_" in cand and "(" in cand:
                        syntax_lines.append(cand.strip())
        for name in names:
            entry = by_name.get(name)
            if entry is None:
                entry = by_name[name] = {
                    "function": name,
                    "category": category,
                    "documented": True,
                    "deprecated": "deprecated" in slug,
                    "description": description,
                    "syntax": "",
                    "kb_url": page["url"],
                    "kb_pages": [],
                }
            entry["kb_pages"].append(page["n"])
            if not entry["syntax"]:
                for s in syntax_lines:
                    pos = s.find(name + "(")
                    if pos == -1:
                        continue
                    start = s.rfind("=", 0, pos)
                    start = start if start != -1 else pos
                    # Cut the syntax at the closing parenthesis that balances the opening one.
                    depth = 0
                    for j in range(pos, len(s)):
                        if s[j] == "(":
                            depth += 1
                        elif s[j] == ")":
                            depth -= 1
                            if depth == 0:
                                entry["syntax"] = s[start : j + 1].strip()
                                break
                    if entry["syntax"]:
                        break
            # Functions named in the syntax line but documented on a shared page
            # (e.g. MM_LEFTOUTERJOIN / MM_RIGHTOUTERJOIN)  all get the page.

    # Pass 2: every MM_ token anywhere in the  KB (changelogs, guides, ...).
    for page in pages:
        for line in page["lines"]:
            line = HYPHEN_BREAK_RE.sub(lambda m: m.group(1) + m.group(2), line)
            for raw_tok in MM_TOKEN_RE.findall(line):
                # 'MM_SETSIZEMM_VUNION' (scrape lost  the separators) -> two tokens
                for tok in [t for t in re.split(r"(?=MM_)", raw_tok) if t]:
                    _note_mention(by_name, tok, page)
    # Drop scrape fragments:  a mentioned-only token that is a strict prefix of a
    # documented one  ('MM_HYPERLI' from a hyphenated 'MM_HYPERLI-NK') is noise.
    for tok in list(by_name):
        e = by_name[tok]
        if e["documented"]:
            continue
        longer = [by_name[o] for o in by_name if o != tok and o.startswith(tok)]
        if any(o["documented"] or len(o["kb_pages"]) >= len(e["kb_pages"]) for o in longer):
            del by_name[tok]
    for entry in by_name.values():
        entry["kb_pages"] = sorted(entry["kb_pages"])
        if not entry["kb_url"] and entry["kb_pages"]:
            first = next(p for p in pages if p["n"] == entry["kb_pages"][0])
            entry["kb_url"] = first["url"]
    return sorted(by_name.values(), key=lambda e: e["function"])


def _note_mention(by_name: dict, tok: str, page: dict) -> None:
    if tok == "MM_" or len(tok) < 5:
        return
    entry = by_name.get(tok)
    if entry is None:
        entry = by_name[tok] = {
            "function": tok,
            "category": "mentioned-only",
            "documented": False,
            "deprecated": False,
            "description": "",
            "syntax": "",
            "kb_url": "",
            "kb_pages": [],
        }
    if page["n"] not in entry["kb_pages"]:
        entry["kb_pages"].append(page["n"])


def mine_supported_excel_functions(pages: list[dict]) -> dict:
    page = next((p for p in pages if p["url"].rstrip("/").endswith("supported-excel-formulas")), None)
    if page is None:
        raise SystemExit("supported-excel-formulas page not found in the .docx")
    categories: "OrderedDict[str, list[str]]" = OrderedDict()
    current = None
    started = False
    for line in page["lines"]:
        if line.startswith("List of supported Excel formulas"):
            started = True
            continue
        if not started:
            continue
        if line.startswith("- "):
            if current is None:
                current = "Uncategorized"
                categories[current] = []
            categories[current].append(line[2:].strip().upper())
        else:
            current = line.strip()
            categories.setdefault(current, [])
    flat = sorted({f for fs in categories.values() for f in fs})
    return {"source_url": page["url"], "kb_page": page["n"], "categories": dict(categories), "functions": flat}


def verify_flags(lines: list[str]) -> None:
    text = "\n".join(lines).lower()
    missing = [f["name"] for f in DOCUMENTED_FLAGS if f["kb_term"].lower() not in text]
    missing += [m["marker"] for m in DOCUMENTED_MARKERS if m["kb_term"].lower() not in text]
    if missing:
        raise SystemExit(f"documented flag(s)/marker(s) not found in the KB text -- table is stale: {missing}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", default=str(DEFAULT_DOCX))
    ap.add_argument("--out-dir", default=str(PACKAGE_ROOT / "references"))
    args = ap.parse_args()
    docx = Path(args.docx)
    if not docx.is_file():
        print(f"docx not found: {docx}", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    lines = read_docx_lines(docx)
    pages = split_pages(lines)
    verify_flags(lines)

    header = {
        "version": "1.0.0",
        "source_document": docx.name,
        "note": "Generated by tools/mine_kb_docx.py from the Milliman Mind knowledge-base scrape. Do not hand-edit; re-run the script against an updated document instead.",
        "kb_pages_total": len(pages),
    }

    functions = mine_functions(pages)
    (out_dir / "mm-function-registry-kb.yaml").write_text(
        yaml.safe_dump({**header, "functions": functions}, sort_keys=False, allow_unicode=True, width=120),
        encoding="utf-8",
    )
    documented = sum(1 for f in functions if f["documented"])
    print(f"mm-function-registry-kb.yaml: {len(functions)} MM_ functions ({documented} with a dedicated article)")

    supported = mine_supported_excel_functions(pages)
    (out_dir / "supported-excel-functions.yaml").write_text(
        yaml.safe_dump({**header, **supported}, sort_keys=False, allow_unicode=True, width=120), encoding="utf-8"
    )
    print(f"supported-excel-functions.yaml: {len(supported['functions'])} native functions in {len(supported['categories'])} categories")

    (out_dir / "mind-flags.yaml").write_text(
        yaml.safe_dump({**header, "flags": DOCUMENTED_FLAGS, "markers": DOCUMENTED_MARKERS}, sort_keys=False, allow_unicode=True, width=120),
        encoding="utf-8",
    )
    print(f"mind-flags.yaml: {len(DOCUMENTED_FLAGS)} flags, {len(DOCUMENTED_MARKERS)} markers (all verified present in the KB text)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
