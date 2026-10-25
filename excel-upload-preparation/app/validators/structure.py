"""STR-*/GRID-* validators: structure/grid checks from
01_Mind_Readiness_Standard.md section 3, backed by the KB's
documented grid-detection algorithm (app/grids.py, "General structure and
guidelines"): grids are blocks of non-empty cells, an alone text cell is
ignored, a text-only first row is the header row, '#Name' above a grid names
it, sheet/workbook names become navigation folders."""
from __future__ import annotations

import re
from typing import Any

from ..grids import all_grids
from ._common import DEFAULT_SHEET_NAME_RE, DEFAULT_WORKBOOK_NAME_RE, finding, fmt_grids, fmt_sites, grid_location

HIDE_MARKER = "&&Hide"
STEP_LABELS_MARKER = "steplabels"

# Flags whose grid the KB documents as header-based (columns are matched by
# header name), so a non-text header row is an error rather than a style issue.
HEADER_REQUIRED_FLAGS = {"input", "reorder", "parameters", "projectsettings", "exportsettings", "inputsettings", "instancekeys", "aocorder", "translations", "calculationsteps"}


def hidden_sheets(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-007: hidden sheets should carry the '&&Hide' naming marker."""
    unmarked = []
    marked_but_visible = []
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            has_marker = HIDE_MARKER.lower() in sheet["name"].lower()
            is_hidden = sheet["state"] != "visible"
            if is_hidden and not has_marker:
                unmarked.append(sheet["name"])
            if has_marker and not is_hidden:
                marked_but_visible.append(sheet["name"])

    if not unmarked and not marked_but_visible:
        return finding("PASS", "All hidden sheets use the '&&Hide' naming marker (or none are hidden).", {"unmarked_hidden_sheets": [], "marked_but_visible_sheets": []})
    return finding(
        "WARNING",
        f"{len(unmarked)} hidden sheet(s) without '&&Hide' in the name, {len(marked_but_visible)} sheet(s) named with '&&Hide' but not actually hidden.",
        {"unmarked_hidden_sheets": unmarked, "marked_but_visible_sheets": marked_but_visible},
    )


def grid_detection_heuristic(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-001: grids as Mind would detect them. A '#Title' cell *inside* a
    detected block is the deterministic signature of two logical grids that
    were merged (no empty row between the first grid and the next title)."""
    per_sheet = {}
    merged = []
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            per_sheet[sheet["name"]] = len(sheet["grids"])
            for g in sheet["grids"]:
                if g.get("inner_title_cells"):
                    merged.append({"sheet": sheet["name"], "cell": g["anchor"], "grid": g["display_name"], "inner_titles": g["inner_title_cells"]})
    if merged:
        return finding(
            "WARNING",
            f"{len(merged)} detected grid(s) contain another '#Title' cell inside their block -- Mind will read them as one grid. "
            f"Insert an empty row before each inner title: {fmt_sites(merged)}",
            {"grid_blocks_per_sheet": per_sheet, "merged_grid_candidates": merged},
            location={"sheet": merged[0]["sheet"], "cell": merged[0]["cell"]},
            evidence="INFERENCE",
        )
    return finding(
        "PASS",
        f"{sum(per_sheet.values())} grid block(s) detected with the documented algorithm; no title cell is trapped inside another grid.",
        {"grid_blocks_per_sheet": per_sheet},
        evidence="INFERENCE",
    )


def grid_naming_heuristic(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-004: grids without a '#Name' title become 'untitled <cell>' in Mind."""
    grids = all_grids(analysis)
    untitled = [g for g in grids if not g.get("name")]
    named = len(grids) - len(untitled)
    if not untitled:
        return finding("PASS", f"All {named} detected grid(s) have a '#Name' title cell.", {"named_grids": named, "untitled_grids": []}, evidence="INFERENCE")
    return finding(
        "WARNING",
        f"{len(untitled)} of {len(grids)} detected grid(s) have no '#Name' title and will appear as 'untitled <cell>' in Mind: {fmt_grids(untitled)}",
        {"named_grids": named, "untitled_grids": [f"{g['sheet']}!{g['ref']}" for g in untitled[:50]]},
        location=grid_location(untitled[0]),
        evidence="INFERENCE",
    )


def inventory_hierarchy(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """GRID-001: workbook/sheet/grid inventory -- always satisfiable from the analysis itself."""
    sheet_count = sum(len(wb["sheets"]) for wb in analysis["workbooks"])
    grids = all_grids(analysis)
    flagged = [g for g in grids if g["flags"]]
    flag_counts: dict[str, int] = {}
    for g in flagged:
        for name in g["flag_names"]:
            flag_counts[name] = flag_counts.get(name, 0) + 1
    return finding(
        "PASS",
        f"Inventoried {len(analysis['workbooks'])} workbook(s), {sheet_count} sheet(s), {len(grids)} grid(s) ({len(flagged)} flagged).",
        {
            "workbook_count": len(analysis["workbooks"]),
            "sheet_count": sheet_count,
            "grid_count": len(grids),
            "flag_counts": flag_counts,
            "grids": [{"sheet": g["sheet"], "ref": g["ref"], "name": g["display_name"], "flags": g["flag_names"]} for g in grids[:200]],
        },
    )


def step_labels_isolated(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """GRID-004: a /StepLabels grid must be alone on its own sheet (KB: 'this
    grid must be alone in a separate spreadsheet or an error will be thrown at import')."""
    violations = []
    found = []
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            step_grids = [g for g in sheet["grids"] if "steplabels" in g["flag_names"] or g.get("looks_like_step_labels")]
            name_marker = STEP_LABELS_MARKER in sheet["name"].replace(" ", "").lower()
            if not step_grids and not name_marker:
                continue
            found.append(sheet["name"])
            other_content = len(sheet["grids"]) - len(step_grids) + len(sheet.get("standalone_text_cells", []))
            if len(step_grids) > 1 or other_content > 0:
                violations.append({"sheet": sheet["name"], "cell": step_grids[0]["anchor"] if step_grids else "A1", "other_blocks": other_content})

    if not found:
        return finding("PASS", "No StepLabels grid detected.", {"step_labels_sheets": []})
    if violations:
        return finding(
            "ERROR",
            f"StepLabels grid shares a sheet with other content (Mind throws an error at import): {fmt_sites(violations)}",
            {"sheets_with_extra_content": violations},
            location={"sheet": violations[0]["sheet"], "cell": violations[0]["cell"]},
        )
    return finding("PASS", "StepLabels grid(s) are alone on their sheet.", {"step_labels_sheets": found})


def str_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-002: alone text cells are ignored by Mind -- surface them so the user
    can confirm they are not expected to become grids/navigation items."""
    cells = []
    orphan_titles = []
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            for c in sheet.get("standalone_text_cells", []):
                (orphan_titles if c.get("looks_like_title") else cells).append(c)
    if not cells and not orphan_titles:
        return finding("PASS", "No standalone text cells (every text cell belongs to a grid or a title).", {"standalone_text_cells": []}, evidence="INFERENCE")
    msg = f"{len(cells)} standalone text cell(s) that Mind will ignore"
    if orphan_titles:
        msg += f"; {len(orphan_titles)} '#Title' cell(s) with no grid directly below them (the title is lost): {fmt_sites(orphan_titles)}"
    if cells:
        msg += f". Standalone cells: {fmt_sites(cells)}"
    first = (orphan_titles or cells)[0]
    return finding(
        "WARNING",
        msg,
        {"standalone_text_cells": [f"{c['sheet']}!{c['cell']}" for c in cells[:100]], "orphan_titles": [f"{c['sheet']}!{c['cell']}" for c in orphan_titles]},
        location={"sheet": first["sheet"], "cell": first["cell"]},
        evidence="INFERENCE",
    )


def str_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-003: a header row is recognized only if it is text-only; a mixed
    first row (text + numbers/formulas) is the failure the doc describes."""
    mixed = []
    required_broken = []
    no_headers = 0
    for g in all_grids(analysis):
        hv = g.get("header_values", [])
        texts = sum(1 for v in hv if isinstance(v, str) and not v.startswith("="))
        if g["header_is_all_text"]:
            continue
        needs_headers = bool(HEADER_REQUIRED_FLAGS & set(g["flag_names"])) and "noheader" not in g["flag_names"] and "noheaders" not in g["flag_names"]
        if needs_headers:
            required_broken.append(g)
        elif 0 < texts < len(hv):
            mixed.append(g)
        else:
            no_headers += 1
    if required_broken:
        return finding(
            "ERROR",
            f"{len(required_broken)} flagged grid(s) whose columns Mind matches by header have a non-text first row: {fmt_grids(required_broken)}",
            {"header_required_but_not_text": [f"{g['sheet']}!{g['ref']}" for g in required_broken], "mixed_header_rows": [f"{g['sheet']}!{g['ref']}" for g in mixed]},
            location=grid_location(required_broken[0]),
        )
    if mixed:
        return finding(
            "WARNING",
            f"{len(mixed)} grid(s) have a first row mixing text with numbers/formulas -- Mind will not recognize headers and will show 'Col n': {fmt_grids(mixed)}",
            {"mixed_header_rows": [f"{g['sheet']}!{g['ref']}" for g in mixed], "grids_without_text_headers": no_headers},
            location=grid_location(mixed[0]),
            evidence="INFERENCE",
        )
    return finding(
        "PASS",
        f"No mixed header rows; {no_headers} grid(s) start with a non-text row (shown as 'Col n' in Mind, which is fine if no headers were intended).",
        {"mixed_header_rows": [], "grids_without_text_headers": no_headers},
        evidence="INFERENCE",
    )


def str_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-005: sheet names are folder names in Mind -- flag Excel defaults."""
    defaults = [s["name"] for wb in analysis["workbooks"] for s in wb["sheets"] if DEFAULT_SHEET_NAME_RE.match(s["name"].strip())]
    if defaults:
        return finding("WARNING", f"Default sheet name(s) will become navigation folders in Mind: {defaults}", {"default_sheet_names": defaults})
    return finding("PASS", "No default (Sheet1-style) sheet names.", {"default_sheet_names": []})


def str_006(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-006: the workbook name is a navigation folder and an instance-key header."""
    problems = []
    for wb in analysis["workbooks"]:
        name = wb.get("file_name") or ""
        low = name.lower()
        if DEFAULT_WORKBOOK_NAME_RE.match(name.strip()):
            problems.append({"workbook": name, "issue": "default Excel workbook name"})
        elif "copy of" in low or low.endswith(" - copy") or " - copy (" in low or low.startswith("new "):
            problems.append({"workbook": name, "issue": "looks like a copy / temporary name"})
        elif any(ch in name for ch in '\\/:*?"<>|'):
            problems.append({"workbook": name, "issue": "contains characters unusable in file names"})
    if problems:
        return finding("WARNING", f"Workbook name should be meaningful (it is used as a folder name and instance-key header in Mind): {problems}", {"workbook_name_issues": problems})
    return finding("PASS", "Workbook name looks deliberate.", {"workbook_names": [wb.get("file_name") for wb in analysis["workbooks"]]})


def broken_defined_name_refs(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """REF-001: defined names that resolve to #REF! are broken, and a name no
    definition exists for (a misspelt name, a name deleted from the workbook)
    is broken too -- Excel shows #NAME?. A formula that references either
    cannot convert in Milliman Mind (the converter reports 'No function found
    ... (not a function : <name>)'). Flag every referencing cell (a hard
    conversion blocker); if broken names exist but nothing references them, warn."""
    import difflib

    from ..formula_utils import mask_strings, undefined_names

    wb0 = analysis["workbooks"][0]
    defined = [str(d["name"]) for d in wb0.get("defined_names", [])]
    tables = [str(t) for s in wb0.get("sheets", []) for t in s.get("tables", [])]
    known = {n.upper() for n in defined} | {t.upper() for t in tables}
    broken = sorted({d["name"] for d in wb0.get("defined_names", []) if "#REF!" in str(d.get("value") or "")})
    name_re = re.compile(r"(?<![A-Za-z0-9_.])(" + "|".join(re.escape(n) for n in broken) + r")(?![A-Za-z0-9_.])") if broken else None
    lit_re = re.compile(r"#REF!")
    sites = []
    missing: dict[str, int] = {}
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        masked = mask_strings(formula)
        kinds = []
        if name_re is not None and name_re.search(masked):
            kinds.append("broken-name")
        if lit_re.search(masked):
            kinds.append("literal-#REF!")
        unknown = undefined_names(formula, known)
        if unknown:
            kinds.append("undefined-name")
            for n in unknown:
                missing[n] = missing.get(n, 0) + 1
        if kinds:
            site = {"sheet": f["sheet"], "cell": f["cell"], "kinds": kinds, "formula": formula[:140]}
            if unknown:
                site["undefined_names"] = unknown
            sites.append(site)
    # the closest existing name, for a misspelt one
    closest = {n: (difflib.get_close_matches(n, defined + tables, n=1, cutoff=0.75) or [None])[0] for n in missing}
    if sites:
        parts = []
        if any("undefined-name" in s["kinds"] for s in sites):
            named = ", ".join(f"{n} ({c} cell(s)" + (f", closest existing name: {closest[n]}" if closest[n] else "") + ")" for n, c in sorted(missing.items(), key=lambda kv: -kv[1])[:8])
            parts.append(f"name(s) that do not exist in the workbook (Excel shows #NAME?): {named}")
        if any(k in s["kinds"] for s in sites for k in ("broken-name", "literal-#REF!")):
            parts.append("#REF! literals or references to #REF! defined names")
        return finding(
            "ERROR",
            f"{len(sites)} cell(s) contain broken references -- {'; '.join(parts)}. Mind's converter cannot compile them "
            f"('No function found ... not a function'). First: {sites[0]['sheet']}!{sites[0]['cell']}. "
            "Fix: a misspelt name is corrected to the existing one; a broken reference is replaced with NA() (an error stays an error) or the cell is deleted if it has no impact. "
            "Where a function needs a cell range (SUMIF(S)/COUNTIF(S)/AVERAGEIF(S)/MAXIFS/MINIFS/OFFSET/ROW/COLUMN/SUBTOTAL), "
            "Excel refuses NA() -- replace that whole call with NA() instead (e.g. =SUMIFS('Data'!#REF!,J:J,\"x\") -> =NA()).",
            {"broken_names": broken, "undefined_names": {n: {"cells": c, "closest": closest[n]} for n, c in missing.items()}, "sites": sites[:80]},
            location={"sheet": sites[0]["sheet"], "cell": sites[0]["cell"]},
        )
    if broken:
        return finding(
            "WARNING",
            f"{len(broken)} broken (#REF!) defined name(s) exist but no formula references them: {broken[:12]}. Consider deleting them.",
            {"broken_names": broken, "sites": []},
            evidence="RECOMMENDATION",
        )
    return finding("PASS", "No broken (#REF!) references.", {"broken_names": [], "sites": []})


_SUM_RE = re.compile(r"^=SUM\(([^()]+)\)$", re.I)
_REF_TOK = r"\$?[A-Za-z]{1,3}\$?\d{1,7}"
_CHAIN_RE = re.compile(r"^=(" + _REF_TOK + r"(?:\+" + _REF_TOK + r")+)$")


def _parse_total(formula: str):
    """A pure total formula -> ('sum', 'A1:A10') or ('chain', ['A1','A2',...]) else None.
    Cross-sheet / mixed / partial expressions are intentionally not treated as totals."""
    f = (formula or "").replace(" ", "")
    m = _SUM_RE.match(f)
    if m:
        rng = m.group(1)
        if ":" in rng and "!" not in rng and "," not in rng:
            return ("sum", rng.replace("$", ""))
    m = _CHAIN_RE.match(f)
    if m:
        return ("chain", [r.replace("$", "") for r in m.group(1).split("+")])
    return None


def totals_reconcile(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """REP-001: a total cell (=SUM(range) or =a+b+c) whose cached value does not equal
    the sum of its addends' cached values -- the numbers do not reconcile (a stale value
    or a total overridden by a hard-coded number). Report-only; never changes numbers."""
    from ..formula_utils import parse_ref
    from ..inventory import _values_workbook, ignored_sheet_names

    wb0 = analysis["workbooks"][0]
    try:
        vals = _values_workbook(wb0["copy_path"], ignored_sheet_names(analysis))
    except Exception:
        vals = None
    if vals is None:
        return finding("NOT_SUPPORTED", "Cached values are unavailable, so totals cannot be reconciled.", {"mismatches": []})

    def num(v):
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    checked = 0
    mismatches = []
    for f in wb0.get("formulas", []):
        parsed = _parse_total(f.get("formula") or "")
        if not parsed:
            continue
        sheet, cell = f["sheet"], f["cell"]
        if sheet not in vals.sheetnames:
            continue
        ws = vals[sheet]
        total_val = num(ws[cell].value)
        if total_val is None:
            continue
        kind, ref = parsed
        addends = []
        if kind == "sum":
            r = parse_ref(ref)
            if not r or r["whole_column"] or r["whole_row"] or r["cells"] > 5000:
                continue
            for rr in range(r["r1"], r["r2"] + 1):
                for cc in range(r["c1"], r["c2"] + 1):
                    addends.append(num(ws.cell(rr, cc).value))
        else:
            for rf in ref:
                try:
                    addends.append(num(ws[rf].value))
                except Exception:
                    addends.append(None)
        nums = [a for a in addends if a is not None]
        if not nums or len(nums) != len(addends):
            continue  # non-numeric addends: cannot reconcile cleanly, skip
        checked += 1
        total = sum(nums)
        tol = max(1e-6, 1e-6 * abs(total))
        if abs(total_val - total) > tol:
            mismatches.append({"sheet": sheet, "cell": cell, "formula": (f.get("formula") or "")[:80],
                               "cell_value": total_val, "sum_of_addends": total, "diff": total_val - total})
    if mismatches:
        return finding(
            "WARNING",
            f"{len(mismatches)} of {checked} checked total(s) do not equal the sum of their addends -- the numbers do not "
            "reconcile (a stale value, or a total overridden by a hard-coded number). Recalculate the workbook or correct the value.",
            {"mismatches": mismatches[:50], "checked": checked},
        )
    return finding("PASS", f"All {checked} checked SUM/addition total(s) reconcile with their addends.", {"mismatches": [], "checked": checked})


from .references import external_links_in_formulas  # noqa: E402,F401  (REF-002)



def lone_text_cells_read(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = 80) -> dict:
    """STR-008: a text cell with nothing around it (no grid, no title above, no label beside it)
    that a formula reads. Mind imports grids; a text cell standing alone is not in the model,
    and the formula gets nothing (seen in a real Mind conversion: a lone text cell G2 that
    Prep had separated from its label; "Name#"&$G$2 came back as "Name#" and 20,000 cells
    followed). A number standing alone is imported; so is a text under a '#' title, next to
    or under a label, or inside a grid (test workbook, 2026-10-08)."""
    from ..prep import reference_index, referenced_by, standalone_labels

    try:
        index = reference_index(analysis)
    except Exception as exc:
        return finding("NOT_SUPPORTED", f"The references could not be indexed ({type(exc).__name__}: {str(exc)[:120]}).", {"sites": []})
    sites = []
    for lab in standalone_labels(analysis):
        owner = referenced_by(index, lab["sheet"], lab["row"], lab["col"])
        if owner:
            sites.append({"sheet": lab["sheet"], "cell": lab["cell"], "text": str(lab["text"])[:80], "read_by": owner})
    observed = {"sites": sites if limit is None else sites[:limit], "sites_count": len(sites)}
    if not sites:
        return finding("PASS", "No text cell standing alone on its sheet is read by a formula.", observed)
    first = sites[0]
    return finding(
        "ERROR",
        f"{len(sites)} text cell(s) standing alone on their sheet -- no grid, no title above, no label beside -- are read by formulas. Mind imports grids, "
        "not loose text cells: the formula gets nothing (confirmed in a real Mind conversion: \"Name#\"&$G$2, with G2 a lone text cell, came back as \"Name#\" "
        f"and 20,000 cells followed). First: {first['sheet']}!{first['cell']} = {first['text']!r}, read by {first['read_by']}. "
        "The Prep action 'Give a lone text cell that formulas read a title of its own' writes a '#' title in the empty cell above it, which makes it a one-cell grid Mind imports "
        "(a number standing alone is imported as it is; a text next to or under a label too).",
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )


def blocking_pivot(p: dict[str, Any]) -> str | None:
    """Why Mind's converter cannot take this pivot table. A client model was rejected ("The method or
    operation is not implemented") with its pivot tables, one of them with a grouped field, and
    accepted with all of them as values (2026-10-18); with only the grouped one as values and the
    plain ones kept, it was rejected again (2026-10-25). So every pivot table blocks."""
    if p.get("grouped_fields"):
        return f"its field {', '.join(p['grouped_fields'][:3])} is grouped"
    return "Mind's converter stops on a pivot table"


def pivot_tables(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """STR-009: the workbook's pivot tables, every one blocking (blocking_pivot). The Prep action turns
    them into formulas or values."""
    pivots = []
    for s in analysis["workbooks"][0].get("sheets", []):
        for p in s.get("pivot_tables", []) or []:
            pivots.append({"sheet": s["name"], "name": p["name"], "ref": p["ref"], "blocking": blocking_pivot(p)})
    observed = {"pivots": pivots, "count": len(pivots), "blocking": len(pivots)}
    if not pivots:
        return finding("PASS", "The workbook has no pivot table.", observed)
    first = pivots[0]
    return finding(
        "ERROR",
        f"{len(pivots)} pivot table(s), first {first['sheet']}!{first['ref']} ({first['name']}): Mind's converter stops on a pivot table. "
        "In a real Mind conversion a workbook was refused (\"The method or operation is not implemented\") with its pivot tables, again with only the one with a grouped field "
        "turned into values (2026-10-18, 2026-10-25), and accepted once all of them were plain cells. "
        "The Prep action 'Pivot tables' turns them into formulas over their source (values where formulas cannot give the same numbers), or into plain values.",
        observed,
        location={"sheet": first["sheet"], "cell": first["ref"].split(":")[0]},
    )
