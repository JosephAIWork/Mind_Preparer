"""RSK-* validators: upload-risk review items. RSK-003/004/005 are
inventory-style (the doc's own text is "Review hyperlinks." / "Review LET
formulas." / "Inventory external named ranges."); RSK-001/002 are the two
BLOCKER checks the doc phrases as FAIL IF conditions, implemented on the
detected grids (KB 'MM_SETSIZE function' example 3 / 'Dynamic resizing of
grids' for resize drivers; 'MM_RANGE function' for array results)."""
from __future__ import annotations

from typing import Any

from ..formula_utils import ref_text, called_functions, parse_ref
from ..grids import all_grids, grid_containing
from ..inventory import mm_calls
from ._common import finding, fmt_grids, fmt_sites, grid_location

RESIZE_FLAGS = {"resize", "resizerow", "resizecolumn"}


def hyperlinks(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    found = {sheet["name"]: [h["cell"] for h in sheet["hyperlinks"]] for wb in analysis["workbooks"] for sheet in wb["sheets"] if sheet["hyperlinks"]}
    total = sum(len(v) for v in found.values())
    return finding("WARNING" if total else "PASS", f"{total} hyperlink(s) found (HTTP/HTTPS links are kept by Mind); review for upload risk." if total else "No hyperlinks found.", {"hyperlinks_by_sheet": found})


def let_function_usage(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    sites = [f"{f['sheet']}!{f['cell']}" for wb in analysis["workbooks"] for f in wb["formulas"] if "LET" in called_functions(f["formula"])]
    return finding("WARNING" if sites else "PASS", f"{len(sites)} formula(s) use LET(), which is not on the KB supported list; rewrite without it." if sites else "No LET() usage found.", {"let_formula_cells": sites[:100]})


def external_named_ranges(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    external_names = [dn["name"] for wb in analysis["workbooks"] for dn in wb["defined_names"] if isinstance(dn.get("value"), str) and "[" in dn["value"]]
    link_parts = [p for wb in analysis["workbooks"] for p in wb["external_link_parts"]]
    total = len(external_names) or len(link_parts)
    return finding(
        "WARNING" if total else "PASS",
        f"{len(external_names)} defined name(s) referencing external workbooks, {len(link_parts)} external link package part(s). Mind asks at upload whether to preserve links or replace linked cells by assumptions."
        if total
        else "No external named ranges or external link parts found.",
        {"external_named_ranges": external_names, "external_link_parts": link_parts},
    )


def rsk_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RSK-001: a grid with two resize drivers (two MM_SETSIZE cells, or
    conflicting row-resize flags) has an ambiguous dynamic-reference group."""
    setsize_by_grid: dict[tuple[str, str], list[str]] = {}
    for c in mm_calls(analysis, "MM_SETSIZE"):
        ref = parse_ref(c["cell"])
        if not ref:
            continue
        sheet_grids = [g for g in all_grids(analysis) if g["sheet"] == c["sheet"]]
        g = grid_containing(sheet_grids, ref["r1"], ref["c1"])
        key = (c["sheet"], g["anchor"] if g else c["cell"])
        setsize_by_grid.setdefault(key, []).append(c["cell"])
    multi_setsize = [{"sheet": k[0], "cell": k[1], "mm_setsize_cells": v} for k, v in setsize_by_grid.items() if len(v) > 1]

    conflicting = []
    for g in all_grids(analysis):
        row_flags = [f["raw"] for f in g["flags"] if f["name"] in ("resize", "resizerow")]
        col_flags = [f["raw"] for f in g["flags"] if f["name"] == "resizecolumn"]
        if len(row_flags) > 1 or len(col_flags) > 1:
            conflicting.append({"sheet": g["sheet"], "cell": g["anchor"], "grid": g["display_name"], "flags": row_flags + col_flags})
    if multi_setsize or conflicting:
        first = (multi_setsize or conflicting)[0]
        return finding(
            "ERROR",
            (f"{len(multi_setsize)} grid(s) contain more than one MM_SETSIZE: {fmt_sites(multi_setsize)}. " if multi_setsize else "")
            + (f"{len(conflicting)} grid(s) carry duplicate resize flags: {[c['flags'] for c in conflicting]}" if conflicting else ""),
            {"grids_with_multiple_mm_setsize": multi_setsize, "grids_with_duplicate_resize_flags": conflicting},
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    return finding("PASS", "No grid has more than one resize driver.", {"grids_with_multiple_mm_setsize": [], "grids_with_duplicate_resize_flags": []})


def rsk_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RSK-002: array formulas whose target range crosses the grid boundary,
    (MM_RANGE placement -- KB: 'must be contained in a table without any
    other formulas or inputs' -- is RNG-001 in validators/kb.py)."""
    oversized = []
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            if not f.get("array_ref"):
                continue
            target = parse_ref(f["array_ref"])
            origin = parse_ref(f["cell"])
            if not target or not origin:
                continue
            sheet_grids = [g for g in all_grids(analysis) if g["sheet"] == f["sheet"]]
            g = grid_containing(sheet_grids, origin["r1"], origin["c1"])
            if g and (target["r2"] > g["last_row"] or target["c2"] > g["last_col"]):
                # the first cell Mind names: just past the grid's last column (or its last row)
                beyond = ref_text(g["last_col"] + 1, target["r1"]) if target["c2"] > g["last_col"] else ref_text(target["c1"], g["last_row"] + 1)
                oversized.append({"sheet": f["sheet"], "cell": f["cell"], "array_ref": f["array_ref"], "grid_ref": g["ref"], "grid": g["display_name"], "first_cell_outside": beyond, "headers": g["ref"].split(":")[0] + ":" + ref_text(g["last_col"], g["first_row"])})
    if oversized:
        first = oversized[0]
        return finding(
            "ERROR",
            f"{len(oversized)} array formula(s) exceed their grid: Mind's converter refuses the model ('the array formula ... exceeds the grid size. "
            f"Ensure that the column headers cover the entire array formula'). First: {first['sheet']}!{first['cell']} spans {first['array_ref']} but its grid "
            f"{first['grid']} stops at {first['grid_ref'].split(':')[-1]} (first cell outside: {first['first_cell_outside']}). "
            "Fix: extend the header row over every column the array covers (a merged header counts for all the columns it spans), or shorten the array. "
            f"All: {fmt_sites(oversized)} (MM_RANGE placement is checked by RNG-001).",
            {"oversized_array_formulas": oversized[:50]},
            location={"sheet": oversized[0]["sheet"], "cell": oversized[0]["cell"]},
            evidence="INFERENCE",
        )
    n_arrays = analysis["features"].get("array_formula_count", 0)
    return finding("PASS", f"{n_arrays} array formula(s) stay inside their grids.", {"oversized_array_formulas": []}, evidence="INFERENCE")
