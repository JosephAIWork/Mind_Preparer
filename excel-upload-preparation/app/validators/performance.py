"""DBG-*  validators: KB  'Performance tips' ("calculation engine's speed
depends on how many cells it needs to evaluate"; "VLOOKUP() will evaluate
every cell while OFFSET() will only offset to a specific cell"; built-in
performance profiler), 'Enabling debug mode',  and the formula-field FAQ
(partial evaluation with F9). The recommendation rules always  PASS with
RECOMMENDATION evidence; the two measurable  ones (large ranges, lookups)
report what they found."""
from __future__ import annotations

from typing import Any

from ..formula_utils import called_functions, cell_refs_in_formula
from ._common import finding, fmt_sites

LOOKUP_FUNCTIONS = {"VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "XLOOKUP", "XMATCH"}
DEFAULT_LARGE_RANGE_CELLS = 50000


def dbg_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    return finding(
        "PASS",
        "Recommendation: during model development enable Mind's Debug mode (Settings > Debug settings) -- runs stop at the first error, locked cells unlock, hidden rows/columns show, and a Debug/Statistics panel appears (KB 'Enabling debug mode'). Requires a Creator account.",
        {"formula_count": analysis["features"].get("formula_count", 0)},
        evidence="RECOMMENDATION",
    )


def dbg_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    return finding(
        "PASS",
        "Recommendation: turn on the Performance profiler (Settings) and run once; the tachometer icon then shows run time by grid (% of total) and by cell, and in Debug mode where the first NaN originates (KB 'Performance tips').",
        {"grid_count": analysis["features"].get("grid_count", 0)},
        evidence="RECOMMENDATION",
    )


def dbg_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """DBG-003:  whole-column/row and  very large range references."""
    threshold = int(config.get("large_range_cells", DEFAULT_LARGE_RANGE_CELLS))
    hits = []
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            for ref in cell_refs_in_formula(f["formula"]):
                if ref["whole_column"] or ref["whole_row"] or (ref["cells"] or 0) > threshold:
                    hits.append({"sheet": f["sheet"], "cell": f["cell"], "reference": (ref["sheet"] + "!" if ref["sheet"] else "") + ref["ref"], "cells": ref["cells"] or "whole column/row"})
                    break
    if hits:
        return finding(
            "WARNING",
            f"{len(hits)} formula(s) reference whole columns/rows or ranges above {threshold} cells -- Mind evaluates every referenced cell; bound them to the grid (MM_TABLE/MM_COLUMN/MM_ROW): {fmt_sites(hits)}",
            {"large_range_formulas": hits[:50], "threshold_cells": threshold},
            location={"sheet": hits[0]["sheet"], "cell": hits[0]["cell"]},
        )
    return finding("PASS", f"No whole-column/row references or ranges above {threshold} cells.", {"large_range_formulas": [], "threshold_cells": threshold})


def dbg_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """DBG-004:  lookup-function usage  review."""
    usage = analysis["features"].get("function_usage", {})
    counts = {fn: usage[fn] for fn in sorted(LOOKUP_FUNCTIONS & set(usage))}
    if not counts:
        return finding("PASS", "No VLOOKUP/HLOOKUP/LOOKUP/MATCH usage.", {"lookup_usage": {}})
    sites = []
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            if LOOKUP_FUNCTIONS & set(called_functions(f["formula"])):
                sites.append({"sheet": f["sheet"], "cell": f["cell"]})
                if len(sites) >= 20:
                    break
    total = sum(counts.values())
    return finding(
        "WARNING",
        f"{total} lookup call(s) {counts}: each evaluates its whole range in Mind. Prefer OFFSET/INDEX to a known position, MM_READTABLE as an array formula, or MM_VLOOKUPOPTIMIZED/MM_HLOOKUPOPTIMIZED/MM_MATCHOPTIMIZED on large tables. E.g. {fmt_sites(sites)}",
        {"lookup_usage": counts, "sample_sites": sites},
        location={"sheet": sites[0]["sheet"], "cell": sites[0]["cell"]} if sites else None,
    )


def dbg_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    cross_sheet = 0
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            if any(r["sheet"] and r["sheet"] != f["sheet"] for r in cell_refs_in_formula(f["formula"])):
                cross_sheet += 1
    return finding(
        "PASS",
        f"Recommendation: {cross_sheet} formula(s) reference other sheets; use Mind's Audit > Tree view of cell dependencies and the formula navigation (Shift+click a reference) to check dependency chains before upload.",
        {"cross_sheet_formulas": cross_sheet, "formula_count": analysis["features"].get("formula_count", 0)},
        evidence="RECOMMENDATION",
    )


def dbg_006(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    return finding(
        "PASS",
        "Recommendation: in Mind's formula field, select part of a formula and press F9 to evaluate it partially (Shift+hover shows a reference's value) -- the same workflow as Excel's Evaluate Formula for debugging after upload.",
        None,
        evidence="RECOMMENDATION",
    )
