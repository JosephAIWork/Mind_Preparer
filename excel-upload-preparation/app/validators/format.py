"""FMT-*  / FORMAT-001  validators, backed by the KB pages "Excel specificities
kept in Mind" (which cell formats/styles survive import) and "General
structure and guidelines" (no theme colours; Excel styles are not kept on
empty cells; outlines are  supported).

Where the PASS/FAIL condition depends on user intent ("... are
intentional"), the  validator inventories the feature deterministically and
reports  WARNING when present, PASS when absent --  never claims to know the
user's intent itself."""
from __future__ import annotations

from typing import Any

from ..inventory import PRESERVED_FORMAT_CATEGORIES
from ._common import finding


def merged_cells(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    found = {sheet["name"]: sheet["merged_cells"] for wb in analysis["workbooks"] for sheet in wb["sheets"] if sheet["merged_cells"]}
    total = sum(len(v) for v in found.values())
    return finding(
        "WARNING" if total else "PASS",
        f"{total} merged range(s) found; merged cells import with the same aspect, but confirm they are intentional." if total else "No merged cells found.",
        {"merged_ranges_by_sheet": found},
    )


def locked_cells(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FMT-005:  KB --  locked cells become non-clickable in Mind only when the
    sheet is protected ('lock the cell ... then Protect your spreadsheet').
    Excel locks every cell by default, so an unprotected sheet has no
    effective locking  and is not reported."""
    found = {sheet["name"]: sheet["locked_cell_count"] for wb in analysis["workbooks"] for sheet in wb["sheets"] if sheet["locked_cell_count"]}
    total = sum(found.values())
    protected = [sheet["name"] for wb in analysis["workbooks"] for sheet in wb["sheets"] if sheet.get("protected")]
    return finding(
        "WARNING" if total else "PASS",
        f"{total} locked cell(s) on protected sheet(s) {sorted(found)} will not be clickable in Mind; confirm that is intended."
        if total
        else ("No effective cell locking (no protected sheets)." if not protected else f"Protected sheet(s) {protected} but no locked non-empty cells."),
        {"locked_cell_counts_by_protected_sheet": found, "protected_sheets": protected},
    )


def comments(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    found = {sheet["name"]: len(sheet["comments"]) for wb in analysis["workbooks"] for sheet in wb["sheets"] if sheet["comments"]}
    total = sum(found.values())
    return finding(
        "WARNING" if total else "PASS",
        f"{total} cell note(s) found; Mind keeps and displays Excel notes -- confirm the user wants them kept." if total else "No cell comments found.",
        {"comment_counts_by_sheet": found},
    )


def _style_totals(analysis: dict[str, Any], key: str) -> dict[str, int]:
    return {sheet["name"]: sheet["style_stats"][key] for wb in analysis["workbooks"] for sheet in wb["sheets"] if sheet.get("style_stats", {}).get(key)}


def fmt_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FMT-001:  number formats  outside the KB's preserved list."""
    categories: dict[str, int] = {}
    samples: list[dict[str, Any]] = []
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            stats = sheet.get("style_stats", {})
            for cat, n in stats.get("number_format_categories", {}).items():
                categories[cat] = categories.get(cat, 0) + n
            for s in stats.get("unsupported_format_samples", []):
                if len(samples) < 25:
                    samples.append({"sheet": sheet["name"], **s})
    unsupported = {k: v for k, v in categories.items() if k not in PRESERVED_FORMAT_CATEGORIES}
    if unsupported:
        first = samples[0] if samples else None
        return finding(
            "WARNING",
            f"{sum(unsupported.values())} cell(s) use number formats outside the documented preserved set (Standard/Number/Text/Boolean/Date): {unsupported}. "
            "They may display differently in Mind.",
            {"format_categories": categories, "unsupported_categories": unsupported, "samples": samples},
            location={"sheet": first["sheet"], "cell": first["cell"]} if first else None,
            expected=sorted(PRESERVED_FORMAT_CATEGORIES),
        )
    return finding("PASS", "All cell number formats are in the documented preserved set.", {"format_categories": categories}, expected=sorted(PRESERVED_FORMAT_CATEGORIES))


def fmt_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FMT-002:  KB --  'We do not support Themes colors'."""
    found = _style_totals(analysis, "theme_color_cells")
    total = sum(found.values())
    return finding(
        "WARNING" if total else "PASS",
        f"{total} cell(s) use theme colours (font or fill), which Mind does not support -- they will fall back to defaults." if total else "No theme-colour styling found.",
        {"theme_color_cells_by_sheet": found},
    )


def fmt_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FMT-003:  KB --  'Excel styles are not supported for empty cells. Add
    quote plus space to your empty cells to keep Excel styles in Mind.'"""
    found = _style_totals(analysis, "empty_styled_cells")
    total = sum(found.values())
    return finding(
        "WARNING" if total else "PASS",
        f"{total} empty cell(s) carry a solid fill that Mind will not keep (add an apostrophe + space to keep the style)." if total else "No styled empty cells found.",
        {"empty_styled_cells_by_sheet": found},
    )


def fmt_006(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FMT-006:  outlines import  at any depth (group/ungroup per level)."""
    found = {}
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            o = sheet.get("outline", {})
            if o.get("row_groups") or o.get("column_groups") or o.get("hidden_rows") or o.get("hidden_columns"):
                found[sheet["name"]] = o
    if found:
        return finding("WARNING", f"Row/column outlines or hidden rows/columns on {sorted(found)}; Mind imports them (grouping works per depth level) -- confirm they are intentional.", {"outline_by_sheet": found})
    return finding("PASS", "No row/column outlines or hidden rows/columns.", {"outline_by_sheet": {}})


def compare_semantic_features(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FORMAT-001:  preservation inventory  of the features the rule names
    (cell types, basic formatting, merged cells, validation, notes,
    outlines). Nothing is compared *against* in an analysis-only run, so the
    inventory is recorded for the post-transformation comparison
    (scripts/compare_packages.py) and reported  as PASS."""
    inventory = {}
    for wb in analysis["workbooks"]:
        for sheet in wb["sheets"]:
            inventory[sheet["name"]] = {
                "merged_ranges": len(sheet["merged_cells"]),
                "data_validations": len(sheet.get("data_validations", [])),
                "conditional_format_rules": sheet.get("conditional_format_rule_count", 0),
                "notes": len(sheet["comments"]),
                "hyperlinks": len(sheet["hyperlinks"]),
                "outline_groups": sheet.get("outline", {}).get("row_groups", 0) + sheet.get("outline", {}).get("column_groups", 0),
                "format_categories": sheet.get("style_stats", {}).get("number_format_categories", {}),
            }
    non_openpyxl = analysis["features"].get("non_openpyxl_parts", [])
    msg = "Preservation inventory recorded (nothing was transformed in this run, so there is nothing to compare against)."
    if non_openpyxl:
        msg += f" {len(non_openpyxl)} package part(s) exist that only Excel itself can preserve on save; outputs are written through Excel."
    return finding("PASS", msg, {"features_by_sheet": inventory, "excel_only_parts": non_openpyxl[:20]})
