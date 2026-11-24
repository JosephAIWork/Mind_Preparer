"""MMX-*  validators: authoring  environment (KB home page: 'Milliman Mind for
Excel ... Available for Microsoft Excel 2007+ for Windows'; 'How to install
MM4Excel')."""
from __future__ import annotations

from typing import Any

from ._common import finding
from .formula import _known_mm_functions

# KB  article category  -> the rule's classification vocabulary
CATEGORY_MAP = {
    "resize": "Resize",
    "loopbatch": "Loop",
    "dimensions": "Dimension",
    "instance-related": "Instance",
    "distribution-related": "Distribution",
    "copulas-functions": "Distribution",
    "audit-helper-functions": "Audit",
    "aocaos_mm": "AOC/AOS",
    "iteration-related": "Iteration",
    "ranges": "Range (dynamic)",
    "sql-like": "SQL-like",
    "optimized-functions": "Lookup (optimized)",
    "goalseek": "Goal seek",
    "specific-actuarial-functions": "Actuarial",
    "other-functions": "Other",
    "Loop / Dimension": "Loop",
}
INPUT_EXPORT_FUNCTIONS = {"MM_INPUTSTATUS": "Input"}


def mmx_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """MMX-001:  MM_ functions  need the MMForExcel add-in to be maintained /
    recalculated in Excel. A '_xludf.' storage prefix on a formula means the
    file was last saved by an Excel that could not resolve the function."""
    mm = analysis["features"].get("mm_functions_used", {})
    prefixes = analysis["features"].get("formula_storage_prefixes", {})
    xludf = prefixes.get("_xludf.", 0)
    xll = prefixes.get("_xll.", 0)
    if not mm:
        return finding("PASS", "No MM_ functions used; the MMForExcel add-in is not required for this workbook.", {"mm_function_calls": 0})
    total = sum(mm.values())
    if xludf:
        return finding(
            "WARNING",
            f"{total} MM_ function call(s) require the MMForExcel add-in, and {xludf} formula(s) carry the '_xludf.' prefix -- the workbook was last saved without the add-in resolving them (cached values are #NAME?). Re-save with MMForExcel installed before upload.",
            {"mm_function_calls": total, "distinct_mm_functions": sorted(mm), "xludf_prefixed_formulas": xludf, "xll_prefixed_formulas": xll},
        )
    return finding(
        "PASS",
        f"{total} MM_ function call(s) across {len(mm)} distinct function(s); the MMForExcel add-in is required to maintain and recalculate this workbook (Mind itself evaluates them server-side)."
        + (f" {xll} formula(s) are stored as '_xll.' add-in calls, i.e. saved with the add-in loaded." if xll else ""),
        {"mm_function_calls": total, "distinct_mm_functions": sorted(mm), "xludf_prefixed_formulas": 0, "xll_prefixed_formulas": xll},
    )


def mmx_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """MMX-002:  which application  last saved the file (docProps/app.xml)."""
    app = analysis["features"].get("application")
    version = analysis["features"].get("app_version")
    observed = {"application": app, "app_version": version}
    if not app:
        return finding("NOT_SUPPORTED", "docProps/app.xml has no <Application> entry; cannot tell which application maintains this workbook.", observed)
    low = app.lower()
    if "macintosh" in low or "mac" in low.split():
        return finding("WARNING", f"Last saved by '{app}' -- MMForExcel is available for Windows Excel only; maintain the workbook in Windows Excel.", observed)
    if "microsoft excel" in low:
        return finding("PASS", f"Last saved by '{app}' (version {version or 'unknown'}).", observed)
    return finding("WARNING", f"Last saved by '{app}', not Windows Excel -- verify formulas/MM_ functions survived and re-save from Windows Excel with MMForExcel.", observed)


def mmx_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """MMX-003:  classify every  MM_ function used, from the mined registries."""
    known = _known_mm_functions()
    mm = analysis["features"].get("mm_functions_used", {})
    classified: dict[str, list[str]] = {}
    unclassified = []
    for fn in sorted(mm):
        info = known.get(fn)
        if fn in INPUT_EXPORT_FUNCTIONS:
            cls = INPUT_EXPORT_FUNCTIONS[fn]
        elif info is None:
            unclassified.append(fn)
            continue
        elif not info.get("documented", True):
            cls = "Other (mentioned in KB, no article)"
        else:
            cls = CATEGORY_MAP.get(info.get("category") or "", info.get("category") or "Other")
        classified.setdefault(cls, []).append(fn)
    if not mm:
        return finding("PASS", "No MM_ functions to classify.", {"classification": {}, "unclassified": []})
    msg = f"Classified {len(mm) - len(unclassified)} of {len(mm)} MM_ function(s): " + ", ".join(f"{k}: {len(v)}" for k, v in sorted(classified.items()))
    if unclassified:
        return finding("WARNING", msg + f". Not in any registry: {unclassified}", {"classification": classified, "unclassified": unclassified})
    return finding("PASS", msg, {"classification": classified, "unclassified": []})
