"""FRM-*/FORMULA-* validators.

FRM-002 checks every function call against two mined lists:
  * MM_ functions: references/mm-function-registry.yaml (doc 02) UNION
    references/mm-function-registry-kb.yaml (the KB scrape,
    tools/mine_kb_docx.py) -- doc 02's own rule 1.1 is "do not invent MM
    function syntax; if a function is not listed, do not auto-generate or
    auto-correct it";
  * native Excel functions: references/supported-excel-functions.yaml, the
    KB page "Supported Excel formulas" (1.3.0 -- previously unavailable).
A name that is on neither list and is not an Excel function at all is a
probable VBA UDF / custom formula and is reported by FORMULA-002 instead.
"""
from __future__ import annotations

import re
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

from ..formula_utils import STORAGE_PREFIX_RE, called_functions, mask_strings
from ._common import finding, fmt_sites
from .excel_functions import EXCEL_FUNCTIONS

REFERENCES = Path(__file__).resolve().parents[2] / "references"
REGISTRY_PATH = REFERENCES / "mm-function-registry.yaml"
KB_REGISTRY_PATH = REFERENCES / "mm-function-registry-kb.yaml"
SUPPORTED_PATH = REFERENCES / "supported-excel-functions.yaml"

# Storage-only tokens Excel writes for dynamic-array features; reviewed by FRM-003, not FRM-002.
DYNAMIC_ARRAY_ARTIFACTS = {"SINGLE", "ANCHORARRAY"}

# Native functions the KB "Supported Excel formulas" scrape omits but real Milliman Mind
# conversion actually accepts (confirmed by uploading models to Mind). The KB page (163)
# predates the dynamic-array functions.
MIND_CONFIRMED_SUPPORTED = {"FILTER"}  # Shlomo_IFRS model converted cleanly using FILTER (Cashflows!M4), 2026-08-30
IMPLICIT_INTERSECTION_RE = re.compile(r"@(?=[A-Za-z$'(])")
SPILL_REF_RE = re.compile(r"\$?[A-Z]{1,3}\$?\d{1,7}#")


@lru_cache(maxsize=1)
def _known_mm_functions() -> dict[str, dict[str, Any]]:
    known: dict[str, dict[str, Any]] = {}
    for path, source in ((REGISTRY_PATH, "02_MM_Function_Registry.md"), (KB_REGISTRY_PATH, "CompleteMindDocn.docx (KB)")):
        if not path.exists():
            continue
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        for f in data.get("functions", []):
            name = (f.get("function") or "").upper()
            if not name:
                continue
            entry = known.setdefault(name, {"sources": [], "category": f.get("category"), "deprecated": False, "documented": True})
            entry["sources"].append(source)
            if f.get("deprecated"):
                entry["deprecated"] = True
            if f.get("documented") is False and len(entry["sources"]) == 1:
                entry["documented"] = False
    return known


@lru_cache(maxsize=1)
def _supported_native_functions() -> set[str]:
    if not SUPPORTED_PATH.exists():
        return set()
    data = yaml.safe_load(SUPPORTED_PATH.read_text(encoding="utf-8")) or {}
    return {f.upper() for f in data.get("functions", [])} | MIND_CONFIRMED_SUPPORTED


def _first_sites(analysis: dict[str, Any], names: set[str], limit_per_name: int = 1) -> dict[str, list[dict[str, Any]]]:
    sites: dict[str, list[dict[str, Any]]] = {n: [] for n in names}
    remaining = set(names)
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            hits = remaining & set(called_functions(f["formula"]))
            for h in hits:
                sites[h].append({"sheet": f["sheet"], "cell": f["cell"], "formula": STORAGE_PREFIX_RE.sub("", f["formula"])[:200]})
                if len(sites[h]) >= limit_per_name:
                    remaining.discard(h)
            if not remaining:
                return sites
    return sites


def unsupported_functions(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    known_mm = _known_mm_functions()
    supported_native = _supported_native_functions()
    usage = analysis["features"].get("function_usage", {})

    unknown_mm = sorted(fn for fn in usage if fn.startswith("MM_") and fn not in known_mm)
    deprecated_mm = sorted(fn for fn in usage if fn.startswith("MM_") and known_mm.get(fn, {}).get("deprecated"))
    unsupported_native = sorted(
        fn for fn in usage if not fn.startswith("MM_") and fn in EXCEL_FUNCTIONS and fn not in supported_native and fn not in DYNAMIC_ARRAY_ARTIFACTS
    )
    not_excel = sorted(fn for fn in usage if not fn.startswith("MM_") and fn not in EXCEL_FUNCTIONS and fn not in DYNAMIC_ARRAY_ARTIFACTS)

    flagged = set(unknown_mm) | set(unsupported_native)
    sites = _first_sites(analysis, flagged, limit_per_name=25) if flagged else {}
    call_sites = [
        {"functions": [fn], **site}
        for fn in sorted(flagged)
        for site in sites.get(fn, [])
    ]
    observed = {
        "unregistered_mm_functions": unknown_mm,
        "deprecated_mm_functions": deprecated_mm,
        "unsupported_native_functions": unsupported_native,
        "not_excel_functions_see_FORMULA_002": not_excel,
        "usage_counts": {fn: usage[fn] for fn in sorted(flagged)},
        "call_sites": call_sites,
        "known_mm_registry_size": len(known_mm),
        "supported_native_list_size": len(supported_native),
    }
    if flagged:
        parts = []
        if unknown_mm:
            parts.append(f"MM_ function(s) not in either mined registry (do not invent syntax): {unknown_mm}")
        if unsupported_native:
            parts.append(f"native Excel function(s) not on the KB 'Supported Excel formulas' list: {unsupported_native}")
        if deprecated_mm:
            parts.append(f"deprecated MM_ function(s): {deprecated_mm}")
        return finding(
            "ERROR",
            "; ".join(parts) + f". First call sites: {fmt_sites(call_sites)}",
            observed,
            location={"sheet": call_sites[0]["sheet"], "cell": call_sites[0]["cell"]} if call_sites else None,
        )
    msg = f"All {len(usage)} distinct function(s) are either registered MM_ functions or on the KB supported-native list."
    if deprecated_mm:
        msg += f" Deprecated MM_ function(s) in use: {deprecated_mm}."
    if not_excel:
        msg += f" Unknown non-Excel name(s) reviewed by FORMULA-002: {not_excel}."
    return finding("WARNING" if deprecated_mm else "PASS", msg, observed)


def frm_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FRM-001: formula inventory."""
    per_sheet = {s["name"]: s["formula_count"] for wb in analysis["workbooks"] for s in wb["sheets"] if s["formula_count"]}
    usage = analysis["features"].get("function_usage", {})
    top = sorted(usage.items(), key=lambda kv: -kv[1])[:25]
    return finding(
        "PASS",
        f"{analysis['features'].get('formula_count', 0)} formula(s) ({analysis['features'].get('array_formula_count', 0)} array formulas) using {len(usage)} distinct function(s); "
        f"{len(analysis['features'].get('mm_functions_used', {}))} distinct MM_ function(s).",
        {"formulas_per_sheet": per_sheet, "top_functions": top, "mm_functions": sorted(analysis["features"].get("mm_functions_used", {}))},
    )


def frm_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FRM-003: dynamic-array / implicit-intersection review (KB 'Implicit
    intersection operator since Office 365')."""
    at_sites, array_sites, spill_sites = [], [], []
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            formula = f["formula"]
            fns = set(called_functions(formula))
            masked = mask_strings(formula)
            if "SINGLE" in fns or IMPLICIT_INTERSECTION_RE.search(masked):
                at_sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:160]})
            if f.get("array_ref"):
                array_sites.append({"sheet": f["sheet"], "cell": f["cell"], "array_ref": f["array_ref"], "formula": formula[:160]})
            if "ANCHORARRAY" in fns or SPILL_REF_RE.search(masked):
                spill_sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:160]})
    observed = {"implicit_intersection_sites": at_sites[:50], "array_formula_sites": array_sites[:50], "spill_reference_sites": spill_sites[:50],
                "counts": {"implicit_intersection": len(at_sites), "array_formulas": len(array_sites), "spill_references": len(spill_sites)}}
    if spill_sites:
        # Mind's own validation rejects every one of these as "Unsupported formula: ANCHORARRAY()"
        # (the token Excel stores for a spilled-range reference such as INDEX(A1#, ...)).
        first = spill_sites[0]
        return finding(
            "ERROR",
            f"{len(spill_sites)} formula(s) reference spilled ranges ('A1#', stored as ANCHORARRAY) -- Mind rejects each of them with "
            f"'Unsupported formula: ANCHORARRAY()'; Prep replaces the reference by the fixed range the spill covers today (Mind has no Office 365 dynamic arrays)"
            + (f"; {len(at_sites)} formula(s) also use the implicit-intersection '@' operator" if at_sites else "")
            + f". Sites: {fmt_sites(spill_sites + at_sites)}",
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if at_sites:
        first = at_sites[0]
        return finding(
            "WARNING",
            f"{len(at_sites)} formula(s) use the implicit-intersection '@' operator; "
            f"Mind resizes arrays with MM_RANGE / MM_GETRANGE, not Office 365 dynamic arrays. Review: {fmt_sites(at_sites)}",
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if array_sites:
        return finding(
            "PASS",
            f"No '@' / spill references; {len(array_sites)} fixed-size array formula(s) present (supported; sizes checked by RSK-002).",
            observed,
        )
    return finding("PASS", "No implicit-intersection operators, spill references or array formulas.", observed)


def _vba_bytes(analysis: dict[str, Any]) -> bytes:
    for wb in analysis["workbooks"]:
        try:
            with zipfile.ZipFile(wb["copy_path"]) as zf:
                if "xl/vbaProject.bin" in zf.namelist():
                    return zf.read("xl/vbaProject.bin")
        except (zipfile.BadZipFile, OSError):
            pass
    return b""


def detect_vba_udf_calls(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FORMULA-002: a called name that is neither an MM_ function nor any
    native Excel function is a user-defined function (VBA or add-in). When
    a VBA project is present its binary is searched (read-only, never
    executed) for a 'Function <name>' declaration to confirm."""
    usage = analysis["features"].get("function_usage", {})
    candidates = sorted(fn for fn in usage if not fn.startswith("MM_") and fn not in EXCEL_FUNCTIONS and fn not in DYNAMIC_ARRAY_ARTIFACTS)
    if not candidates:
        return finding("PASS", "Every called function is a native Excel or MM_ function; no VBA user-defined functions are called.", {"udf_candidates": []})

    vba = _vba_bytes(analysis)
    vba_lower = vba.lower()
    confirmed = []
    for name in candidates:
        n = name.lower().encode("ascii", errors="ignore")
        n16 = name.lower().encode("utf-16-le", errors="ignore")
        if vba and (b"function " + n in vba_lower or n16.lower() in vba_lower or n in vba_lower):
            confirmed.append(name)
    sites = _first_sites(analysis, set(candidates))
    call_sites = [{"functions": [fn], **sites[fn][0]} for fn in candidates if sites.get(fn)]
    observed = {"udf_candidates": candidates, "confirmed_in_vba_project": confirmed, "has_vba": analysis["features"]["has_vba"], "call_sites": call_sites}
    loc = {"sheet": call_sites[0]["sheet"], "cell": call_sites[0]["cell"]} if call_sites else None
    if analysis["features"]["has_vba"]:
        return finding(
            "ERROR",
            f"{len(candidates)} called function(s) are not Excel or MM_ functions and the workbook has a VBA project -- VBA does not run in Mind: {candidates} "
            f"(found by name in the VBA binary: {confirmed or 'none'}). Sites: {fmt_sites(call_sites)}",
            observed,
            location=loc,
        )
    return finding(
        "WARNING",
        f"{len(candidates)} called function(s) are neither native Excel nor MM_ functions (custom C# formula registered in Mind, or a missing add-in?): {candidates}. Sites: {fmt_sites(call_sites)}",
        observed,
        location=loc,
    )


def vba_presence(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FRM-004's actual FAIL condition is 'VBA logic is required' -- not merely
    present. Mind silently ignores VBA it can't run; it does not block upload
    over it. Presence alone can't tell us whether the workbook's *correctness*
    depends on a macro running, so this is a WARNING (confirm no needed
    behavior depends on it) rather than a blocking ERROR."""
    has_vba = analysis["features"]["has_vba"]
    return finding(
        "WARNING" if has_vba else "PASS",
        "Workbook contains a VBA project. This does not block upload -- Mind ignores VBA it "
        "can't run -- but confirm nothing in the workbook depends on the macro actually "
        "executing, since it won't in Mind (UDF calls are checked by FORMULA-002)."
        if has_vba
        else "No VBA project detected.",
        {"has_vba": has_vba, "vba_project_bytes": analysis["features"].get("vba_project_bytes", 0)},
    )
