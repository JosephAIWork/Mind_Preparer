"""Real change-set application (CLAUDE_CODE_PROMPT.md phase 8), scoped to
the handful of rules the mined rule set itself marks as safe to
auto-correct (see CHANGELOG.md -- almost everything else needs a human
judgment call by the rules' own design). Every apply function:

  * works on a **fresh copy**, never the file it was handed directly
    in place without a hash check first (non-negotiable rule #1);
  * returns a dict with `output_path`, the write `method`, whether the
    output was verified to open in Excel, and a change-log entry;
  * never claims recalculation succeeded -- call app.recalc.recalculate()
    separately afterward (see app/modes -- VERIFY is a distinct step).

1.3.0: cell edits are written by **Excel through COM** whenever Excel is
available (`Range.Formula` / `Range.FormulaArray` on the copy, then Save).
A pure-Python (openpyxl) save drops package parts it doesn't model and the
result can fail to open in Excel -- see app/excel_com.py. The openpyxl path
is kept as an explicitly-labelled reduced-fidelity fallback.

The change log is an append-only JSON sidecar next to the output file,
never written into the workbook itself (instructions/reporting.md,
README.md's "Stores the change log in the application, not in the
workbook").
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import openpyxl

from .excel_com import FILE_FORMAT_CODES, close_quietly, com_available, excel_session, open_for_write, open_workbook, save_in_place, verify_opens_in_excel
from .inventory import loop_definitions, make_immutable_copy, sha256_of


def _changelog_path(output_path: Path) -> Path:
    return output_path.with_name(output_path.name + ".changelog.json")


def append_change_log(output_path: Path, entry: dict[str, Any]) -> Path:
    log_path = _changelog_path(output_path)
    entries = []
    if log_path.exists():
        entries = json.loads(log_path.read_text(encoding="utf-8"))
    entries.append({**entry, "timestamp": datetime.now(timezone.utc).isoformat()})
    log_path.write_text(json.dumps(entries, indent=2), encoding="utf-8")
    return log_path


def convert_output_format(source_path: Path, work_dir: Path, target_suffix: str) -> dict[str, Any]:
    """FILE-002: Excel COM Save-As to .xlsx/.xlsm. Used both to fix FILE-002
    findings and to convert an uploaded .xlsb before analysis can even run
    (openpyxl can't read .xlsb at all)."""
    if target_suffix not in ("xlsx", "xlsm"):
        return {"status": "ERROR", "message": f"unsupported target format: {target_suffix}"}
    if not com_available():
        return {"status": "NOT_SUPPORTED", "message": "pywin32 is not available; cannot convert via Excel."}

    source_path = Path(source_path).resolve()
    copy_path, source_sha256 = make_immutable_copy(source_path, work_dir)
    output_path = copy_path.with_suffix(f".{target_suffix}")

    def _save_as(excel: Any) -> None:
        wb = None
        try:
            wb = open_workbook(excel, copy_path, read_only=False)
            wb.SaveAs(str(output_path), FileFormat=FILE_FORMAT_CODES[target_suffix])
        finally:
            close_quietly(wb)

    try:
        with excel_session() as excel:
            _save_as(excel)
        entry = {
            "rule_id": "FILE-002",
            "action": "convert_output_format",
            "method": "excel_com",
            "source_path": str(source_path),
            "source_sha256": source_sha256,
            "output_path": str(output_path),
            "output_sha256": sha256_of(output_path),
        }
        append_change_log(output_path, entry)
        return {"status": "APPLIED", "method": "excel_com", "output_path": output_path, "change_log_entry": entry}
    except Exception as exc:  # pragma: no cover -- requires a real Excel install to exercise
        return {"status": "ERROR", "message": f"Excel COM Save-As failed: {exc}"}


def apply_formula_edits(copy_path: Path, edits: list[dict[str, Any]], prefer_excel: bool = True) -> dict[str, Any]:
    """Write `edits` ([{sheet, cell, after}]) into `copy_path` in place --
    the caller guarantees it is already a disposable copy. Excel COM when
    available (full fidelity), openpyxl otherwise (reduced fidelity)."""
    copy_path = Path(copy_path)
    applied: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []
    warnings: list[str] = []

    if prefer_excel and com_available():

        def _edit(excel: Any) -> None:
            wb = None
            try:
                wb = open_for_write(excel, copy_path)
                for edit in edits:
                    rng = None
                    try:
                        rng = wb.Worksheets(edit["sheet"]).Range(edit["cell"])
                        if rng.HasArray:
                            rng.CurrentArray.FormulaArray = edit["after"]
                        else:
                            rng.Formula = edit["after"]
                        applied.append(edit)
                    except Exception as exc:
                        failed.append({**edit, "error": str(exc)[:200]})
                    finally:
                        rng = None
                save_in_place(wb, copy_path)
            finally:
                close_quietly(wb)

        try:
            with excel_session() as excel:
                _edit(excel)
            return {"method": "excel_com", "applied": applied, "failed": failed, "warnings": warnings}
        except Exception as exc:
            warnings.append(f"Excel COM edit failed ({str(exc)[:160]}); fell back to openpyxl.")
            applied, failed = [], []

    is_xlsm = copy_path.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(copy_path, data_only=False, keep_vba=is_xlsm)
    for edit in edits:
        try:
            wb[edit["sheet"]][edit["cell"]].value = edit["after"]
            applied.append(edit)
        except Exception as exc:
            failed.append({**edit, "error": str(exc)[:200]})
    wb.save(copy_path)
    wb.close()
    warnings.append(
        "Written with openpyxl: package parts it cannot model (data model, customXml, slicers, controls...) "
        "are dropped and the file may not open in Excel."
    )
    return {"method": "openpyxl", "applied": applied, "failed": failed, "warnings": warnings}


def plan_loop_case_fix(analysis: dict[str, Any]) -> tuple[dict[str, str], list[dict[str, Any]]]:
    """Compute (renames, edits) for LOOP-002 without touching any file.
    Renames every case-variant spelling of a loop name (in any MM_-function
    string argument) to the single most-used casing."""
    defs = loop_definitions(analysis)
    by_lower: dict[str, dict[str, int]] = {}
    for name, sites in defs.items():
        by_lower.setdefault(name.lower(), {})[name] = len(sites)

    renames: dict[str, str] = {}
    for variants in by_lower.values():
        if len(variants) < 2:
            continue
        canonical = max(variants.items(), key=lambda kv: (kv[1], kv[0]))[0]
        for variant in variants:
            if variant != canonical:
                renames[variant] = canonical
    if not renames:
        return {}, []

    pattern = re.compile(r'"(' + "|".join(re.escape(v) for v in renames) + r')"')

    def _sub(match: re.Match) -> str:
        return f'"{renames[match.group(1)]}"'

    edits = []
    for wb_entry in analysis["workbooks"]:
        for f in wb_entry["formulas"]:
            formula = f["formula"]
            if "MM_" not in formula.upper():
                continue
            new_value, count = pattern.subn(_sub, formula)
            if count:
                edits.append({"sheet": f["sheet"], "cell": f["cell"], "before": formula, "after": new_value})
    return renames, edits


def fix_loop_case_mismatch(source_path: Path, work_dir: Path, analysis: dict[str, Any], prefer_excel: bool = True) -> dict[str, Any]:
    """LOOP-002: rewrite every case-variant spelling of a loop name to the
    single most-used casing. Narrow and genuinely unambiguous -- loop names
    are case-sensitive in Mind, so this can only make previously-broken
    references consistent, never change which loop a formula refers to."""
    renames, edits = plan_loop_case_fix(analysis)
    if not renames:
        return {"status": "NOT_APPLICABLE", "message": "No loop-name case mismatches found."}

    source_path = Path(source_path).resolve()
    copy_path, source_sha256 = make_immutable_copy(source_path, work_dir)
    outcome = apply_formula_edits(copy_path, edits, prefer_excel=prefer_excel)

    verification = verify_opens_in_excel(copy_path) if com_available() else {"opens": None}
    entry = {
        "rule_id": "LOOP-002",
        "action": "fix_loop_case_mismatch",
        "method": outcome["method"],
        "renames": renames,
        "source_path": str(source_path),
        "source_sha256": source_sha256,
        "output_path": str(copy_path),
        "output_sha256": sha256_of(copy_path),
        "changed_cells": outcome["applied"],
        "failed_cells": outcome["failed"],
        "verified_opens_in_excel": verification.get("opens"),
        "warnings": outcome["warnings"],
    }
    append_change_log(copy_path, entry)
    status = "APPLIED" if outcome["applied"] and not outcome["failed"] else ("PARTIAL" if outcome["applied"] else "ERROR")
    return {
        "status": status,
        "method": outcome["method"],
        "output_path": copy_path,
        "verified_opens_in_excel": verification.get("opens"),
        "warnings": outcome["warnings"],
        "message": (
            f"{len(outcome['applied'])} cell(s) rewritten via {outcome['method']}"
            + (f"; {len(outcome['failed'])} failed" if outcome["failed"] else "")
        ),
        "change_log_entry": entry,
    }
