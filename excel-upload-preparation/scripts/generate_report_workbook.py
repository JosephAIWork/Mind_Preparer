#!/usr/bin/env python3
"""Run  PLAN_MODE analysis  on a workbook and produce the reviewable Excel
outputs: a standalone report (.xlsx) and the analyzed workbook with the
Mind_Readiness_Report sheets appended -- the latter written by Excel through
COM when available and re-opened  by Excel to verify it loads (see
app/excel_report.py). See scripts/validate_workbook.py for the  raw JSON form
(scripting/automation)."""
from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.config import load_config  # noqa: E402
from app.excel_report import build_report_workbook, build_standalone_report  # noqa: E402
from app.modes import plan_mode  # noqa: E402
from app.rules_engine import RulesEngine  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input")
    p.add_argument("--output", help="Workbook copy + report sheets. Defaults to <input>_mind_report<ext> next to the input file")
    p.add_argument("--standalone", help="Standalone report .xlsx. Defaults to <input>_mind_readiness_report.xlsx next to the input file")
    p.add_argument("--no-excel", action="store_true", help="Do not use Excel COM (reduced-fidelity openpyxl write, unverified)")
    args = p.parse_args()

    input_path = Path(args.input)
    if not input_path.is_file():
        print(f"file not found: {input_path}")
        return 2

    # Keep  the source's  own extension (xlsx stays xlsx, xlsm stays xlsm) so a
    # macro-enabled workbook doesn't silently lose its VBA project on save.
    output_path = Path(args.output) if args.output else input_path.with_name(f"{input_path.stem}_mind_report{input_path.suffix}")
    standalone_path = Path(args.standalone) if args.standalone else input_path.with_name(f"{input_path.stem}_mind_readiness_report.xlsx")

    config = load_config()
    engine = RulesEngine()

    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(tmp)
        result = plan_mode.run(input_path, work_dir, config, engine=engine)
        copy_path = Path(result["workbook_analysis"]["source"]["copy_path"])
        build_standalone_report(result["validation_report"], standalone_path, rules_by_id=engine.rules, source_name=input_path.name)
        built = build_report_workbook(
            copy_path, result["validation_report"], output_path, rules_by_id=engine.rules, prefer_excel=not args.no_excel, source_name=input_path.name
        )

    print(f"Overall status: {result['validation_report']['status']}")
    print(f"Wrote standalone report: {standalone_path}")
    print(f"Wrote workbook copy + report: {built.path} (written by {built.method}; opens in Excel: {built.verified_opens_in_excel})")
    for w in built.warnings:
        print(f"WARNING: {w}")
    return 0 if result["validation_report"]["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
