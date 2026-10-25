#!/usr/bin/env python3
"""Run every active rule against one xlsx/xlsm file and print a schema-valid
ValidationReport. Deterministic only: a rule without a real check reports
NOT_SUPPORTED, never a silently-skipped PASS.
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.config import load_config  # noqa: E402
from app.modes import plan_mode  # noqa: E402
from app.models import ValidationReport  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input", nargs="?")
    p.add_argument("--output")
    p.add_argument("--work-dir")
    args = p.parse_args()

    if not args.input:
        print(json.dumps({"status": "ERROR", "message": "input path is required"}, indent=2))
        return 2
    input_path = Path(args.input)
    if not input_path.is_file():
        print(json.dumps({"status": "ERROR", "message": f"file not found: {input_path}"}, indent=2))
        return 2

    config = load_config()

    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(args.work_dir) if args.work_dir else Path(tmp)
        result = plan_mode.run(input_path, work_dir, config)
        report = ValidationReport.model_validate(result["validation_report"])
        payload = json.dumps(report.model_dump(mode="json"), indent=2)

    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0 if report.status.value == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
