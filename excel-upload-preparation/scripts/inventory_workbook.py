#!/usr/bin/env python3
"""Build  a schema-valid  WorkbookAnalysis for one xlsx/xlsm file.

Always operates on a fresh copy of the source (never modifies it). Prints
the analysis JSON to stdout, or writes it to --output.
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.inventory import build_analysis  # noqa: E402
from app.models import WorkbookAnalysis  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input", nargs="?")
    p.add_argument("--output")
    p.add_argument("--work-dir", help="Where to place the immutable copy (default: a temp dir)")
    args = p.parse_args()

    if not args.input:
        print(json.dumps({"status": "ERROR", "message": "input path is required"}, indent=2))
        return 2

    input_path = Path(args.input)
    if not input_path.is_file():
        print(json.dumps({"status": "ERROR", "message": f"file not found: {input_path}"}, indent=2))
        return 2

    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(args.work_dir) if args.work_dir else Path(tmp)
        analysis = build_analysis(input_path, work_dir, analysis_id=str(uuid.uuid4()))
        validated = WorkbookAnalysis.model_validate(analysis)
        payload = json.dumps(validated.model_dump(mode="json"), indent=2)

    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
