#!/usr/bin/env python3
"""Recalculation  adapter (CLAUDE_CODE_PROMPT.md  phase 9): drives the
locally-installed Excel via win32com (app/recalc.py) to recalculate a fresh
copy of the workbook -- never the original -- and reports any formula
errors found. NOT_SUPPORTED if Excel/pywin32 isn't available here;  never
fakes a PASS.
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.inventory import make_immutable_copy  # noqa: E402
from app.recalc import recalculate  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input", nargs="?")
    p.add_argument("--output")
    args = p.parse_args()

    if not args.input:
        print(json.dumps({"status": "ERROR", "message": "input path is required"}, indent=2))
        return 2
    input_path = Path(args.input)
    if not input_path.is_file():
        print(json.dumps({"status": "ERROR", "message": f"file not found: {input_path}"}, indent=2))
        return 2

    with tempfile.TemporaryDirectory() as tmp:
        copy_path, source_sha256 = make_immutable_copy(input_path, Path(tmp))
        result = recalculate(copy_path)
        result["source_sha256"] = source_sha256
        payload = json.dumps(result, indent=2)

    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0 if result["status"] == "PASS" else (2 if result["status"] == "NOT_SUPPORTED" else 1)


if __name__ == "__main__":
    raise SystemExit(main())
