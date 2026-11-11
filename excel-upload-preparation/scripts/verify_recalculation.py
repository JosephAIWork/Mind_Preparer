#!/usr/bin/env python3
"""Verifies  a workbook  recalculates cleanly through the trusted Excel COM
adapter (app/recalc.py) -- same mechanism as recalculate_workbook.py, framed
as a yes/no check rather than "perform and report". Always works on a fresh
copy, never the original.
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
        copy_path, _ = make_immutable_copy(input_path, Path(tmp))
        result = recalculate(copy_path)
        payload = json.dumps(
            {
                "recalculates_cleanly": result["status"] == "PASS",
                "status": result["status"],
                "message": result["message"],
                "formula_error_count": len(result["formula_errors"]),
                "formula_errors": result["formula_errors"],
                "mmforexcel_loaded": result.get("mmforexcel_loaded", False),
                "addin_gap_errors": result.get("addin_gap_errors", []),
            },
            indent=2,
        )

    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0 if result["status"] == "PASS" else (2 if result["status"] == "NOT_SUPPORTED" else 1)


if __name__ == "__main__":
    raise SystemExit(main())
