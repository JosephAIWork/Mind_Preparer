#!/usr/bin/env python3
"""Pure  OOXML package-part  diff between two xlsx/xlsm files (e.g. a source
and a candidate output). Read-only, no mutation -- see app/preservation.py.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.preservation import compare_packages  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("before")
    p.add_argument("after")
    p.add_argument("--output")
    args = p.parse_args()

    before, after = Path(args.before), Path(args.after)
    for path in (before, after):
        if not path.is_file():
            print(json.dumps({"status": "ERROR", "message": f"file not found: {path}"}, indent=2))
            return 2

    result = compare_packages(before, after)
    payload = json.dumps(result, indent=2)
    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0 if result["identical"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
