"""Did the preparation change what the model computes?  (app/numbers_check.py from the command line)

Recalculates the ORIGINAL workbook and a PREPARED version of it in Excel and
compares every formula cell, through the rows / columns Prep inserted and the
sheets it renamed (read from the change logs written next to every output):

    python scripts/compare_versions.py --original model.xlsm --prepared model_v4.xlsm \
        --changelog apply/model.xlsm.changelog.json apply/v2/model.xlsm.changelog.json --out report.json

In the app the same check is POST /api/sessions/{id}/numbers-check, and the
automatic fixer runs it before it starts on a prepared version.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.numbers_check import Moves, compare_workbooks  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--original", type=Path, required=True)
    ap.add_argument("--prepared", type=Path, required=True)
    ap.add_argument("--changelog", type=Path, nargs="*", default=[], help="the .changelog.json of every Apply, oldest first")
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()
    batches = [entry.get("applied", []) for path in args.changelog for entry in json.loads(path.read_text(encoding="utf-8"))]
    res = compare_workbooks(args.original, args.prepared, Moves(batches))
    res["regressions"] = len(res.get("regressions", []))
    if args.out:
        args.out.write_text(json.dumps(res, indent=1, ensure_ascii=False, default=str), encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps({k: v for k, v in res.items() if k != "samples"}, indent=1, ensure_ascii=False, default=str))
    for kind, items in (res.get("samples") or {}).items():
        for item in items[:8]:
            print(kind, item)
    return 0 if res.get("clean") else 1


if __name__ == "__main__":
    sys.exit(main())
