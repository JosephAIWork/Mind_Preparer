"""Mind's loop values (from the probe's logs.json) against the macro's values (check_loop_plan.py's check.json).

    python scripts/mind_loop_compare.py <probe dir>/logs.json <check.json> --target OUTPUT!A2:J6 --first 2

check.json: expected[<source block>] = one list of values per pass (pass 1 = loop value `first`).
logs.json: differences[] = {sheet, cell, excel, mind} for every cell where Mind's value differs from the
workbook's cached value (with the MM for Excel add-in the cached values are the FIRST pass, so pass 1's
row usually does not appear in the log: it is taken from the cached value, i.e. from Excel's check).
Prints a table per pass: cell, macro value, Mind value, same / differs; exit 0 when every Mind value
seen matches the macro and every pass after the first appears in the log.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path


def col_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_target(text: str) -> tuple[str, int, int, int, int]:
    m = re.match(r"^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", text.strip(), re.I)
    if not m:
        raise SystemExit(f"target must look like OUTPUT!A2:J6, got {text}")
    return m.group(1), col_num(m.group(2)), int(m.group(3)), col_num(m.group(4)), int(m.group(5))


def num(text: str):
    t = text.strip().replace("−", "-")
    if t.lower() in ("", "null", "none", "empty", "(empty)", "#n/a", "n/a"):
        return None
    try:
        return float(t.replace(",", ""))
    except ValueError:
        return t


def same(a, b, tol: float = 1e-6) -> bool:
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        if math.isnan(a) or math.isnan(b):
            return False
        return abs(a - b) <= tol * max(1.0, abs(a), abs(b))
    if isinstance(a, (int, float)) != isinstance(b, (int, float)):
        # a date string against a number, or text against a number: compare the text forms loosely
        return str(a).strip()[:10] == str(b).strip()[:10]
    return str(a).strip() == str(b).strip()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("logs")
    ap.add_argument("check")
    ap.add_argument("--target", default="OUTPUT!A2:J6", help="the MM_RESULT block: one row per pass")
    ap.add_argument("--first", type=int, default=2, help="the first loop value (its pass is row 1 of the block)")
    ap.add_argument("--source", default=None, help="the source block key in check.json (default: its only key)")
    args = ap.parse_args()
    logs = json.loads(Path(args.logs).read_text(encoding="utf-8"))
    chk = json.loads(Path(args.check).read_text(encoding="utf-8"))
    expected = chk["expected"]
    source = args.source or next(iter(expected))
    passes = expected[source]
    sheet, c1, r1, c2, r2 = parse_target(args.target)
    mind = {}
    for d in logs.get("differences", []):
        if d.get("sheet", "").strip().lower() == sheet.lower() and d.get("cell"):
            mind[d["cell"].upper()] = num(d["mind"])
    n_rows = r2 - r1 + 1
    if n_rows != len(passes):
        print(f"note: the block has {n_rows} rows but the macro check holds {len(passes)} passes")
    ok = True
    seen_rows = 0
    for i in range(min(n_rows, len(passes))):
        row = r1 + i
        exp_row = passes[i]
        print(f"\n== pass {i + 1} (loop value {args.first + i}) -> {sheet}!row {row}")
        any_seen = False
        for j in range(c2 - c1 + 1):
            col = ""
            n = c1 + j
            while n > 0:
                n, rem = divmod(n - 1, 26)
                col = chr(65 + rem) + col
            cell = f"{col}{row}"
            exp = exp_row[j] if j < len(exp_row) else None
            if isinstance(exp, str) and re.match(r"^\d{4}-\d{2}-\d{2}", exp):
                exp = exp[:10]
            got = mind.get(cell, "(not in the log: same as the cached value)")
            if cell in mind:
                any_seen = True
                verdict = "same" if same(exp, mind[cell]) else "DIFFERS"
                if verdict == "DIFFERS":
                    ok = False
            else:
                verdict = "-"
            print(f"  {cell:<6} macro {str(exp)[:28]:<30} mind {str(got)[:40]:<42} {verdict}")
        if any_seen:
            seen_rows += 1
    print(f"\nrows with Mind values in the log: {seen_rows} of {n_rows} (pass 1 is normally absent: it equals the cached values)")
    print("RESULT:", "every Mind value in the log matches the macro" if ok else "some Mind values DIFFER from the macro")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
