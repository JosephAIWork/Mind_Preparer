"""Mind's loop instances of an output grid (CSV from the grid panel: Export grid > Export all dimension) against
Excel's values for each loop value (scripts/excel_loop_values.py), cell by cell.

    python scripts/mind_grid_compare.py <values.json> <grid.csv> --anchor C13 [--tol 1e-9] [--out compare.json]

The CSV holds every instance one after the other (instance 1 first), each with the grid's header rows
(written once or per instance); --anchor is the workbook cell of the first DATA cell of the CSV (the
grid's first value row, first column). The block in values.json must contain the anchor. Numbers are
compared with a relative tolerance; text and errors as text.
"""
from __future__ import annotations

import argparse
import csv
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


def col_name(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", ""))
        except ValueError:
            return None
    return None


def same(a, b, tol: float) -> bool:
    if a in (None, "") and b in (None, "", 0, 0.0, "0"):
        return True  # an empty cell shows as 0 in Mind's grid
    x, y = num(a), num(b)
    if x is not None and y is not None:
        return math.isclose(x, y, rel_tol=tol, abs_tol=tol)
    return str(a).strip() == str(b).strip()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("values")
    ap.add_argument("csv")
    ap.add_argument("--anchor", required=True)
    ap.add_argument("--tol", type=float, default=1e-9)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    vals = json.loads(Path(a.values).read_text(encoding="utf-8"))
    m = re.match(r"^'?(.+?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", vals["block"])
    b_c1, b_r1 = col_num(m.group(2)), int(m.group(3))
    am = re.match(r"^([A-Z]+)(\d+)$", a.anchor.upper())
    a_c, a_r = col_num(am.group(1)), int(am.group(2))
    rows = list(csv.reader(Path(a.csv).open(encoding="utf-8-sig", newline="")))
    width = max(len(r) for r in rows)
    # data rows: every row whose cells all read as numbers (or are empty); header rows are text
    data = [r for r in rows if r and all(c.strip() == "" or num(c) is not None for c in r)]
    passes = [str(k) for k in vals["values"]]
    per = len(data) // len(passes) if passes else 0
    rep: dict = {"values": a.values, "csv": a.csv, "anchor": a.anchor, "csv_rows": len(rows), "data_rows": len(data),
                 "instances": len(passes), "rows_per_instance": per, "width": width, "compared": 0, "differ": 0, "per_instance": {}, "samples": []}
    if per * len(passes) != len(data):
        rep["warning"] = f"{len(data)} data rows do not split evenly into {len(passes)} instances"
    for i, k in enumerate(passes):
        block = vals["passes"][k]
        n_cmp = n_diff = 0
        for j in range(per):
            mind_row = data[i * per + j]
            r = a_r + j
            for c_off, mv in enumerate(mind_row):
                c = a_c + c_off
                bi, bj = r - b_r1, c - b_c1
                if not (0 <= bi < len(block) and 0 <= bj < len(block[bi])):
                    continue
                ev = block[bi][bj]
                n_cmp += 1
                if not same(ev, mv, a.tol):
                    n_diff += 1
                    if len(rep["samples"]) < 60:
                        rep["samples"].append({"instance": i + 1, "loop_value": k, "cell": f"{col_name(c)}{r}", "excel": ev, "mind": mv})
        rep["per_instance"][k] = {"compared": n_cmp, "differ": n_diff}
        rep["compared"] += n_cmp
        rep["differ"] += n_diff
    text = json.dumps(rep, indent=1, default=str)
    if a.out:
        Path(a.out).write_text(text, encoding="utf-8")
    print(f"{rep['instances']} instances x {per} rows: {rep['compared']} cells compared, {rep['differ']} differ" + (f" ({rep['warning']})" if rep.get("warning") else ""))
    for s in rep["samples"][:15]:
        print("  ", s)
    return 0 if rep["differ"] == 0 and not rep.get("warning") else 1


if __name__ == "__main__":
    raise SystemExit(main())
