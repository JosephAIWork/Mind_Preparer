"""Mind's value-comparison log -> the tables of ResultCalc.xlsx, any size.

Mind's log is one line per cell:
    <workbook> > <sheet> > <cell>: Excel value: <x> Mind value: <y>
ResultCalc.xlsx takes those lines in Sheet1!A and builds, with table formulas, a
summary per workbook / sheet / column / kind, and a merge per workbook / sheet /
kind. The formulas stop at the table's last row; this script does the same with
no limit, from a text file (one line per row) or from ResultCalc's Sheet1 column A.

    python scripts/mind_log_table.py <log.txt | ResultCalc.xlsx> [--out table.csv] [--sheet-filter Diag]

Prints the merged table (what you paste to me) and, with --out, writes every
line parsed with its kind, as CSV.
"""
from __future__ import annotations

import argparse
import csv
import re
from collections import OrderedDict, defaultdict
from pathlib import Path

LINE_RE = re.compile(r"^(?P<wb>.+?) > (?P<sheet>.+?) > (?P<cell>\$?[A-Z]{1,3}\$?\d+): Excel value: ?(?P<excel>.*?) ?Mind value: (?P<mind>.*)$", re.S)  # the Excel value may be blank
COL_RE = re.compile(r"([A-Z]{1,3})(\d+)")


def _num(text: str) -> float | None:
    t = text.strip()
    if t in ("", "None"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def kind_of(excel: str, mind: str) -> str:
    """The same words ResultCalc uses."""
    e, m = excel.strip(), mind.strip()
    if (e, m.lower()) in (("1", "true"), ("0", "false")):
        return "Same (a boolean: Excel's side shows 1/0, Mind's True/False)"
    en, mn = _num(e), _num(m)
    if m.endswith("+") and _num(m[:-1]) is not None and e == m[:-1]:
        return "Same (a text such as '80+': the comparison's Excel side reads it as the number)"
    if en is not None and mn is not None and abs(en - mn) < 1e-5:
        return "Same (rounding noise below 1e-5)"
    if m.lower() == "nan":
        return "Mind returns NaN"
    if e == "" and m != "":
        return "Excel blank, Mind has value"
    if mn is not None and en is not None:
        if mn == 0 and en != 0:
            return "Mind returns 0"
        return "Values differ"
    if mn is None and en is not None:
        return "Mind returns text"
    if en is None and mn is not None:
        return "Mind returns number"
    return "Values differ"


def read_lines(source: Path) -> list[str]:
    if source.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl

        ws = openpyxl.load_workbook(source, read_only=True, data_only=True)["Sheet1"]
        return [str(r[0]) for r in ws.iter_rows(min_col=1, max_col=1, values_only=True) if r[0] not in (None, "", "Data")]
    return [ln.rstrip("\r\n") for ln in source.read_text(encoding="utf-8", errors="replace").splitlines() if ln.strip()]


def ranges(cells: list[tuple[str, int]]) -> str:
    """J80, K80, O80:O84 ... from (column, row) pairs: consecutive rows of one column joined."""
    by_col: dict[str, list[int]] = defaultdict(list)
    for col, row in cells:
        by_col[col].append(row)
    out = []
    for col, rows in by_col.items():
        rows = sorted(set(rows))
        start = prev = rows[0]
        for r in rows[1:] + [None]:
            if r is not None and r == prev + 1:
                prev = r
                continue
            out.append(f"{col}{start}" if start == prev else f"{col}{start}:{col}{prev}")
            if r is not None:
                start = prev = r
    return ", ".join(out)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("source")
    ap.add_argument("--out")
    ap.add_argument("--sheet-filter", default=None, help="only this sheet (e.g. Diag)")
    args = ap.parse_args()
    rows = []
    skipped = 0
    for line in read_lines(Path(args.source)):
        m = LINE_RE.match(line)
        if not m:
            skipped += 1
            continue
        c = COL_RE.match(m.group("cell").replace("$", ""))
        rows.append({"wb": m.group("wb").strip(), "sheet": m.group("sheet").strip(), "cell": m.group("cell").replace("$", ""), "col": c.group(1), "row": int(c.group(2)),
                     "excel": m.group("excel").strip(), "mind": m.group("mind").strip(), "kind": kind_of(m.group("excel"), m.group("mind"))})
    if args.sheet_filter:
        rows = [r for r in rows if r["sheet"] == args.sheet_filter]
    merged: dict[tuple[str, str, str], list[dict]] = OrderedDict()
    for r in rows:
        merged.setdefault((r["wb"], r["sheet"], r["kind"]), []).append(r)
    print(f"{len(rows)} cell(s) parsed, {skipped} line(s) not understood\n")
    print("Merge Key | WB Name | Sheet | Error Type | Columns Merged | # Columns | # Cells | Rows Affected | Cell References")
    for (wb, sheet, kind), items in merged.items():
        cols = sorted({r["col"] for r in items}, key=lambda c: (len(c), c))
        print(f"{wb} | {sheet} | {kind} | {', '.join(cols)} | {len(cols)} | {len(items)} | {ranges([('', r['row']) for r in items]).replace(':', '–')} | {ranges([(r['col'], r['row']) for r in items])}")
    if args.sheet_filter:
        print("\nCell | Excel | Mind")
        for r in rows:
            print(f"{r['cell']} | {r['excel'][:60]} | {r['mind'][:60]}")
    if args.out:
        with open(args.out, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=["wb", "sheet", "cell", "col", "row", "kind", "excel", "mind"])
            w.writeheader()
            w.writerows(rows)
        print("\nwritten:", args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
