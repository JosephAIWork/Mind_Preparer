"""Append one policy to a loop version's input table, through Excel, to show the loop growing in Mind.

    python scripts/add_policy_row.py <loop version .xlsm> <out .xlsm> --sheet "<input table sheet>" [--copy-row 6] [--id-col A]

Copies the last data row of the table one row down (values only: the macro's table holds constants),
gives the new row a new id in --id-col (the copied id with "_7" appended, or +1 for a number), saves the
workbook as <out>. Everything else is untouched, so the new instance's results must equal the copied
policy's results. Excel with the MM for Excel add-in keeps the cached values sound.
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.excel_com import close_quietly, excel_session, open_for_write  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--sheet", required=True, help="the sheet of the loop's input table")
    ap.add_argument("--copy-row", type=int, default=6)
    ap.add_argument("--id-col", default="A")
    a = ap.parse_args()
    src, out = Path(a.src).resolve(), Path(a.out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, out)
    with excel_session() as excel:
        book = open_for_write(excel, out)
        try:
            ws = book.Worksheets(a.sheet)
            last_col = ws.UsedRange.Columns.Count
            r, n = a.copy_row, a.copy_row + 1
            src_rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, last_col))
            dst_rng = ws.Range(ws.Cells(n, 1), ws.Cells(n, last_col))
            dst_rng.Value = src_rng.Value  # values only
            old_id = ws.Range(f"{a.id_col}{r}").Value
            new_id = old_id + 1 if isinstance(old_id, (int, float)) else f"{old_id}_{n}"
            ws.Range(f"{a.id_col}{n}").Value = new_id
            excel.CalculateFull()
            book.Save()
            print(f"row {n} of {a.sheet} = row {r} with {a.id_col}{n} = {new_id!r} ({last_col} columns)")
        finally:
            close_quietly(book)
    print("written", out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
