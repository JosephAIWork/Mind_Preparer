"""Turn every pivot table of a workbook into plain cells (values and formats), in a copy, through Excel.

    python scripts/pivots_to_values.py <workbook> <out workbook> [--only "Sheet1!PivotTable4,CFs!PivotTable6"]

Mind's converter stopped on a client model with "The method or operation is not implemented" -- the only
parts no accepted workbook had were its 8 pivot tables (one with a grouped field) and a web add-in. Mind
imports values only, so a pivot table as values is what Mind would read: each pivot's whole range (its
report filters included) is copied as values and number formats to a scratch sheet, the pivot is
removed, and the copy is pasted back where it was. Formulas that read those cells keep reading the same
values; GETPIVOTDATA would not (the script refuses a workbook that uses it). Prints what it did as JSON.
"""
from __future__ import annotations

import json
import shutil
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.excel_com import close_quietly, excel_session, open_for_write, save_in_place  # noqa: E402

XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122


def main() -> int:
    args = [a for a in sys.argv[1:] if a != "--only"]
    only = None
    if "--only" in sys.argv:
        only = {s.strip() for s in sys.argv[sys.argv.index("--only") + 1].split(",") if s.strip()}
        args.remove(sys.argv[sys.argv.index("--only") + 1])
    src, out = Path(args[0]).resolve(), Path(args[1]).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, out)
    t0 = time.time()
    done: list[dict] = []
    with excel_session(mm_addin=False) as excel:
        excel.DisplayAlerts = False
        wb = open_for_write(excel, out)
        try:
            calc = excel.Calculation
            excel.Calculation = -4135  # manual while cells move
            for n in range(1, int(wb.Worksheets.Count) + 1):
                ws = wb.Worksheets(n)
                if ws.UsedRange.Find(What="GETPIVOTDATA", LookIn=-4123, LookAt=2) is not None:  # in formulas, part of the text
                    raise SystemExit(f"{ws.Name} uses GETPIVOTDATA: turning its pivot tables into values would break it")
            scratch = wb.Worksheets.Add()
            for n in range(1, int(wb.Worksheets.Count) + 1):
                ws = wb.Worksheets(n)
                if ws.Name == scratch.Name:
                    continue
                pts = ws.PivotTables()
                for k in range(int(pts.Count), 0, -1):
                    pt = pts.Item(k)
                    if only is not None and f"{ws.Name}!{pt.Name}" not in only:
                        continue
                    rng = pt.TableRange2
                    address, name = str(rng.Address), str(pt.Name)
                    rows, cols = int(rng.Rows.Count), int(rng.Columns.Count)
                    scratch.Cells.Clear()
                    rng.Copy()
                    scratch.Range("A1").PasteSpecial(XL_PASTE_VALUES)
                    scratch.Range("A1").PasteSpecial(XL_PASTE_FORMATS)
                    excel.CutCopyMode = False
                    rng.Clear()  # clearing the whole range removes the pivot table
                    scratch.Range(scratch.Cells(1, 1), scratch.Cells(rows, cols)).Copy(ws.Range(address))
                    excel.CutCopyMode = False
                    left = int(ws.PivotTables().Count)
                    done.append({"sheet": str(ws.Name), "pivot": name, "range": address.replace("$", ""), "pivots_left_on_sheet": left})
            scratch.Delete()
            excel.Calculation = calc
            excel.CalculateFull()
            save_in_place(wb, out)
        finally:
            close_quietly(wb)
    print(json.dumps({"out": str(out), "converted": done, "seconds": round(time.time() - t0, 1)}, ensure_ascii=True, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
