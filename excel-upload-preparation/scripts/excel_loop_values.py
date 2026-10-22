"""Play a Mind loop in Excel, value by value, and read a block of cells for each: what Mind's loop instances
should show for that block, computed by Excel from the very file uploaded to Mind.

    python scripts/excel_loop_values.py <workbook> --driver "Projection!C3" --first 1 --last 19 \
        --block "Macro!A11:J138" --out values.json [--errors]

A copy of the workbook is opened in a hidden Excel of its own (macros off, no add-in: the driver cell's
MM_LOOP formula is replaced by each loop value, which is what MM_LOOP gives for that instance). For each
value: write it into the driver cell, recalculate (one full recalculation first, then
Application.Calculate per value, as the workbook's own macro does), read the block. With --errors, every formula cell holding an error in the whole workbook is
counted per pass (Range.SpecialCells), with the first addresses. Nothing is saved.

values.json: {"driver", "block", "values": [k...], "passes": {k: [[row values]...]}, "errors": {k: {...}}}
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.excel_com import close_quietly, excel_pid, excel_session, kill_excel, open_workbook  # noqa: E402

XL_CALC_MANUAL = -4135
XL_FORMULAS, XL_ERRORS = -4123, 16
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826245: "#GETTING_DATA",
          -2146826239: "#SPILL!", -2146826238: "#CONNECT!", -2146826237: "#BLOCKED!", -2146826236: "#UNKNOWN!",
          -2146826235: "#FIELD!", -2146826234: "#CALC!"}


def split_ref(ref: str) -> tuple[str, str]:
    m = re.match(r"^'?(.+?)'?!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$", ref.strip(), re.I)
    if not m:
        raise SystemExit(f"not a sheet!range reference: {ref}")
    return m.group(1).replace("''", "'"), m.group(2)


def plain(v):
    if isinstance(v, int) and v in ERRORS:
        return ERRORS[v]
    return v


def error_cells(wb) -> dict:
    total, first = 0, []
    for ws in wb.Worksheets:
        try:
            rng = ws.UsedRange.SpecialCells(XL_FORMULAS, XL_ERRORS)
        except Exception:  # no error cell on this sheet
            continue
        n = int(rng.Count)
        total += n
        if len(first) < 40:
            first.append(f"{ws.Name}!{str(rng.Address).replace('$', '')[:200]} ({n})")
    return {"count": total, "first": first}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--driver", required=True)
    ap.add_argument("--first", type=int, required=True)
    ap.add_argument("--last", type=int, required=True)
    ap.add_argument("--block", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--errors", action="store_true")
    a = ap.parse_args()
    d_sheet, d_cell = split_ref(a.driver)
    b_sheet, b_range = split_ref(a.block)
    work = Path(tempfile.mkdtemp(prefix="mind_ready_loop_"))
    copy = work / Path(a.workbook).name
    shutil.copy2(a.workbook, copy)
    out: dict = {"workbook": str(a.workbook), "driver": a.driver, "block": a.block, "values": list(range(a.first, a.last + 1)), "passes": {}, "errors": {}, "seconds": {}}

    def run(excel) -> None:
        wb = None
        try:
            wb = open_workbook(excel, copy, read_only=False)
            excel.Calculation = XL_CALC_MANUAL
            driver = wb.Worksheets(d_sheet).Range(d_cell)
            out["driver_formula"] = driver.Formula
            driver.Value2 = a.first
            excel.CalculateFull()  # every cell from scratch once, as Mind computes them
            block = wb.Worksheets(b_sheet).Range(b_range)
            for k in out["values"]:
                t0 = time.time()
                driver.Value2 = k
                excel.Calculate()
                out["passes"][str(k)] = [[plain(v) for v in row] for row in block.Value2]
                if a.errors:
                    out["errors"][str(k)] = error_cells(wb)
                out["seconds"][str(k)] = round(time.time() - t0, 1)
                print(f"pass {k}: {out['seconds'][str(k)]} s" + (f", {out['errors'][str(k)]['count']} error cells" if a.errors else ""), flush=True)
        finally:
            close_quietly(wb)

    with excel_session(mm_addin=False) as excel:
        pid = excel_pid(excel)
        try:
            run(excel)
        except Exception:
            kill_excel(pid)
            raise
    shutil.rmtree(work, ignore_errors=True)
    Path(a.out).write_text(json.dumps(out, indent=1, default=str), encoding="utf-8")
    print("saved", a.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
