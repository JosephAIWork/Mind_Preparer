"""Time what Excel itself takes on a workbook, step by step, the way the app drives it.

    python scripts/perf_excel_bench.py <workbook> [--out perf.json] [--no-addin] [--threads off|on]
                                       [--steps calc,full,rebuild,read,save,reopen]

A fresh hidden Excel (app.excel_com.excel_session) opens a copy of the workbook and every step is timed:
open, Calculate (F9), CalculateFull (Ctrl+Alt+F9), CalculateFullRebuild (Ctrl+Shift+Alt+F9), reading
every used range's values and formulas (what app/recalc.py does), Save, and opening the saved file again
in a second Excel (what verify_opens_in_excel does). Also reports Excel's calculation settings, so the
app's numbers can be set against plain Excel's. The source file is never changed.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.excel_com import close_quietly, excel_session, mm_addin_loaded, open_workbook  # noqa: E402
from app.settings import SETTINGS  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--out", default=None)
    ap.add_argument("--no-addin", action="store_true")
    ap.add_argument("--threads", choices=("as-is", "on", "off"), default="as-is")
    ap.add_argument("--steps", default="calc,full,rebuild,read,save,reopen")
    ap.add_argument("--work", default=str(SETTINGS.data_folder / "perf" / "work"))
    a = ap.parse_args()
    steps = [x.strip() for x in a.steps.split(",") if x.strip()]
    src = Path(a.workbook).resolve()
    work = Path(a.work)
    work.mkdir(parents=True, exist_ok=True)
    copy = work / f"bench_{src.name}"
    t = time.perf_counter()
    shutil.copy2(src, copy)
    out: dict = {"workbook": str(src), "mb": round(src.stat().st_size / 2**20, 1), "addin": not a.no_addin, "threads": a.threads, "timings": {}}
    out["timings"]["copy"] = round(time.perf_counter() - t, 2)

    def timed(name, fn):
        t0 = time.perf_counter()
        r = fn()
        out["timings"][name] = round(time.perf_counter() - t0, 2)
        print(f"{name:>16}: {out['timings'][name]:8.2f} s", flush=True)
        return r

    t0 = time.perf_counter()
    with excel_session(mm_addin=not a.no_addin) as excel:
        out["timings"]["start_excel"] = round(time.perf_counter() - t0, 2)

        def run():
            out["mm_addin_loaded"] = mm_addin_loaded(excel)
            if a.threads != "as-is":
                excel.MultiThreadedCalculation.Enabled = a.threads == "on"
            mt = excel.MultiThreadedCalculation
            out["excel"] = {"version": excel.Version, "mt_enabled": bool(mt.Enabled), "mt_threads": int(mt.ThreadCount), "mt_mode": int(mt.ThreadMode)}
            wb = timed("open", lambda: open_workbook(excel, copy, read_only=False))
            try:
                out["excel"].update({
                    "calculation": int(excel.Calculation), "calculate_before_save": bool(excel.CalculateBeforeSave),
                    "force_full_calculation": bool(wb.ForceFullCalculation), "sheets": int(wb.Worksheets.Count),
                    "calculation_state": int(excel.CalculationState),
                })
                if "calc" in steps:
                    timed("calculate", excel.Calculate)
                if "full" in steps:
                    timed("calculate_full", excel.CalculateFull)
                if "rebuild" in steps:
                    timed("calc_full_rebuild", excel.CalculateFullRebuild)
                if "calc" in steps:
                    timed("calculate_again", excel.Calculate)
                if "read" in steps:
                    def read():
                        cells = 0
                        per = {}
                        for sh in wb.Worksheets:
                            t1 = time.perf_counter()
                            ur = sh.UsedRange
                            v = ur.Value
                            f = ur.Formula
                            n = int(ur.Rows.Count) * int(ur.Columns.Count)
                            cells += n
                            per[sh.Name] = {"cells": n, "s": round(time.perf_counter() - t1, 2), "address": str(ur.Address)}
                            v = f = ur = None
                            sh = None
                        out["used_cells"] = cells
                        out["read_per_sheet"] = dict(sorted(per.items(), key=lambda kv: -kv[1]["s"])[:12])
                    timed("read_used_ranges", read)
                if "save" in steps:
                    timed("save", wb.Save)
            finally:
                close_quietly(wb)
                wb = None

        run()
    out["timings"]["session_total"] = round(time.perf_counter() - t0, 2)
    if "reopen" in steps:
        t1 = time.perf_counter()
        with excel_session(mm_addin=not a.no_addin) as excel2:
            def reopen():
                wb2 = open_workbook(excel2, copy, read_only=True)
                close_quietly(wb2)
            reopen()
        out["timings"]["reopen_fresh_excel"] = round(time.perf_counter() - t1, 2)
        print(f"{'reopen_fresh':>16}: {out['timings']['reopen_fresh_excel']:8.2f} s", flush=True)
    text = json.dumps(out, indent=1, default=str)
    if a.out:
        Path(a.out).parent.mkdir(parents=True, exist_ok=True)
        Path(a.out).write_text(text, encoding="utf-8")
    print(text)
    copy.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
