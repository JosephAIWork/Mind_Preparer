"""Stage 1: read every sheet's FormulaR1C1 grid through Excel (read-only, macros off,
calculation manual) and store it compactly for the structure analysis.

Per sheet: an int32 grid of "shape ids" (0 empty, 1 number, 2 text/bool, >=10 a formula's
R1C1 text id) and an int64 grid of constant hashes (0 where empty or formula).
Output: <out>/<tag>.pkl  {"file", "sheets": [{name, r0, c0, shape, const}], "texts": [r1c1...]}
"""
from __future__ import annotations

import pickle
import sys
import time
from pathlib import Path

import numpy as np
import pythoncom
import win32com.client as win32

CHUNK = 4000


def extract(path: str, out: Path, tag: str) -> None:
    t0 = time.time()
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.AutomationSecurity = 3          # macros off
    xl.EnableEvents = False
    texts: dict[str, int] = {}
    sheets = []
    try:
        blank = xl.Workbooks.Add()
        xl.Calculation = -4135          # manual, before the real file opens
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
        blank.Close(False)
        print(f"{tag}: opened in {time.time() - t0:.0f} s", flush=True)
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            r0, c0 = ur.Row, ur.Column
            nr, nc = ur.Rows.Count, ur.Columns.Count
            if nr * nc > 60_000_000:
                print(f"  {ws.Name}: {nr}x{nc} too big, skipped", flush=True)
                continue
            shape = np.zeros((nr, nc), dtype=np.int32)
            const = np.zeros((nr, nc), dtype=np.int64)
            for start in range(0, nr, CHUNK):
                n = min(CHUNK, nr - start)
                rng = ws.Range(ws.Cells(r0 + start, c0), ws.Cells(r0 + start + n - 1, c0 + nc - 1))
                data = rng.FormulaR1C1
                if not isinstance(data, tuple):
                    data = ((data,),)
                for i, row in enumerate(data):
                    for j, v in enumerate(row):
                        if v is None or v == "":
                            continue
                        if isinstance(v, str) and v.startswith("="):
                            sid = texts.get(v)
                            if sid is None:
                                sid = texts[v] = len(texts) + 10
                            shape[start + i, j] = sid
                        else:
                            # FormulaR1C1 hands numbers back as text ("100", "0.05")
                            num = None
                            if isinstance(v, (int, float)) and not isinstance(v, bool):
                                num = float(v)
                            elif isinstance(v, str):
                                try:
                                    num = float(v)
                                except ValueError:
                                    pass
                            shape[start + i, j] = 1 if num is not None else 2
                            const[start + i, j] = hash(num if num is not None else v) or 1
            sheets.append({"name": ws.Name, "r0": r0, "c0": c0, "shape": shape, "const": const})
            print(f"  {ws.Name}: {nr}x{nc}, {int((shape >= 10).sum())} formulas", flush=True)
        wb.Close(False)
    finally:
        xl.Quit()
    inv = [None] * (len(texts) + 10)
    for k, v in texts.items():
        inv[v] = k
    with open(out / f"{tag}.pkl", "wb") as fh:
        pickle.dump({"file": path, "sheets": sheets, "texts": inv}, fh, protocol=4)
    print(f"{tag}: done in {time.time() - t0:.0f} s, {len(texts)} distinct R1C1 formulas", flush=True)


if __name__ == "__main__":
    out = Path(__file__).parent / "data"
    out.mkdir(exist_ok=True)
    for arg in sys.argv[1:]:
        tag, path = arg.split("=", 1)
        extract(path, out, tag)
