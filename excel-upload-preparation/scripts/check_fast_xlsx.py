"""Prove app/fast_xlsx.py changes nothing: for a workbook, compare

  1. the formula load with the fast parser vs openpyxl as shipped (every cell's value, type, style ids,
     merged ranges, validations, defined names), and
  2. the data_only view built from the formula load vs a real data_only load (every cell's value).

    python scripts/check_fast_xlsx.py <workbook> [--ignore "Sheet A,Sheet B"]

Exit 0 when everything is identical. Runs the shipped openpyxl in a child process
(MINDREADY_FAST_XLSX=0) so the two loads cannot share any state.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import pickle
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def fingerprint(wb, data_only: bool) -> dict:
    """Per sheet, a hash of every stored cell (coordinate, type, value, style array)."""
    out = {}
    for ws in wb.worksheets:
        h = hashlib.sha256()
        cells = getattr(ws, "_cells", {})
        for key in sorted(cells):
            c = cells[key]
            v = c.value
            text = getattr(v, "text", None)
            ref = getattr(v, "ref", None)
            h.update(f"{key}|{type(v).__name__}|{v if text is None else text}|{ref}|{c.data_type}|{tuple(c._style)}\n".encode("utf-8", "replace"))
        h.update(str(sorted(str(m) for m in ws.merged_cells.ranges)).encode())
        h.update(str([(str(d.sqref), d.type, d.formula1) for d in ws.data_validations.dataValidation]).encode())
        out[ws.title] = (len(cells), h.hexdigest()[:20])
    out["__names__"] = sorted((n, d.attr_text) for n, d in wb.defined_names.items())
    return out


def values_of(wb) -> dict:
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        cells = getattr(ws, "_cells", None)
        if cells is None:
            continue
        out[name] = {k: c.value for k, c in cells.items()}
    return out


def child(path: str, ignore: list[str], out: str) -> None:
    os.environ["MINDREADY_FAST_XLSX"] = "0"
    from app import inventory

    t = time.perf_counter()
    f = inventory.load_workbook_selective(path, ignore, data_only=False)
    tf = time.perf_counter() - t
    t = time.perf_counter()
    v = inventory.load_workbook_selective(path, ignore, data_only=True)
    tv = time.perf_counter() - t
    Path(out).write_bytes(pickle.dumps({"fp": fingerprint(f, False), "values": values_of(v), "t_formulas": tf, "t_values": tv}))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--ignore", default="")
    ap.add_argument("--child", default=None)
    a = ap.parse_args()
    ignore = [x for x in a.ignore.split(",") if x]
    if a.child:
        child(a.workbook, ignore, a.child)
        return 0
    tmp = Path(tempfile.mkdtemp()) / "shipped.pkl"
    print("loading with openpyxl as shipped (child process) ...", flush=True)
    subprocess.run([sys.executable, __file__, a.workbook, "--ignore", a.ignore, "--child", str(tmp)], check=True)
    shipped = pickle.loads(tmp.read_bytes())

    from app import fast_xlsx, inventory

    print("parser:", fast_xlsx.install(), flush=True)
    t = time.perf_counter()
    f = inventory.load_workbook_selective(a.workbook, ignore, data_only=False)
    tf = time.perf_counter() - t
    view = fast_xlsx.values_view(f)
    assert view is not None, "the fast load kept no cached values"
    fp = fingerprint(f, False)
    report = {"formula_load_s": {"shipped": round(shipped["t_formulas"], 1), "fast": round(tf, 1)},
              "values_load_s": {"shipped": round(shipped["t_values"], 1), "fast": 0.0}, "differences": []}
    for k in sorted(set(fp) | set(shipped["fp"])):
        if fp.get(k) != shipped["fp"].get(k):
            report["differences"].append({"formula_load": k, "fast": fp.get(k), "shipped": shipped["fp"].get(k)})
    mine = values_of(view)
    for sheet, cells in shipped["values"].items():
        got = mine.get(sheet, {})
        bad = [(k, cells[k], got.get(k)) for k in cells if got.get(k) != cells[k] or type(got.get(k)) is not type(cells[k])]
        extra = [k for k in got if k not in cells]
        if bad or extra:
            report["differences"].append({"values": sheet, "differ": len(bad), "extra": len(extra), "sample": [str(x) for x in bad[:5]] + [str(x) for x in extra[:3]]})
    print(json.dumps(report, indent=1, default=str))
    return 1 if report["differences"] else 0


if __name__ == "__main__":
    sys.exit(main())
