"""Compare two workbooks cell by cell: every formula (shared formulas expanded) and every value Excel stored.

    python scripts/compare_final_workbooks.py <before.xlsm> <after.xlsm> [--out diff.json] [--ignore "Sheet A,Sheet B"]

Made for before/after runs of a speed change, where the final workbooks must be identical. One workbook
is in memory at a time: the first is reduced to a hash per cell and freed before the second is loaded.
Prints counts and samples of formula differences, value differences, and cells present in one only.
"""
from __future__ import annotations

import argparse
import gc
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.inventory import load_workbook_selective  # noqa: E402  (installs the fast reader: formulas + cached values)

NUM_TOL = 1e-9


def _norm(v):
    if isinstance(v, float):
        return round(v, 9) if abs(v) < 1e15 else v
    if hasattr(v, "text"):  # ArrayFormula / DataTableFormula
        return f"{type(v).__name__}:{getattr(v, 'ref', '')}:{v.text}"
    return v


def cells(path: Path, ignore: list[str]):
    wb = load_workbook_selective(path, ignore)
    try:
        for ws in wb.worksheets:
            if getattr(ws, "mind_ready_ignored", False) or not hasattr(ws, "_cells"):
                continue
            cached = getattr(ws, "mind_ready_cached", {})
            for (r, c), cell in ws._cells.items():
                content = _norm(cell._value)
                stored = _norm(cached.get((r, c))) if (r, c) in cached else None
                if content is None and stored is None:
                    continue
                yield ws.title, r, c, content, stored
    finally:
        wb.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("before")
    ap.add_argument("after")
    ap.add_argument("--out", default=None)
    ap.add_argument("--ignore", default="")
    a = ap.parse_args()
    ignore = [s.strip() for s in a.ignore.split(",") if s.strip()]
    t0 = time.time()
    first: dict[tuple[str, int, int], tuple[int, int]] = {}
    for sheet, r, c, content, stored in cells(Path(a.before), ignore):
        first[(sheet, r, c)] = (hash(repr(content)), hash(repr(stored)))
    gc.collect()
    t1 = time.time()
    formula_diff, value_diff, only_after, samples = 0, 0, 0, {"formula": [], "value": [], "only_after": [], "only_before": []}
    seen = set()
    for sheet, r, c, content, stored in cells(Path(a.after), ignore):
        key = (sheet, r, c)
        h = first.get(key)
        if h is None:
            only_after += 1
            if len(samples["only_after"]) < 15:
                samples["only_after"].append([sheet, r, c, repr(content)[:120]])
            continue
        seen.add(key)
        if h[0] != hash(repr(content)):
            formula_diff += 1
            if len(samples["formula"]) < 15:
                samples["formula"].append([sheet, r, c, repr(content)[:160]])
        elif h[1] != hash(repr(stored)):
            value_diff += 1
            if len(samples["value"]) < 15:
                samples["value"].append([sheet, r, c, repr(content)[:100], repr(stored)[:60]])
    only_before = [k for k in first if k not in seen]
    samples["only_before"] = [list(k) for k in only_before[:15]]
    out = {
        "before": a.before, "after": a.after, "cells_before": len(first), "cells_compared": len(seen),
        "formula_or_content_differences": formula_diff, "stored_value_differences": value_diff,
        "only_in_after": only_after, "only_in_before": len(only_before),
        "identical": formula_diff == value_diff == only_after == len(only_before) == 0,
        "seconds": {"before": round(t1 - t0, 1), "after": round(time.time() - t1, 1)}, "samples": samples,
    }
    text = json.dumps(out, indent=1, ensure_ascii=True, default=str)
    if a.out:
        Path(a.out).write_text(text, encoding="utf-8")
    print(text)
    return 0 if out["identical"] else 1


if __name__ == "__main__":
    sys.exit(main())
