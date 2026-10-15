"""Does the Mind loop plan give what the macro gives? Checked in Excel, pass by pass.

    python scripts/check_loop_plan.py <workbook.xlsm> [--macro Name] [--first a] [--last b] [--name N] [--work dir]

1. The plan is built (app.loop_builder) on a copy of the workbook and written into a second copy
   through the app's own Apply (Excel COM), as the "Write it into a new version" button does.
2. With --macro, the macro is run in Excel on the first copy (macros allowed for this one
   workbook, nothing else) and the block it writes per pass is read back: the expected values.
3. Each pass is then played in Excel on the loop version: the driver cells get the k-th value of
   their MM_LOOP range (what MM_LOOP returns for loop value k), Excel recalculates, and the source
   cells of every MM_RESULT block are read. They must equal the macro's k-th block.

Excel here has no Mind add-in, so MM_LOOP itself cannot be evaluated; this checks that the plan
points every driver cell at the right column, every result cell at the right source for the
right loop value, and that the model computes the same numbers from them as from the macro's
copies. The run in Mind (scripts/mind_upload.py --run) is the other half.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import sys
import time
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.excel_com import close_quietly, excel_session, open_for_write, save_in_place  # noqa: E402
from app.formula_utils import num_to_col, parse_ref  # noqa: E402
from app.inventory import build_analysis  # noqa: E402
from app.loop_builder import build_loop_plan, describe  # noqa: E402
from app.model_map import build_model_map  # noqa: E402
from app.prep import apply_operations  # noqa: E402
from app.recalc import XL_ERROR_CODES  # noqa: E402

_PH_RE = re.compile(r"<([A-Za-z_]\w*)>")


def _split(cell: str) -> tuple[str, str]:
    sheet, ref = cell.rsplit("!", 1)
    return sheet.strip().strip("'").replace("''", "'"), ref


def _block(ws: Any, ref: str) -> list[Any]:
    v = ws.Range(ref).Value
    if not isinstance(v, tuple):
        return [v]
    return [x for row in v for x in row]


def _same(a: Any, b: Any) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        if a in XL_ERROR_CODES or b in XL_ERROR_CODES:
            return a == b
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9)
    if a in (None, "") and b in (None, ""):
        return True
    return a == b


def _show(v: Any) -> str:
    if isinstance(v, int) and v in XL_ERROR_CODES:
        return XL_ERROR_CODES[v]
    if v is None:
        return "empty"
    return str(v)[:24]


def run_macro(path: Path, macro: str, targets: list[tuple[str, str]]) -> dict[str, list[Any]]:
    """Run `macro` in Excel on `path` (macros allowed for this workbook only) and read back the blocks."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    out: dict[str, list[Any]] = {}
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow: this copy's macro may run
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0)
        try:
            t0 = time.time()
            excel.Run(f"'{path.name}'!{macro}")
            print(f"macro {macro} ran in {time.time() - t0:.1f}s")
            for sheet, ref in targets:
                out[f"{sheet}!{ref}"] = _block(wb.Worksheets(sheet), ref)
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
            wb = None
    finally:
        if excel is not None:
            excel.Quit()
        excel = None
        pythoncom.CoUninitialize()
    return out


def play_passes(path: Path, plan: dict[str, Any]) -> dict[int, dict[str, list[Any]]]:
    """For every loop value k: give each driver cell the k-th value of its MM_LOOP range, recalculate,
    read the source blocks of the MM_RESULT outputs. {k: {"Sheet!From": [values]}}"""
    loops = [l for l in plan["loops"] if l["status"] == "ready"]
    outs = [o for o in plan["outputs"] if o["status"] == "ready"]
    if len(loops) != 1:
        raise SystemExit(f"this check plays one loop at a time ({len(loops)} ready)")
    lp = loops[0]
    got: dict[int, dict[str, list[Any]]] = {}

    def _play(excel: Any) -> None:
        wb = open_for_write(excel, path)
        try:
            excel.Calculation = -4135  # manual: one recalculation per pass
            for k in range(1, lp["count"] + 1):
                for c in lp["cells"]:
                    ws = wb.Worksheets(c["sheet"])
                    if lp["kind"] in ("row_of_table", "column_of_table", "list"):
                        r = parse_ref(c.get("range_fixed") or c["range"])
                        src = wb.Worksheets(r["sheet"])
                        cells = [(col, row) for row in range(r["r1"], r["r2"] + 1) for col in range(r["c1"], r["c2"] + 1)]
                        col, row = cells[k - 1]
                        value = src.Cells(row, col).Value
                        if value is None:
                            ws.Range(c["ref"]).ClearContents()
                        else:
                            ws.Range(c["ref"]).Value = value
                    elif lp["kind"] == "counter":
                        ws.Range(c["ref"]).Value = lp["first"] + k - 1
                    else:
                        raise SystemExit(f"cannot play a {lp['kind']} loop here")
                excel.CalculateFull()
                got[k] = {o["from"]: _block(wb.Worksheets(_split(o["from"])[0]), _split(o["from"])[1]) for o in outs}
            excel.Calculation = -4105
        finally:
            close_quietly(wb)

    with excel_session() as excel:
        _play(excel)
    return got


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("workbook")
    ap.add_argument("--macro", default=None, help="the macro to run for the expected values (else the block's stored values are used)")
    ap.add_argument("--name", default=None)
    ap.add_argument("--first", default=None)
    ap.add_argument("--last", default=None)
    ap.add_argument("--work", default=None)
    a = ap.parse_args()
    src = Path(a.workbook).resolve()
    work = Path(a.work) if a.work else src.parent / f"{src.stem}_loop_check"
    work.mkdir(parents=True, exist_ok=True)
    macro_copy = work / f"{src.stem}_macro{src.suffix}"
    shutil.copy2(src, macro_copy)

    print("== plan")
    analysis = build_analysis(macro_copy, work / "analysis", src.stem)
    plan = build_loop_plan(analysis, build_model_map(analysis), name=a.name, first=a.first, last=a.last)
    print(describe(plan))
    (work / "plan.json").write_text(json.dumps(plan, indent=1, default=str, ensure_ascii=False), encoding="utf-8")
    outs = [o for o in plan["outputs"] if o["status"] == "ready"]
    lp = next((l for l in plan["loops"] if l["status"] == "ready"), None)
    if not outs or lp is None:
        print("nothing to check: no ready loop with a per-pass block")
        return 2

    print("== write the loop version (the app's Apply, through Excel)")
    res = apply_operations(macro_copy, work / "apply", plan["operations"])
    print(res["message"], "| verified opens:", res.get("verified_opens_in_excel"))
    if res["status"] == "ERROR":
        for f in res["failed"][:5]:
            print("  failed:", f.get("sheet"), f.get("cell"), f.get("error"))
        return 1
    loop_version = Path(res["output_path"])
    print("  ->", loop_version)

    # the blocks the macro writes, pass by pass (target with the loop value substituted)
    def _targets(o: dict[str, Any]) -> list[tuple[str, str]]:
        t = []
        for k in range(1, lp["count"] + 1):
            cell = _PH_RE.sub(str(lp["first"] + k - 1), o["target"])
            t.append(_split(cell))
        return t

    expected: dict[str, list[list[Any]]] = {}
    if a.macro:
        print(f"== run the macro {a.macro} in Excel (expected values)")
        all_targets = [t for o in outs for t in _targets(o)]
        blocks = run_macro(macro_copy, a.macro, all_targets)
        for o in outs:
            expected[o["from"]] = [blocks[f"{s}!{r}"] for s, r in _targets(o)]
    else:
        print("== expected values: what the block holds now (no --macro given)")

        def _read(excel: Any) -> None:
            from app.excel_com import open_workbook

            wb = open_workbook(excel, macro_copy, read_only=True)
            try:
                for o in outs:
                    expected[o["from"]] = [_block(wb.Worksheets(s), r) for s, r in _targets(o)]
            finally:
                close_quietly(wb)

        with excel_session() as excel:
            _read(excel)

    print(f"== play the {lp['count']} passes on the loop version")
    got = play_passes(loop_version, plan)
    bad = 0
    total = 0
    for o in outs:
        for k in range(1, lp["count"] + 1):
            exp, val = expected[o["from"]][k - 1], got[k][o["from"]]
            diffs = [(i, e, v) for i, (e, v) in enumerate(zip(exp, val)) if not _same(e, v)]
            total += len(exp)
            bad += len(diffs)
            status = "same" if not diffs else f"{len(diffs)} differ"
            print(f"  value {k} (= {lp['first'] + k - 1}): {_PH_RE.sub(str(lp['first'] + k - 1), o['target'])} <- {o['from']}: {status}")
            for i, e, v in diffs[:4]:
                r = parse_ref(_split(o["from"])[1])
                width = r["c2"] - r["c1"] + 1
                cell = f"{num_to_col(r['c1'] + i % width)}{r['r1'] + i // width}"
                print(f"      {cell}: macro {_show(e)} | loop {_show(v)}")
    print(f"== {total - bad} of {total} cells the same as the macro" + ("" if not bad else f"; {bad} differ"))
    (work / "check.json").write_text(json.dumps({"expected": expected, "got": {str(k): v for k, v in got.items()}, "same": total - bad, "differ": bad}, default=str), encoding="utf-8")
    return 0 if not bad else 1


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    raise SystemExit(main())
