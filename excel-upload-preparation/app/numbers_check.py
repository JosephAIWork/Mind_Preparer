"""1.8.0 -- did the preparation change what the model computes?

Prep inserts rows, writes titles and renames sheets. Excel moves ordinary
references along, but not everything a model reads is an ordinary reference:
a 3-D reference (=SUM('LoB 1:>>'!K65)) keeps its row when a row is inserted on
one of its sheets, and INDIRECT("'"&$E$2&"'!H"&n) is text. On one of the test
models (PVFP) the standard Prep changed 2,462 computed values that way -- 1,781
of them into errors -- and nothing said so.

This module recalculates the ORIGINAL workbook and a PREPARED version of it in
one Excel session and compares every formula cell, following each cell through
the rows / columns that were inserted and the sheets that were renamed (the
change log of every Apply says which). It reports:

  good -> error     a value the original computed is an error now
  good -> another   a value the original computed is a different value now
  error -> good     an error of the original that is a value now

A formula Prep rewrote on purpose (REF-001: #REF! -> NA()) is an error before
and after. Cells that change by themselves (NOW, TODAY, RAND) are left out.
The automatic fixer (app/autofix.py) takes the good -> error list as cells it
must not give a fallback value: that would hide what Prep broke.
"""
from __future__ import annotations

import bisect
import time
from pathlib import Path
from typing import Any, Callable, Iterable

from .autofix import NUMERIC_TOL, VOLATILE_FUNCS, _same, _show, a1, parse_r1c1
from .excel_com import close_quietly, com_available, excel_session, open_workbook
from .formula_utils import col_to_num
from .recalc import XL_ERROR_CODES

MAX_REGRESSIONS = 300_000  # good -> error cells handed to the fixer (the count is always complete)
MAX_SAMPLES = 60


class Moves:
    """Where a cell of the original sits in the prepared version: the sheet's
    new name, and the rows / columns inserted before it, Apply after Apply."""

    def __init__(self, batches: Iterable[Iterable[dict[str, Any]]] = ()):
        self.batches: list[dict[str, Any]] = []
        for ops in batches:
            rows: dict[str, list[int]] = {}
            cols: dict[str, list[int]] = {}
            renames: dict[str, str] = {}
            for op in ops:
                kind = op.get("op")
                if kind == "insert_row":
                    rows.setdefault(op["sheet"], []).append(int(op["row"]))
                elif kind == "insert_column":
                    cols.setdefault(op["sheet"], []).append(col_to_num(str(op["column"])))
                elif kind == "rename_sheet":
                    renames[op["sheet"]] = op["after"]
            for v in list(rows.values()) + list(cols.values()):
                v.sort()
            if rows or cols or renames:
                self.batches.append({"rows": rows, "cols": cols, "renames": renames})

    def forward(self, sheet: str, row: int, col: int) -> tuple[str, int, int]:
        """An original cell -> the same cell in the prepared version. Inside one
        Apply the rows are inserted bottom-up at the coordinates of the version
        it started from, so a row moves down by one for every insert at or above it."""
        for b in self.batches:
            row += bisect.bisect_right(b["rows"].get(sheet, ()), row)
            col += bisect.bisect_right(b["cols"].get(sheet, ()), col)
            sheet = b["renames"].get(sheet, sheet)
        return sheet, row, col

    def summary(self) -> dict[str, int]:
        return {
            "row_inserts": sum(len(v) for b in self.batches for v in b["rows"].values()),
            "column_inserts": sum(len(v) for b in self.batches for v in b["cols"].values()),
            "sheet_renames": sum(len(b["renames"]) for b in self.batches),
        }


def _read_formula_cells(wb: Any, tell: Callable[[str], None]) -> dict[str, dict[tuple[int, int], tuple[Any, bool]]]:
    """{sheet: {(row, col): (value, volatile)}} for every formula cell. Constants
    do not compute: what Prep writes as text (titles) is not a model value."""
    out: dict[str, dict[tuple[int, int], tuple[Any, bool]]] = {}
    volatile: dict[str, bool] = {}
    for ws in wb.Worksheets:
        name = str(ws.Name)
        tell(name)
        ur = ws.UsedRange
        r0, c0, nr, nc = int(ur.Row), int(ur.Column), int(ur.Rows.Count), int(ur.Columns.Count)
        ur = None
        cells: dict[tuple[int, int], tuple[Any, bool]] = {}
        per = max(1, min(2000, 600_000 // max(1, nc)))
        for start in range(0, nr, per):
            rows = min(per, nr - start)
            rng = ws.Range(ws.Cells(r0 + start, c0), ws.Cells(r0 + start + rows - 1, c0 + nc - 1))
            try:
                values = rng.Value2
                if not isinstance(values, tuple):
                    values = ((values,),)
                if not any(any(v is not None for v in row) for row in values):
                    continue
                formulas = rng.FormulaR1C1
            finally:
                rng = None
            if not isinstance(formulas, tuple):
                formulas = ((formulas,),)
            for i, vrow in enumerate(values):
                frow = formulas[i]
                for j, v in enumerate(vrow):
                    f = frow[j]
                    if f.__class__ is str and f[:1] == "=":
                        vol = volatile.get(f)
                        if vol is None:
                            vol = volatile[f] = bool(parse_r1c1(f).funcs & VOLATILE_FUNCS)
                        cells[(r0 + start + i, c0 + j)] = (v, vol)
        out[name] = cells
    return out


def _is_error(v: Any) -> bool:
    return v.__class__ is int and v in XL_ERROR_CODES


def compare_workbooks(original: Path, prepared: Path, moves: Moves, progress: Callable[..., None] | None = None) -> dict[str, Any]:
    """Recalculate both and compare every formula cell of the original with the
    cell it became.

    -> {ran, verdict, clean, counts, by_sheet, samples, moves, seconds,
        regressions: [(sheet, row, col, original value)] in the PREPARED version's coordinates}"""
    if not com_available():
        return {"ran": False, "clean": None, "verdict": "Excel is not available here: the numbers could not be compared.", "counts": {}, "regressions": []}
    original, prepared = Path(original).resolve(), Path(prepared).resolve()
    t0 = time.time()

    def tell(message: str) -> None:
        if progress is not None:
            try:
                progress("compare", message, None)
            except Exception:
                pass

    try:
        with excel_session() as excel:
            def one(path: Path, what: str) -> Any:
                # one after the other: the two usually carry the same file name, and Excel
                # cannot hold two workbooks of the same name open at once
                wb = None
                try:
                    tell(f"Opening {what} in Excel")
                    wb = open_workbook(excel, path, read_only=True)
                    if wb is None:
                        raise RuntimeError(f"Excel could not open {path.name}")
                    try:
                        excel.Calculation = -4135  # manual: one explicit full rebuild
                    except Exception:
                        pass
                    tell(f"Recalculating {what}")
                    excel.CalculateFullRebuild()
                    return _read_formula_cells(wb, lambda sheet: tell(f"Reading {what}: {sheet}"))
                finally:
                    close_quietly(wb)
                    wb = None

            before = one(original, "the original")
            after = one(prepared, "the current version")
    except Exception as exc:
        return {"ran": False, "clean": None, "verdict": f"The numbers could not be compared: {type(exc).__name__}: {str(exc)[:200]}", "counts": {}, "regressions": []}

    counts = {"compared": 0, "same": 0, "good_to_error": 0, "good_to_other": 0, "error_to_good": 0, "error_to_error": 0, "missing": 0, "missing_good": 0, "volatile_skipped": 0}
    samples: dict[str, list[dict[str, Any]]] = {"good_to_error": [], "good_to_other": [], "error_to_good": [], "missing": []}
    by_sheet: dict[str, dict[str, int]] = {}
    regressions: list[tuple[str, int, int, Any]] = []
    for sheet, cells in before.items():
        per_sheet: dict[str, int] = {}
        for (row, col), (v0, vol0) in cells.items():
            s1, r1, c1 = moves.forward(sheet, row, col)
            got = after.get(s1, {}).get((r1, c1))
            if got is None:
                # the formula is gone: Prep replaced it by a value on purpose (a dead #REF! formula), or cleared it
                counts["missing"] += 1
                if not _is_error(v0):
                    counts["missing_good"] += 1  # a formula that computed a value is not a formula any more
                if not _is_error(v0) and len(samples["missing"]) < MAX_SAMPLES:
                    samples["missing"].append({"original": f"{sheet}!{a1(row, col)}", "now": f"{s1}!{a1(r1, c1)}", "was": _show(v0)})
                continue
            v1, vol1 = got
            if vol0 or vol1:
                counts["volatile_skipped"] += 1
                continue
            counts["compared"] += 1
            if _is_error(v0):
                kind = "error_to_error" if _is_error(v1) else "error_to_good"
            elif _is_error(v1):
                kind = "good_to_error"
            elif _same(v0, v1, NUMERIC_TOL):
                kind = "same"
            else:
                kind = "good_to_other"
            counts[kind] += 1
            if kind in ("same", "error_to_error"):
                continue
            per_sheet[kind] = per_sheet.get(kind, 0) + 1
            if kind == "good_to_error" and len(regressions) < MAX_REGRESSIONS:
                regressions.append((s1, r1, c1, v0))
            if len(samples[kind]) < MAX_SAMPLES:
                samples[kind].append({"original": f"{sheet}!{a1(row, col)}", "now": f"{s1}!{a1(r1, c1)}", "was": _show(v0), "is": _show(v1)})
        if per_sheet:
            by_sheet[sheet] = per_sheet
    changed = counts["good_to_error"] + counts["good_to_other"] + counts["missing_good"]
    if changed:
        verdict = (
            f"The preparation changed {changed:,} computed value(s): {counts['good_to_error']:,} became errors, "
            f"{counts['good_to_other']:,} became other values"
            + (f", {counts['missing_good']:,} formula(s) that computed a value are no longer formulas" if counts["missing_good"] else "")
            + f" ({counts['compared']:,} formula cells compared with the original)."
        )
    else:
        verdict = f"The preparation changed no computed value ({counts['compared']:,} formula cells compared with the original)."
    return {
        "ran": True, "clean": changed == 0, "verdict": verdict, "counts": counts, "by_sheet": by_sheet, "samples": samples,
        "moves": moves.summary(), "seconds": round(time.time() - t0, 1), "regressions": regressions,
    }
