"""Did  the preparation  change what the model computes?

Prep inserts rows, writes titles and renames sheets. Excel moves ordinary
references along, but not everything a model reads is an ordinary reference:
a 3-D reference (=SUM('Line 1:>>'!K65)) keeps  its row when a row is inserted on
one of its sheets,  and INDIRECT("'"&$E$2&"'!H"&n) is text. On one  real model
the standard Prep changed 2,462 computed  values that way -- 1,781 of them into
errors -- and nothing  said so.

This module recalculates the ORIGINAL workbook  and a PREPARED version of it in
one Excel session and compares every formula cell, following each cell through
the rows / columns that were inserted and the sheets that were renamed (the
change log of every Apply says which). It reports:

  good -> error     a value the original computed is an error now
  good -> another   a value the original computed is a different value now
  error ->  good     an error of the original that is a value now

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

from .autofix import NUMERIC_TOL, VOLATILE_FUNCS, _same, _show, a1, parse_r1c1, ref_rect
from .excel_com import close_quietly, com_available, excel_session, open_workbook
from .formula_utils import col_to_num
from .plain_codes import plain_error
from .recalc import XL_ERROR_CODES

MAX_REGRESSIONS = 300_000  # good  -> error  cells handed to the fixer (the count is always complete)
MAX_SAMPLES = 60


class Moves:
    """Where  a cell  of the original sits in the prepared version: the sheet's
    new name, and the rows / columns inserted before it, Apply after Apply."""

    def __init__(self, batches: Iterable[Iterable[dict[str, Any]]] = ()):
        self.batches: list[dict[str, Any]] = []
        # formulas a title replaced ON PURPOSE (a formula caption above  a table, 1.14.1): (batches before, sheet, row, col)
        self._replaced_raw: list[tuple[int, str, int, int]] = []
        self._replaced: set[tuple[str, int, int]] | None = None
        # the neutral path /  file name texts Prep wrote in  place of CELL("filename") (FRM-012): the model
        # owner's own  location is not kept, so a text built on them differs  from the original only by that
        self.neutral_texts: list[str] = []
        self._frozen_raw: list[tuple[int, str, int, int]] = []
        self._frozen: set[tuple[str, int, int]] | None = None
        self._written_raw: list[tuple[int, str, int, int, int, int]] = []  # (batches before,  sheet, r1, r2, c1, c2)
        self._written: list[tuple[str, int, int, int, int]] | None = None
        for ops in batches:
            rows: dict[str, list[int]] = {}
            banded: dict[str, list[tuple[int, int, int]]] = {}  # row inserts that shift only some columns
            cols: dict[str, list[int]] = {}
            renames: dict[str, str] = {}
            for op in ops:
                kind = op.get("op")
                if kind == "insert_row" and op.get("columns"):
                    a, _, b = str(op["columns"]).partition(":")
                    banded.setdefault(op["sheet"], []).append((int(op["row"]), col_to_num(a), col_to_num(b or a)))
                elif kind == "insert_row":
                    rows.setdefault(op["sheet"], []).append(int(op["row"]))
                elif kind == "insert_column":
                    cols.setdefault(op["sheet"], []).append(col_to_num(str(op["column"])))
                elif kind == "rename_sheet":
                    renames[op["sheet"]] = op["after"]
                elif kind in ("set_formula", "set_array_formula") and op.get("rule_id") == "FRM-012" and "neutral" in str(op.get("note") or ""):
                    after = str(op.get("after") or "")
                    if after.startswith('="') and after.endswith('"') and len(after) > 4:
                        self.neutral_texts.append(after[2:-1].replace('""', '"'))
                elif kind == "set_value" and str(op.get("before") or "").startswith("=") and str(op.get("after") or "").startswith("#"):
                    from .formula_utils import parse_ref
                    at = parse_ref(str(op.get("cell") or "").replace("$", ""))
                    if at and not at.get("whole_column") and not at.get("whole_row"):
                        self._replaced_raw.append((len(self.batches), str(op["sheet"]), at["r1"], at["c1"]))
                if str(kind or "").startswith(("set_", "clear")) and (op.get("cell") or op.get("range")):
                    # 18 Nov 2026: the cells Prep wrote. Those it froze on purpose (they asked Excel about
                    # their own file, FRM-012): a text built on them differs from the original by the file's
                    # location only. Every other write: a text that reads it is never let off.
                    from .formula_utils import parse_ref
                    at = parse_ref(str(op.get("range") or op.get("cell") or "").replace("$", ""))
                    n = len(self.batches)
                    if at and kind in ("set_formula", "set_array_formula", "set_value") and op.get("rule_id") == "FRM-012" \
                            and not at.get("whole_column") and not at.get("whole_row") and (at["r2"] - at["r1"] + 1) * (at["c2"] - at["c1"] + 1) <= 10_000:
                        self._frozen_raw.extend((n, str(op["sheet"]), r, c) for r in range(at["r1"], at["r2"] + 1) for c in range(at["c1"], at["c2"] + 1))
                    elif at:
                        self._written_raw.append((n, str(op["sheet"]), at["r1"], at["r2"], at["c1"], at["c2"]))
                    else:
                        self._written_raw.append((n, str(op["sheet"]), 1, 1_048_576, 1, 16_384))  # not readable: the whole sheet
            for v in list(rows.values()) + list(cols.values()):
                v.sort()
            if rows or banded or cols or renames:
                self.batches.append({"rows": rows, "banded": banded, "cols": cols, "renames": renames})

    def neutralized(self, value: Any) -> Any:
        """A text with the  neutral path / file name texts of the plan named as <folder> / <file> (as _without_location does)."""
        if not self.neutral_texts or not isinstance(value, str):
            return value
        for text in sorted(set(self.neutral_texts), key=len, reverse=True):
            token = "<file>" if text.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm")) else "<folder>"
            # a folder's closing separator stays outside the token (the location stripper leaves it too)
            value = value.replace(text if token == "<file>" else text.rstrip("\\/"), token)
        return value

    def is_replaced(self, sheet: str, row: int, col: int) -> bool:
        """True for a cell of the PREPARED version whose formula the plan replaced by a title on purpose."""
        if self._replaced is None:
            self._replaced = {self.forward(sh, r, c, start=n) for n, sh, r, c in self._replaced_raw}
        return (sheet, row, col) in self._replaced

    def frozen(self) -> set[tuple[str, int, int]]:
        """The cells, in the PREPARED version's coordinates, that Prep froze on purpose (FRM-012)."""
        if self._frozen is None:
            self._frozen = {self.forward(sh, r, c, start=n) for n, sh, r, c in self._frozen_raw}
        return self._frozen

    def written(self) -> list[tuple[str, int, int, int, int]]:
        """Every other range Prep wrote, (sheet, r1, r2, c1, c2) in the PREPARED version's coordinates."""
        if self._written is None:
            self._written = []
            for n, sh, r1, r2, c1, c2 in self._written_raw:
                s, a, b = self.forward(sh, r1, c1, start=n)
                _, d, e = self.forward(sh, r2, c2, start=n)
                self._written.append((s, a, d, b, e))
        return self._written

    def forward(self, sheet: str, row: int, col: int, start: int = 0) -> tuple[str, int, int]:
        """An original cell -> the same cell in the  prepared version. Inside one
        Apply the  rows are inserted bottom-up at the coordinates  of the version
        it  started from,  so a row moves down by one for every insert at or above it (an insert that
        shifts only some columns: when the cell's column is one of them)."""
        for b in self.batches[start:]:
            row += bisect.bisect_right(b["rows"].get(sheet, ()), row) + sum(1 for k, c1, c2 in b.get("banded", {}).get(sheet, ()) if k <= row and c1 <= col <= c2)
            col += bisect.bisect_right(b["cols"].get(sheet, ()), col)
            sheet = b["renames"].get(sheet, sheet)
        return sheet, row, col

    def summary(self) -> dict[str, int]:
        return {
            "row_inserts": sum(len(v) for b in self.batches for v in list(b["rows"].values()) + list(b.get("banded", {}).values())),
            "column_inserts": sum(len(v) for b in self.batches for v in b["cols"].values()),
            "sheet_renames": sum(len(b["renames"]) for b in self.batches),
        }


def _read_formula_cells(wb: Any, tell: Callable[[str], None], formulas_out: dict[str, dict[tuple[int, int], str]] | None = None) -> dict[str, dict[tuple[int, int], tuple[Any, bool]]]:
    """{sheet:  {(row, col):  (value, volatile)}} for every formula cell. Constants
    do not compute: what Prep writes as text (titles) is not a model value.
    formulas_out, when given, receives every formula cell's R1C1 text."""
    out: dict[str, dict[tuple[int, int], tuple[Any, bool]]] = {}
    volatile: dict[str, bool] = {}
    texts: dict[str, str] = {}  # one string object per  distinct formula text
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
                        if formulas_out is not None:
                            formulas_out.setdefault(name, {})[(r0 + start + i, c0 + j)] = texts.setdefault(f, f)
        out[name] = cells
    return out


NameKey = tuple  # (sheet  | None,  NAME): a name kept for one sheet, or for the whole workbook (None)
Rect = tuple     # (sheet,  r1, r2,  c1, c2)


def _read_names(wb: Any) -> dict[NameKey, list[Rect] | None]:
    """{(sheet  | None,  NAME): the cells it points to}; None for a name whose cells cannot be told
    (a constant, OFFSET / INDIRECT, another workbook): a text read through it is never let off."""
    out: dict[NameKey, list[Rect] | None] = {}
    try:
        names = list(wb.Names)
    except Exception:
        return out
    for n in names[:20_000]:
        try:
            full = str(n.Name)
            target = str(n.RefersToR1C1 or "")
        except Exception:
            continue
        scope, _, bare = full.rpartition("!")
        sheet = scope.strip("'").replace("''", "'") if scope else None
        parsed = parse_r1c1(target)
        rects: list[Rect] = []
        ok = bool(parsed.refs) and not parsed.dynamic and not parsed.external and not parsed.names
        for spec, a, b in parsed.refs:
            if not spec or spec[0] != "sheet":
                ok = False
                break
            r1, r2, c1, c2 = ref_rect((spec, a, b), 1, 1)
            rects.append((spec[1], r1, r2, c1, c2))
        out[(sheet, bare.upper())] = rects if ok else None
    return out


MAX_TRACE_STEPS = 2_000_000  # total  work spent  following text cells back to a frozen cell
MAX_TRACED = 500             # text  cells followed  back (the rest are counted as changed, as before)
BIG_RANGE = 5_000            # a  range larger  than this is searched through the sheet's formula cells


def _overlaps(rect: Rect, rects: Iterable[Rect]) -> bool:
    sh, r1, r2, c1, c2 = rect
    return any(s == sh and r1 <= d and a <= r2 and c1 <= e and b <= c2 for s, a, d, b, e in rects)


def _reading_frozen(starts: list[tuple[str, int, int]], formulas: dict[str, dict[tuple[int, int], str]],
                    names: dict[NameKey, list[Rect] | None], frozen: set[tuple[str, int, int]],
                    written: list[Rect] = (), changed: set[tuple[str, int, int]] = frozenset()) -> set[tuple[str, int, int]]:
    """The  cells of  `starts` (PREPARED coordinates) whose value comes from a cell Prep froze on
    purpose and from nothing else that moved: everything they read -- directly, through other
    formulas, through a defined name -- is  followed, and a cell is let off only when

      * at least one  cell read is frozen,
      * no other cell  read was written by Prep (`written`) or  computes another value now (`changed`),
      * every read can be told: no INDIRECT  / OFFSET / table, no 3-D or  outside reference, no
        name whose cells cannot be told, and the work stays inside the budget.

    Anything else leaves the cell out, so it still counts as changed."""
    hits: set[tuple[str, int, int]] = set()
    budget = MAX_TRACE_STEPS
    frozen_by_sheet: dict[str, list[tuple[int, int]]] = {}
    for sh, r, c in frozen:
        frozen_by_sheet.setdefault(sh, []).append((r, c))
    starts_set = set(starts)
    for start in starts[:MAX_TRACED]:
        if budget <= 0:
            break
        seen = {start}
        todo = [start]
        found = False
        bad = False
        while todo and not bad:
            if budget <= 0:
                bad = True
                break
            sheet, row, col = todo.pop()
            key0 = (sheet, row, col)
            if key0 != start and key0 in changed and key0 not in starts_set:
                bad = True  # it reads a cell that computes another value now, for another reason
                break
            fx = formulas.get(sheet, {}).get((row, col))
            if fx is None:
                continue
            parsed = parse_r1c1(fx)
            if parsed.dynamic or parsed.external:
                bad = True
                break
            rects: list[Rect] = []
            for spec, a, b in parsed.refs:
                if spec is not None and spec[0] != "sheet":
                    bad = True  # a 3-D reference
                    break
                r1, r2, c1, c2 = ref_rect((spec, a, b), row, col)
                rects.append((sheet if spec is None else spec[1], r1, r2, c1, c2))
            for nm in parsed.names:
                target = names[(sheet, nm)] if (sheet, nm) in names else names.get((None, nm), None)
                if target is None:
                    bad = True  # a name that cannot be followed (or a function this parser read as a name)
                    break
                rects.extend(target)
            if bad:
                break
            for rect in rects:
                sh, r1, r2, c1, c2 = rect
                budget -= 1
                if any(r1 <= fr <= r2 and c1 <= fc <= c2 for fr, fc in frozen_by_sheet.get(sh, ())):
                    found = True
                if _overlaps(rect, written):
                    bad = True
                    break
                cells = formulas.get(sh)
                if not cells:
                    continue
                area = (r2 - r1 + 1) * (c2 - c1 + 1)
                budget -= min(area, len(cells))
                if area <= BIG_RANGE:
                    inside = [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in cells]
                else:
                    inside = [(r, c) for (r, c) in cells if r1 <= r <= r2 and c1 <= c <= c2]
                for r, c in inside:
                    key = (sh, r, c)
                    if key in frozen:
                        continue  # the frozen cell itself: its own text is the  neutral one, on purpose
                    if key not in seen:
                        seen.add(key)
                        todo.append(key)
        if found and not bad:
            hits.add(start)
    return hits


def _is_error(v: Any) -> bool:
    return v.__class__ is int and v in XL_ERROR_CODES


def _without_location(v: Any, path: Path) -> Any:
    """A  text value  with the workbook's own folder and file name taken out:
    =CELL("filename") and what is built on it differ between two copies of the
    same workbook only because the copies sit in different folders."""
    if not isinstance(v, str) or len(v) < 3:
        return v
    low = v.lower()
    for piece, token in ((str(path.parent).lower(), "<folder>"), (path.name.lower(), "<file>"), (path.stem.lower(), "<file>")):
        if piece and piece in low:
            i = low.index(piece)
            v = v[:i] + token + v[i + len(piece):]
            low = v.lower()
    return v


def compare_workbooks(original: Path, prepared: Path, moves: Moves, progress: Callable[..., None] | None = None) -> dict[str, Any]:
    """Recalculate  both and  compare every formula cell of the original with the
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

    frozen = moves.frozen()
    after_formulas: dict[str, dict[tuple[int, int], str]] = {}
    after_names: dict[NameKey, list[Rect] | None] = {}
    try:
        with excel_session() as excel:
            def one(path: Path, what: str, keep: bool = False) -> Any:
                # one after the  other: the two usually carry the same file name, and Excel
                # cannot  hold two workbooks of the same  name open at once
                wb = None
                try:
                    tell(f"Opening {what} in Excel")
                    wb = open_workbook(excel, path, read_only=True)
                    if wb is None:
                        raise RuntimeError(f"Excel could not open {path.name}")
                    try:
                        excel.Calculation = -4135  # manual: one explicit full  rebuild
                    except Exception:
                        pass
                    tell(f"Recalculating {what}")
                    excel.CalculateFullRebuild()
                    if keep:
                        after_names.update(_read_names(wb))
                    return _read_formula_cells(wb, lambda sheet: tell(f"Reading {what}: {sheet}"), after_formulas if keep else None)
                finally:
                    close_quietly(wb)
                    wb = None

            before = one(original, "the original")
            after = one(prepared, "the current version", keep=bool(frozen))
    except Exception as exc:
        # Excel's error in plain words, never pywin32's tuple: the screens show  the verdict as it is.
        # Any other error  keeps its type in front: its text alone may say nothing.
        said = plain_error(exc, with_type=True)
        return {"ran": False, "clean": None, "verdict": f"The numbers could not be compared: {said[:300]}", "counts": {}, "regressions": []}

    counts = {"compared": 0, "same": 0, "good_to_error": 0, "good_to_other": 0, "error_to_good": 0, "error_to_error": 0, "missing": 0, "missing_good": 0, "replaced_on_purpose": 0, "volatile_skipped": 0, "frozen_location": 0}
    samples: dict[str, list[dict[str, Any]]] = {"good_to_error": [], "good_to_other": [], "error_to_good": [], "missing": []}
    by_sheet: dict[str, dict[str, int]] = {}
    regressions: list[tuple[str, int, int, Any]] = []
    moved: set[tuple[str, int, int]] = set()  # cells (PREPARED coordinates) that compute another value now
    held: list[tuple[str, str, int, int, str, int, int, Any, Any]] = []  # text -> other text, followed back to a frozen cell below
    for sheet, cells in before.items():
        per_sheet: dict[str, int] = by_sheet.setdefault(sheet, {})
        for (row, col), (v0, vol0) in cells.items():
            s1, r1, c1 = moves.forward(sheet, row, col)
            got = after.get(s1, {}).get((r1, c1))
            if got is None and not _is_error(v0) and moves.is_replaced(s1, r1, c1):
                # a formula caption the plan replaced by its title on purpose: expected, not a lost formula
                counts["replaced_on_purpose"] += 1
                continue
            if got is None:
                # the formula is gone: Prep replaced it by a value on purpose (a dead #REF! formula), or cleared it
                counts["missing"] += 1
                if not _is_error(v0):
                    counts["missing_good"] += 1  # a  formula that computed a value is not a formula any more
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
            elif _same(v0, v1, NUMERIC_TOL) or (isinstance(v0, str) and isinstance(v1, str) and _without_location(v0, original) == _without_location(moves.neutralized(v1), prepared)):
                kind = "same"
            else:
                kind = "good_to_other"
            if kind == "good_to_other" and frozen and isinstance(v0, str) and isinstance(v1, str):
                held.append((sheet, s1, r1, c1, kind, row, col, v0, v1))
                continue
            counts[kind] += 1
            if kind in ("same", "error_to_error"):
                continue
            per_sheet[kind] = per_sheet.get(kind, 0) + 1
            moved.add((s1, r1, c1))
            if kind == "good_to_error" and len(regressions) < MAX_REGRESSIONS:
                regressions.append((s1, r1, c1, v0))
            if len(samples[kind]) < MAX_SAMPLES:
                samples[kind].append({"original": f"{sheet}!{a1(row, col)}", "now": f"{s1}!{a1(r1, c1)}", "was": _show(v0), "is": _show(v1)})
    if held:
        # 18 Nov 2026: a text that differs only because it reads a cell Prep froze on purpose
        # (its own file location, FRM-012) is expected; numbers are never let off this way
        tell("Following changed texts back to the cells Prep froze")
        let_off = _reading_frozen([(s1, r1, c1) for _, s1, r1, c1, *_ in held], after_formulas, after_names, frozen, moves.written(), moved)
        for sheet, s1, r1, c1, kind, row, col, v0, v1 in held:
            if (s1, r1, c1) in let_off:
                counts["frozen_location"] += 1
                continue
            counts[kind] += 1
            by_sheet[sheet][kind] = by_sheet[sheet].get(kind, 0) + 1
            if len(samples[kind]) < MAX_SAMPLES:
                samples[kind].append({"original": f"{sheet}!{a1(row, col)}", "now": f"{s1}!{a1(r1, c1)}", "was": _show(v0), "is": _show(v1)})
    by_sheet = {k: v for k, v in by_sheet.items() if v}
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
    if counts["replaced_on_purpose"]:
        verdict += f" {counts['replaced_on_purpose']:,} formula caption(s) were replaced by their title on purpose (expected)."
    if counts["frozen_location"]:
        verdict += f" {counts['frozen_location']:,} text(s) built on the workbook's own folder or file name now show the neutral name Prep wrote on purpose (expected)."
    return {
        "ran": True, "clean": changed == 0, "verdict": verdict, "counts": counts, "by_sheet": by_sheet, "samples": samples,
        "moves": moves.summary(), "seconds": round(time.time() - t0, 1), "regressions": regressions,
    }
