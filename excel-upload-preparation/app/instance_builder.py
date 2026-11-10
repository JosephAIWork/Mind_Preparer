"""Mind instances -- the macro's run loop as one copy of the model per key.

A model that runs one policy at a time has a macro that sets a row number, recalculates and adds
the pass's results to a block (one block per group, product, tariff ...). The Mind loop
(app.loop_builder) writes that loop as MM_LOOP / MM_RESULT in one workbook. This module writes it
as Mind *instances*: one copy of the model per value of a key column of the model-point table, a
loop inside each copy over that copy's rows, and a second small MAIN workbook (the one workbook
Mind never copies) that holds the keys and reads every copy's results. Both files go to one Mind
project.

What the plan does to the model (a hand-built prototype of the same transformation converted, ran
and matched the macro's sums by key in a real Mind project on 2026-10-28; this builder's own
output converted and ran there on 2026-10-29, 0 differences from Excel's sums by key):
- the model-point table (the range the loop driver indexes through INDEX(range, driver, ...))
  becomes a split grid: a key column as its first column holding every row's key, the header
  "Instance key", the title "#<Name> /InstanceSplit" directly above it, an empty row above the
  title (rows and the column are inserted when needed; in each copy Mind keeps the rows whose
  key is the copy's key). Filled columns touching the table on its left are part of the grid in
  Mind's eyes, so the key column goes before them too;
- the loop's count cell gets =COUNTA(<the key column>): each copy counts its own rows;
- a free cell gets the biggest count (the loop size), and the driver cell
  =MM_LOOPINSTANCE("<loop>", <loop size cell>, <count cell>): 1..count in every copy (plus the
  header row when the indexed range starts on it, so every pass reads a data row);
- every formula cell of the block the macro copies out on every pass becomes
  =MM_RESULT(<the cell of the calculation sheet it reads>): the sum over the copy's loop, which
  is what the macro's accumulation gave;
- the macro's paste area is emptied, and a free cell shows the copy's key (MM_INSTANCEKEY).

The main workbook: sheet "Instances" with "#<key> instances /InstanceKeys" (header = the model's
file name, one row per key), sheet "Summary" with "#Results by <key>" (one row per key, one
column per row of the output block, MM_INSTANCE by key) and "#All copies" (the whole block as plain
links to the model, with a Total row: in Mind a plain reference to a copied workbook is the sum over
the copies -- checked in a real Mind project on 29 October 2026, the prototype's main workbook).

Copies from loops (build_instance_plan(..., loops=[...])): for a macro that runs the whole model once
per scenario and cohort, one copy per combination of the chosen loops' values; each copy reads its
values from a small split grid, the macro's other loops stay MM_LOOPs inside every copy (checked in
Mind on 10 November 2026: 24 copies of an IFRS 17 tool, each equal to Excel with its drivers set).

    python -m app.instance_builder <workbook.xlsm> [--json] [--key K] [--keys all] [--name N] [--loops A,B] [--results Sheet!A1:B9]
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

from .formula_utils import cell_refs_in_formula, col_to_num, find_calls, num_to_col, parse_ref
from .loop_builder import MAX_CLEARS, _cells_of, _dropdown_at, _held, _qs, _sheet_max, _split, loop_name_from

ACTION_ID = "mind_instances"
RULE_ID = "INST-001"
KEY_HEADER = "Instance key"
SPLIT_FLAG = "/InstanceSplit"
SAMPLE_ROWS = 5000  # rows of the table read to find the key candidates
MIN_DISTINCT, MAX_DISTINCT = 2, 200  # distinct values a key column holds
MAX_MAIN_CELLS = 30_000  # formulas one grid of the main workbook holds at most
MAX_LEFT_COLUMNS = 10  # filled columns left of the table the plan takes into the grid at most
KEY_WORDS_RE = re.compile(r"tariff|product|group|portfolio|segment|class|category|type|country|fund|treaty|cohort", re.I)
# the macro's own key: WorksheetFunction.VLookup(<loop variable>, Range("<table>"), k, ...) -> column k of the table
_VLOOKUP_RE = re.compile(r"""VLookup\(\s*(?P<var>[A-Za-z_]\w*)\s*,\s*(?:(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*)?Range\(\s*"(?P<rng>[^"]+)"\s*\)(?:\.Value2?)?\s*,\s*(?P<k>\d+)""", re.I)
# the list of keys the macro runs through: WorksheetFunction.Match(<key>, Range("<list>"), 0)
_MATCH_RE = re.compile(r"""Match\(\s*[^,()]+?\s*,\s*(?:(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*)?Range\(\s*"(?P<rng>[^"]+)"\s*\)""", re.I)


# --- small helpers --------------------------------------------------------------------------------
def _op(op: str, sheet: str, **kw: Any) -> dict[str, Any]:
    return {"op": op, "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": sheet, **kw}


def _plain(text: Any, default: str = "Key") -> str:
    """A header as a grid title can carry it: no '#', '/' or '&' (Mind reads them as a title or a flag)."""
    s = re.sub(r"[#/&]+", " ", str(text or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return s or default


def _safe_stem(text: Any, default: str = "model") -> str:
    """A file stem of letters, digits and underscores (Mind reads the /InstanceKeys header as a file name)."""
    s = re.sub(r"[^A-Za-z0-9_]+", "_", str(text or "")).strip("_")
    return s or default


def _json_value(v: Any) -> Any:
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def _empty(v: Any) -> bool:
    return v is None or v == ""


def _xref(file: str, sheet: str, ref: str) -> str:
    """'[model.xlsm]Sheet'!$A$1: a reference to another workbook in the same folder."""
    return f"'[{file}]{sheet.replace(chr(39), chr(39) * 2)}'!{ref}"


def _abs(sheet: str, col: int, row: int) -> str:
    return f"{_qs(sheet)}!${num_to_col(col)}${row}"


def _norm_arg(arg: str) -> str:
    """A formula argument as a comparable text: no '$', no quotes, upper case."""
    s = arg.strip().replace("$", "")
    if s.startswith("@"):
        s = s[1:]
    if "!" in s:
        sheet, ref = _split(s)
        return f"{sheet}!{ref}".upper()
    return s.upper()


def _resolve_range(arg: str, default_sheet: str, names: Any) -> tuple[str, dict[str, Any], str | None] | None:
    """(sheet, parsed range, defined name) for a literal range or a defined name, else None."""
    a = arg.strip()
    if a.startswith("@"):
        a = a[1:]
    r = parse_ref(a)
    name = None
    if r is None:
        ref = names.ref(a) if names else None
        if not ref:
            return None
        r = parse_ref(ref)
        if r is None:
            return None
        name = names.pretty(a)
    return (r["sheet"] or default_sheet), r, name


# --- the loop ----------------------------------------------------------------------------------------
def _pick_loop(loop_plan: dict[str, Any]) -> tuple[dict[str, Any] | None, str | None, str]:
    """The loop the instances are built on: the first counter loop of the loop plan that has a
    count cell. -> (loop entry or None, a note when it cannot be used, the plan's status)."""
    if not loop_plan.get("found"):
        return None, "No run loop in the macros: nothing to build.", "not_built"
    loops = [le for le in loop_plan.get("loops") or [] if le.get("status") != "not_built"]
    for le in loops:
        if le.get("kind") == "counter" and le.get("status") == "ready" and le.get("cells"):
            return le, None, "ready"
    if not loops:
        return None, "The macro's loop cannot be written as a Mind loop (the Mind loop panel says why): nothing to build.", "not_built"
    le = loops[0]
    if le.get("kind") == "counter":
        return le, (f"The loop {le['name']} counts {le.get('first')} to {le.get('last')} but no cell holds the count. Choose a count cell on the Mind loop panel first: "
                    "each copy counts its own rows in that cell."), "needs_decision"
    how = le.get("how") or le.get("kind") or "another kind of loop"
    return le, (f"The loop {le['name']} is not a counter over the rows of a table ({how}). The instances are built from a loop that sets a row number "
                "(1, 2, 3 ...) in a cell and reads a cell that holds the count."), "needs_decision"


# --- the table -----------------------------------------------------------------------------------
def _find_table(analysis: dict[str, Any], le: dict[str, Any], names: Any) -> dict[str, Any] | None:
    """The model-point table: the union of the INDEX (else OFFSET) ranges whose row argument is the
    loop driver -- its defined name, or its address with or without '$' -- on the sheet most of
    those formulas read. None when no formula indexes a range by the driver."""
    wb0 = analysis["workbooks"][0]
    drv_sheet, drv_ref = _split(str(le.get("driver") or ""))
    if not drv_sheet or not parse_ref(drv_ref):
        return None
    drv_ref = drv_ref.replace("$", "").upper()
    wanted = {f"{drv_sheet}!{drv_ref}".upper()}
    if le.get("driver_name"):
        wanted.add(str(le["driver_name"]).upper())
    hits: dict[str, list[dict[str, Any]]] = {}
    for f in wb0.get("formulas") or []:
        formula = str(f.get("formula") or "")
        up = formula.upper()
        if "INDEX(" not in up and "OFFSET(" not in up:
            continue
        for fn in ("INDEX", "OFFSET"):
            for args in find_calls(formula, fn):
                if len(args) < 2:
                    continue
                row_arg = _norm_arg(args[1])
                if row_arg not in wanted and not (row_arg == drv_ref and str(f["sheet"]).upper() == drv_sheet.upper()):
                    continue
                got = _resolve_range(args[0], f["sheet"], names)
                if not got:
                    continue
                sheet, r, name = got
                if fn == "OFFSET":
                    if r["r1"] != r["r2"]:
                        continue  # OFFSET from a block: not a row picked by the driver
                    r = {**r, "r1": r["r1"] + 1, "r2": 1_048_576, "whole_column": True}  # the rows below the anchor
                hits.setdefault(sheet, []).append({"r": r, "name": name, "fn": fn, "at": f"{f['sheet']}!{f['cell']}"})
    if not hits:
        return None
    sheet = max(hits, key=lambda s: (sum(1 for h in hits[s] if h["fn"] == "INDEX"), len(hits[s])))
    rows = hits[sheet]
    if any(h["fn"] == "INDEX" for h in rows):
        rows = [h for h in rows if h["fn"] == "INDEX"]
    max_col, max_row = _sheet_max(analysis, sheet)
    r1 = min(h["r"]["r1"] for h in rows)
    r2 = min(max(h["r"]["r2"] for h in rows), max_row)
    c1 = min(h["r"]["c1"] for h in rows)
    c2 = min(max(h["r"]["c2"] for h in rows), max_col)
    if r2 < r1 or c2 < c1:
        return None
    used = Counter(h["name"] for h in rows if h["name"])
    return {"sheet": sheet, "r1": r1, "r2": r2, "c1": c1, "c2": c2, "formulas": len(rows), "name": used.most_common(1)[0][0] if used else None,
            "cells": sorted({h["at"] for h in rows})[:6], "how": rows[0]["fn"]}


def _profile(analysis: dict[str, Any], sheet: str, row: int, c1: int, c2: int) -> tuple[int, int, int]:
    """(text cells, number cells, filled cells) in a row, over the table's first twelve columns."""
    texts = nums = filled = 0
    for c in range(c1, min(c2, c1 + 11) + 1):
        v = _held(analysis, sheet, c, row)
        if _empty(v):
            continue
        filled += 1
        if isinstance(v, str) and not v.startswith("="):
            texts += 1
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            nums += 1
    return texts, nums, filled


def _header_rows(analysis: dict[str, Any], sheet: str, r1: int, c1: int, c2: int) -> tuple[int | None, int]:
    """(header row, first data row): the header is the text row directly above the first row the
    driver indexes, or that first row itself when it is text above numbers."""
    if r1 >= 2:
        t, n, _ = _profile(analysis, sheet, r1 - 1, c1, c2)
        if t and not n:
            return r1 - 1, r1
    t, n, _ = _profile(analysis, sheet, r1, c1, c2)
    _, n2, _ = _profile(analysis, sheet, r1 + 1, c1, c2)
    if t and not n and n2:
        return r1, r1 + 1
    return None, r1


def _row_empty(analysis: dict[str, Any], sheet: str, row: int, c1: int, c2: int, ignore: tuple[int, int] | None = None) -> bool:
    """Is the row empty across the columns (an ignored cell aside)?"""
    if row < 1:
        return True
    for c in range(c1, c2 + 1):
        if ignore and (row, c) == ignore:
            continue
        if not _empty(_held(analysis, sheet, c, row)):
            return False
    return True


def _column_empty(analysis: dict[str, Any], sheet: str, col: int, r1: int, r2: int) -> bool:
    """Is the column empty over rows r1..r2? (read from the analysis copy's own cells, so an empty
    column costs one lookup per row)"""
    from .inventory import open_cached

    try:
        cells = getattr(open_cached(analysis)[sheet], "_cells", None)
    except Exception:
        cells = None
    if cells is None or not hasattr(cells, "get"):
        return all(_empty(_held(analysis, sheet, col, r)) for r in range(r1, r2 + 1))
    return all(_empty(getattr(cells.get((r, col)), "value", None)) for r in range(r1, r2 + 1))


def _filled_columns_left(analysis: dict[str, Any], sheet: str, c1: int, r1: int, r2: int) -> int:
    """How many columns directly left of column c1 hold something over rows r1..r2 (the table's
    header and data rows). In Mind, cells touching a grid are part of it: the key column has to go
    before them, or it is not the grid's first column and the split keys on the wrong column."""
    n = 0
    while c1 - n - 1 >= 1 and n < MAX_LEFT_COLUMNS and not _column_empty(analysis, sheet, c1 - n - 1, r1, r2):
        n += 1
    return n


def _scan_candidates(copy_path: str, sheet: str, c1: int, c2: int, first_row: int, stop: int) -> dict[str, Any]:
    """One pass over the first SAMPLE_ROWS rows of the table (openpyxl read_only): per column, is it
    text, how many distinct values, how many filled rows."""
    import openpyxl

    cols = range(c1, c2 + 1)
    distinct: dict[int, set] = {c: set() for c in cols}
    texty = {c: True for c in cols}
    filled = {c: 0 for c in cols}
    sampled = 0
    last = first_row - 1
    wb = openpyxl.load_workbook(copy_path, read_only=True, data_only=True, keep_links=False)
    try:
        ws = wb[sheet]
        stop = min(stop, ws.max_row or stop)
        for i, row in enumerate(ws.iter_rows(min_row=first_row, max_row=min(stop, first_row + SAMPLE_ROWS - 1), min_col=c1, max_col=c2, values_only=True)):
            any_v = False
            for j, v in enumerate(row):
                if _empty(v):
                    continue
                any_v = True
                c = c1 + j
                filled[c] += 1
                if isinstance(v, str):
                    if len(distinct[c]) <= MAX_DISTINCT:
                        distinct[c].add(v)
                else:
                    texty[c] = False
            if any_v:
                sampled += 1
                last = first_row + i
    finally:
        wb.close()
    need = max(1, int(0.8 * sampled))
    candidates = [c for c in cols if texty[c] and filled[c] >= need and MIN_DISTINCT <= len(distinct[c]) <= MAX_DISTINCT]
    return {"candidates": candidates, "distinct": {c: len(distinct[c]) for c in cols}, "texty": texty, "filled": filled, "sampled": sampled, "last_sampled": last, "stop": stop}


def _read_column(copy_path: str, sheet: str, c1: int, c2: int, col: int, first_row: int, stop: int) -> tuple[list[Any], int]:
    """(the values of one column of the table, the table's last filled row), read whole."""
    import openpyxl

    values: list[Any] = []
    last = first_row - 1
    empty_run = 0
    wb = openpyxl.load_workbook(copy_path, read_only=True, data_only=True, keep_links=False)
    try:
        ws = wb[sheet]
        for i, row in enumerate(ws.iter_rows(min_row=first_row, max_row=stop, min_col=c1, max_col=c2, values_only=True)):
            if any(not _empty(v) for v in row):
                last = first_row + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run > 5000:
                    break
            values.append(row[col - c1])
    finally:
        wb.close()
    return values[: max(0, last - first_row + 1)], last


# --- the key ---------------------------------------------------------------------------------------
def _vba_text(analysis: dict[str, Any]) -> str:
    from . import model_map as _mm

    try:
        return "\n".join(code for _name, code in _mm.read_vba(analysis["workbooks"][0]["copy_path"]))
    except Exception:
        return ""


def _macro_key_column(vba: str, loop_var: str, table: dict[str, Any], names: Any) -> int | None:
    """The column the macro itself looks the key up in: VLookup(<loop var>, Range("<table>"), k) over
    a range that overlaps the table -> that range's k-th column."""
    for m in _VLOOKUP_RE.finditer(vba):
        if m.group("var").lower() != str(loop_var or "").lower():
            continue
        got = _resolve_range(m.group("rng"), m.group("sheet") or table["sheet"], names)
        if not got:
            continue
        sheet, r, _name = got
        if sheet != table["sheet"] or r["r2"] < table["r1"] or r["r1"] > table["r2"] or r["c2"] < table["c1"] or r["c1"] > table["c2"]:
            continue
        col = r["c1"] + int(m.group("k")) - 1
        if table["c1"] <= col <= table["c2"]:
            return col
    return None


def _macro_key_list(analysis: dict[str, Any], vba: str, table: dict[str, Any], names: Any) -> list[Any] | None:
    """The list of keys the macro runs through (Match(<key>, Range("<list>"), 0)): its texts, in order."""
    from .inventory import values_cell

    for m in _MATCH_RE.finditer(vba):
        got = _resolve_range(m.group("rng"), m.group("sheet") or table["sheet"], names)
        if not got:
            continue
        sheet, r, _name = got
        if r["whole_column"] or r["whole_row"] or (r["r2"] - r["r1"] + 1) * (r["c2"] - r["c1"] + 1) > 2000:
            continue
        if sheet == table["sheet"] and not (r["r2"] < table["r1"] or r["r1"] > table["r2"]):
            continue  # the table itself, not a list
        out: list[Any] = []
        for c, row in _cells_of(sheet, f"{num_to_col(r['c1'])}{r['r1']}:{num_to_col(r['c2'])}{r['r2']}"):
            v = _held(analysis, sheet, c, row)
            if isinstance(v, str) and v.startswith("="):
                try:
                    v = values_cell(analysis, sheet, row, c)
                except Exception:
                    v = None
            if not _empty(v):
                out.append(_json_value(v))
        if out:
            return out
    return None


def _choose_key(table: dict[str, Any], headers: dict[int, Any], candidates: list[int], macro_col: int | None, user_key: Any) -> tuple[int | None, str, str | None]:
    """(the key column, how it was chosen, a note when the user's key was not found)."""
    if user_key not in (None, ""):
        want = str(user_key).strip()
        for c in range(table["c1"], table["c2"] + 1):
            h = headers.get(c)
            if isinstance(h, str) and h.strip().lower() == want.lower():
                return c, "user", None
        if re.fullmatch(r"[A-Za-z]{1,3}", want):
            c = col_to_num(want.upper())
            if table["c1"] <= c <= table["c2"]:
                return c, "user", None
        return None, "user", f"No column of the table is called {want!r} (a header text or a column letter between {num_to_col(table['c1'])} and {num_to_col(table['c2'])})."
    if macro_col is not None:
        return macro_col, "macro", None
    for c in candidates:
        h = headers.get(c)
        if isinstance(h, str) and KEY_WORDS_RE.search(h):
            return c, "heuristic", None
    return None, "heuristic", None


# --- free cells and shifting -----------------------------------------------------------------------
def _free_row(analysis: dict[str, Any], sheet: str, near_row: int, c1: int, c2: int, avoid: list[tuple[int, int, int, int]]) -> int:
    """A row near `near_row` that is empty, with empty rows above and below, across columns c1..c2,
    and outside the rectangles to avoid (r1, c1, r2, c2): the loop size and the copy's key go there."""
    _max_col, max_row = _sheet_max(analysis, sheet)

    def ok(r: int) -> bool:
        if r < 2:
            return False
        for rr in (r - 1, r, r + 1):
            if not _row_empty(analysis, sheet, rr, c1, c2):
                return False
            for a1, b1, a2, b2 in avoid:
                if a1 <= rr <= a2 and not (b2 < c1 or b1 > c2):
                    return False
        return True

    for r in list(range(near_row + 1, near_row + 80)) + list(range(near_row - 1, 1, -1)):
        if ok(r):
            return r
    return max_row + 2


class _Shift:
    """Addresses after the plan's inserts on the table sheet: rows from `at` move down by `rows`,
    columns from `col` move right by one."""

    def __init__(self, sheet: str, at: int, rows: int, col: int | None):
        self.sheet, self.at, self.rows, self.col = sheet, at, rows, col

    @property
    def any(self) -> bool:
        return bool(self.rows or self.col)

    def cell(self, sheet: str, col: int, row: int) -> tuple[int, int]:
        if sheet != self.sheet:
            return col, row
        if self.rows and row >= self.at:
            row += self.rows
        if self.col and col >= self.col:
            col += 1
        return col, row

    def ref(self, sheet: str, col: int, row: int) -> str:
        c, r = self.cell(sheet, col, row)
        return f"{num_to_col(c)}{r}"

    def touches(self, op: dict[str, Any]) -> bool:
        """Does the operation write the table sheet, or a formula that reads it?"""
        if not self.any:
            return False
        if op["sheet"] == self.sheet:
            return True
        after = op.get("after")
        if isinstance(after, str) and after.startswith("="):
            return any((r.get("sheet") or op["sheet"]) == self.sheet for r in cell_refs_in_formula(after))
        return False


# --- the plan --------------------------------------------------------------------------------------
def build_instance_plan(analysis: dict[str, Any], model_map: dict[str, Any], *, loop_plan: dict[str, Any] | None = None, key: Any = None, keys: str = "present",
                        model_name: str | None = None, loops: list[str] | None = None, results: str | None = None) -> dict[str, Any]:
    """The Mind instances for the macro's run loop: what goes where, with the operations that write it.

    Two ways to make the copies:
    - by a key column of the model-point table (the default): one copy per key, a loop inside each
      copy over the copy's own rows. `key`: the key column -- a header text or a column letter of the
      table (default: the macro's own lookup column, else a header that names a group, else the text
      column with the fewest distinct values). `keys`: "present" = the key values that have rows;
      "all" = every value of the list the macro runs through, when it can be found.
    - by loops (`loops`: loop names or driver cells of the Mind loop plan): one copy per combination
      of the values of those loops -- a scenario x a cohort for instance. Each copy reads its own
      values from a small split grid; the macro's other loops stay MM_LOOPs inside every copy.
      `results`: the block the main workbook shows for every copy (default: the formula block of the
      model's output sheet).
    `model_name`: the model's file stem as it will be uploaded (default <workbook>_instances).
    Every plan lists `loop_options`: the loops copies can be made from, and why the others cannot."""
    from .loop_builder import build_loop_plan

    wanted = [str(x).strip() for x in (loops or []) if str(x or "").strip()]
    if loop_plan is None:
        # one loop plan per macro (Excel runs each on its own): the loops asked for pick theirs
        loop_plan = build_loop_plan(analysis, model_map)
        plans = [loop_plan] + [build_loop_plan(analysis, model_map, macro=m["name"]) for m in loop_plan.get("macros") or [] if not m.get("chosen")]
        if wanted:
            loop_plan = pick_loop_plan(plans, wanted)
        options = loop_options_of(plans) if len(plans) > 1 else loop_options(loop_plan)
    else:
        options = loop_options(loop_plan)
    if wanted:
        plan = _build_loops_plan(analysis, model_map, loop_plan, wanted, results=results, model_name=model_name)
    else:
        plan = _build_table_plan(analysis, model_map, loop_plan=loop_plan, key=key, keys=keys, model_name=model_name)
        plan["mode"] = "table"
        usable = [o["name"] for o in options if o["usable"]]
        if plan["status"] != "ready" and usable:
            plan["notes"].append(f"Or make one copy per value of a loop ({', '.join(usable)}): choose the loops under \"Copies from loops\".")
    plan["loop_options"] = options
    return plan


def _build_table_plan(analysis: dict[str, Any], model_map: dict[str, Any], *, loop_plan: dict[str, Any] | None = None, key: Any = None, keys: str = "present",
                      model_name: str | None = None) -> dict[str, Any]:
    """The copies by a key column of the model-point table (see build_instance_plan)."""
    from .grids import all_grids, grid_containing
    from .loop_builder import build_loop_plan
    from .model_map import _Names

    wb0 = analysis["workbooks"][0]
    names = _Names(wb0.get("defined_names", []) or [])
    if loop_plan is None:
        loop_plan = build_loop_plan(analysis, model_map)
    stem = _safe_stem(wb0.get("file_name") or Path(str(wb0.get("copy_path") or "model")).stem)
    model_stem = _safe_stem(model_name) if model_name else f"{stem}_instances"
    ext = str(wb0.get("file_type") or Path(str(wb0.get("copy_path") or "")).suffix.lstrip(".") or "xlsx").lower()
    main: dict[str, Any] = {"file": f"{stem}_Main.xlsx", "model_file": f"{model_stem}.{ext}", "keys": [], "grids": []}
    plan: dict[str, Any] = {"found": bool(loop_plan.get("found")), "status": "not_built", "notes": [], "loop": None, "table": None, "key": None, "split_grid": None,
                            "outputs": [], "clears": [], "operations": [], "operation_count": 0,
                            "counts": {"key_values": 0, "result_cells": 0, "cleared": 0, "inserted_rows": 0}, "main": main, "how_it_runs": None}
    notes = plan["notes"]

    le, why, status = _pick_loop(loop_plan)
    if le is None:
        plan["status"] = status
        notes.append(why or "nothing to build")
        return plan
    count_text = str(le["cells"][0]["range"]) if le.get("cells") else ""
    cnt_sheet, cnt_ref = _split(count_text)
    cnt = parse_ref(cnt_ref) if cnt_sheet else None
    drv_sheet, drv_ref = _split(str(le.get("driver") or ""))
    drv = parse_ref(drv_ref) if drv_sheet else None
    plan["loop"] = {"name": le["name"], "driver": le.get("driver"), "count_cell": count_text or None, "max_cell": None, "formula": None, "macro": le.get("macro"), "kind": le.get("kind"),
                    "first": le.get("first"), "last": le.get("last")}
    if why:
        plan["status"] = status
        notes.append(why)
        return plan
    if not cnt or not drv or cnt["cells"] != 1 or drv["cells"] != 1:
        plan["status"] = "needs_decision"
        notes.append(f"The loop {le['name']} has no single count cell and driver cell the plan can write ({count_text!r}, {le.get('driver')!r}).")
        return plan

    # --- the table -----------------------------------------------------------------------------------
    table = _find_table(analysis, le, names)
    if not table:
        plan["status"] = "needs_decision"
        notes.append(f"No formula reads a row of a table through the loop driver {le.get('driver')} (INDEX(range, {le.get('driver_name') or le.get('driver')}, ...)): "
                     "the model-point table could not be found, so there is nothing to split into copies.")
        return plan
    sheet, c1, c2 = table["sheet"], table["c1"], table["c2"]
    header_row, first_row = _header_rows(analysis, sheet, table["r1"], c1, c2)
    # filled columns touching the table on its left belong to the grid in Mind's eyes: the table
    # starts there, so the key column is inserted before them (and the title sits above it)
    n_left = _filled_columns_left(analysis, sheet, c1, header_row or first_row, max(table["r2"], first_row))
    if n_left:
        notes.append(f"{n_left} filled column(s) left of the table ({num_to_col(c1 - n_left)}{'' if n_left == 1 else ':' + num_to_col(c1 - 1)}) touch it, so Mind reads them as part of the grid: "
                     "the key column goes before them.")
        c1 -= n_left
        table = {**table, "c1": c1, "widened_by": n_left}
    copy_path = str(wb0["copy_path"])
    scan = _scan_candidates(copy_path, sheet, c1, c2, first_row, table["r2"])
    headers = {c: _held(analysis, sheet, c, header_row) for c in range(c1, c2 + 1)} if header_row else {}
    sampled = table["r2"] - first_row + 1 > SAMPLE_ROWS  # the counts come from the first SAMPLE_ROWS rows
    candidates = [{"column": num_to_col(c), "header": _json_value(headers.get(c)), "distinct": scan["distinct"][c], "sampled": sampled} for c in scan["candidates"]]
    vba = _vba_text(analysis)
    macro_col = _macro_key_column(vba, str(le.get("loop_var") or ""), table, names)
    key_col, chosen_by, key_note = _choose_key(table, headers, scan["candidates"], macro_col, key)
    if key_col is None and chosen_by == "heuristic" and scan["candidates"]:
        key_col = min(scan["candidates"], key=lambda c: (scan["distinct"][c], c))  # the fewest distinct values
    stop = scan["stop"]
    if key_col is None:
        plan["status"] = "needs_decision"
        plan["table"] = {"sheet": sheet, "ref": f"{num_to_col(c1)}{first_row}:{num_to_col(c2)}{max(first_row, scan['last_sampled'])}", "header_row": header_row, "first_row": first_row,
                         "last_row": scan["last_sampled"], "rows": max(0, scan["last_sampled"] - first_row + 1), "columns": c2 - c1 + 1, "name": table["name"]}
        plan["key"] = {"column": None, "header": None, "distinct": 0, "counts": {}, "candidates": candidates, "chosen_by": chosen_by}
        notes.append(key_note or f"No text column of the table {_qs(sheet)}!{num_to_col(c1)}:{num_to_col(c2)} holds {MIN_DISTINCT} to {MAX_DISTINCT} distinct values: name the key column.")
        return plan
    column_values, last_row = _read_column(copy_path, sheet, c1, c2, key_col, first_row, stop)
    last_row = max(last_row, scan["last_sampled"])
    column_values = column_values + [None] * (last_row - first_row + 1 - len(column_values))
    rows = last_row - first_row + 1
    if rows <= 0:
        plan["status"] = "needs_decision"
        notes.append(f"The table {_qs(sheet)}!{num_to_col(c1)}{first_row}:{num_to_col(c2)} holds no rows.")
        return plan
    present = Counter(v for v in column_values if not _empty(v))
    order = list(dict.fromkeys(v for v in column_values if not _empty(v)))  # first appearance
    macro_list = _macro_key_list(analysis, vba, table, names)
    if macro_list:
        listed = [v for v in macro_list if v in present]
        order = listed + [v for v in order if v not in listed]
    if keys == "all":
        if macro_list:
            key_values = list(dict.fromkeys(macro_list)) + [v for v in order if v not in macro_list]
            without = [v for v in key_values if v not in present]
            if without:
                notes.append(f"{len(without)} key(s) of the list the macro runs through have no row in the table: their copies compute nothing ({', '.join(str(v) for v in without[:6])}{', …' if len(without) > 6 else ''}).")
        else:
            key_values = order
            notes.append("The list of keys the macro runs through was not found: the keys are the values present in the table.")
    else:
        key_values = order
    missing = sum(1 for v in column_values if _empty(v))
    if missing:
        notes.append(f"{missing} row(s) of the table have no {_plain(headers.get(key_col), 'key')}: no copy keeps them.")
    if not scan["texty"].get(key_col, True):
        notes.append(f"The key column {num_to_col(key_col)} holds numbers as well as text: Mind compares the keys as they are written.")
    counts = {str(_json_value(k)): present[k] for k in key_values if k in present}
    max_count = max(present.values()) if present else 0
    key_header = _json_value(headers.get(key_col))
    word = _plain(key_header, "key")
    plan["table"] = {"sheet": sheet, "ref": f"{num_to_col(c1)}{first_row}:{num_to_col(c2)}{last_row}", "header_row": header_row, "first_row": first_row, "last_row": last_row,
                     "rows": rows, "columns": c2 - c1 + 1, "name": table["name"], "indexed_by": table["cells"]}
    for cand in candidates:
        if cand["column"] == num_to_col(key_col):
            cand.update(distinct=len(present), sampled=False)  # the chosen column was read whole
    plan["key"] = {"column": num_to_col(key_col), "header": key_header, "distinct": len(present), "counts": counts, "candidates": candidates, "chosen_by": chosen_by, "missing": missing,
                   "values": [_json_value(v) for v in key_values]}
    # MM_LOOPINSTANCE counts 1..count in every copy. INDEX(range, i) reads row i of the range: when
    # the range starts on the header row (the macro ran 2..n), the driver adds that row, so every
    # pass reads a data row; the macro's own first row beyond that is not kept (every row joins)
    rows_before = max(0, first_row - table["r1"])  # range rows above the first data row (1 when the header row is inside the range)
    first_n, last_n = le.get("first"), le.get("last")
    macro_rows = (last_n - first_n + 1) if isinstance(first_n, int) and isinstance(last_n, int) else last_n
    if macro_rows not in (None, rows):
        notes.append(f"The macro runs {macro_rows} row(s); the table holds {rows}: every row of the table joins its copy's loop.")
    if rows_before:
        notes.append(f"The range the loop indexes starts on the table's header row (the macro started at {first_n}): the driver adds {rows_before} to the loop index, so every pass reads a data row.")
    if isinstance(first_n, int) and first_n - 1 - rows_before > 0:
        notes.append(f"The macro starts at {first_n}, {first_n - 1 - rows_before} data row(s) into the table: here every row of the table joins its copy's loop (the loop starts at the first row).")
    elif isinstance(first_n, int) and first_n - 1 - rows_before < 0:
        notes.append(f"The macro's first pass ({first_n}) read the table's header row: in the copies the loop starts at the first data row.")

    # --- the layout: the key column first, a title with an empty row above it -----------------------
    insert_col = key_col != c1
    grids_here = [g for g in all_grids(analysis) if g.get("sheet") == sheet]
    g = grid_containing(grids_here, first_row, c1)
    old_title: tuple[int, int, str] | None = None
    if g and g.get("title_cell") and header_row and parse_ref(g["title_cell"]) and parse_ref(g["title_cell"])["r1"] == header_row - 1:
        tc = parse_ref(g["title_cell"])
        old_title = (tc["r1"], tc["c1"], str(g.get("title") or ""))
    left = max(1, c1 - 1)
    if header_row is None:
        at, extra_header = first_row, 1  # a header row is made for the table (Mind needs text headers)
        h_old = first_row
    else:
        at, extra_header = header_row, 0
        h_old = header_row
    ignore = (old_title[0], old_title[1]) if old_title else None
    e1 = _row_empty(analysis, sheet, h_old - 1, left, c2, ignore)
    e2 = _row_empty(analysis, sheet, h_old - 2, left, c2)
    if h_old >= 2 and e1 and (h_old == 2 or e2):
        k = 0
    elif h_old == 1 or e1:
        k = 1
    else:
        k = 2
    n_rows = k + extra_header
    sh = _Shift(sheet, at, n_rows, c1 if insert_col else None)
    header_new = h_old + k if extra_header == 0 else first_row + k  # the header row after the inserts
    title_row = header_new - 1
    key_new_col = c1  # the key column is the table's first column after the inserts
    first_new, last_new = first_row + n_rows, last_row + n_rows
    if old_title and SPLIT_FLAG.lower() in old_title[2].lower():
        title = old_title[2]
    elif old_title and old_title[2].startswith("#"):
        title = f"{old_title[2].rstrip()} {SPLIT_FLAG}"
    else:
        base = loop_name_from(table["name"] or sheet)
        taken = {str(x.get("name") or "").lower() for x in all_grids(analysis)}
        if base.lower() in taken and not (g and str(g.get("name") or "").lower() == base.lower()):
            base += "_rows"
        title = f"#{base} {SPLIT_FLAG}"
    ops: list[dict[str, Any]] = []
    for _ in range(n_rows):
        ops.append(_op("insert_row", sheet, row=at, before=f"row {at}", after=f"empty row inserted at {at} (the table moves down one row)",
                       note="room for the split grid's title: the title row and the row above it must be empty (Excel shifts every reference automatically)"))
    if insert_col:
        ops.append(_op("insert_column", sheet, column=num_to_col(c1), before=f"column {num_to_col(c1)}", after=f"empty column inserted at {num_to_col(c1)} (the table moves right one column)",
                       note=f"the key column must be the first column of the split grid; the {word} values are copied into it (Excel shifts every reference automatically)"))
    title_cell = f"{num_to_col(key_new_col)}{title_row}"
    if old_title:
        oc, orow = sh.cell(sheet, old_title[1], old_title[0])  # above the inserted rows: only the column insert moves it
        old_ref = f"{num_to_col(oc)}{orow}"
        if old_ref != title_cell:
            ops.append(_op("clear_cell", sheet, cell=old_ref, before=old_title[2], after=None, note="the table's title moves above its new first column (the key column)"))
    ops.append(_op("set_value", sheet, cell=title_cell, before=old_title[2] if old_title and f"{num_to_col(old_title[1])}{old_title[0]}" == title_cell else None, after=title,
                   note=f"the split grid's title: in each copy Mind keeps the rows whose first column holds the copy's {word}"))
    header_cell = f"{num_to_col(key_new_col)}{header_new}"
    key_header_now = headers.get(key_col)
    if insert_col or extra_header or not (isinstance(key_header_now, str) and key_header_now.strip() and not key_header_now.startswith("=")):
        ops.append(_op("set_value", sheet, cell=header_cell, before=None if insert_col or extra_header else _json_value(key_header_now), after=KEY_HEADER, note="the header of the key column (a grid's first row is its header)"))
    if extra_header:
        for c in range(c1, c2 + 1):
            if not insert_col and c == key_col:
                continue  # the key column's header is "Instance key", written above
            cc, _r = sh.cell(sheet, c, first_row)
            ops.append(_op("set_value", sheet, cell=f"{num_to_col(cc)}{header_new}", before=None, after=f"Column {num_to_col(c)}", note="a text header for the table (Mind needs one over every column of a grid)"))
        notes.append(f"The table had no header row: one is made above it with the column letters as headers (row {header_new} of {_qs(sheet)}).")
    elif n_left:
        # the columns taken into the grid on the left may have no header text: a grid's first row is its header
        for c in range(c1, c1 + n_left):
            if (not insert_col and c == key_col) or not _empty(headers.get(c)):
                continue
            cc, _r = sh.cell(sheet, c, header_row)
            ops.append(_op("set_value", sheet, cell=f"{num_to_col(cc)}{header_new}", before=None, after=f"Column {num_to_col(c)}", note="a text header for a column that touches the table (a grid's first row is its header)"))
    if insert_col:
        key_range = f"{num_to_col(key_new_col)}{first_new}:{num_to_col(key_new_col)}{last_new}"
        ops.append(_op("set_range_values", sheet, ref=key_range, cell=key_range, values=[[_json_value(v)] for v in column_values], before=None, after=f"{rows} {word} value(s)",
                       note=f"the key of every row, copied from column {num_to_col(key_col)} ({word})"))
    plan["split_grid"] = {"title_cell": f"{_qs(sheet)}!{title_cell}", "title": title, "header_cell": f"{_qs(sheet)}!{header_cell}", "inserted_rows": n_rows,
                          "inserted_column": num_to_col(c1) if insert_col else None, "key_range": f"{_qs(sheet)}!{num_to_col(key_new_col)}{first_new}:{num_to_col(key_new_col)}{last_new}",
                          "reused_title": bool(old_title)}
    if old_title:
        notes.append(f"The table already had the title {old_title[2]!r}: it is kept, with {SPLIT_FLAG} added.")

    # --- the count, the loop size, the driver ----------------------------------------------------------
    key_abs = f"{_qs(sheet)}!${num_to_col(key_new_col)}${first_new}:${num_to_col(key_new_col)}${last_new}"
    cnt_c, cnt_r = sh.cell(cnt_sheet, cnt["c1"], cnt["r1"])
    count_formula = f"=COUNTA({key_abs})"
    ops.append(_op("set_formula", cnt_sheet, cell=f"{num_to_col(cnt_c)}{cnt_r}", before=_held(analysis, cnt_sheet, cnt["c1"], cnt["r1"]), after=count_formula,
                   note=f"the count of the loop {le['name']}: in each copy, the rows of that copy (COUNTA of the key column counts the copy's rows in Mind)"))
    avoid: list[tuple[int, int, int, int]] = []
    out_blocks: list[tuple[str, dict[str, Any], dict[str, Any]]] = []  # (sheet, parsed from block, loop plan output)
    for o in loop_plan.get("outputs") or []:
        src = str(o.get("from") or "")
        if not src or "each loop pass" not in str(o.get("when") or ""):
            continue
        f_sheet, f_ref = _split(src)
        fr = parse_ref(f_ref) if f_sheet else None
        if fr and not fr["whole_column"] and not fr["whole_row"]:
            out_blocks.append((f_sheet, fr, o))
            if f_sheet == cnt_sheet:
                avoid.append((fr["r1"], fr["c1"], fr["r2"], fr["c2"]))
    clear_areas = [c for c in loop_plan.get("clears") or [] if c.get("status") == "ready" and c.get("sheet")]
    for c in clear_areas:
        if c["sheet"] == cnt_sheet:
            avoid.append((c["r1"], c["c1"], c["r2"], c["c2"]))
    label_col = max(1, cnt["c1"] - 1)
    free = _free_row(analysis, cnt_sheet, cnt["r1"], label_col, label_col + 3, avoid)
    max_c, max_r = sh.cell(cnt_sheet, label_col + 1, free)
    lab_c, _ = sh.cell(cnt_sheet, label_col, free)
    ops.append(_op("set_value", cnt_sheet, cell=f"{num_to_col(lab_c)}{max_r}", before=None, after=f"Loop size: rows of the biggest {word}", note="label of the loop size cell"))
    ops.append(_op("set_value", cnt_sheet, cell=f"{num_to_col(max_c)}{max_r}", before=None, after=max_count, note=f"the loop size: the most rows one {word} has ({max_count}); every copy loops over at most that many"))
    max_abs = _abs(cnt_sheet, max_c, max_r)
    cnt_abs = _abs(cnt_sheet, cnt_c, cnt_r)
    drv_formula = f'=MM_LOOPINSTANCE("{le["name"]}",{max_abs},{cnt_abs})' + (f"+{rows_before}" if rows_before else "")
    drv_c, drv_r = sh.cell(drv_sheet, drv["c1"], drv["r1"])
    drv_before = _held(analysis, drv_sheet, drv["c1"], drv["r1"])
    ops.append(_op("set_formula", drv_sheet, cell=f"{num_to_col(drv_c)}{drv_r}", before=drv_before, after=drv_formula,
                   note=f"the loop {le['name']}: in each copy it runs 1 to the copy's row count (in Excel it shows {1 + rows_before})"
                   + (f"; the range indexed starts on its header row, so {rows_before} is added to read the data rows" if rows_before else "")))
    dropdown = _dropdown_at(analysis, drv_sheet, drv_ref.replace("$", ""))
    if dropdown:
        ops.append(_op("remove_validation", drv_sheet, cell=drv_ref.replace("$", ""), before=dropdown, after=None, note="Mind treats a cell with a dropdown as an input and drops its formula, so the loop would be lost"))
        notes.append(f"{le.get('driver')} had {dropdown}: it is removed, because Mind drops the formula of a cell with a dropdown.")
    plan["loop"].update({"max_cell": f"{_qs(cnt_sheet)}!{num_to_col(max_c)}{max_r}", "count_cell": f"{_qs(cnt_sheet)}!{num_to_col(cnt_c)}{cnt_r}", "formula": drv_formula, "count_formula": count_formula,
                         "key_cell": None, "driver": f"{_qs(drv_sheet)}!{num_to_col(drv_c)}{drv_r}", "offset": rows_before})

    # --- the output block: the sum over the copy's loop -------------------------------------------------
    written: set[tuple[str, str]] = set()
    block_for_main: dict[str, Any] | None = None
    for f_sheet, fr, o in out_blocks:
        cells = 0
        left_alone: list[str] = []
        example = None
        src_sheet = None
        box: tuple[int, int, int, int] | None = None  # the cells written: (first col, first row, last col, last row), as the workbook is today
        for col, row in _cells_of(f_sheet, fr["ref"]):
            held = _held(analysis, f_sheet, col, row)
            if not (isinstance(held, str) and held.startswith("=")):
                if not _empty(held):
                    left_alone.append(f"{num_to_col(col)}{row}")
                continue
            ref = next((r for r in cell_refs_in_formula(held) if r.get("sheet") and r["sheet"] != f_sheet and r["cells"] == 1), None)
            if ref is None:
                left_alone.append(f"{num_to_col(col)}{row}")
                continue
            rc, rr = sh.cell(ref["sheet"], ref["c1"], ref["r1"])
            formula = f"=MM_RESULT({_qs(ref['sheet'])}!{num_to_col(rc)}{rr})"
            tc, tr = sh.cell(f_sheet, col, row)
            target = f"{num_to_col(tc)}{tr}"
            ops.append(_op("set_formula", f_sheet, cell=target, before=held, after=formula, note=f"the sum over the copy's loop {le['name']} of what {_qs(ref['sheet'])}!{num_to_col(rc)}{rr} holds on each pass (the macro added each pass here)"))
            written.add((f_sheet, target.upper()))
            cells += 1
            src_sheet = src_sheet or ref["sheet"]
            example = example or f"{_qs(f_sheet)}!{target} {formula}"
            box = (min(box[0], col), min(box[1], row), max(box[2], col), max(box[3], row)) if box else (col, row, col, row)
        entry = {"block": f"{_qs(f_sheet)}!{fr['ref']}", "source_sheet": src_sheet, "cells": cells, "example": example, "status": "ready" if cells else "not_built",
                 "why": None if cells else "no cell of the block reads a cell of another sheet: nothing to sum", "left": left_alone[:12], "left_count": len(left_alone)}
        if left_alone:
            notes.append(f"{len(left_alone)} cell(s) of the block {_qs(f_sheet)}!{fr['ref']} keep what they hold (a value, or a formula with no cell of another sheet to sum): {', '.join(left_alone[:6])}{', …' if len(left_alone) > 6 else ''}.")
        plan["outputs"].append(entry)
        plan["counts"]["result_cells"] += cells
        if block_for_main is None and cells and box:
            # the main workbook reads the cells that sum (the macro's copy range may carry header rows above them)
            bc1, br1, bc2, br2 = box
            c1b, r1b = sh.cell(f_sheet, bc1, br1)
            c2b, r2b = sh.cell(f_sheet, bc2, br2)
            head = [_held(analysis, f_sheet, c, br1 - 1) for c in range(bc1, bc2 + 1)] if br1 > 1 else []
            has_head = any(isinstance(v, str) and v.strip() and not v.startswith("=") for v in head)
            headers_b = [(v.strip() if isinstance(v, str) and v.strip() and not v.startswith("=") else f"Column {num_to_col(bc1 + i)}") for i, v in enumerate(head)] if has_head else [f"Column {num_to_col(c)}" for c in range(bc1, bc2 + 1)]
            labels = []
            for i, row in enumerate(range(br1, br2 + 1)):
                v = _held(analysis, f_sheet, bc1 - 1, row) if bc1 > 1 else None
                labels.append(str(_json_value(v)).strip() if not _empty(v) and not (isinstance(v, str) and v.startswith("=")) else f"Row {i + 1}")
            block_for_main = {"sheet": f_sheet, "ref": f"{num_to_col(c1b)}{r1b}:{num_to_col(c2b)}{r2b}", "c1": c1b, "r1": r1b, "c2": c2b, "r2": r2b, "rows": r2b - r1b + 1, "cols": c2b - c1b + 1,
                              "headers": headers_b, "row_labels": labels, "header_row": has_head}
    if not out_blocks:
        notes.append("The macro copies no block out on every pass: the model's results stay where the calculation puts them (read them with MM_RESULT or MM_INSTANCE).")

    # --- the paste area: emptied, as the macro does before a run ------------------------------------------
    from .inventory import open_cached

    try:
        wb = open_cached(analysis)
    except Exception:
        wb = None
    readers: list[str] = []
    block_cells = {(s, f"{num_to_col(col)}{row}") for s, fr, _o in out_blocks for col, row in _cells_of(s, fr["ref"])}
    for c in clear_areas:
        n = over = 0
        if wb is not None and c["sheet"] in wb.sheetnames:
            ws = wb[c["sheet"]]
            for (row, col), cell in sorted(ws._cells.items()):
                if not (c["r1"] <= row <= c["r2"] and c["c1"] <= col <= c["c2"]) or _empty(cell.value):
                    continue
                tc, tr = sh.cell(c["sheet"], col, row)
                target = f"{num_to_col(tc)}{tr}"
                if (c["sheet"], target.upper()) in written:
                    continue
                if n >= MAX_CLEARS:
                    over += 1
                    continue
                ops.append(_op("clear_cell", c["sheet"], cell=target, before=_json_value(cell.value), after=None, note="the macro's paste area: the copies' results live in the block, nothing from an earlier run stays"))
                n += 1
        area = f"{_qs(c['sheet'])}!{num_to_col(c['c1'])}{c['r1']}:{num_to_col(c['c2'])}{c['r2']}"
        plan["clears"].append({"area": area, "cells": n, "status": "ready", "over": over, "why": f"{over} more filled cells were left (limit {MAX_CLEARS})" if over else None})
        plan["counts"]["cleared"] += n
        # cells outside the block that read the paste area are left as they are
        want = c["sheet"].lower()
        for f in wb0.get("formulas") or []:
            if (f["sheet"], str(f["cell"]).upper()) in block_cells:
                continue
            formula = str(f.get("formula") or "")
            if f["sheet"] != c["sheet"] and want not in formula.lower():
                continue  # another sheet's formula reads this area only by naming the sheet
            for r in cell_refs_in_formula(formula):
                if (r.get("sheet") or f["sheet"]) != c["sheet"]:
                    continue
                r1, r2 = r.get("r1") or 1, r.get("r2") or 1_048_576
                cc1, cc2 = r.get("c1") or 1, r.get("c2") or 16_384
                if r1 <= c["r2"] and r2 >= c["r1"] and cc1 <= c["c2"] and cc2 >= c["c1"]:
                    readers.append(f"{_qs(f['sheet'])}!{f['cell']}")
                    break
    if readers:
        notes.append(f"{len(readers)} cell(s) outside the block read the macro's paste area ({', '.join(readers[:6])}{', …' if len(readers) > 6 else ''}): left as they are; in Mind that area is empty.")

    # --- the copy's key, shown ------------------------------------------------------------------------------
    first_key = str(_json_value(key_values[0])) if key_values else ""
    kl_c, _ = sh.cell(cnt_sheet, label_col + 2, free)
    kv_c, _ = sh.cell(cnt_sheet, label_col + 3, free)
    key_formula = f'=MM_INSTANCEKEY("{first_key.replace(chr(34), chr(34) * 2)}")'
    ops.append(_op("set_value", cnt_sheet, cell=f"{num_to_col(kl_c)}{max_r}", before=None, after=f"This copy's {word}", note="label of the key cell"))
    ops.append(_op("set_formula", cnt_sheet, cell=f"{num_to_col(kv_c)}{max_r}", before=None, after=key_formula, note=f"the copy's {word} (in Excel: the first key)"))
    plan["loop"]["key_cell"] = f"{_qs(cnt_sheet)}!{num_to_col(kv_c)}{max_r}"

    # --- after the inserts: the operations that write the table sheet, or read it, wait for them ---------------
    if sh.any:
        for o in ops:
            if o["op"] in ("insert_row", "insert_column", "remove_validation"):
                continue
            if sh.touches(o):
                o["after_inserts"] = True
    plan["operations"] = ops
    plan["operation_count"] = len(ops)
    plan["counts"].update({"key_values": len(key_values), "inserted_rows": n_rows, "inserted_columns": 1 if insert_col else 0, "operations": len(ops)})
    plan["status"] = "ready"

    # --- the main workbook -------------------------------------------------------------------------------
    main["keys"] = [_json_value(v) for v in key_values]
    main["key_header"] = word
    main["model_stem"] = model_stem
    main["block"] = block_for_main
    main["grids"] = [{"title": f"#{word} instances /InstanceKeys", "sheet": "Instances", "rows": len(key_values), "cols": 1,
                      "what": f"the keys: Mind makes one copy of {main['model_file']} per row (the header is the model's file name)"}]
    if block_for_main:
        b = block_for_main
        if len(key_values) * b["rows"] <= MAX_MAIN_CELLS:
            main["grids"].append({"title": f"#Results by {word}", "sheet": "Summary", "rows": len(key_values), "cols": b["rows"] + 1,
                                  "what": f"one row per {word}, one column per row of {_qs(b['sheet'])}!{b['ref']}: the copy's last column of the block (MM_INSTANCE)"})
        else:
            notes.append(f"The grid of results by {word} would hold {len(key_values) * b['rows']} formulas (limit {MAX_MAIN_CELLS}): it is left out of the main workbook.")
        if b["rows"] * b["cols"] <= MAX_MAIN_CELLS:
            # the whole block as plain links to the model: in Mind a plain reference to a copied workbook is
            # the sum over the copies (checked in a real Mind project, 29 October 2026)
            main["grids"].append({"title": "#All copies", "sheet": "Summary", "rows": b["rows"] + 2, "cols": b["cols"] + 1,
                                  "what": f"the whole block {_qs(b['sheet'])}!{b['ref']} as plain links to the model (in Mind: the sum over the {len(key_values)} copies), with a Total row"})
        else:
            notes.append(f"The block {_qs(b['sheet'])}!{b['ref']} has {b['rows'] * b['cols']} cells (limit {MAX_MAIN_CELLS}): the main workbook does not show it as a whole.")
    text = (f"In Mind: one copy of {main['model_file']} per {word} ({len(key_values)} copies, listed in {main['file']}). In each copy the grid {title.split(' /')[0]} keeps the copy's rows, "
            f"{le['name']} loops over them (MM_LOOPINSTANCE, at most {max_count} passes)")
    if plan["counts"]["result_cells"]:
        text += f", and {plan['outputs'][0]['block']} sums each pass's results (MM_RESULT, {plan['counts']['result_cells']} cells)"
    text += f". {main['file']} reads every copy's results (MM_INSTANCE) and adds them up."
    if le.get("macro"):
        text += f" The macro {le['macro']} is not needed in Mind."
    plan["how_it_runs"] = text
    return plan


# --- copies from loops: one copy per combination of the values of one or more loops ------------------
LOOP_COPY_KINDS = ("counter", "list")  # a driver that takes the loop value itself, or the items of a list
MAX_COPIES = 1000
PARAMS_NAME = "Instance values"
_PART_RE = re.compile(r"[A-Za-z0-9.\-]{1,24}")


def loop_options(loop_plan: dict[str, Any]) -> list[dict[str, Any]]:
    """The loops of the Mind loop plan, each with whether copies can be made from it: a ready loop whose
    one driver cell takes the loop value itself (a counter) or the items of a list."""
    out: list[dict[str, Any]] = []
    for le in loop_plan.get("loops") or []:
        usable = le.get("status") == "ready" and le.get("kind") in LOOP_COPY_KINDS and len(le.get("cells") or []) == 1
        why = None
        if le.get("status") != "ready":
            why = le.get("why") or "the loop needs a decision first (Mind loop panel)"
        elif not usable and str(le.get("kind") or "").endswith("_of_table"):
            why = "the driver reads a row or a column of a table: copies are made from a loop that counts, or runs through a list"
        elif not usable:
            why = "the driver takes a value worked out from the loop's counter: copies are made from a loop whose driver takes the value itself, or the items of a list"
        out.append({"name": le.get("name"), "driver": le.get("driver"), "kind": le.get("kind"), "count": le.get("count"), "first": le.get("first"), "last": le.get("last"),
                    "macro": le.get("macro"), "usable": usable, "why": why})
    return out


def _matches(le: dict[str, Any], w: str) -> bool:
    from .loop_builder import _driver_key

    return str(le.get("name") or "").lower() == w.lower() or _driver_key(str(le.get("driver") or "")) == _driver_key(w)


def pick_loop_plan(loop_plans: list[dict[str, Any]], wanted: list[str]) -> dict[str, Any]:
    """Of the Mind loop plans of a model's macros (one each: Excel runs each macro on its own), the one
    whose own loops hold every loop asked for; else the first (the plan then says the loops are of
    different macros)."""
    for lp in loop_plans:
        own = [le for le in lp.get("loops") or [] if not le.get("other_macro")]
        if wanted and all(any(_matches(le, w) for le in own) for w in wanted):
            return lp
    return loop_plans[0]


def loop_options_of(loop_plans: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """loop_options over the loop plans of every macro: each loop as its own macro's plan has it, in the
    order of the first plan's loops."""
    first = loop_plans[0].get("loops") or []
    own: dict[str, dict[str, Any]] = {}
    for lp in loop_plans:
        for o, le in zip(loop_options(lp), lp.get("loops") or []):
            if not le.get("other_macro"):
                own.setdefault(str(le.get("driver")), o)
    return [own.get(str(le.get("driver"))) or o for le, o in zip(first, loop_options(loop_plans[0]))]


def _loop_values(analysis: dict[str, Any], le: dict[str, Any]) -> list[Any] | None:
    """The values the loop's driver takes, in the macro's order: a counter's first..last, a list's items."""
    if le.get("kind") == "counter" and isinstance(le.get("first"), int) and isinstance(le.get("last"), int):
        return list(range(le["first"], le["last"] + 1))
    if le.get("kind") == "list" and le.get("cells"):
        from .inventory import values_cell

        rng = str(le["cells"][0].get("range_fixed") or le["cells"][0].get("range") or "")
        sheet, ref = _split(rng)
        r = parse_ref(ref) if sheet else None
        if not r or r["whole_column"] or r["whole_row"] or r["cells"] > MAX_COPIES:
            return None
        out: list[Any] = []
        for col, row in _cells_of(sheet, ref):
            v = _held(analysis, sheet, col, row)
            if isinstance(v, str) and v.startswith("="):
                try:
                    v = values_cell(analysis, sheet, row, col)
                except Exception:
                    v = None
            if not _empty(v):
                out.append(_json_value(v))
        return out or None
    return None


def _key_part(name: str, value: Any, k: int) -> str:
    """One loop's part of a copy's key: the loop's name and its value ('Scenario0', 'CohortRng2021'), or
    the value's position when the value is long or holds other characters ('GroupName3')."""
    if isinstance(value, bool):
        value = int(value)
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = "" if value is None else str(value)
    return f"{name}{s if _PART_RE.fullmatch(s) else k}"


def _key_parts(name: str, values: list[Any]) -> list[str]:
    """One loop's parts of the copies' keys, one per value, all different: by value when that gives
    different parts, else by position for every value of the loop ('Scenario1', 'Scenario2', ...) --
    a value written by its position could meet a value equal to that position, and a list may repeat a
    value."""
    parts = [_key_part(name, v, k) for k, v in enumerate(values, 1)]
    if len(set(parts)) == len(parts):
        return parts
    return [f"{name}{k}" for k in range(1, len(values) + 1)]


def _same_formula(a: Any, b: Any) -> bool:
    from .formula_utils import STORAGE_PREFIX_RE

    if not (isinstance(a, str) and isinstance(b, str)):
        return a == b
    return STORAGE_PREFIX_RE.sub("", a).replace(" ", "").upper() == STORAGE_PREFIX_RE.sub("", b).replace(" ", "").upper()


def _results_block(analysis: dict[str, Any], model_map: dict[str, Any], results: str | None, notes: list[str]) -> dict[str, Any] | None:
    """The block the main workbook shows for every copy: the range asked for, else the biggest run of
    consecutive formula rows on the model's output sheet (the sheet with the most formulas among the
    sheets the model map names as outputs). Its headers are the text row above it."""
    wb0 = analysis["workbooks"][0]
    sheetnames = {s["name"] for s in wb0.get("sheets") or []}
    chosen_by = "user"
    if results:
        sheet, ref = _split(str(results))
        r = parse_ref(ref) if sheet else None
        if not r or r["whole_column"] or r["whole_row"] or sheet not in sheetnames:
            notes.append(f"The results range {results!r} is not a range of a sheet of the model (Sheet!A1:D10): the main workbook shows no results.")
            return None
        c1, r1, c2, r2 = r["c1"], r["r1"], r["c2"], r["r2"]
    else:
        chosen_by = "outputs"
        sheet = None
        best: tuple[int, int, int, int, int] | None = None  # (cells, first row, last row, first col, last col)
        for s in sorted((model_map.get("outputs") or {}).get("sheets") or [], key=lambda x: -(x.get("formulas") or 0)):
            rows: dict[int, set[int]] = {}
            for f in wb0.get("formulas") or []:
                if f["sheet"] != s.get("sheet"):
                    continue
                p = parse_ref(str(f["cell"]))
                if p:
                    rows.setdefault(p["r1"], set()).add(p["c1"])
            run: list[int] = []
            for row in sorted(rows) + [None]:
                if row is not None and run and row == run[-1] + 1:
                    run.append(row)
                    continue
                if run:
                    cells = sum(len(rows[x]) for x in run)
                    cols = set().union(*(rows[x] for x in run))
                    if best is None or cells > best[0]:
                        best = (cells, run[0], run[-1], min(cols), max(cols))
                run = [row] if row is not None else []
            if best:
                sheet = s["sheet"]
                break
        if not best or not sheet:
            notes.append("The model map names no output sheet with formulas: the main workbook shows no results (name a results range).")
            return None
        _cells_n, r1, r2, c1, c2 = best
    head = [_held(analysis, sheet, c, r1 - 1) for c in range(c1, c2 + 1)] if r1 > 1 else []
    has_head = any(isinstance(v, str) and v.strip() and not v.startswith("=") for v in head)
    headers: list[str] = []
    for i in range(c2 - c1 + 1):
        v = head[i] if has_head else None
        h = _plain(v, f"Column {num_to_col(c1 + i)}") if isinstance(v, str) and not v.startswith("=") else f"Column {num_to_col(c1 + i)}"
        n, base = 2, h
        while h in headers or h in ("Key", "Row"):
            h, n = f"{base} ({n})", n + 1
        headers.append(h)
    return {"sheet": sheet, "ref": f"{num_to_col(c1)}{r1}:{num_to_col(c2)}{r2}", "c1": c1, "r1": r1, "c2": c2, "r2": r2, "rows": r2 - r1 + 1, "cols": c2 - c1 + 1,
            "headers": headers, "row_labels": [f"Row {r}" for r in range(r1, r2 + 1)], "header_row": has_head, "chosen_by": chosen_by}


def _build_loops_plan(analysis: dict[str, Any], model_map: dict[str, Any], loop_plan: dict[str, Any], wanted: list[str], *, results: str | None = None,
                      model_name: str | None = None) -> dict[str, Any]:
    """One copy per combination of the values of the loops asked for (see build_instance_plan).

    In the model: a split grid "#Instance values /InstanceSplit" below the used range of the first
    loop's sheet -- the key and the loops' values, one row per copy (in each copy Mind keeps the copy's
    row); each of those loops' driver cell reads its value from it with =INDEX(<column>,1); the
    macro's other ready loops are written as MM_LOOPs (as the Mind loop plan has them), so they run
    inside every copy; a cell shows the copy's key (MM_INSTANCEKEY). The main workbook lists the keys
    and shows the results block of every copy, one row per copy and row of the block (MM_INSTANCE)."""
    from itertools import product

    from .grids import all_grids, grids_with_flag
    from .loop_builder import ACTION_ID as LOOP_ACTION_ID, _driver_key

    wb0 = analysis["workbooks"][0]
    stem = _safe_stem(wb0.get("file_name") or Path(str(wb0.get("copy_path") or "model")).stem)
    model_stem = _safe_stem(model_name) if model_name else f"{stem}_instances"
    ext = str(wb0.get("file_type") or Path(str(wb0.get("copy_path") or "")).suffix.lstrip(".") or "xlsx").lower()
    main: dict[str, Any] = {"file": f"{stem}_Main.xlsx", "model_file": f"{model_stem}.{ext}", "keys": [], "grids": [], "layout": "by_key_rows"}
    plan: dict[str, Any] = {"found": bool(loop_plan.get("found")), "mode": "loops", "status": "not_built", "notes": [], "loop": None, "table": None, "key": None, "split_grid": None,
                            "copies_from": [], "kept_loops": [], "results": None, "outputs": [], "clears": [], "operations": [], "operation_count": 0,
                            "counts": {"key_values": 0, "result_cells": 0, "cleared": 0, "inserted_rows": 0}, "main": main, "how_it_runs": None}
    notes = plan["notes"]
    if not loop_plan.get("found"):
        notes.append("No run loop in the macros: nothing to build.")
        return plan
    entries = list(loop_plan.get("loops") or [])
    chosen: list[dict[str, Any]] = []
    for w in wanted:
        le = next((e for e in entries if str(e.get("name") or "").lower() == w.lower() or _driver_key(str(e.get("driver") or "")) == _driver_key(w)), None)
        if le is None:
            plan["status"] = "needs_decision"
            notes.append(f"No loop is called {w!r} (the loops: {', '.join(str(e.get('name')) for e in entries) or 'none'}).")
            return plan
        if le not in chosen:
            chosen.append(le)
    runs = sorted({str(le.get("run") or le.get("macro") or "") for le in chosen if le.get("run") or le.get("macro")})
    if len(runs) > 1:
        plan["status"] = "needs_decision"
        notes.append(f"The loops {', '.join(str(le['name']) for le in chosen)} belong to different macros ({', '.join(runs)}). Excel runs each macro on its own, "
                     "so copies are made from the loops of one macro at a time.")
        return plan
    opts = {o["name"]: o for o in loop_options(loop_plan)}
    values: list[list[Any]] = []
    for le in chosen:
        o = opts.get(le.get("name")) or {}
        vals = _loop_values(analysis, le) if o.get("usable") else None
        if not vals:
            plan["status"] = "needs_decision"
            notes.append(f"Copies cannot be made from the loop {le.get('name')}: {o.get('why') or 'its values could not be read'}.")
            return plan
        values.append(vals)
    n = 1
    for vals in values:
        n *= len(vals)
    if n > MAX_COPIES:
        plan["status"] = "needs_decision"
        notes.append(f"{' x '.join(str(len(v)) for v in values)} = {n} copies, above the limit of {MAX_COPIES}: keep one of these loops a loop.")
        return plan
    names = [str(le["name"]) for le in chosen]
    parts = [_key_parts(names[i], vals) for i, vals in enumerate(values)]
    combos = list(product(*[list(zip(parts[i], vals)) for i, vals in enumerate(values)]))
    keys = ["_".join(p for p, _v in combo) for combo in combos]
    rows_values = [[v for _p, v in combo] for combo in combos]
    if len(set(keys)) != len(keys):  # cannot happen (each loop's parts are unique): a guard, Mind needs one row per copy
        plan["status"] = "needs_decision"
        notes.append("Two copies would get the same key: the copies cannot be told apart.")
        return plan

    # --- where the split grid goes: below the used range of the first loop's sheet ---------------------
    p_sheet, p_ref = _split(str(chosen[0]["driver"]))
    first_drv = parse_ref(p_ref) if p_sheet else None
    if not first_drv:
        plan["status"] = "needs_decision"
        notes.append(f"The driver cell of the loop {names[0]} ({chosen[0].get('driver')}) is not one cell.")
        return plan
    left = max(1, min(parse_ref(_split(str(le["driver"]))[1])["c1"] for le in chosen if _split(str(le["driver"]))[0] == p_sheet) - 1)
    right = left + len(chosen)
    _max_col, max_row = _sheet_max(analysis, p_sheet)
    top = max_row + 3  # two empty rows below the used range
    while top < max_row + 400 and not all(_row_empty(analysis, p_sheet, r, max(1, left - 1), right + 1) for r in range(top - 1, top + 5 + len(keys))):
        top += 1
    if top + 4 + len(keys) > 1_048_576:
        plan["status"] = "needs_decision"
        notes.append(f"{_qs(p_sheet)} has no free rows below its used range for the grid of the copies' values.")
        return plan
    taken = {str(g.get("name") or "").lower() for g in all_grids(analysis)}
    grid_name = PARAMS_NAME if PARAMS_NAME.lower() not in taken else f"{PARAMS_NAME} {stem}"
    title = f"#{grid_name} {SPLIT_FLAG}"
    title_row, header_row, first_row, last_row = top + 2, top + 3, top + 4, top + 3 + len(keys)
    L, R = num_to_col(left), num_to_col(right)
    ops: list[dict[str, Any]] = []
    ops.append(_op("set_value", p_sheet, cell=f"{L}{top}", before=None, after="This copy's key", note="label of the key cell"))
    ops.append(_op("set_formula", p_sheet, cell=f"{num_to_col(left + 1)}{top}", before=None, after=f'=MM_INSTANCEKEY("{keys[0]}")', note="the copy's key (in Excel: the first key)"))
    ops.append(_op("set_value", p_sheet, cell=f"{L}{title_row}", before=None, after=title,
                   note="the split grid's title: in each copy Mind keeps the one row whose first column holds the copy's key"))
    block = [[KEY_HEADER, *names]] + [[k, *vals] for k, vals in zip(keys, rows_values)]
    ops.append(_op("set_range_values", p_sheet, ref=f"{L}{header_row}:{R}{last_row}", cell=f"{L}{header_row}:{R}{last_row}", values=block, before=None,
                   after=f"{len(keys)} copies: the key and the value of {', '.join(names)}", note="one row per copy: its key, then the value each loop's driver takes in it"))

    # --- each chosen loop's driver reads its value from the copy's row ------------------------------------
    for i, le in enumerate(chosen):
        d_sheet, d_ref = _split(str(le["driver"]))
        d = parse_ref(d_ref)
        col = num_to_col(left + 1 + i)
        formula = f"=INDEX({_qs(p_sheet)}!${col}${first_row}:${col}${last_row},1)"
        before = _held(analysis, d_sheet, d["c1"], d["r1"])
        ops.append(_op("set_formula", d_sheet, cell=f"{num_to_col(d['c1'])}{d['r1']}", before=before, after=formula,
                       note=f"{le['name']} of this copy: the one row Mind keeps in the copy (in Excel: the first copy's value, {rows_values[0][i]})"))
        dropdown = _dropdown_at(analysis, d_sheet, d_ref.replace("$", ""))
        if dropdown:
            ops.append(_op("remove_validation", d_sheet, cell=d_ref.replace("$", ""), before=dropdown, after=None, note="Mind treats a cell with a dropdown as an input and drops its formula"))
            notes.append(f"{le['driver']} had {dropdown}: it is removed, because Mind drops the formula of a cell with a dropdown.")
        plan["copies_from"].append({"name": le["name"], "driver": le["driver"], "kind": le.get("kind"), "values": values[i], "count": len(values[i]), "formula": formula, "macro": le.get("macro")})

    # --- the macro's other loops stay loops inside every copy ---------------------------------------------
    chosen_names = set(names)
    kept_ops = 0
    other_macros: dict[str, list[str]] = {}
    for le in entries:
        if le["name"] in chosen_names:
            continue
        if le.get("other_macro"):
            other_macros.setdefault(str(le.get("run") or le.get("macro")), []).append(str(le["name"]))
            continue
        if le.get("status") != "ready":
            notes.append(f"The loop {le['name']} is not written ({le.get('why') or 'it needs a decision on the Mind loop panel'}): in Mind its driver keeps one value.")
            continue
        mine = [o for o in loop_plan.get("operations") or [] if o.get("action_id") == LOOP_ACTION_ID and o.get("loop") == le["name"]]
        todo = [o for o in mine if not (o["op"] == "set_formula" and _same_formula(o.get("before"), o.get("after")))]
        ops.extend(todo)
        kept_ops += len(todo)
        plan["kept_loops"].append({"name": le["name"], "driver": le.get("driver"), "count": le.get("count"), "written": bool(todo), "already": bool(mine) and not todo})
    for m, lnames in other_macros.items():
        notes.append(f"The loops of the macro {m} ({', '.join(lnames)}) are not part of these copies: Excel runs that macro on its own, with these drivers fixed.")
    # the areas the macro clears before a run, as the Mind loop plan clears them (asked for, or read by a block)
    cleared = [o for o in loop_plan.get("operations") or [] if o.get("action_id") == LOOP_ACTION_ID and o.get("op") == "clear_cell" and not o.get("loop")]
    if cleared:
        ops.extend(cleared)
        plan["counts"]["cleared"] = len(cleared)
        plan["clears"] = [c for c in loop_plan.get("clears") or [] if c.get("included")]
        notes.append(f"{len(cleared)} cell(s) the macro clears before a run are emptied, as the Mind loop plan does: what an earlier run left there would stay in every copy.")
    # anything else that still names a chosen loop (an MM_RESULT of an earlier loop build) loses its loop
    drivers = {(_split(str(le["driver"]))[0], _split(str(le["driver"]))[1].replace("$", "").upper()) for le in chosen}
    pat = re.compile(r'MM_(?:LOOP|RESULT)\w*\s*\([^"]*"(' + "|".join(re.escape(x) for x in names) + r')"', re.I)
    stale = [f"{_qs(f['sheet'])}!{f['cell']}" for f in wb0.get("formulas") or [] if pat.search(str(f.get("formula") or "")) and (f["sheet"], str(f["cell"]).upper()) not in drivers]
    if stale:
        notes.append(f"{len(stale)} cell(s) still name the loop {' / '.join(names)} in an MM_ function ({', '.join(stale[:6])}{', …' if len(stale) > 6 else ''}): "
                     "that loop is a copy now, so read those values in the main workbook instead.")
    inputs = [g for g in grids_with_flag(analysis, "input")]
    if inputs:
        g0 = _plain(str(inputs[0].get("title") or inputs[0].get("name") or "").split("/")[0], "Input")
        notes.append(f"{len(inputs)} input grid(s) (/Input) stay in the copied model: every copy starts with the values they hold now. In Mind a file named "
                     f"'<key>-<grid>.csv' (for example '{keys[0]}-{g0}.csv') changes one copy only; a file without a key in its name is not matched.")
    if any(le.get("macro") for le in chosen):
        notes.append("Every copy computes. If the macro skips some passes, or writes out only some of them, read only the copies it would keep.")

    # --- the main workbook: the keys, and the results block of every copy ---------------------------------
    word = " and ".join(names)
    res = _results_block(analysis, model_map, results, notes)
    main.update(keys=keys, key_header=word, model_stem=model_stem, block=res)
    main["grids"] = [{"title": f"#{word} instances /InstanceKeys", "sheet": "Instances", "rows": len(keys), "cols": 1,
                      "what": f"the keys: Mind makes one copy of {main['model_file']} per row (the header is the model's file name)"}]
    if res:
        cells = len(keys) * res["rows"] * res["cols"]
        if cells <= MAX_MAIN_CELLS:
            main["grids"].append({"title": f"#Results by {word}", "sheet": "Summary", "rows": len(keys) * res["rows"], "cols": res["cols"] + 2,
                                  "what": f"one row per copy and row of {_qs(res['sheet'])}!{res['ref']}: what the copy computes there (MM_INSTANCE by key)"})
            plan["counts"]["result_cells"] = cells
        else:
            notes.append(f"The results of every copy would take {cells} formulas (limit {MAX_MAIN_CELLS}): name a smaller results range.")
        plan["results"] = {"block": f"{_qs(res['sheet'])}!{res['ref']}", "rows": res["rows"], "cols": res["cols"], "chosen_by": res["chosen_by"]}
        notes.append(f"A plain link to the copied model would add every copy up ({word} together), so the main workbook shows each copy's own values only.")

    plan["key"] = {"column": L, "header": KEY_HEADER, "distinct": len(keys), "counts": {}, "candidates": [], "chosen_by": "loops", "missing": 0, "values": keys}
    plan["split_grid"] = {"title_cell": f"{_qs(p_sheet)}!{L}{title_row}", "title": title, "header_cell": f"{_qs(p_sheet)}!{L}{header_row}", "inserted_rows": 0, "inserted_column": None,
                          "key_range": f"{_qs(p_sheet)}!{L}{first_row}:{L}{last_row}", "reused_title": False, "key_cell": f"{_qs(p_sheet)}!{num_to_col(left + 1)}{top}"}
    plan["operations"] = ops
    plan["operation_count"] = len(ops)
    plan["counts"].update({"key_values": len(keys), "inserted_columns": 0, "operations": len(ops), "kept_loop_operations": kept_ops})
    plan["status"] = "ready"
    text = (f"In Mind: one copy of {main['model_file']} per combination of {word} ({len(keys)} copies, listed in {main['file']}). Each copy reads its own "
            f"{word} from the grid {grid_name} ({_qs(p_sheet)}!{L}{title_row}), which keeps only the copy's row.")
    kept = [k["name"] for k in plan["kept_loops"]]
    if kept:
        text += f" {', '.join(kept)} {'stays a loop' if len(kept) == 1 else 'stay loops'} inside every copy."
    if res and len(main["grids"]) > 1:
        text += f" {main['file']} shows {_qs(res['sheet'])}!{res['ref']} of every copy (MM_INSTANCE)."
    macros = sorted({le["macro"] for le in chosen if le.get("macro")})
    if macros:
        text += f" The macro {', '.join(macros)} is not needed in Mind."
    plan["how_it_runs"] = text
    return plan


# --- the main workbook -------------------------------------------------------------------------------
def write_main_workbook(plan: dict[str, Any], model_path: str | Path, main_path: str | Path) -> dict[str, Any]:
    """Write the main workbook next to the model, through Excel: the keys grid, the results by key,
    and the whole block as plain links (in Mind: the sum over the copies). Keys are written as they
    are (a text key stays text). `model_path` must be named plan["main"]["model_file"] and sit in the folder
    `main_path` goes to, so Excel stores the references as a file of the same folder."""
    from .excel_com import close_quietly, excel_session, excel_text, open_workbook

    main = plan.get("main") or {}
    # the long form of both paths: Excel stores the link as the bare file name only when the model and the
    # main workbook sit in the same folder spelled the same way (a short 8.3 name such as SOMEUS~1 differs)
    model_path, main_path = Path(model_path).resolve(), Path(main_path).resolve()
    if model_path.name != main.get("model_file"):
        raise ValueError(f"the model must be named {main.get('model_file')!r} (it is {model_path.name!r}): the keys grid names that file")
    if not model_path.is_file():
        raise FileNotFoundError(str(model_path))
    if model_path.parent.resolve() != main_path.parent.resolve():
        raise ValueError("the main workbook must be written in the model's folder")
    keys = list(main.get("keys") or [])
    if not keys:
        raise ValueError("no keys: nothing to write")
    word = main.get("key_header") or "key"
    stem = main.get("model_stem") or model_path.stem
    block = main.get("block")
    titles = {g["title"] for g in main.get("grids") or []}
    formulas = 0
    grids = 0
    main_path.parent.mkdir(parents=True, exist_ok=True)
    if main_path.exists():
        main_path.unlink()
    with excel_session(mm_addin=True) as excel:
        try:
            calc_mode = excel.Calculation
            excel.Calculation = -4135  # manual while writing
        except Exception:
            calc_mode = None
        model = wb = None
        try:
            model = open_workbook(excel, model_path, read_only=True)  # open, so '[model]Sheet'! resolves to a file of the same folder
            excel.SheetsInNewWorkbook = 1
            wb = excel.Workbooks.Add()
            inst = wb.Worksheets(1)
            inst.Name = "Instances"
            inst.Range("B2").Value2 = f"One copy of {stem} per {word}. Mind copies {model_path.name} once per key below; this file itself is not copied."
            inst.Range("B4").Value2 = f"#{word} instances /InstanceKeys"
            inst.Range("B5").Value2 = stem
            # the keys as they are: a text key stays text ("001" is not 1), through excel_text, as in the model's key column
            inst.Range(f"B6:B{5 + len(keys)}").Value2 = tuple((excel_text(k),) for k in keys)
            inst.Columns("B:B").ColumnWidth = 28
            grids += 1
            summ = wb.Worksheets.Add(After=inst)
            summ.Name = "Summary"
            row = 2
            if block and main.get("layout") == "by_key_rows":
                # copies made from loops: one row per copy and row of the block, each cell what that copy
                # computes there (MM_INSTANCE by key); no sum over the copies (it would add scenarios up)
                b = block
                summ.Range(f"B{row}").Value2 = (f"What each copy computes in {b['sheet']}!{b['ref']}, one row per copy and row of that block. "
                                                "In Excel every row shows the one copy Excel has (the first key).")
                row += 2
                if f"#Results by {word}" in titles:
                    summ.Range(f"B{row}").Value2 = f"#Results by {word}"
                    last = num_to_col(3 + b["cols"])
                    summ.Range(f"B{row + 1}:{last}{row + 1}").Value2 = (("Key", "Row") + tuple(excel_text(str(h)) for h in b["headers"]),)
                    first = row + 2
                    n_rows = len(keys) * b["rows"]
                    summ.Range(f"B{first}:C{first + n_rows - 1}").Value2 = tuple((excel_text(k), excel_text(str(lab))) for k in keys for lab in b["row_labels"])
                    rows = []
                    for i in range(n_rows):
                        m = i % b["rows"]
                        r = first + i
                        rows.append(tuple(f"=MM_INSTANCE({_xref(model_path.name, b['sheet'], f'${num_to_col(b['c1'] + c)}${b['r1'] + m}')},$B{r})" for c in range(b["cols"])))
                    summ.Range(f"D{first}:{last}{first + n_rows - 1}").Formula = tuple(rows)
                    formulas += n_rows * b["cols"]
                    grids += 1
            elif block:
                b = block
                summ.Range(f"B{row}").Value2 = f"Each copy's results ({b['sheet']}!{b['ref']} of that copy) by {word}, then all copies added up. In Excel every grid shows the one copy Excel has."
                row += 2
                last_col = num_to_col(b["c2"])
                heads = ("Row",) + tuple(excel_text(str(h)) for h in b["headers"])
                labels = tuple((excel_text(str(x)),) for x in b["row_labels"])
                if f"#Results by {word}" in titles:
                    summ.Range(f"B{row}").Value2 = f"#Results by {word}"
                    summ.Range(f"B{row + 1}").Value2 = word
                    summ.Range(f"C{row + 1}:{num_to_col(2 + b['rows'])}{row + 1}").Value2 = tuple(excel_text(str(x)) for x in b["row_labels"])
                    summ.Range(f"B{row + 2}:B{row + 1 + len(keys)}").Value2 = tuple((excel_text(k),) for k in keys)  # the keys as values, apart from the formulas
                    rows = []
                    for i, k in enumerate(keys):
                        r = row + 2 + i
                        cells = []
                        for m in range(b["rows"]):
                            source = _xref(model_path.name, b["sheet"], f"${last_col}${b['r1'] + m}")  # the block row's last column, in the copy
                            cells.append(f"=MM_INSTANCE({source},$B{r})")
                        rows.append(tuple(cells))
                    summ.Range(f"C{row + 2}:{num_to_col(2 + b['rows'])}{row + 1 + len(keys)}").Formula = tuple(rows)
                    formulas += len(keys) * b["rows"]
                    grids += 1
                    row = row + 2 + len(keys) + 2
                if "#All copies" in titles:
                    # the whole block as plain links to the model: in Mind a plain reference to a copied workbook is
                    # the sum over the copies (checked in a real Mind project, 29 October 2026); a Total row under it
                    summ.Range(f"B{row}").Value2 = "#All copies"
                    summ.Range(f"B{row + 1}:{num_to_col(2 + b['cols'])}{row + 1}").Value2 = heads
                    summ.Range(f"B{row + 2}:B{row + 1 + b['rows']}").Value2 = labels
                    links = tuple(tuple("=" + _xref(model_path.name, b["sheet"], f"{num_to_col(b['c1'] + c)}{b['r1'] + m}") for c in range(b["cols"])) for m in range(b["rows"]))
                    summ.Range(f"C{row + 2}:{num_to_col(2 + b['cols'])}{row + 1 + b['rows']}").Formula = links
                    tr = row + 2 + b["rows"]
                    summ.Range(f"B{tr}").Value2 = "Total"
                    summ.Range(f"C{tr}:{num_to_col(2 + b['cols'])}{tr}").FormulaR1C1 = f"=SUM(R[-{b['rows']}]C:R[-1]C)"
                    formulas += b["rows"] * b["cols"] + b["cols"]
                    grids += 1
            else:
                summ.Range("B2").Value2 = "The macro copies no block out on every pass: read the copies' results with MM_INSTANCE."
            summ.Columns("B:B").ColumnWidth = 24
            if calc_mode is not None:
                try:
                    excel.Calculation = calc_mode
                except Exception:
                    pass
            excel.DisplayAlerts = False
            wb.SaveAs(str(main_path), FileFormat=51)  # xlOpenXMLWorkbook
        finally:
            close_quietly(wb)
            close_quietly(model)
            wb = model = None
    return {"grids": grids, "formulas": formulas, "path": str(main_path)}


def zip_for_mind(model_path: str | Path, main_path: str | Path, zip_path: str | Path) -> Path:
    """Both files in one zip, stored flat (the two files the user uploads to one Mind project)."""
    zip_path = Path(zip_path)
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in (Path(model_path), Path(main_path)):
            z.write(p, p.name)
    return zip_path


# --- words for a person -----------------------------------------------------------------------------
def describe(plan: dict[str, Any]) -> str:
    lines = [f"status: {plan['status']}"]
    lp = plan.get("loop")
    if lp:
        lines.append(f"loop {lp['name']}  driver {lp.get('driver')}  count cell {lp.get('count_cell')}  loop size cell {lp.get('max_cell')}")
        if lp.get("formula"):
            lines.append(f"   {lp['formula']}")
    t = plan.get("table")
    if t:
        lines.append(f"table {t['sheet']}!{t['ref']}  {t['rows']} rows x {t['columns']} columns  header row {t.get('header_row')}")
    k = plan.get("key")
    if k and k.get("column"):
        lines.append(f"key column {k['column']} ({k.get('header')!r}, {k['distinct']} values, chosen by {k['chosen_by']}): " + ", ".join(f"{v}: {n}" for v, n in list(k["counts"].items())[:12]) + (" …" if len(k["counts"]) > 12 else ""))
    elif k:
        lines.append("key: " + (", ".join(f"{c['column']} {c['header']!r} ({c['distinct']})" for c in k.get("candidates") or []) or "no candidate"))
    for c in plan.get("copies_from") or []:
        lines.append(f"copies from loop {c['name']} ({c.get('count')} values)  driver {c.get('driver')}  {c.get('formula')}")
    if plan.get("kept_loops"):
        lines.append("loops kept inside every copy: " + ", ".join(k["name"] for k in plan["kept_loops"]))
    sg = plan.get("split_grid")
    if sg:
        lines.append(f"split grid: title {sg['title']!r} in {sg['title_cell']}, header {sg['header_cell']}, keys in {sg['key_range']}  (rows inserted: {sg['inserted_rows']}, column inserted: {sg['inserted_column']})")
    for o in plan.get("outputs") or []:
        lines.append(f"output {o['block']}  [{o['status']}]  {o['cells']} MM_RESULT cells" + (f"  e.g. {o['example']}" if o.get("example") else "") + (f"  {o['why']}" if o.get("why") else ""))
    for c in plan.get("clears") or []:
        lines.append(f"clear {c['area']}  {c['cells']} filled cells")
    m = plan.get("main") or {}
    if m.get("grids"):
        lines.append(f"main workbook {m['file']} (model file {m['model_file']}): " + "; ".join(f"{g['title']} {g['rows']}x{g['cols']}" for g in m["grids"]))
    lines.append(f"operations: {plan.get('operation_count', 0)}")
    if plan.get("how_it_runs"):
        lines.append(plan["how_it_runs"])
    for n in plan.get("notes") or []:
        lines.append(n)
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    import argparse
    import tempfile

    from .inventory import build_analysis
    from .model_map import build_model_map

    ap = argparse.ArgumentParser(description="The macro's run loop as Mind instances: one copy of the model per key, and the main workbook.")
    ap.add_argument("workbook")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--key", default=None, help="the key column: a header text or a column letter of the model-point table")
    ap.add_argument("--keys", choices=("present", "all"), default="present", help="present = keys that have rows; all = every key of the list the macro runs through")
    ap.add_argument("--name", default=None, help="the model's file stem as it will be uploaded (default <workbook>_instances)")
    ap.add_argument("--loops", default=None, help="copies from these loops (comma separated names or driver cells) instead of a key column")
    ap.add_argument("--results", default=None, help="with --loops: the block the main workbook shows for every copy (Sheet!A1:B9)")
    ap.add_argument("--work", default=None)
    a = ap.parse_args(argv)
    work = Path(a.work) if a.work else Path(tempfile.mkdtemp(prefix="mind_instances_"))
    analysis = build_analysis(Path(a.workbook), work, Path(a.workbook).stem)
    loops = [x.strip() for x in a.loops.split(",") if x.strip()] if a.loops else None
    plan = build_instance_plan(analysis, build_model_map(analysis), key=a.key, keys=a.keys, model_name=a.name, loops=loops, results=a.results)
    out = sys.__stdout__
    try:
        out.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    out.write((json.dumps(plan, indent=1, default=str, ensure_ascii=False) if a.json else describe(plan)) + "\n")
    out.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
