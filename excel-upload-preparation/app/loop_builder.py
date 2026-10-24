"""Mind loop -- the macro's run loop written as Mind formulas.

The Model map (app.model_map) says which cell a macro sets on every pass (the driver), what it
sets it from (a row of a table, an item of a list, a counter), the cells that bound the run, and
where each pass's results are written. This module turns that into the Mind construct that runs
through the same values without the macro:

- every driver cell gets ``=MM_LOOP("Name", range)``: the loop, one value per row of the table
  (or item of the list), so everything that reads the driver is computed for every value;
- the block the macro writes on every pass gets ``=MM_RESULT(source, "Name", k)`` row by row:
  row k reads the source cells for the k-th loop value, as the macro's k-th pass wrote them;
- the area the macro clears before a run is emptied of what an earlier run left there.

Everything comes from the map and the workbook. Where the macro does something these two
functions cannot say (a block whose place depends on a counter the macro keeps, a counted loop
with no cell holding the count, a block that depends on several nested loops) the plan says so
and builds nothing for it. Loop names are the driver's defined name, else the source sheet's
name; both are overridable.

    python -m app.loop_builder <workbook.xlsm> [--json] [--name N] [--first a] [--last b]
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from .formula_utils import col_to_num, num_to_col, parse_ref

ACTION_ID = "mind_loop"
RULE_ID = "LOOP-001"
MAX_RESULT_CELLS = 60_000  # MM_RESULT formulas one plan writes at most (items x cells of the block)
MAX_CLEARS = 5_000

_PH_RE = re.compile(r"<([A-Za-z_]\w*)>")
_VAR_ASSIGN_RE = re.compile(r"^\s*(?:Dim\s+[^:]+:\s*)?([A-Za-z_]\w*)\s*=")
_OFFSET_RE = re.compile(r"\.Offset\(\s*(?P<r>[^,]+?)\s*,\s*(?P<c>[^)]+?)\s*\)", re.IGNORECASE)
_RANGE_NAME_RE = re.compile(r"""Range\(\s*"(?P<name>[^"]+)"\s*\)(?:\.Value2?)?""", re.IGNORECASE)
_SAFE_EXPR_RE = re.compile(r"^[\d\s+\-*/().]+$")


# --- small helpers --------------------------------------------------------------------------------
def loop_name_from(text: str) -> str:
    """A loop name Mind accepts everywhere: letters, digits and underscores."""
    name = re.sub(r"[^A-Za-z0-9_]+", "_", str(text or "")).strip("_")
    return name or "Loop"


def _qs(sheet: str) -> str:
    """The sheet name as a formula writes it: quoted unless it is a plain word."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


def _split(cell: str) -> tuple[str | None, str]:
    """'Sheet'!A1:B2 -> (Sheet, A1:B2)."""
    if "!" not in cell:
        return None, cell
    sheet, ref = cell.rsplit("!", 1)
    return sheet.strip().strip("'").replace("''", "'"), ref.strip()


def _cells_of(sheet: str, ref: str) -> list[tuple[int, int]]:
    """(col, row) of every cell of a literal range, row-major."""
    r = parse_ref(f"'{sheet}'!{ref}")
    if not r or r["whole_column"] or r["whole_row"]:
        return []
    return [(c, row) for row in range(r["r1"], r["r2"] + 1) for c in range(r["c1"], r["c2"] + 1)]


def _reader_of_area(wb: Any, block: str, area: dict[str, Any]) -> str | None:
    """The first cell of `block` ('Sheet'!A1:B2, an output the macro copies) whose formula reads a cell of
    `area` (a cleared area: sheet, r1, c1, r2, c2), or None. Such a block adds what an earlier run left
    there to every loop value: the macro clears it first, Mind would not."""
    from .formula_utils import cell_refs_in_formula

    sheet, ref = _split(block)
    if wb is None or sheet is None or sheet not in wb.sheetnames:
        return None
    ws = wb[sheet]
    for col, row in _cells_of(sheet, ref):
        f = ws.cell(row=row, column=col).value
        if not isinstance(f, str) or not f.startswith("="):
            continue
        for r in cell_refs_in_formula(f):
            if (r.get("sheet") or sheet) != area["sheet"]:
                continue
            r1, r2 = r.get("r1") or 1, r.get("r2") or 1_048_576
            c1, c2 = r.get("c1") or 1, r.get("c2") or 16_384
            if r1 <= area["r2"] and r2 >= area["r1"] and c1 <= area["c2"] and c2 >= area["c1"]:
                return f"{_qs(sheet)}!{num_to_col(col)}{row}"
    return None


def _held(analysis: dict[str, Any], sheet: str, col: int, row: int) -> Any:
    from .inventory import cell_value

    try:
        return cell_value(analysis, sheet, row, col)
    except Exception:
        return None


def _in_sqref(sqref: str, col: int, row: int) -> bool:
    for part in str(sqref or "").replace("$", "").split():
        r = parse_ref(part)
        if r and r["c1"] <= col <= r["c2"] and r["r1"] <= row <= r["r2"]:
            return True
    return False


def _x14_dropdowns(copy_path: str, sheet: str) -> list[tuple[str, str]]:
    """(sqref, source) of the dropdowns Excel keeps in the sheet's extension block (a list taken from
    another sheet): openpyxl does not read those."""
    import posixpath
    import zipfile

    try:
        with zipfile.ZipFile(copy_path) as z:
            wbx = z.read("xl/workbook.xml").decode("utf-8", "replace")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
            rid = None
            for m in re.finditer(r"<sheet\b[^>]*>", wbx):
                tag = m.group(0)
                nm = re.search(r'name="([^"]*)"', tag)
                if nm and nm.group(1).replace("&amp;", "&").replace("&apos;", "'") == sheet:
                    rm = re.search(r'r:id="([^"]*)"', tag)
                    rid = rm.group(1) if rm else None
                    break
            if not rid:
                return []
            target = None
            for m in re.finditer(r"<Relationship\b[^>]*>", rels):
                if f'Id="{rid}"' in m.group(0):
                    tm = re.search(r'Target="([^"]*)"', m.group(0))
                    target = tm.group(1) if tm else None
            if not target:
                return []
            path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            xml = z.read(path).decode("utf-8", "replace")
    except Exception:
        return []
    out = []
    for m in re.finditer(r"<x14:dataValidation\b.*?</x14:dataValidation>", xml, re.S):
        block = m.group(0)
        if 'type="list"' not in block:
            continue
        sq = re.search(r"<xm:sqref>([^<]*)</xm:sqref>", block)
        f = re.search(r"<xm:f>([^<]*)</xm:f>", block)
        if sq:
            out.append((sq.group(1), f.group(1) if f else ""))
    return out


def _dropdown_at(analysis: dict[str, Any], sheet: str, ref: str) -> str | None:
    """The list dropdown (data validation) on a cell, if it has one. Mind reads a cell with a
    dropdown as an input and drops its formula: an MM_LOOP written there is lost (seen in Mind on
    2026-10-16 with a model's two driver cells and a probe workbook)."""
    from .inventory import open_cached

    r = parse_ref(ref)
    if not r:
        return None
    col, row = r["c1"], r["r1"]
    try:
        ws = open_cached(analysis)[sheet]
        for dv in getattr(ws.data_validations, "dataValidation", None) or []:
            if _in_sqref(str(dv.sqref), col, row):
                return f"a dropdown list ({dv.formula1})" if dv.type == "list" else f"a data validation ({dv.type})"
    except Exception:
        pass
    try:
        copy_path = analysis["workbooks"][0]["copy_path"]
    except Exception:
        return None
    for sqref, src in _x14_dropdowns(copy_path, sheet):
        if _in_sqref(sqref, col, row):
            return f"a dropdown list ({src})"
    return None


def _number_at(analysis: dict[str, Any], sheet: str, ref: str) -> tuple[Any, str | None]:
    """A cell's number: its constant, else the value Excel last stored for its formula."""
    from .inventory import values_cell

    r = parse_ref(ref)
    if not r:
        return None, None
    held = _held(analysis, sheet, r["c1"], r["r1"])
    if isinstance(held, str) and held.startswith("="):
        try:
            v = values_cell(analysis, sheet, r["r1"], r["c1"])
        except Exception:
            v = None
        return v, held
    return held, None


def _as_int(v: Any) -> int | None:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        try:
            return _as_int(float(v.strip()))
        except ValueError:
            return None
    return None


def _sheet_max(analysis: dict[str, Any], sheet: str) -> tuple[int, int]:
    """(max column, max row) of a sheet's used range, from the inventory's dimensions."""
    for s in analysis["workbooks"][0].get("sheets", []):
        if s["name"] == sheet:
            dims = s.get("dimensions") or s.get("dims")
            r = parse_ref(str(dims)) if dims else None
            if r and not r["whole_column"] and not r["whole_row"]:
                return r["c2"], r["r2"]
    return 16384, 1048576


def _eval_shift(expr: str, var: str, value: int) -> int | None:
    """`i - 1` at i = value, for the small arithmetic VBA writes inside Offset()."""
    text = re.sub(rf"\b{re.escape(var)}\b", str(value), expr)
    if not _SAFE_EXPR_RE.fullmatch(text):
        return None
    try:
        return int(eval(text, {"__builtins__": {}}, {}))  # noqa: S307 -- digits and + - * / ( ) only, checked above
    except Exception:
        return None


def _tight(expr: str) -> str:
    """The expression without spaces, except inside quoted sheet names."""
    parts = expr.split("'")
    return "'".join(part if i % 2 else part.replace(" ", "") for i, part in enumerate(parts))


def _bound_var(bound: dict[str, Any]) -> str | None:
    """The VBA variable a bound was read into (`num_debut = Worksheets(...)`) -- the macro's
    own clearing / writing statements use it, so the map's placeholders may carry it."""
    m = _VAR_ASSIGN_RE.match(str(bound.get("text") or ""))
    return m.group(1) if m else None


def _bound_value(analysis: dict[str, Any], bound: dict[str, Any], override: Any) -> tuple[int | None, dict[str, Any]]:
    """The number a loop bound stands for, and where it came from."""
    if override not in (None, ""):
        n = _as_int(override)
        return n, {"kind": "chosen", "value": n}
    if bound.get("kind") == "literal":
        n = _as_int(bound.get("value"))
        return n, {"kind": "literal", "value": n}
    if bound.get("kind") == "cell" and bound.get("cell"):
        sheet, ref = _split(bound["cell"])
        if sheet:
            v, formula = _number_at(analysis, sheet, ref)
            n = _as_int(v)
            return n, {"kind": "cell", "cell": bound["cell"], "name": bound.get("name"), "value": n, "formula": formula}
    return None, {"kind": "unknown", "text": bound.get("text")}


# --- the loops -------------------------------------------------------------------------------------

def _anchored_column(sheet: str, col: str, first: int, key_col: str | None) -> str:
    """'Sheet'!X<first>:INDEX('Sheet'!X:X, MM_LASTROW('Sheet'!<key><first>)): the column from the loop's
    first row to the last row of the table, as Mind sizes it (the key column's grid), and as Excel sees
    it today (the key column's last filled row). MM_COLUMN(X2) alone collapses to one cell in Excel when
    X has an empty cell inside the loop rows, and MM_LOOP then reads that single cell as a count."""
    q = _qs(sheet)
    return f"{q}!{col}{first}:INDEX({q}!{col}:{col},MM_LASTROW({q}!{key_col}{first}))"


def _anchored_row(sheet: str, row: int, first: int, key_row: int | None) -> str:
    """The row counterpart: 'Sheet'!<first col><row>:INDEX('Sheet'!<row>:<row>, MM_LASTCOLUMN('Sheet'!<first col><key row>))."""
    q = _qs(sheet)
    c = num_to_col(first)
    return f"{q}!{c}{row}:INDEX({q}!{row}:{row},MM_LASTCOLUMN({q}!{c}{key_row}))"



def _clear_stale_loop(analysis: dict[str, Any], drv_sheet: str, drv_cell: tuple[int, int], name: str, extra_ops: list[dict[str, Any]]) -> None:
    """A driver cell the plan leaves out must not keep an MM_LOOP of this loop from an earlier build
    (it would be a definition with another length): clear it."""
    dc, dr = drv_cell
    held = _held(analysis, drv_sheet, dc, dr)
    if isinstance(held, str) and re.search(rf'MM_LOOP\s*\(\s*"{re.escape(name)}"', held, re.I):
        extra_ops.append({"op": "clear_cell", "sheet": drv_sheet, "cell": f"{num_to_col(dc)}{dr}", "before": held, "after": None,
                          "note": f"loop {name}: this column of the table holds nothing, the loop written here earlier is removed"})



def _growth_note(analysis: dict[str, Any], sheet: str, row: int, col: int) -> str:
    """Which ways of adding rows reach the loop: a new upload always; Mind's Input Manager only imports into
    grids flagged /Input."""
    from .grids import all_grids, grid_containing

    try:
        g = grid_containing([x for x in all_grids(analysis) if x.get("sheet") == sheet], row, col)
    except Exception:
        g = None
    if g and "input" in (g.get("flag_names") or []):
        return "rows join the loop when a workbook with more rows is uploaded, and when the Input Manager imports into this /Input grid"
    return f"rows join the loop when a workbook with more rows is uploaded; rows imported through Mind's Input Manager only reach a grid flagged /Input, and this table in {_qs(sheet)} has no such title"


def _text_header_at(analysis: dict[str, Any], d: dict[str, Any], first: int, last: int) -> bool:
    """Is the loop's first value the table's header row (column), text above numbers? Only for table loops."""
    src_sheet, src_ref = _split(str(d.get("set_from") or ""))
    if not src_sheet or last <= first:
        return False
    var = re.escape(f"<{d['loop_var']}>")
    if re.search(rf"[A-Za-z]{var}", src_ref):  # a row per loop value
        head = [_held(analysis, src_sheet, c, first) for c in range(1, 6)]
        below = [_held(analysis, src_sheet, c, first + 1) for c in range(1, 6)]
    elif re.search(rf"{var}\d", src_ref):  # a column per loop value
        head = [_held(analysis, src_sheet, first, r) for r in range(1, 6)]
        below = [_held(analysis, src_sheet, first + 1, r) for r in range(1, 6)]
    else:
        return False
    num = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)  # noqa: E731
    return any(isinstance(v, str) and v.strip() for v in head) and any(num(v) for v in below) and not any(num(v) for v in head)


def _dynamic_ok(analysis: dict[str, Any], sheet: str, cells: list[tuple[int, int]]) -> bool:
    """Does the column (row) hold any value in the loop's cells? A column the table leaves empty
    gets no loop (the macro copies nothing there, and MM_COLUMN of an empty column is no range)."""
    for col, row in cells:
        v = _held(analysis, sheet, col, row)
        if v is not None and v != "":
            return True
    return False


def _table_loop(analysis: dict[str, Any], d: dict[str, Any], name: str, first: int, last: int, dynamic: bool = True) -> dict[str, Any]:
    """The macro copies one row (or column) of a table into a driver row (or column): every
    driver cell takes its own column (row) of the table as an MM_LOOP range.

    `dynamic` (default): the range is MM_COLUMN(first cell) / MM_ROW(first cell), the form the
    Formula Generation Guide shows (=MM_LOOP("...", MM_COLUMN(A2))): Mind reads the column to the
    last row of the table's grid, so rows added to the table (Input Manager, a new upload) join the
    loop. The fixed range stays in `range_fixed` for the Excel check and the plan's counts. A
    column with no value in the loop's rows gets no loop: MM_COLUMN of an empty column is no range,
    and the macro copies nothing there."""
    var = d["loop_var"]
    src_sheet, src_ref = _split(d["set_from"])
    drv_sheet, drv_ref = _split(d["cell"])
    if not src_sheet or not drv_sheet:
        return {"status": "not_built", "why": "the source or the driver has no sheet"}
    m = re.fullmatch(r"\$?([A-Za-z]+|<\w+>)\$?(\d+|<\w+>)(?::\$?([A-Za-z]+|<\w+>)\$?(\d+|<\w+>))?", src_ref.replace("$", ""))
    if not m:
        return {"status": "not_built", "why": f"cannot read the source range {d['set_from']}"}
    c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3) or m.group(1), m.group(4) or m.group(2)
    ph = f"<{var}>"
    drv_cells = _cells_of(drv_sheet, drv_ref)
    if not drv_cells:
        return {"status": "not_built", "why": f"cannot read the driver range {d['cell']}"}
    max_col, max_row = _sheet_max(analysis, src_sheet)
    cells: list[dict[str, Any]] = []
    skipped = 0
    empty_cols: list[str] = []
    extra_ops: list[dict[str, Any]] = []
    key_col = key_row = None
    cut_note = ""
    if dynamic and r1 == ph and r2 == ph and ph not in (c1, c2):
        for col in range(col_to_num(c1.upper()), col_to_num(c2.upper()) + 1):
            if all(_held(analysis, src_sheet, col, r) not in (None, "") for r in range(first, last + 1)):
                key_col = num_to_col(col)
                break
        if key_col is None:
            dynamic = False  # no column of the table is filled on every loop row: nothing to anchor the length to
        else:
            kc = col_to_num(key_col)
            end = last
            while end + 1 <= max_row and _held(analysis, src_sheet, kc, end + 1) not in (None, ""):
                end += 1
            if end > last:
                dynamic = False
                cut_note = f"the loop stops at row {last}, before the end of the table (row {end}): the ranges stay fixed, rows added later do not join the loop"
    elif dynamic and c1 == ph and c2 == ph and ph not in (r1, r2):
        for row in range(int(r1), int(r2) + 1):
            if all(_held(analysis, src_sheet, c, row) not in (None, "") for c in range(first, last + 1)):
                key_row = row
                break
        if key_row is None:
            dynamic = False
        else:
            end = last
            while end + 1 <= max_col and _held(analysis, src_sheet, end + 1, key_row) not in (None, ""):
                end += 1
            if end > last:
                dynamic = False
                cut_note = f"the loop stops at column {num_to_col(last)}, before the end of the table (column {num_to_col(end)}): the ranges stay fixed, columns added later do not join the loop"
    if r1 == ph and r2 == ph and ph not in (c1, c2):
        # a row per loop value: the loop value is the row number; one range per column
        cols = list(range(col_to_num(c1.upper()), col_to_num(c2.upper()) + 1))
        for i, col in enumerate(cols):
            if i >= len(drv_cells):
                break
            if col > max_col:
                skipped += 1
                continue
            fixed = f"{_qs(src_sheet)}!{num_to_col(col)}{first}:{num_to_col(col)}{last}"
            if dynamic and not _dynamic_ok(analysis, src_sheet, [(col, r) for r in range(first, last + 1)]):
                empty_cols.append(num_to_col(col))
                _clear_stale_loop(analysis, drv_sheet, drv_cells[i], name, extra_ops)
                continue
            rng = _anchored_column(src_sheet, num_to_col(col), first, key_col) if dynamic else fixed
            dc, dr = drv_cells[i]
            cells.append({"cell": f"{_qs(drv_sheet)}!{num_to_col(dc)}{dr}", "sheet": drv_sheet, "ref": f"{num_to_col(dc)}{dr}", "range": rng, "range_fixed": fixed, "formula": f'=MM_LOOP("{name}",{rng})',
                          "header": _held(analysis, src_sheet, col, 1) if first > 1 else None})
        unit, axis = "row", "columns"
        example_range = _anchored_column(src_sheet, c1.upper(), first, key_col) if dynamic else f"{_qs(src_sheet)}!{c1}{first}:{c1}{last}"
    elif c1 == ph and c2 == ph and ph not in (r1, r2):
        # a column per loop value: the loop value is the column number; one range per row
        rows = list(range(int(r1), int(r2) + 1))
        for i, row in enumerate(rows):
            if i >= len(drv_cells):
                break
            if row > max_row:
                skipped += 1
                continue
            fixed = f"{_qs(src_sheet)}!{num_to_col(first)}{row}:{num_to_col(last)}{row}"
            if dynamic and not _dynamic_ok(analysis, src_sheet, [(c, row) for c in range(first, last + 1)]):
                empty_cols.append(str(row))
                _clear_stale_loop(analysis, drv_sheet, drv_cells[i], name, extra_ops)
                continue
            rng = _anchored_row(src_sheet, row, first, key_row) if dynamic else fixed
            dc, dr = drv_cells[i]
            cells.append({"cell": f"{_qs(drv_sheet)}!{num_to_col(dc)}{dr}", "sheet": drv_sheet, "ref": f"{num_to_col(dc)}{dr}", "range": rng, "range_fixed": fixed, "formula": f'=MM_LOOP("{name}",{rng})', "header": None})
        unit, axis = "column", "rows"
        example_range = _anchored_row(src_sheet, int(r1), first, key_row) if dynamic else f"{_qs(src_sheet)}!{num_to_col(first)}{r1}:{num_to_col(last)}{r1}"
    else:
        return {"status": "not_built", "why": f"the loop value is not a plain row or column number in {d['set_from']}"}
    notes = []
    if skipped:
        notes.append(f"{skipped} {axis} of the source past the sheet's used range hold nothing: no loop written for them (the macro copies them empty)")
    if empty_cols:
        notes.append(f"{len(empty_cols)} {axis} of the table hold nothing in {unit}s {first} to {last} ({', '.join(empty_cols[:8])}{', …' if len(empty_cols) > 8 else ''}): no loop written for them (MM_COLUMN of an empty column is no range; the macro copies nothing there)")
    if cut_note:
        notes.append(cut_note)
    if dynamic:
        key = f"column {key_col}" if unit == "row" else f"row {key_row}"
        notes.append(f"the ranges follow the table: every range runs from {unit} {first} to the last {unit} of {key} of {_qs(src_sheet)} (MM_LASTROW / MM_LASTCOLUMN, which Mind reads from the table's grid), so {unit}s added to the table join the loop")
        notes.append(f"Excel behaviour may differ from Mind behaviour for these formulas: in Excel the add-in stops at the first empty cell of {key}, which today gives {unit}s {first} to {last}; Mind reads the table's grid")
        notes.append(_growth_note(analysis, src_sheet, first if unit == "row" else (key_row or first), col_to_num(key_col) if (unit == "row" and key_col) else first))
    elif not dynamic and (key_col is None and key_row is None) and unit:
        pass
    if extra_ops:
        notes.append(f"{len(extra_ops)} driver cell(s) kept a loop from an earlier build over an empty column: cleared ({', '.join(o['cell'] for o in extra_ops[:6])})")
    return {"kind": f"{unit}_of_table", "status": "ready" if cells else "not_built", "cells": cells, "source_sheet": src_sheet, "unit": unit, "example_range": example_range, "notes": notes, "dynamic": dynamic, "extra_ops": extra_ops,
            "how": f"{d['cell']} takes {unit} <k> of {_qs(src_sheet)} ({unit}s {first} to {last}" + (", and every further one the table gets" if dynamic else "") + f"): each cell its own column of the table, through MM_LOOP"}


def _list_loop(analysis: dict[str, Any], d: dict[str, Any], name: str, first: int, last: int, dynamic: bool = True) -> dict[str, Any]:
    """The macro walks a list with Range("start").Offset(0, i - 1) (or Offset(i - 1, 0)):
    the driver cell gets MM_LOOP over the stretch of the list the loop covers. `dynamic`: when the
    stretch reaches the last filled cell of the list, MM_ROW / MM_COLUMN from its first cell, so
    items added to the list join the loop in Mind."""
    var = d["loop_var"]
    m = _OFFSET_RE.search(str(d.get("statement") or ""))
    src_sheet, src_ref = _split(d["set_from"])
    r = parse_ref(f"'{src_sheet}'!{src_ref}") if src_sheet else None
    if not m or not r:
        return {"status": "not_built", "why": f"cannot read the list the macro walks in: {d.get('statement')}"}
    row_expr, col_expr = m.group("r"), m.group("c")
    along_rows = re.search(rf"\b{re.escape(var)}\b", row_expr) is not None
    expr = row_expr if along_rows else col_expr
    s_first, s_last = _eval_shift(expr, var, first), _eval_shift(expr, var, last)
    if s_first is None or s_last is None:
        return {"status": "not_built", "why": f"cannot evaluate the offset {expr!r} at {var} = {first} and {last}"}
    notes: list[str] = []
    if along_rows:
        fixed = f"{_qs(src_sheet)}!{num_to_col(r['c1'])}{r['r1'] + s_first}:{num_to_col(r['c1'])}{r['r1'] + s_last}"
        beyond = _held(analysis, src_sheet, r["c1"], r["r1"] + s_last + 1)
        dyn = f"MM_COLUMN({_qs(src_sheet)}!{num_to_col(r['c1'])}{r['r1'] + s_first})"
    else:
        fixed = f"{_qs(src_sheet)}!{num_to_col(r['c1'] + s_first)}{r['r1']}:{num_to_col(r['c1'] + s_last)}{r['r1']}"
        beyond = _held(analysis, src_sheet, r["c1"] + s_last + 1, r["r1"])
        dyn = f"MM_ROW({_qs(src_sheet)}!{num_to_col(r['c1'] + s_first)}{r['r1']})"
    use_dynamic = dynamic and (beyond is None or beyond == "")
    rng = dyn if use_dynamic else fixed
    if dynamic and not use_dynamic:
        notes.append(f"the macro stops before the end of the list ({fixed}; the list goes on): the range stays fixed, items added to the list do not join the loop")
    elif use_dynamic:
        notes.append(f"the range follows the list: {dyn} reads to the end of the list's grid in Mind, so items added to it join the loop")
    drv_sheet, drv_ref = _split(d["cell"])
    cell = {"cell": d["cell"], "sheet": drv_sheet, "ref": drv_ref, "range": rng, "range_fixed": fixed, "formula": f'=MM_LOOP("{name}",{rng})', "header": None}
    return {"kind": "list", "status": "ready", "cells": [cell], "source_sheet": src_sheet, "unit": "item", "example_range": rng, "notes": notes, "dynamic": use_dynamic,
            "how": f"{d['cell']} takes each item of {fixed}" + (" and every item added after it" if use_dynamic else "") + " in turn, through MM_LOOP"}


def _counter_loop(analysis: dict[str, Any], d: dict[str, Any], name: str, first: int, last: int, first_src: dict[str, Any], last_src: dict[str, Any], count_cell: str | None) -> dict[str, Any]:
    """The driver is the loop counter itself (a cell the macro sets to its counter): MM_LOOP over
    a cell holding the count gives 1..n; another first value needs the IndexStart / IndexEnd
    arguments; without a cell holding the count the plan asks for one."""
    drv_sheet, drv_ref = _split(d["cell"])
    notes: list[str] = []
    ops_extra: list[dict[str, Any]] = []
    if last_src.get("kind") == "cell" and last_src.get("formula") is None:
        n_cell = last_src["cell"]
        n_text = f"{_qs(_split(n_cell)[0])}!{_split(n_cell)[1]}"
    elif count_cell:
        sheet, ref = _split(count_cell)
        if not sheet or not parse_ref(ref):
            return {"status": "not_built", "why": f"the count cell {count_cell!r} is not Sheet!Cell"}
        n_text = f"{_qs(sheet)}!{ref}"
        ops_extra.append({"op": "set_value", "sheet": sheet, "cell": ref, "before": _held(analysis, sheet, parse_ref(ref)["c1"], parse_ref(ref)["r1"]), "after": last, "note": f"the count of the loop {name} (chosen cell)"})
        notes.append(f"{count_cell} gets the count {last} (you chose the cell)")
    else:
        return {"status": "needs_decision", "kind": "counter", "cells": [], "notes": [],
                "why": f"MM_LOOP needs a cell that holds the count ({last}); the macro's last value is {last_src.get('text') or last!r}, not a cell. Name a free cell for it (count cell) and the plan writes {last} there.",
                "how": f"{d['cell']} is the loop counter itself ({first} to {last})"}
    if first == 1:
        formula = f'=MM_LOOP("{name}",{n_text})'
    else:
        if first_src.get("kind") == "cell":
            f_sheet, f_ref = _split(first_src["cell"])
            f_text = f"{_qs(f_sheet)}!{f_ref}"
        else:
            f_text = str(first)
        formula = f'=MM_LOOP("{name}",{n_text},{first},{f_text},{n_text})'
        notes.append(f"the run starts at {first}, not 1: the IndexStart / IndexEnd arguments of MM_LOOP say so (verified in Mind: the loop then has {last - first + 1} instances, {first} to {last}, and MM_RESULT counts them from 1)")
    cell = {"cell": d["cell"], "sheet": drv_sheet, "ref": drv_ref, "range": n_text, "formula": formula, "header": None}
    return {"kind": "counter", "status": "ready", "cells": [cell], "source_sheet": None, "unit": "value", "example_range": n_text, "notes": notes, "extra_ops": ops_extra,
            "how": f"{d['cell']} counts from {first} to {last} ({n_text} holds the count), through MM_LOOP"}


def _derived_loop(analysis: dict[str, Any], d: dict[str, Any], names: Any, name: str, first: int, last: int, last_src: dict[str, Any], count_cell: str | None) -> dict[str, Any]:
    """The driver is computed from the counter (e.g. year = first_year + i - 1): the
    statement's right-hand side with MM_LOOP in the counter's place."""
    st = str(d.get("statement") or "")
    if "=" not in st:
        return {"status": "not_built", "why": f"cannot read the statement {st!r}"}
    rhs = st.split("=", 1)[1].strip()
    var = d["loop_var"]
    if last_src.get("kind") == "cell" and last_src.get("formula") is None:
        n_text = f"{_qs(_split(last_src['cell'])[0])}!{_split(last_src['cell'])[1]}"
        extra: list[dict[str, Any]] = []
        note = None
    elif count_cell:
        sheet, ref = _split(count_cell)
        if not sheet or not parse_ref(ref):
            return {"status": "not_built", "why": f"the count cell {count_cell!r} is not Sheet!Cell"}
        n_text = f"{_qs(sheet)}!{ref}"
        extra = [{"op": "set_value", "sheet": sheet, "cell": ref, "before": _held(analysis, sheet, parse_ref(ref)["c1"], parse_ref(ref)["r1"]), "after": last - first + 1, "note": f"the count of the loop {name} (chosen cell)"}]
        note = f"{count_cell} gets the count {last - first + 1} (you chose the cell)"
    else:
        return {"status": "needs_decision", "kind": "derived", "cells": [], "notes": [],
                "why": f"MM_LOOP needs a cell that holds the count ({last - first + 1}); the loop runs {var} = {first} to {last} with no such cell. Name a free cell for it (count cell) and the plan writes the count there.",
                "how": f"{d['cell']} is computed from the counter: {rhs}"}

    def _name_to_ref(m: re.Match) -> str:
        ref = names.ref(m.group("name")) if names else None
        if not ref:
            return m.group(0)
        sheet, r = _split(ref)
        return f"{_qs(sheet)}!{r.replace('$', '')}" if sheet else r.replace("$", "")

    expr = _RANGE_NAME_RE.sub(_name_to_ref, rhs)
    loop_text = f'MM_LOOP("{name}",{n_text})' if first == 1 else f'(MM_LOOP("{name}",{n_text})+{first - 1})'
    expr = re.sub(rf"\b{re.escape(var)}\b", loop_text, expr)
    if "Range(" in expr or "Cells(" in expr:
        return {"status": "not_built", "why": f"the statement reads a range the plan cannot name: {rhs}"}
    drv_sheet, drv_ref = _split(d["cell"])
    cell = {"cell": d["cell"], "sheet": drv_sheet, "ref": drv_ref, "range": n_text, "formula": "=" + _tight(expr), "header": None}
    return {"kind": "derived", "status": "ready", "cells": [cell], "source_sheet": None, "unit": "value", "example_range": n_text, "notes": [note] if note else [], "extra_ops": extra,
            "how": f"{d['cell']} is computed from the counter ({first} to {last}): {rhs}"}


# --- the plan --------------------------------------------------------------------------------------

def _complete_header(analysis: dict[str, Any], t_sheet: str, t_cells: list[tuple[int, int]], from_sheet: str, src_cells: list[tuple[int, int]], loop_name: str) -> tuple[list[dict[str, Any]], str | None]:
    """The row above an MM_RESULT block, completed so Mind reads the block as ONE grid.

    Seen in real Mind (2026-10-15): a grid's first row is its header; an empty cell in that row
    splits the block into several grids, and a grid that then starts on the block's first row shows
    the first loop pass as its header. Every header cell that is empty, a number or a date gets the
    caption the source block carries above it (the row above the source), else the source cell's
    address. Text headers and formulas are left alone. Returns (operations, note)."""
    if not t_cells or not src_cells or len(t_cells) != len(src_cells):
        return [], None
    top = min(r for _, r in t_cells)
    if top <= 1:
        return [], f"the block starts on row 1: no header row above it, so Mind takes its first row (loop value 1) as the header"
    src_top = min(r for _, r in src_cells)
    pairs = sorted({(tc, sc) for (sc, sr), (tc, tr) in zip(src_cells, t_cells) if tr == top and sr == src_top})
    ops: list[dict[str, Any]] = []
    for tc, sc in pairs:
        held = _held(analysis, t_sheet, tc, top - 1)
        if isinstance(held, str) and held.startswith("="):
            continue  # a formula: not ours to touch
        if isinstance(held, str) and held.strip() and not re.match(r"^\d{4}-\d{2}-\d{2}", held.strip()):
            continue  # a text header already (a date, read as text, is not a caption)
        caption: Any = _held(analysis, from_sheet, sc, src_top - 1) if src_top > 1 else None
        if not (isinstance(caption, str) and caption.strip() and not caption.startswith("=")):
            caption = f"{from_sheet}!{num_to_col(sc)}{src_top}"
        ops.append({"op": "set_value", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": t_sheet, "cell": f"{num_to_col(tc)}{top - 1}", "before": held, "after": caption.strip(),
                    "note": f"header of the {loop_name} block for Mind (a grid's first row is its header; an empty cell there splits the block)"})
    if not ops:
        return [], None
    cells = ", ".join(o["cell"] for o in ops)
    return ops, f"the header row {top - 1} above the block is completed in {cells} (Mind reads a grid's first row as its header; an empty cell there splits the block into several grids)"


def build_loop_plan(analysis: dict[str, Any], model_map: dict[str, Any], *, names: dict[str, str] | None = None, name: str | None = None,
                    first: Any = None, last: Any = None, count_cell: str | None = None, clear_old: bool = False, dynamic: bool = True, max_result_cells: int = MAX_RESULT_CELLS) -> dict[str, Any]:
    """The Mind loop for the map's run loop: what goes where, with the operations that write it.

    `names` maps a driver cell to the loop name wanted; `name` is the shorthand for the first
    loop; `first` / `last` override the bounds read from the workbook; `count_cell` is where a
    counted loop's size may be written when the workbook has no cell for it; `clear_old` also
    empties what an earlier run left in the areas the macro clears (listed either way)."""
    from .model_map import _Names

    wb0 = analysis["workbooks"][0]
    nm = _Names(wb0.get("defined_names", []) or [])
    lp = model_map.get("loop") or {}
    plan: dict[str, Any] = {"found": bool(lp.get("found")), "loops": [], "outputs": [], "clears": [], "operations": [], "notes": [], "counts": {"loop_cells": 0, "result_cells": 0, "cleared": 0, "extra": 0}, "how_it_runs": None}
    if not lp.get("found"):
        plan["notes"].append("No run loop in the macros: nothing to build." + (" The item cell(s) say which sheet copy is which item." if lp.get("item_cells") else ""))
        return plan
    drivers = lp.get("drivers") or []
    names = dict(names or {})
    if name and drivers:
        names.setdefault(drivers[0]["cell"], name)
    ops: list[dict[str, Any]] = []
    built: dict[str, dict[str, Any]] = {}  # loop var -> loop entry
    placeholders: dict[str, int] = {}  # bound variables the macro's statements use -> their value
    for idx, d in enumerate(drivers):
        first_n, first_src = _bound_value(analysis, d.get("from") or {}, first if idx == 0 else None)
        last_n, last_src = _bound_value(analysis, d.get("to") or {}, last if idx == 0 else None)
        loop_name = loop_name_from(names.get(d["cell"]) or d.get("name") or (_split(d["set_from"])[0] if d.get("set_from") else None) or d["loop_var"])
        entry: dict[str, Any] = {"name": loop_name, "driver": d["cell"], "driver_name": d.get("name"), "label": d.get("label"), "macro": d.get("macro"), "loop_var": d["loop_var"],
                                 "source": d.get("set_from"), "first": first_n, "last": last_n, "first_from": first_src, "last_from": last_src, "count": None, "cells": [], "notes": [], "status": "ready", "kind": None, "how": d.get("how")}
        if first_n is None or last_n is None:
            entry.update(status="needs_decision", why="the first or last value of the loop could not be read from the workbook (" + ", ".join(x for x in (first_src.get("text"), last_src.get("text")) if x) + "): type them in")
            plan["loops"].append(entry)
            continue
        if last_n < first_n:
            entry.update(status="needs_decision", why=f"the last value ({last_n}) is before the first ({first_n}): the macro would do nothing")
            plan["loops"].append(entry)
            continue
        set_from0 = str(d.get("set_from") or "")
        if first_src.get("kind") != "chosen" and set_from0 and f"<{d['loop_var']}>" in set_from0 and _text_header_at(analysis, d, first_n, last_n):
            # the macro starts on the table's header row (it runs it too); Mind would make the header instance 1
            entry["notes"].append(f"the macro starts at {first_n}, the table's header: the loop starts at {first_n + 1} (choose first = {first_n} to keep the header pass)")
            entry["header_skipped"] = first_n
            first_n += 1
            entry["first"] = first_n
        entry["count"] = last_n - first_n + 1
        for b, v in ((d.get("from") or {}, first_n), (d.get("to") or {}, last_n)):
            bv = _bound_var(b)
            if bv:
                placeholders[bv] = v
        set_from = str(d.get("set_from") or "")
        how = str(d.get("how") or "")
        if set_from and f"<{d['loop_var']}>" in set_from:
            res = _table_loop(analysis, d, loop_name, first_n, last_n, dynamic=dynamic)
        elif how == "the loop value itself":
            res = _counter_loop(analysis, d, loop_name, first_n, last_n, first_src, last_src, count_cell)
        elif set_from and _OFFSET_RE.search(str(d.get("statement") or "")):
            res = _list_loop(analysis, d, loop_name, first_n, last_n, dynamic=dynamic)
        elif set_from and "through" in how:
            res = _derived_loop(analysis, d, nm, loop_name, first_n, last_n, last_src, count_cell)
        else:
            res = {"status": "not_built", "why": f"the macro sets the driver in a way the plan cannot write as MM_LOOP: {d.get('statement')}"}
        pre_notes = list(entry["notes"])
        entry.update({k: v for k, v in res.items() if k != "extra_ops"})
        entry["notes"] = pre_notes + list(entry.get("notes") or [])
        if res.get("status") == "ready":
            if entry["kind"] in ("row_of_table", "column_of_table") and first_n >= 1:
                # the first value is the table's header row? (text in the first filled cells, numbers below)
                src_sheet = res["source_sheet"]
                head = [_held(analysis, src_sheet, c, first_n) for c in range(1, 6)]
                below = [_held(analysis, src_sheet, c, first_n + 1) for c in range(1, 6)] if last_n > first_n else []
                if any(isinstance(v, str) and v.strip() for v in head) and any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in below) and not any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in head):
                    entry["notes"].append(f"the first value is row {first_n} of {_qs(src_sheet)}, its header row: the macro runs it too (its results are the first block). Set first to {first_n + 1} to leave it out.")
            for c in res["cells"]:
                r = parse_ref(c["ref"])
                before = _held(analysis, c["sheet"], r["c1"], r["r1"]) if r else None
                note = f"loop {loop_name}: the driver takes its value from {c['range']}"
                if isinstance(before, str) and before.startswith("="):
                    note += " (held a formula; the macro overwrote it with values on every pass)"
                ops.append({"op": "set_formula", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": c["sheet"], "cell": c["ref"], "before": before, "after": c["formula"], "note": note})
                dropdown = _dropdown_at(analysis, c["sheet"], c["ref"])
                if dropdown:
                    ops.append({"op": "remove_validation", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": c["sheet"], "cell": c["ref"], "before": dropdown, "after": None,
                                "note": f"loop {loop_name}: Mind treats a cell with a dropdown as an input and drops its formula, so the loop would be lost"})
                    entry["notes"].append(f"{_qs(c['sheet'])}!{c['ref']} had {dropdown}: it is removed, because Mind drops the formula of a cell with a dropdown and the loop would be lost.")
                    plan["counts"]["extra"] += 1
            for o in res.get("extra_ops") or []:
                ops.append({"action_id": ACTION_ID, "rule_id": RULE_ID, **o})
                plan["counts"]["extra"] += 1
            plan["counts"]["loop_cells"] += len(res["cells"])
            built[d["loop_var"]] = entry
        plan["loops"].append(entry)

    # --- the outputs: one block per loop value -> MM_RESULT row by row ------------------------------
    n_built = len(built)
    for o in (model_map.get("outputs") or {}).get("cells") or []:
        target = str(o.get("cell") or "")
        vars_in = _PH_RE.findall(target)
        loop_vars = sorted({v for v in vars_in if v in built})
        entry_o: dict[str, Any] = {"target": target, "from": o.get("from"), "written_by": o.get("written_by"), "when": o.get("when"), "status": "ready", "cells": 0, "items": 0, "example": None, "loop": None}
        when = str(o.get("when") or "")
        if "cleared" in when:
            # the area the macro empties before a run: what an earlier run left there goes
            sub = target
            for v, val in placeholders.items():
                sub = sub.replace(f"<{v}>", str(val))
            for lv, le in built.items():
                sub = sub.replace(f"<{lv}>", str(le["first"]))
            sheet, ref = _split(sub)
            r = parse_ref(f"'{sheet}'!{ref}") if sheet and "<" not in sub else None
            if not r:
                entry_o.update(status="not_built", why=f"cannot place the cleared area {target}")
                plan["clears"].append(entry_o)
                continue
            entry_o["resolved"] = sub
            plan["clears"].append({**entry_o, "sheet": sheet, "r1": r["r1"], "r2": r["r2"], "c1": r["c1"], "c2": r["c2"]})
            continue
        if not o.get("from") or not loop_vars:
            if "copy / paste" in when:
                entry_o.update(status="nothing_to_do", why="a values copy made once after the loop: in Mind the sheet it copies from holds the results; read it there")
            elif "same place" in when and built:
                ex = ",".join(f'"{le["name"]}",k' for le in built.values())
                entry_o.update(status="not_built", why=f"the macro writes every pass's block at a place it computes itself (a tariff, a counter): MM_RESULT cannot follow that. In Mind the results stay in {o.get('from') or 'the source block'} as one value per loop value; read one with =MM_RESULT(cell,{ex}).")
            elif built:
                ex = ",".join(f'"{le["name"]}",k' for le in built.values())
                entry_o.update(status="not_built", why=f"not written per loop value in a way the plan can place; in Mind read {o.get('from') or 'the source'} with =MM_RESULT(cell,{ex}).")
            else:
                entry_o.update(status="not_built", why="no loop was built")
            plan["outputs"].append(entry_o)
            continue
        if n_built > 1 and len(loop_vars) < n_built:
            ex = ",".join(f'"{le["name"]}",k{i + 1}' for i, le in enumerate(built.values()))
            entry_o.update(status="not_built", why=f"the block is placed by {', '.join(loop_vars)} only, but {n_built} nested loops run: MM_RESULT with one loop sums over the others. Read a block with =MM_RESULT(cell,{ex}).")
            plan["outputs"].append(entry_o)
            continue
        lv = loop_vars[0]
        le = built[lv]
        from_sheet, from_ref = _split(str(o["from"]))
        src_cells = _cells_of(from_sheet, from_ref) if from_sheet else []
        if not src_cells:
            entry_o.update(status="not_built", why=f"cannot read the source block {o['from']}")
            plan["outputs"].append(entry_o)
            continue
        sub = target
        for v, val in placeholders.items():
            sub = sub.replace(f"<{v}>", str(val))
        count = le["count"]
        total = count * len(src_cells)
        if total > max_result_cells:
            entry_o.update(status="not_built", items=count, cells=total, why=f"{count} loop values x {len(src_cells)} cells = {total} MM_RESULT formulas, above the limit of {max_result_cells}; narrow the bounds or build this block by hand")
            plan["outputs"].append(entry_o)
            continue
        written = 0
        first_cell = None
        first_block: list[tuple[int, int]] = []
        first_block_sheet = ""
        for k in range(1, count + 1):
            value = le["first"] + k - 1
            tgt = sub.replace(f"<{lv}>", str(value))
            t_sheet, t_ref = _split(tgt)
            t_cells = _cells_of(t_sheet, t_ref) if t_sheet else []
            if not t_cells or len(t_cells) != len(src_cells):
                entry_o.update(status="not_built", why=f"the block {tgt} and its source {o['from']} are not the same shape")
                written = -1
                break
            if k == 1:
                first_block, first_block_sheet = t_cells, t_sheet
            for (sc, sr), (tc, tr) in zip(src_cells, t_cells):
                formula = f'=MM_RESULT({_qs(from_sheet)}!${num_to_col(sc)}${sr},"{le["name"]}",{k})'
                before = _held(analysis, t_sheet, tc, tr)
                ops.append({"op": "set_formula", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": t_sheet, "cell": f"{num_to_col(tc)}{tr}", "before": before, "after": formula,
                            "note": f"loop {le['name']} value {k} (= {value}): what {_qs(from_sheet)}!{num_to_col(sc)}{sr} holds for it"})
                if first_cell is None:
                    first_cell = f"{_qs(t_sheet)}!{num_to_col(tc)}{tr} {formula}"
                written += 1
        if written < 0:
            plan["outputs"].append(entry_o)
            continue
        entry_o.update(items=count, cells=written, example=first_cell, loop=le["name"])
        if le.get("dynamic"):
            entry_o["rows_note"] = (f"written for the {count} loop values present now: Mind computes every instance the table gets (open the source grid, Partial run, "
                                    f"the {le['name']} box), but rows added later have no row in this block")
            plan["notes"].append(f"{target}: {entry_o['rows_note']}")
        plan["counts"]["result_cells"] += written
        header_ops, header_note = _complete_header(analysis, first_block_sheet, first_block, from_sheet, src_cells, le["name"])
        if header_ops:
            ops.extend(header_ops)
            plan["counts"]["extra"] += len(header_ops)
            entry_o["header_cells"] = [f"{o['sheet']}!{o['cell']}" for o in header_ops]
        if header_note:
            entry_o["note"] = header_note
            plan["notes"].append(f"{target}: {header_note}")
        plan["outputs"].append(entry_o)

    # --- the cleared areas: drop what an earlier run left outside the rows the plan writes ---------
    from .inventory import open_cached

    written_cells = {(op["sheet"], op["cell"].upper()) for op in ops if op["op"] == "set_formula"}
    try:
        wb = open_cached(analysis)
    except Exception:
        wb = None
    for c in plan["clears"]:
        if c.get("status") != "ready" or wb is None or c["sheet"] not in wb.sheetnames:
            continue
        ws = wb[c["sheet"]]
        r2 = min(c["r2"], ws.max_row)
        c2 = min(c["c2"], ws.max_column)
        # an output block that reads the cleared area (e.g. A13 = ... + OFFSET(V13, ...) over the
        # per-item results the macro pastes) would add the earlier run's values to every loop
        # value in Mind; the macro clears them first, so the plan does too, without being asked
        reader = next((rd for o in plan["outputs"] if o.get("from") for rd in [_reader_of_area(wb, o["from"], c)] if rd), None)
        clear = clear_old or reader is not None
        if reader and not clear_old:
            c["read_by"] = reader
        n = 0
        over = 0
        for row in ws.iter_rows(min_row=c["r1"], max_row=r2, min_col=c["c1"], max_col=c2):
            for cell in row:
                if cell.value is None or cell.value == "":
                    continue
                ref = f"{num_to_col(cell.column)}{cell.row}"
                if (c["sheet"], ref) in written_cells:
                    continue
                if n >= MAX_CLEARS:
                    over += 1
                    continue
                if clear:
                    ops.append({"op": "clear_cell", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": c["sheet"], "cell": ref, "before": cell.value, "after": None, "note": "the macro clears this area before a run; nothing from an earlier run stays"})
                n += 1
        c["cells"] = n
        c["included"] = bool(clear and n)
        if over:
            c["why"] = f"{over} more filled cells in the area were left (limit {MAX_CLEARS})"
        if clear:
            plan["counts"]["cleared"] += n
        if c.get("read_by") and n:
            plan["notes"].append(f"{c['resolved'] if c.get('resolved') else c['target']} is cleared although not asked: {c['read_by']} reads it, so what an earlier run left there would be added to every loop value in Mind (the macro clears it before its loop)")

    plan["operations"] = ops
    ready = [le for le in plan["loops"] if le["status"] == "ready"]
    if ready:
        parts = []
        for le in ready:
            parts.append(f"{le['name']} is a loop of {le['count']} values ({le['how']})")
        outs = [o for o in plan["outputs"] if o["status"] == "ready"]
        text = "In Mind: " + "; ".join(parts) + "."
        for o in outs:
            text += f" {_PH_RE.sub('k', o['target'])} reads {o['from']} for loop value k through MM_RESULT ({o['items']} blocks)."
        text += " One run computes every value; the macro " + ", ".join(sorted({le["macro"] for le in ready if le.get("macro")})) + " is not needed in Mind."
        plan["how_it_runs"] = text
    if any(le["status"] == "needs_decision" for le in plan["loops"]):
        plan["notes"].append("A loop needs a decision before it can be written: see its line.")
    return plan


def describe(plan: dict[str, Any]) -> str:
    lines = []
    if not plan["found"]:
        return "\n".join(plan["notes"] or ["no run loop"])
    for le in plan["loops"]:
        lines.append(f"loop {le['name']}  driver {le['driver']}  {le['first']} -> {le['last']}  [{le['status']}]")
        if le.get("why"):
            lines.append(f"   {le['why']}")
        for n in le.get("notes") or []:
            lines.append(f"   note: {n}")
        for c in (le.get("cells") or [])[:4]:
            lines.append(f"   {c['cell']}  {c['formula']}")
        if len(le.get("cells") or []) > 4:
            lines.append(f"   ... {len(le['cells']) - 4} more cells")
    for o in plan["outputs"]:
        lines.append(f"output {o['target']} <- {o['from']}  [{o['status']}]" + (f"  {o['cells']} cells in {o['items']} blocks" if o.get("cells") else ""))
        if o.get("example"):
            lines.append(f"   e.g. {o['example']}")
        if o.get("why"):
            lines.append(f"   {o['why']}")
    for c in plan["clears"]:
        lines.append(f"clear {c.get('resolved') or c['target']}  {c.get('cells', 0)} filled cells  [{'written' if c.get('included') else 'listed only' if c['status'] == 'ready' else c['status']}]")
    lines.append(f"operations: {len(plan['operations'])}")
    if plan.get("how_it_runs"):
        lines.append(plan["how_it_runs"])
    for n in plan["notes"]:
        lines.append(n)
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    import argparse
    import tempfile

    from .inventory import build_analysis
    from .model_map import build_model_map

    ap = argparse.ArgumentParser(description="The macro's run loop as Mind formulas: what goes where.")
    ap.add_argument("workbook")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--name", default=None, help="loop name for the first loop")
    ap.add_argument("--first", default=None)
    ap.add_argument("--last", default=None)
    ap.add_argument("--count-cell", default=None, help="Sheet!Cell to hold the count of a counted loop that has none")
    ap.add_argument("--clear-old", action="store_true", help="also empty the areas the macro clears before a run")
    ap.add_argument("--fixed-ranges", action="store_true", help="fixed ranges (X2:X6) instead of MM_COLUMN / MM_ROW that follow the table")
    ap.add_argument("--work", default=None)
    a = ap.parse_args(argv)
    work = Path(a.work) if a.work else Path(tempfile.mkdtemp(prefix="mind_loop_"))
    analysis = build_analysis(Path(a.workbook), work, Path(a.workbook).stem)
    plan = build_loop_plan(analysis, build_model_map(analysis), name=a.name, first=a.first, last=a.last, count_cell=a.count_cell, clear_old=a.clear_old, dynamic=not a.fixed_ranges)
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
