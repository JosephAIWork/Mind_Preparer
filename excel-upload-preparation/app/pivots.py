"""Pivot tables, through Excel: turn one into plain cells, or into formulas over its source.

Mind's converter stops on a workbook with pivot tables ("The method or operation is not implemented");
with its pivot tables turned into values the same file is accepted. Mind imports values only, so a pivot table as
plain cells is what Mind reads -- but frozen: it does not follow its source in a Mind loop. As formulas
it does: every value, subtotal and grand-total cell becomes SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS /
MINIFS over the pivot's source range, with the cell's row and column items (and the report filters)
as the conditions -- what Excel's own PivotCell says the cell adds up.

The formulas are written only where they give back exactly what the pivot showed: after writing, Excel
recalculates and every cell is compared with the pivot's value. A single difference (a grouped field, a
hidden item, a value filter, a calculated field, "show values as" ...) puts the whole pivot back as
values, and the reason is reported. Labels are always plain cells.
"""
from __future__ import annotations

import re
import time
from typing import Any

XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_PASTE_COMMENTS = -4144
XL_PASTE_VALIDATION = 6
XL_CELL_TYPE_FORMULAS = -4123  # Range.SpecialCells
XL_DATABASE = 1  # PivotCache.SourceType: a worksheet range
XL_NO_ADDITIONAL_CALCULATION = -4143
# PivotCell.PivotCellType
XL_CELL_VALUE, XL_CELL_SUBTOTAL, XL_CELL_GRAND_TOTAL = 0, 2, 3
# PivotField.Function -> the conditional function, and whether it needs the value column first
FUNCTIONS = {-4157: "SUMIFS", -4112: "COUNTIFS", -4106: "AVERAGEIFS", -4136: "MAXIFS", -4139: "MINIFS"}
PLAIN = {"SUMIFS": "SUM", "COUNTIFS": "COUNTA", "AVERAGEIFS": "AVERAGE", "MAXIFS": "MAX", "MINIFS": "MIN"}
MAX_FORMULA_CELLS = 20_000  # above this a pivot becomes values: that many SUMIFS slow Excel and Mind down
TOLERANCE = 1e-9
# Value2 gives an error cell as this number (pywin32): written back as the error, not as the number
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}

_NUMBER_RE = re.compile(r"[+-]?\d+(\.\d+)?([eE][+-]?\d+)?")
_SOURCE_RE = re.compile(r"^(?:'?(?P<sheet>(?:[^'!]|'')+)'?!)?R(?P<r1>\d+)C(?P<c1>\d+):R(?P<r2>\d+)C(?P<c2>\d+)$")


def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _criterion(value: Any) -> str:
    """A SUMIFS condition matching exactly `value`: numbers as numbers, text as "=text" with the
    wildcards escaped, an empty item as "" (empty cells)."""
    if value is None or value == "" or value == "(blank)":
        return '""'
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(float(value)).rstrip("0").rstrip(".") if isinstance(value, float) else str(value)
    text = str(value)
    if _NUMBER_RE.fullmatch(text):
        return text  # a number Excel gave as text: compared as a number (not '1_000', 'nan', 'inf' ...)
    text = text.replace("~", "~~").replace("*", "~*").replace("?", "~?").replace('"', '""')
    return f'"={text}"'


def _items(collection: Any) -> list[tuple[str, Any]]:
    """(source field name, item value) of every item of a PivotCell's RowItems / ColumnItems."""
    out = []
    for k in range(1, int(collection.Count) + 1):
        item = collection.Item(k)
        field = item.Parent
        out.append((str(field.SourceName), item.SourceName))
    return out


def formula_plan(pt: Any, excel_ws: Any) -> tuple[list[tuple[int, int, str]] | None, str | None]:
    """[(row, col, formula)] for every value cell of the pivot, or (None, why not)."""
    cache = pt.PivotCache()
    if int(cache.SourceType) != XL_DATABASE:
        return None, "its source is not a worksheet range (an external source or the data model)"
    m = _SOURCE_RE.match(str(pt.SourceData).strip())
    if not m:
        return None, f"its source ({pt.SourceData}) is not a plain range"
    src_sheet = (m.group("sheet") or str(excel_ws.Name)).replace("''", "'")
    r1, c1, r2, c2 = (int(m.group(k)) for k in ("r1", "c1", "r2", "c2"))
    if int(pt.CalculatedFields().Count):
        return None, "it has a calculated field"
    # Hidden items, label / value filters and slicers leave rows out of the pivot that the
    # formulas would still add up; today's numbers can match by chance (a hidden item adding 0).
    # A slicer on a field left out of the layout hides that field's items: HiddenFields covers it.
    # (Not the report filters: picking one item hides the others; CurrentPage handles those below.)
    try:
        values_field = str(pt.DataPivotField.Name)
    except Exception:
        values_field = None
    for fields in (pt.RowFields, pt.ColumnFields, pt.HiddenFields):
        for k in range(1, int(fields.Count) + 1):
            pf = fields.Item(k)
            name = str(pf.Name)
            if name == values_field:  # the "Values" field of a pivot with several values: no items of its own
                continue
            try:
                hidden = int(pf.HiddenItems.Count) or int(pf.PivotFilters.Count)
            except Exception:
                hidden = True
            if hidden:
                return None, f"its field {name} has hidden items or a filter"
    src = excel_ws.Parent.Worksheets(src_sheet)
    headers = src.Range(src.Cells(r1, c1), src.Cells(r1, c2)).Value2[0]
    column_of = {str(h): c1 + i for i, h in enumerate(headers) if h is not None}

    def column(name: str) -> str:
        c = column_of.get(name)
        if c is None:
            raise KeyError(name)
        return f"{_quote_sheet(src_sheet)}!${_col(c)}${r1 + 1}:${_col(c)}${r2}"

    pages: list[tuple[str, Any]] = []
    for k in range(1, int(pt.PageFields.Count) + 1):
        pf = pt.PageFields.Item(k)
        if bool(getattr(pf, "EnableMultiplePageItems", False)):
            return None, f"its report filter {pf.Name} allows several items"
        current = str(pf.CurrentPage.Name)
        if current != "(All)":
            pages.append((str(pf.SourceName), pf.CurrentPage.SourceName))

    body = pt.DataBodyRange
    if body is None:
        return None, "it has no values"
    rows, cols = int(body.Rows.Count), int(body.Columns.Count)
    if rows * cols > MAX_FORMULA_CELLS:
        return None, f"it has {rows * cols:,} value cells (more than {MAX_FORMULA_CELLS:,} formulas would slow Excel and Mind down)"
    top, left = int(body.Row), int(body.Column)
    out: list[tuple[int, int, str]] = []
    try:
        for i in range(rows):
            for j in range(cols):
                cell = excel_ws.Cells(top + i, left + j)
                pc = cell.PivotCell
                kind = int(pc.PivotCellType)
                if kind not in (XL_CELL_VALUE, XL_CELL_SUBTOTAL, XL_CELL_GRAND_TOTAL):
                    return None, f"{_col(left + j)}{top + i} is a custom subtotal"
                df = pc.DataField
                if int(df.Calculation) != XL_NO_ADDITIONAL_CALCULATION:
                    return None, f"its value {df.Name} is shown as a share or a difference"
                fn = FUNCTIONS.get(int(df.Function))
                if fn is None:
                    return None, f"its value {df.Name} uses a summary Excel's conditional functions do not have"
                conditions = _items(pc.RowItems) + _items(pc.ColumnItems) + pages
                parts = [f"{column(name)},{_criterion(value)}" for name, value in conditions]
                value_col = column(str(df.SourceName))
                if not parts:  # a grand total with no condition: the plain function (Excel wants a condition in *IFS)
                    formula = f"={PLAIN[fn]}({value_col})"
                elif fn == "COUNTIFS":
                    formula = f"=COUNTIFS({value_col},\"<>\"{''.join(',' + p for p in parts)})"
                else:
                    formula = f"={fn}({value_col}{''.join(',' + p for p in parts)})"
                out.append((top + i, left + j, formula))
    except KeyError as exc:
        return None, f"the source has no column named {exc.args[0]}"
    return out, None


def _same(a: Any, b: Any) -> bool:
    if a == b:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= TOLERANCE * max(1.0, abs(a), abs(b))
    return a in (None, "") and b in (None, "", 0)


def _restore(ws: Any, cells: list[tuple[int, int]], shown: dict[tuple[int, int], Any]) -> None:
    """Put back what the pivot showed, an error cell as its error (#N/A ...), not as Value2's number.
    Excel busy for an instant (RPC_E_CALL_REJECTED) gets the same cell again."""
    for r, c in cells:
        v = shown[(r, c)]
        for attempt in range(4):
            try:
                cell = ws.Cells(r, c)
                if type(v) is int and v in ERRORS:
                    cell.Formula = ERRORS[v]  # typed error text is stored as the error constant
                else:
                    cell.Value2 = v
                break
            except Exception:
                if attempt == 3:
                    raise
                time.sleep(0.5 * (attempt + 1))


def _formula_of(cell: Any) -> str:
    try:
        return str(cell.Formula2)  # keeps a dynamic-array formula as it is
    except Exception:
        return str(cell.Formula)


def _set_formula(cell: Any, formula: str) -> None:
    try:
        cell.Formula2 = formula
    except Exception:
        cell.Formula = formula


def convert(excel: Any, wb: Any, ws: Any, pivot_name: str, to: str) -> dict[str, Any]:
    """Turn the pivot `pivot_name` of `ws` into plain cells (to="values") or formulas (to="formulas",
    falling back to values when the formulas would not give the same numbers). -> what was done."""
    pt = ws.PivotTables(pivot_name)
    rng = pt.TableRange2
    address = str(rng.Address).replace("$", "")
    rows, cols = int(rng.Rows.Count), int(rng.Columns.Count)
    plan, why = None, None
    if to == "formulas":
        try:
            plan, why = formula_plan(pt, ws)
        except Exception as exc:  # an answer Excel gives for a layout this module does not know
            plan, why = None, f"Excel could not describe its cells ({str(exc)[:120]})"
    shown = None
    if plan:
        shown = {(r, c): ws.Cells(r, c).Value2 for r, c, _ in plan}
    # TableRange2 is the rectangle around the report filters and the body: the cells beside the
    # filters and the row between are the user's own (a pivot cell never holds a formula). Their
    # formulas are written back after the copy, so they keep following their inputs.
    extra: list[tuple[int, int, str]] = []
    if rows * cols > 1:  # SpecialCells on one cell would search the whole sheet
        try:
            extra = [(int(x.Row), int(x.Column), _formula_of(x)) for x in rng.SpecialCells(XL_CELL_TYPE_FORMULAS)]
        except Exception:  # "no cells were found"
            extra = []
    # values and formats (and comments and validation), through a scratch sheet: a pivot cannot be pasted over
    scratch = wb.Worksheets.Add()
    try:
        rng.Copy()
        scratch.Range("A1").PasteSpecial(XL_PASTE_VALUES)
        scratch.Range("A1").PasteSpecial(XL_PASTE_FORMATS)
        for kind in (XL_PASTE_COMMENTS, XL_PASTE_VALIDATION):
            try:
                scratch.Range("A1").PasteSpecial(kind)
            except Exception:  # nothing of that kind to paste
                pass
        excel.CutCopyMode = False
        rng.Clear()  # clearing the whole range removes the pivot table
        # the pivot is gone now: its values exist only on the scratch sheet, so Excel busy
        # for an instant gets the same copy again rather than leaving the area empty
        for attempt in range(4):
            try:
                scratch.Range(scratch.Cells(1, 1), scratch.Cells(rows, cols)).Copy(ws.Range(address))
                break
            except Exception:
                if attempt == 3:
                    raise
                time.sleep(0.5 * (attempt + 1))
        excel.CutCopyMode = False
        for r, c, f in extra:
            _set_formula(ws.Cells(r, c), f)
    finally:
        alerts = excel.DisplayAlerts
        excel.DisplayAlerts = False
        scratch.Delete()
        excel.DisplayAlerts = alerts
    out = {"pivot": pivot_name, "range": address, "to": "values", "formulas": 0}
    if not plan:
        if to == "formulas":
            out["why_values"] = why
        return out
    # A cell the pivot showed empty (a combination with no data) stays empty: its formula
    # would give 0, and AVERAGE, COUNT, ="" and Mind's comparisons treat 0 and empty differently.
    plan = [(r, c, f) for r, c, f in plan if shown[(r, c)] not in (None, "")]
    if not plan:
        out["why_values"] = "it shows no values"
        return out
    written: list[tuple[int, int]] = []
    try:
        for r, c, f in plan:
            ws.Cells(r, c).Formula = f
            written.append((r, c))
        ws.Calculate()
        moved = [(r, c) for r, c, _ in plan if not _same(shown[(r, c)], ws.Cells(r, c).Value2)]
    except Exception as exc:
        # the pivot is already gone, so a retry could not start again: the cells go back to its
        # values here rather than keeping formulas nobody checked
        _restore(ws, written, shown)
        out["why_values"] = f"writing its formulas failed ({str(exc)[:120]}) -- kept as values"
        return out
    if moved:
        _restore(ws, written, shown)
        r, c = moved[0]
        out["why_values"] = f"the formulas did not give back {len(moved)} of its {len(plan)} values (e.g. {_col(c)}{r}: grouped items, hidden items or a filter)"
        return out
    out.update(to="formulas", formulas=len(plan))
    return out
