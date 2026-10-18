"""Pivot tables, through Excel (1.11.7): turn one into plain cells, or into formulas over its source.

Mind's converter stopped on Phoenix TVOG with "The method or operation is not implemented"; with its pivot
tables turned into values the same file was accepted. Mind imports values only, so a pivot table as
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
from typing import Any

XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_DATABASE = 1  # PivotCache.SourceType: a worksheet range
XL_NO_ADDITIONAL_CALCULATION = -4143
# PivotCell.PivotCellType
XL_CELL_VALUE, XL_CELL_SUBTOTAL, XL_CELL_GRAND_TOTAL = 0, 2, 3
# PivotField.Function -> the conditional function, and whether it needs the value column first
FUNCTIONS = {-4157: "SUMIFS", -4112: "COUNTIFS", -4106: "AVERAGEIFS", -4136: "MAXIFS", -4139: "MINIFS"}
PLAIN = {"SUMIFS": "SUM", "COUNTIFS": "COUNTA", "AVERAGEIFS": "AVERAGE", "MAXIFS": "MAX", "MINIFS": "MIN"}
MAX_FORMULA_CELLS = 20_000  # above this a pivot becomes values: that many SUMIFS slow Excel and Mind down
TOLERANCE = 1e-9

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
    try:
        float(text)
        return text  # a number Excel gave as text: compared as a number
    except ValueError:
        pass
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
    # values and formats, through a scratch sheet: a pivot cannot be pasted over
    scratch = wb.Worksheets.Add()
    try:
        rng.Copy()
        scratch.Range("A1").PasteSpecial(XL_PASTE_VALUES)
        scratch.Range("A1").PasteSpecial(XL_PASTE_FORMATS)
        excel.CutCopyMode = False
        top, left = int(rng.Row), int(rng.Column)
        rng.Clear()  # clearing the whole range removes the pivot table
        scratch.Range(scratch.Cells(1, 1), scratch.Cells(rows, cols)).Copy(ws.Range(address))
        excel.CutCopyMode = False
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
    for r, c, f in plan:
        ws.Cells(r, c).Formula = f
    ws.Calculate()
    moved = [(r, c) for r, c, _ in plan if not _same(shown[(r, c)], ws.Cells(r, c).Value2)]
    if moved:
        for r, c, _ in plan:
            ws.Cells(r, c).Value2 = shown[(r, c)]
        r, c = moved[0]
        out["why_values"] = f"the formulas did not give back {len(moved)} of its {len(plan)} values (e.g. {_col(c)}{r}: grouped items, hidden items or a filter)"
        return out
    out.update(to="formulas", formulas=len(plan))
    return out
