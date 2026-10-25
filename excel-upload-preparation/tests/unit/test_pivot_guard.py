"""1.11.6: pivot tables -- the scan records where they sit; the plan never writes into one or inserts a
row / column through one (Excel refuses both, and the same change used to fail every round)."""
from pathlib import Path

import pytest

from app.excel_com import com_available
from app.prep import DEFAULT_PIVOT_MODE, PIVOT_MODES, _drop_pivot_targets, plan_pivot_tables, pivot_ranges
from app.validators.structure import pivot_tables

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")


def _analysis(pivots):
    return {"workbooks": [{"sheets": [{"name": "Sheet1", "pivot_tables": pivots}, {"name": "Calc"}]}]}


def test_writes_and_inserts_through_a_pivot_go_to_the_skip_list():
    pivots = pivot_ranges(_analysis([{"name": "PT1", "ref": "B3:F40"}]))
    ops = [
        {"op": "set_value", "sheet": "Sheet1", "cell": "C6", "after": "#Average of BEL_Net"},   # inside
        {"op": "set_value", "sheet": "Sheet1", "cell": "H6", "after": "#Next to it"},           # beside: kept
        {"op": "insert_row", "sheet": "Sheet1", "row": 7},                                       # cuts it
        {"op": "set_value", "sheet": "Sheet1", "cell": "H7", "after": "#T", "after_inserts": True, "insert_at": 7},  # waits for that insert
        {"op": "insert_row", "sheet": "Sheet1", "row": 45},                                      # below it: kept
        {"op": "set_value", "sheet": "Calc", "cell": "C6", "after": "x"},                       # another sheet: kept
    ]
    skipped: list[str] = []
    kept = _drop_pivot_targets(ops, skipped, pivots)
    assert [(o["sheet"], o.get("cell") or o.get("row")) for o in kept] == [("Sheet1", "H6"), ("Sheet1", 45), ("Calc", "C6")]
    assert len(skipped) == 2 and "pivot table B3:F40" in skipped[0] and "inserting a row" in skipped[1]


@needs_excel
def test_the_scan_records_a_pivot_table_built_by_excel(tmp_path):
    from app.excel_com import close_quietly, excel_session
    from app.inventory import build_analysis

    path = tmp_path / "pivot.xlsx"
    with excel_session() as excel:
        wb = excel.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = "Data"
            ws.Range("A1:B5").Value2 = (("Product", "Amount"), ("A", 1), ("B", 2), ("A", 3), ("C", 4))
            out = wb.Worksheets.Add()
            out.Name = "Pivot"
            cache = wb.PivotCaches().Create(1, "Data!R1C1:R5C2")
            pt = cache.CreatePivotTable(out.Range("C5"), "PT1")
            pt.PivotFields("Product").Orientation = 1  # rows
            pt.AddDataField(pt.PivotFields("Amount"))
            wb.SaveAs(str(path), 51)
        finally:
            close_quietly(wb)
    analysis = build_analysis(path, tmp_path / "work", "pivot-test")
    sheet = next(s for s in analysis["workbooks"][0]["sheets"] if s["name"] == "Pivot")
    assert sheet["pivot_tables"] and sheet["pivot_tables"][0]["name"] == "PT1"
    assert sheet["pivot_tables"][0]["ref"].startswith("C5:")


def test_every_pivot_blocks_and_the_setting_says_formulas_or_values():
    """2026-10-25: the client model with only its grouped pivot as values (plain ones kept) was
    rejected by Mind's converter too."""
    def analysis(grouped):
        return {"workbooks": [{"sheets": [{"name": "S", "pivot_tables": [
            {"name": "Plain", "ref": "A1:C9", "grouped_fields": []},
            {"name": "Grouped", "ref": "E1:F9", "grouped_fields": ["Age"] if grouped else []},
        ]}]}]}
    for grouped in (True, False):
        f = pivot_tables(None, analysis(grouped), {})
        assert f["status"] == "ERROR" and f["observed"]["blocking"] == 2
    assert pivot_tables(None, {"workbooks": [{"sheets": [{"name": "S"}]}]}, {})["status"] == "PASS"
    assert DEFAULT_PIVOT_MODE == "formulas" and set(PIVOT_MODES) == {"formulas", "values"}
    ops, skipped = plan_pivot_tables(analysis(True), {}, "formulas")
    assert [o["pivot"] for o in ops] == ["Plain", "Grouped"] and {o["op"] for o in ops} == {"pivot_to_formulas"} and not skipped
    assert "is grouped" in ops[1]["note"] and "stops on a pivot table" in ops[0]["note"]
    ops, _ = plan_pivot_tables(analysis(False), {}, "values")
    assert {o["op"] for o in ops} == {"pivot_to_values"}
    # the guard never drops the pivot operations themselves
    assert len(_drop_pivot_targets(ops, [], pivot_ranges(analysis(False)))) == 2


@needs_excel
def test_prep_turns_a_plain_pivot_into_formulas_and_a_grouped_one_into_values(tmp_path):
    import itertools

    import openpyxl

    from app.excel_com import close_quietly, excel_session
    from app.prep import apply_operations

    path = tmp_path / "pivots.xlsx"
    with excel_session(mm_addin=False) as excel:
        wb = excel.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = "Data"
            rows = [("Product", "Age", "Amount")] + [(p, a, i + 1) for i, (p, a) in enumerate(itertools.product(["A", "B"], [21, 35, 44, 52]))]
            ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), 3)).Value2 = rows
            out = wb.Worksheets.Add()
            out.Name = "Pivots"
            cache = wb.PivotCaches().Create(1, f"Data!R1C1:R{len(rows)}C3")
            plain = cache.CreatePivotTable(out.Range("B3"), "Plain")
            plain.PivotFields("Product").Orientation = 1
            plain.AddDataField(plain.PivotFields("Amount"))
            cache2 = wb.PivotCaches().Create(1, f"Data!R1C1:R{len(rows)}C3")
            grouped = cache2.CreatePivotTable(out.Range("F3"), "Grouped")
            grouped.PivotFields("Age").Orientation = 1
            grouped.AddDataField(grouped.PivotFields("Amount"))
            grouped.PivotFields("Age").DataRange.Cells(1, 1).Group(20, 59, 10)
            wb.SaveAs(str(path), 51)
        finally:
            close_quietly(wb)
    ops = [
        {"op": "pivot_to_formulas", "action_id": "pivot_tables", "rule_id": "STR-009", "sheet": "Pivots", "pivot": "Plain", "cell": "B3", "range": "B3:C6"},
        {"op": "pivot_to_formulas", "action_id": "pivot_tables", "rule_id": "STR-009", "sheet": "Pivots", "pivot": "Grouped", "cell": "F3", "range": "F3:G8"},
    ]
    res = apply_operations(path, tmp_path / "apply", ops)
    assert res["status"] == "APPLIED", res
    done = {o["pivot"]: o["result"] for o in res["applied"]}
    assert done["Plain"]["to"] == "formulas" and done["Plain"]["formulas"] == 3  # A, B and the grand total
    assert done["Grouped"]["to"] == "values" and "did not give back" in done["Grouped"]["why_values"]
    wb = openpyxl.load_workbook(res["output_path"])
    assert not wb["Pivots"]._pivots
    assert wb["Pivots"]["C4"].value.startswith("=SUMIFS(Data!$C$2:$C$9,Data!$A$2:$A$9,") and wb["Pivots"]["C6"].value == "=SUM(Data!$C$2:$C$9)"
    values = openpyxl.load_workbook(res["output_path"], data_only=True)["Pivots"]
    assert (values["C4"].value, values["C5"].value, values["C6"].value) == (10, 26, 36)


def test_getpivotdata_anywhere_keeps_every_pivot_as_a_pivot():
    analysis = {"workbooks": [{"sheets": [{"name": "Pivot", "pivot_tables": [{"name": "PT1", "ref": "A3:C9", "grouped_fields": ["Age"]}]}],
                               "formulas": [{"sheet": "Summary", "cell": "C5", "formula": '=GETPIVOTDATA("Sum of BEL",Pivot!$A$3,"Line","Motor")'}]}]}
    ops, skipped = plan_pivot_tables(analysis, {}, "formulas")
    assert ops == [] and "GETPIVOTDATA on Summary" in skipped[0] and "#REF!" in skipped[0]
    assert "Mind cannot read it" in skipped[0]  # a grouped pivot: the user is told it still blocks Mind
    # no GETPIVOTDATA: the grouped pivot is converted as before
    analysis["workbooks"][0]["formulas"] = [{"sheet": "Summary", "cell": "C5", "formula": "=Pivot!B4"}]
    ops, _ = plan_pivot_tables(analysis, {}, "formulas")
    assert [o["pivot"] for o in ops] == ["PT1"]


def test_criterion_treats_only_plain_digits_as_a_number():
    from app.pivots import _criterion

    assert _criterion("1.5") == "1.5" and _criterion("-2e3") == "-2e3"
    assert _criterion("1_000") == '"=1_000"' and _criterion("nan") == '"=nan"' and _criterion("inf") == '"=inf"'


def test_restore_writes_an_error_back_as_the_error():
    from app.pivots import _restore

    class Cell:
        def __init__(self):
            self.Formula = self.Value2 = None

    cells: dict = {}

    class Sheet:
        def Cells(self, r, c):
            return cells.setdefault((r, c), Cell())

    _restore(Sheet(), [(1, 1), (1, 2)], {(1, 1): -2146826246, (1, 2): 3.0})
    assert cells[(1, 1)].Formula == "#N/A" and cells[(1, 1)].Value2 is None
    assert cells[(1, 2)].Value2 == 3.0


def test_unprotect_runs_before_the_pivots_and_titles_follow_column_and_repeated_row_inserts():
    from app.prep import _ordered

    ops = [
        {"op": "pivot_to_values", "sheet": "S", "pivot": "PT1", "cell": "A1"},
        {"op": "unprotect_sheet", "sheet": "S"},
        {"op": "insert_row", "sheet": "S", "row": 6},
        {"op": "set_value", "sheet": "S", "cell": "D6", "after": "#T", "after_inserts": True, "insert_at": 6},
        {"op": "insert_column", "sheet": "S", "column": "B"},  # the assistant's: left of the title
        {"op": "insert_row", "sheet": "S", "row": 6},          # the assistant's: the same row again
        {"op": "set_value", "sheet": "T", "cell": "D6", "after": "#U", "after_inserts": True, "insert_at": 6},
    ]
    ordered = _ordered(ops)
    assert [o["op"] for o in ordered[:2]] == ["unprotect_sheet", "pivot_to_values"]
    cells = {o["after"]: o["cell"] for o in ordered if o["op"] == "set_value"}
    assert cells["#T"] == "E7"  # two blank rows at 6 and 7, the grid at 8; one column right
    assert cells["#U"] == "D6"  # another sheet: untouched


@needs_excel
def test_convert_keeps_empty_cells_empty_the_users_formulas_and_refuses_hidden_items(tmp_path):
    from app.excel_com import close_quietly, excel_session
    from app.pivots import convert

    with excel_session(mm_addin=False) as excel:
        wb = excel.Workbooks.Add()
        try:
            data = wb.Worksheets(1)
            data.Name = "Data"
            rows = [("Line", "Year", "Run", "Amount"), ("Motor", 2024, "Base", 1), ("Motor", 2025, "Base", 2), ("Home", 2024, "Base", 3),
                    ("Home", 2024, "Shock", 0)]
            data.Range(data.Cells(1, 1), data.Cells(len(rows), 4)).Value2 = rows
            ws = wb.Worksheets.Add()
            ws.Name = "Pivots"
            src = f"Data!R1C1:R{len(rows)}C4"
            pt = wb.PivotCaches().Create(1, src).CreatePivotTable(ws.Range("A4"), "Sparse")
            pt.PivotFields("Line").Orientation = 1
            pt.PivotFields("Year").Orientation = 2  # columns: Home has no 2025
            pt.PivotFields("Run").Orientation = 3   # a report filter in A2:B2
            pt.AddDataField(pt.PivotFields("Amount"))
            ws.Range("D2").Formula = "=Data!D2*2"   # the user's own cell beside the filter
            hidden = wb.PivotCaches().Create(1, src).CreatePivotTable(ws.Range("K4"), "Hidden")
            hidden.PivotFields("Run").Orientation = 1
            hidden.AddDataField(hidden.PivotFields("Amount"))
            hidden.PivotFields("Run").PivotItems("Shock").Visible = False  # adds 0 today: would match by chance
            two = wb.PivotCaches().Create(1, src).CreatePivotTable(ws.Range("P4"), "TwoValues")
            two.PivotFields("Line").Orientation = 1
            two.AddDataField(two.PivotFields("Amount"), "Sum A", -4157)
            two.AddDataField(two.PivotFields("Year"), "Max Y", -4136)  # a "Values" field among the columns
            sliced = wb.PivotCaches().Create(1, src).CreatePivotTable(ws.Range("U4"), "Sliced")
            sliced.PivotFields("Line").Orientation = 1
            sliced.AddDataField(sliced.PivotFields("Amount"))
            slicer = wb.SlicerCaches.Add2(sliced, "Run")  # Run is not in the layout: only the slicer filters it
            slicer.Slicers.Add(ws, Name="RunSlicer", Caption="Run", Top=10, Left=600, Width=100, Height=100)
            slicer.SlicerItems("Shock").Selected = False
            excel.Calculate()
            body = pt.DataBodyRange
            empty = [(int(c.Row), int(c.Column)) for c in body if c.Value2 is None]
            assert empty and excel.Intersect(ws.Range("D2"), pt.TableRange2) is not None  # inside the rectangle, not the pivot
            done = convert(excel, wb, ws, "Sparse", "formulas")
            assert done["to"] == "formulas", done
            assert all(not ws.Cells(r, c).HasFormula and ws.Cells(r, c).Value2 is None for r, c in empty)
            assert ws.Range("D2").Formula == "=Data!D2*2"
            done = convert(excel, wb, ws, "Hidden", "formulas")
            assert done["to"] == "values" and "hidden items" in done["why_values"]
            assert convert(excel, wb, ws, "TwoValues", "formulas")["to"] == "formulas"
            done = convert(excel, wb, ws, "Sliced", "formulas")
            assert done["to"] == "values" and "Run has hidden items" in done["why_values"]
        finally:
            wb.Saved = True
            close_quietly(wb)
