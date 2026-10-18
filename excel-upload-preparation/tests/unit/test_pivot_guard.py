"""1.11.6: pivot tables -- the scan records where they sit; the plan never writes into one or inserts a
row / column through one (Excel refuses both, and the same change used to fail every round)."""
from pathlib import Path

import pytest

from app.excel_com import com_available
from app.prep import _drop_pivot_targets, plan_pivot_tables, pivot_ranges
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


def test_the_rule_blocks_only_a_grouped_pivot_and_the_setting_decides_the_rest():
    def analysis(grouped):
        return {"workbooks": [{"sheets": [{"name": "S", "pivot_tables": [
            {"name": "Plain", "ref": "A1:C9", "grouped_fields": []},
            {"name": "Grouped", "ref": "E1:F9", "grouped_fields": ["Age"] if grouped else []},
        ]}]}]}
    assert pivot_tables(None, analysis(True), {})["status"] == "ERROR"
    assert pivot_tables(None, analysis(False), {})["status"] == "WARNING"
    assert pivot_tables(None, {"workbooks": [{"sheets": [{"name": "S"}]}]}, {})["status"] == "PASS"
    ops, skipped = plan_pivot_tables(analysis(True), {}, "blocking")
    assert [(o["op"], o["pivot"]) for o in ops] == [("pivot_to_formulas", "Grouped")] and "Plain" in skipped[0]
    ops, _ = plan_pivot_tables(analysis(True), {}, "formulas")
    assert [o["pivot"] for o in ops] == ["Plain", "Grouped"] and {o["op"] for o in ops} == {"pivot_to_formulas"}
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
