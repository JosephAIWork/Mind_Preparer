"""1.11.6: pivot tables -- the scan records where they sit; the plan never writes into one or inserts a
row / column through one (Excel refuses both, and the same change used to fail every round)."""
from pathlib import Path

import pytest

from app.excel_com import com_available
from app.prep import _drop_pivot_targets, pivot_ranges

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
