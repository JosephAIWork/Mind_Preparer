"""1.12.1 (the Mind team's notes): a row Prep inserts for one grid shifts only that grid's columns, so a
grid beside it on the same rows stays where it is. A whole row is still inserted when nothing at or
below the row sits beside the grid (exactly what it did before)."""
from pathlib import Path

import openpyxl
import pytest

from app import mind_loop as ml
from app.config import load_config
from app.excel_com import com_available
from app.modes import plan_mode
from app.numbers_check import Moves
from app.prep import _resolve_insert_at, apply_operations, insert_columns, insert_moves, plan_actions
from app.rules_engine import RulesEngine

LABEL = "SYNTHETIC_TEST_FIXTURE"


def _run(tmp_path, source, sub="work"):
    return plan_mode.run(source, tmp_path / sub, load_config(), engine=RulesEngine())


def _sep(tmp_path, source):
    res = _run(tmp_path, source)
    plan = {a["id"]: a for a in plan_actions(res["workbook_analysis"], res["validation_report"])}
    return plan["separate_merged_grids"]


@pytest.fixture
def side_by_side_xlsx(tmp_path: Path) -> Path:
    """'#Second' trapped inside the left grid (B4:C9) at row 7; the right grid (E4:F12) spans row 7."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Side"
    ws["A1"] = LABEL
    ws["B3"], ws["E3"] = "#Left", "#Right"
    ws["B4"], ws["C4"] = "Year", "Rate"
    ws["B5"], ws["C5"] = 1, 0.5
    ws["B6"], ws["C6"] = 2, 0.6
    ws["B7"] = "#Second"
    ws["B8"], ws["C8"] = 3, 0.7
    ws["B9"], ws["C9"] = 4, 0.8
    ws["E4"], ws["F4"] = "Age", "Qx"
    for i, r in enumerate(range(5, 13)):
        ws.cell(r, 5, 40 + i)
        ws.cell(r, 6, 0.001 * (i + 1))
    ws["H5"] = "=SUM(F5:F12)"
    ws["H6"] = "=C5*2"
    ws["H8"] = "=C8+F8"
    p = tmp_path / "side_by_side.xlsx"
    wb.save(p)
    wb.close()
    return p


@pytest.fixture
def alone_xlsx(tmp_path: Path) -> Path:
    """The same left grid with nothing beside it."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alone"
    ws["A1"] = LABEL
    ws["B3"] = "#Left"
    ws["B4"], ws["C4"] = "Year", "Rate"
    ws["B5"], ws["C5"] = 1, 0.5
    ws["B6"], ws["C6"] = 2, 0.6
    ws["B7"] = "#Second"
    ws["B8"], ws["C8"] = 3, 0.7
    p = tmp_path / "alone.xlsx"
    wb.save(p)
    wb.close()
    return p


def test_the_insert_shifts_only_the_grids_own_columns_when_another_grid_sits_beside_it(tmp_path, side_by_side_xlsx):
    sep = _sep(tmp_path, side_by_side_xlsx)
    ops = sep["operations"]
    assert len(ops) == 1, (ops, sep["skipped"])  # before 1.12.1 it was skipped: a whole row would split '#Right'
    o = ops[0]
    assert o["row"] == 7 and o["columns"] == "B:C" and insert_columns(o) == (2, 3)
    assert "only columns B:C shift down" in o["after"]


def test_a_grid_alone_on_its_rows_still_gets_a_whole_row(tmp_path, alone_xlsx):
    ops = _sep(tmp_path, alone_xlsx)["operations"]
    assert len(ops) == 1 and "columns" not in ops[0] and insert_columns(ops[0]) is None


def test_cells_are_followed_only_in_the_columns_that_moved():
    banded = {"op": "insert_row", "sheet": "S", "row": 7, "columns": "B:C"}
    whole = {"op": "insert_row", "sheet": "S", "row": 3}
    assert insert_moves(banded, 2) and insert_moves(banded, 3) and not insert_moves(banded, 5)
    moves = Moves([[banded, whole]])
    assert moves.forward("S", 8, 3) == ("S", 10, 3)  # both inserts at or above it
    assert moves.forward("S", 8, 6) == ("S", 9, 6)  # the whole row only
    assert moves.forward("S", 2, 3) == ("S", 2, 3)
    assert moves.summary()["row_inserts"] == 2
    f = ml.row_map([3, (7, 2, 3)])
    assert f(8, 3) == 10 and f(8, 6) == 9 and f(8) == 9
    # a title written after the inserts lands one row lower only for the inserts in its column
    title = {"op": "set_value", "sheet": "S", "cell": "E9", "after": "#T", "after_inserts": True, "insert_at": 9}
    assert _resolve_insert_at(title, [banded, whole])["cell"] == "E10"
    title_b = {**title, "cell": "B9"}
    assert _resolve_insert_at(title_b, [banded, whole])["cell"] == "B11"


@pytest.mark.skipif(not com_available(), reason="needs Excel")
def test_excel_shifts_the_band_and_leaves_the_grid_beside_it_and_every_value(tmp_path, side_by_side_xlsx):
    ops = _sep(tmp_path, side_by_side_xlsx)["operations"]
    res = apply_operations(side_by_side_xlsx, tmp_path / "out", ops, prefer_excel=True)
    assert res["status"] == "APPLIED", res.get("message")
    out = Path(res["output_path"])
    wb = openpyxl.load_workbook(out)
    ws = wb["Side"]
    assert ws["B7"].value is None and ws["B8"].value == "#Second" and ws["C9"].value == 0.7  # the band moved
    assert [ws.cell(r, 5).value for r in range(4, 13)] == ["Age"] + [40 + i for i in range(8)]  # '#Right' did not
    assert ws["F7"].value == pytest.approx(0.003)
    assert ws["H8"].value == "=C9+F8"  # Excel followed the moved cell, not the one beside it
    assert ws["H6"].value == "=C5*2" and ws["H5"].value == "=SUM(F5:F12)"
    wb.close()
    # the same values everywhere, with the moved cells followed (Run in Mind's numbers check)
    a = _run(tmp_path, side_by_side_xlsx, "cmp")["workbook_analysis"]["workbooks"][0]
    gate = ml.numbers_gate(side_by_side_xlsx, out, res["applied"], {"workbooks": [a]}, tmp_path / "numbers")
    assert gate["ran"] and gate["match"], gate.get("listed")


def test_no_row_is_inserted_inside_a_range_read_by_position():
    # Tokyo IFRS 17 tool (2026-11-07): =INDEX($D$29:$H$32,Scenario+1,4) -- a title row inserted at row 31
    # stretched the table to D29:H33 and the INDEX read another row ("New" became "Old")
    from app.prep import insert_blocker

    index = {"rects": {"Ctrl": [(29, 32, 4, 8, "Ctrl!G27")]}, "whole_cols": {}, "whole_rows": {}, "positional": {}, "text_addressed": {}}
    assert "Ctrl!G27" in insert_blocker(index, "Ctrl", 31, (4, 8))
    assert "Ctrl!G27" in insert_blocker(index, "Ctrl", 31, None)  # a whole row too
    assert insert_blocker(index, "Ctrl", 29, (4, 8)) is None  # at the first row the whole table moves down: same cells
    assert insert_blocker(index, "Ctrl", 31, (10, 12)) is None  # beside the table: it does not move
    assert insert_blocker(index, "Ctrl", 33, None) is None  # below it
