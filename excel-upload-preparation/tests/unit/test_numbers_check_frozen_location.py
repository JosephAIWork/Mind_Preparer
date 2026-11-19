"""18 Nov 2026: Palermo's numbers check after Prep counted 4 text cells in 'Saved down'
as changed. They show a file path built on two cells Prep froze on purpose (each asked
Excel for its own file name, FRM-012). Such a text is expected to differ; a number never
is let off this way, and a text that does not read a frozen cell is still a change."""
from pathlib import Path

import openpyxl
import pytest

from app.excel_com import com_available
from app.numbers_check import Moves, _reading_frozen, compare_workbooks
from app.prep import apply_operations

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")

FREEZE = {"op": "set_formula", "action_id": "freeze_environment_formulas", "rule_id": "FRM-012", "sheet": "S", "cell": "C5",
          "before": "=CELL(...)", "after": '="mode.xlsx"', "note": "a neutral file name"}


def test_the_moves_know_the_frozen_cells_after_inserts():
    moves = Moves([[{"op": "insert_row", "sheet": "S", "row": 2}, FREEZE,
                    {"op": "set_value", "sheet": "S", "cell": "B1", "rule_id": "OTHER", "after": "x"}]])
    assert moves.frozen() == {("S", 6, 3)}  # C5 moved down by the insert at row 2; the other write is not a freeze


def test_a_text_is_followed_back_through_formulas_and_names_only():
    formulas = {"S": {(5, 3): '="mode.xlsx"',
                      (3, 3): '="Run "&MID(R5C3,1,4)',     # reads the frozen cell
                      (4, 3): '=Directory&"\\"',          # reads it through the name Directory -> C3
                      (7, 3): '=R1C2&""',                  # reads a plain constant
                      (8, 3): '=INDIRECT("C5")&""'}}       # reads it in a way the text does not show
    names = {(None, "DIRECTORY"): [("S", 3, 3, 3, 3)]}
    starts = [("S", 3, 3), ("S", 4, 3), ("S", 7, 3), ("S", 8, 3)]
    assert _reading_frozen(starts, formulas, names, {("S", 5, 3)}) == {("S", 3, 3), ("S", 4, 3)}
    assert _reading_frozen(starts, formulas, names, set()) == set()


FROZEN = {("S", 5, 3)}


def test_offset_a_name_of_another_sheet_and_a_name_that_cannot_be_told_are_not_let_off():
    formulas = {"S": {(5, 3): '="mode.xlsx"', (6, 3): '="x"',
                      (1, 1): '=OFFSET(R5C3,1,0)&""',          # shows C6, not the frozen C5
                      (1, 2): '=Directory&"x"',                # S has no Directory of its own; T's points at C5
                      (1, 3): '=Rate&R5C3'}}                   # Rate is a constant name: its cells cannot be told
    names = {("T", "DIRECTORY"): [("S", 5, 5, 3, 3)], (None, "RATE"): None}
    assert _reading_frozen([("S", 1, 1), ("S", 1, 2), ("S", 1, 3)], formulas, names, FROZEN) == set()
    # T's own formula reading its own Directory is let off
    formulas["T"] = {(1, 1): '=Directory&"x"'}
    assert _reading_frozen([("T", 1, 1)], formulas, names, FROZEN) == {("T", 1, 1)}


def test_a_text_that_also_reads_a_changed_or_written_cell_is_not_let_off():
    formulas = {"S": {(5, 3): '="mode.xlsx"', (2, 2): "=R1C1*2",
                      (3, 3): '="Run "&MID(R5C3,1,4)&" total "&R2C2',   # also reads B2, which changed
                      (4, 3): '="Run "&MID(R5C3,1,4)&R9C9'}}             # also reads I9, which Prep wrote
    assert _reading_frozen([("S", 3, 3)], formulas, {}, FROZEN, changed={("S", 2, 2)}) == set()
    assert _reading_frozen([("S", 4, 3)], formulas, {}, FROZEN, written=[("S", 9, 9, 9, 9)]) == set()
    assert _reading_frozen([("S", 3, 3), ("S", 4, 3)], formulas, {}, FROZEN) == {("S", 3, 3), ("S", 4, 3)}


def test_a_freeze_over_a_range_freezes_every_cell_and_other_writes_are_kept():
    moves = Moves([[dict(FREEZE, op="set_array_formula", cell=None, range="C5:D6"),
                    {"op": "set_value", "sheet": "S", "cell": "B1", "rule_id": "T", "after": "x"},
                    {"op": "set_value", "sheet": "S", "cell": "A9", "rule_id": "FRM-012", "before": "=x", "after": "#title"}]])
    assert moves.frozen() == {("S", 5, 3), ("S", 5, 4), ("S", 6, 3), ("S", 6, 4), ("S", 9, 1)}
    assert ("S", 1, 1, 2, 2) in moves.written()
    assert moves.is_replaced("S", 9, 1)  # still known as a title written in place of a formula


def test_a_cycle_and_a_whole_column_read_end():
    formulas = {"S": {(1, 1): "=R2C1", (2, 1): "=R1C1", (3, 2): "=C3&\"\""}}
    assert _reading_frozen([("S", 1, 1)], formulas, {}, {("S", 9, 9)}) == set()
    assert _reading_frozen([("S", 3, 2)], formulas, {}, {("S", 40, 3)}) == {("S", 3, 2)}


@needs_excel
def test_texts_built_on_a_frozen_file_name_are_expected_and_numbers_are_not(tmp_path):
    src = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["B1"] = "abc"
    ws["C5"] = '=MID(CELL("filename",A1),FIND("[",CELL("filename",A1))+1,FIND("]",CELL("filename",A1))-FIND("[",CELL("filename",A1))-1)'
    ws["C3"] = '="Run "&MID(C5,1,4)'
    ws["C4"] = '=Directory&"\\"'
    ws["C6"] = "=CODE(C5)"  # a number built on the frozen cell: still a change
    ws["C7"] = '=B1&""'
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names["Directory"] = DefinedName("Directory", attr_text="S!$C$3")
    wb.save(src)
    wb.close()
    out = apply_operations(src, tmp_path / "prep", [dict(FREEZE), {"op": "set_value", "sheet": "S", "cell": "B1", "after": "abd", "action_id": "t", "rule_id": "T"}])
    res = compare_workbooks(src, Path(out["output_path"]), Moves([out["applied"]]))
    assert res["ran"], res["verdict"]
    assert res["counts"]["frozen_location"] == 2, res
    changed = {s["original"] for s in res["samples"]["good_to_other"]}
    assert changed == {"S!C6", "S!C7"}, res["samples"]
    assert res["clean"] is False and "neutral name Prep wrote on purpose" in res["verdict"]
