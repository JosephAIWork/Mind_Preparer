"""1.14.0: a grid title never moves a cell (Joseph, 10 November 2026).

The automatic namer inserted cells over one grid's columns to make room for its title; on the IFRS 17
tool's Result_CF that pushed the key column (A1:A8, read as a grid of its own) one row down while the
block beside it (B2:CI8) stayed, so every key sat beside the wrong row. Now: a key column with a
caption on top and the block that reads it are ONE table (grids._merge_key_columns); the title goes
into the caption cell ('#' + its text) or the empty cell above the top-left cell, never anywhere that
needs a row insert; a formula caption, a grid starting on row 1 and a taken cell above are reported
under "could not name" with what to do by hand. The apply path still carries out an old plan's
insert_row."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from app import prep
from app.excel_com import com_available
from app.grids import all_grids, detect_grids
from app.inventory import build_analysis

needs_excel = pytest.mark.skipif(not com_available(), reason="requires pywin32 + an installed Excel")


def _wb(path: Path, sheets: dict[str, dict[str, object]]) -> Path:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, cells in sheets.items():
        ws = wb.create_sheet(name)
        for ref, v in cells.items():
            ws[ref] = v
    wb.save(path)
    wb.close()
    return path


def _analysis(tmp_path: Path, sheets: dict[str, dict[str, object]]) -> dict:
    return build_analysis(_wb(tmp_path / "m.xlsx", sheets), tmp_path / "w", "t")


def _titles(analysis, **kw):
    return prep.plan_create_grid_titles(analysis, {"findings": []}, **kw)


# Result_CF in miniature: a caption in A1, the header row 2, the keys of each row in column A (the
# sheet names the row's lookups use) and the block beside them that reads $A3.
RESULT_CF = {
    "A1": "Running scenario", "A2": "csv file", "B2": "ID", "C2": "Scenario",
    "A3": "BBA_01", "A4": "BBA_02", "A5": "VFA_01",
    "B3": '=$A3&"_1"', "B4": '=$A4&"_1"', "B5": '=$A5&"_1"',
    "C3": "=LEN($A3)", "C4": "=LEN($A4)", "C5": "=LEN($A5)",
}


def test_a_key_column_with_a_caption_on_top_and_its_block_are_one_table(tmp_path):
    analysis = _analysis(tmp_path, {"R": RESULT_CF})
    grids = all_grids(analysis)
    assert [g["ref"] for g in grids] == ["A2:C5"]
    table = grids[0]
    assert table["caption_cell"] == "A1" and table["caption_value"] == "Running scenario"
    assert table["merged_from"] == ["A1:A5", "B2:C5"]
    assert table["header_values"] == ["csv file", "ID", "Scenario"] and table["header_is_all_text"]
    assert table["name"] is None and table["anchor"] == "A2"


def test_the_table_is_named_once_into_its_caption_cell_and_nothing_moves(tmp_path):
    analysis = _analysis(tmp_path, {"R": RESULT_CF})
    ops, skipped = _titles(analysis)
    assert [o["op"] for o in ops] == ["set_value"]
    op = ops[0]
    assert op["sheet"] == "R" and op["cell"] == "A1" and op["before"] == "Running scenario" and op["after"] == "#Running scenario"
    assert "the caption becomes the title" in op["note"] and "key column A1:A5" in op["note"] and "block B2:C5" in op["note"]
    assert not skipped


def test_keys_that_are_formulas_merge_when_the_block_reads_the_column(tmp_path):
    """Result_CF's keys are formulas (=Ctrl_01!D35): the column counts as the key column because the
    block's formulas read it ($A3), on this sheet -- not Other!A3, not BA3."""
    keys_as_formulas = {**RESULT_CF, "A3": "=Ctrl!A1", "A4": "=Ctrl!A2", "A5": "=Ctrl!A3"}
    analysis = _analysis(tmp_path, {"R": keys_as_formulas, "Ctrl": {"A1": "BBA_01", "A2": "BBA_02", "A3": "VFA_01"}})
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "R"] == ["A2:C5"]
    not_read = {**keys_as_formulas, "B3": "=Ctrl!A1&BA3", "B4": "=Ctrl!A2", "B5": "=Ctrl!A3", "C3": "=1", "C4": "=2", "C5": "=3"}
    analysis = _analysis(tmp_path, {"R": not_read, "Ctrl": {"A1": "BBA_01", "A2": "BBA_02", "A3": "VFA_01"}})
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "R"] == ["A1:A5", "B2:C5"]  # two grids, as Mind reads them


def test_two_lists_that_only_touch_are_not_joined():
    # different lengths: the left list runs on below the block
    occupied = {(1, 1): "Names", (2, 1): "ann", (3, 1): "bob", (4, 1): "cy", (5, 1): "dee", (2, 2): "Ages", (3, 2): 30, (4, 2): 31}
    grids, _ = detect_grids("S", occupied)
    assert [g["ref"] for g in grids] == ["A1:A5", "B2:B4"] and not any(g.get("caption_cell") for g in grids)
    # the block starts two rows below the column's top (two cells on top, not one caption)
    occupied = {(1, 1): "Cap", (2, 1): "Sub", (3, 1): "k1", (4, 1): "k2", (3, 2): 1, (4, 2): 2}
    grids, _ = detect_grids("S", occupied)
    assert [g["ref"] for g in grids] == ["A1:A4", "B3:B4"]
    # the column is already titled: it is a grid of its own by its author's choice
    occupied = {(1, 1): "#Keys", (2, 1): "Cap", (3, 1): "k1", (4, 1): "k2", (3, 2): 1, (4, 2): 2}
    grids, _ = detect_grids("S", occupied)
    assert [(g["ref"], g["name"]) for g in grids] == [("A2:A4", "Keys"), ("B3:B4", None)]
    # the column is two cells wide
    occupied = {(1, 1): "Cap", (1, 2): "x", (2, 1): "k1", (2, 2): "y", (3, 1): "k2", (3, 2): "z", (2, 3): 1, (3, 3): 2}
    grids, _ = detect_grids("S", occupied)
    assert [g["ref"] for g in grids] == ["A1:B3", "C2:C3"]


def test_an_empty_cell_above_gets_the_made_up_title(tmp_path):
    analysis = _analysis(tmp_path, {"D": {"B3": "Age", "C3": "Rate", "B4": 30, "C4": 0.01, "B5": 31, "C5": 0.02}})
    ops, skipped = _titles(analysis)
    assert [(o["op"], o["cell"], o["before"], o["after"]) for o in ops] == [("set_value", "B2", None, "#Age Rate")]
    assert not skipped


def test_a_formula_caption_is_reported_not_replaced(tmp_path):
    formula_caption = {**RESULT_CF, "A1": '="Running scenario ["&E1&"]"', "E1": "Base"}
    analysis = _analysis(tmp_path, {"R": formula_caption})
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "R"] == ["A2:C5"]
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"]
    mine = [s for s in skipped if s.startswith("R!A2:C5")]
    assert len(mine) == 1 and mine[0].startswith('R!A2:C5: the cell above it, A1, holds a formula (="Running scenario ["&E1&"]")') and "Move that formula somewhere else by hand" in mine[0], skipped


def test_nothing_is_ever_inserted_for_a_title(tmp_path):
    analysis = _analysis(tmp_path, {
        "Top": {"A1": "X", "B1": "Y", "A2": 1, "B2": 2},  # starts on row 1
        "CF": {"B2": 1, "C2": "Demographic Assumptions", "C3": "Gender", "D3": "M", "C4": "Smoker", "D4": "NS"},  # a heading row (a grid) right above
        "Stack": {"B1": "H", "C1": "K", "B2": 1, "C2": 2, "B3": 3, "C3": 4, "B5": "H2", "C5": "K2", "B6": 5, "C6": 6},  # a grid right above the title cell, same first column
        "Beside": {"A1": "H", "B1": "K", "C1": "L", "A2": 1, "B2": 2, "C2": 3, "A3": 3, "B3": 4, "C3": 5, "B5": "H2", "C5": "K2", "B6": 5, "C6": 6},  # the grid above starts in another column
    })
    ops, skipped = _titles(analysis)
    assert not any(o["op"] == "insert_row" for o in ops)
    assert not any(o.get("after_inserts") or o.get("insert_at") or o.get("columns") for o in ops)
    by_sheet = {o["sheet"]: o for o in ops}
    assert "Top" not in by_sheet and any(s.startswith("Top!A1:B2: starts on row 1") and "Insert an empty row above it by hand" in s for s in skipped), skipped
    assert "CF" not in by_sheet or by_sheet["CF"]["cell"] != "C2"
    assert any(s.startswith("CF!C3:D4: the cell above (C2) holds 'Demographic Assumptions' (part of the grid untitled B2, B2:C2)") and "nothing is moved" in s for s in skipped), skipped
    assert any(s.startswith("Stack!B5:C6: a title in B4 would touch B3") and "one grid" in s for s in skipped), skipped
    assert by_sheet["Beside"]["cell"] == "B4"  # A1:C3 is read down column A: a title in B4 stands on its own


def test_a_label_the_plan_moves_into_the_title_may_sit_right_above_it(tmp_path):
    analysis = _analysis(tmp_path, {"D": {"B2": "Premium table", "B4": "Key", "C4": "Value", "B5": "a", "C5": 1}})
    ops, skipped = _titles(analysis)
    assert [(o["op"], o["cell"], o["after"]) for o in ops] == [("set_value", "B3", "#Premium table"), ("clear_cell", "B2", None)]
    # ... but not when something reads the label (it stays, and would swallow the title below it)
    analysis = _analysis(tmp_path, {"D": {"B2": "Premium table", "B4": "Key", "C4": "Value", "B5": "a", "C5": 1}, "T": {"A1": "=D!B2"}})
    ops, skipped = _titles(analysis)
    assert not ops and any(s.startswith("D!B4:C5: a title in B3 would touch B2") for s in skipped), skipped


def test_a_cell_beside_the_title_is_moved_into_it_or_reported(tmp_path):
    """Mind reads left to right: a cell right beside the title cell is read first and runs on into
    the title. The old namer wrote such titles (the anonymized IFRS 17 tool: 33 grids with a
    '#Title' trapped inside after its own re-scan). A lone label nothing reads is moved into the
    title; a label something reads, or a grid starting on that row, is reported."""
    analysis = _analysis(tmp_path, {
        "L": {"B2": "Mortality", "C3": "Age", "D3": "Qx", "C4": 40, "D4": 0.001},  # lone label beside the title cell C2
        "R": {"B2": "Mortality", "C3": "Age", "D3": "Qx", "C4": 40, "D4": 0.001, "F1": "=B2"},  # ... that a formula reads
        "G": {"A2": "k", "B2": 1, "D3": "Age", "E3": "Qx", "D4": 40, "E4": 0.001, "C2": "x"},  # a grid A2:C2 ends right beside D2
    })
    ops, skipped = _titles(analysis)
    by_sheet: dict[str, list] = {}
    for o in ops:
        by_sheet.setdefault(o["sheet"], []).append((o["op"], o["cell"], o["after"]))
    assert by_sheet["L"] == [("set_value", "C2", "#Mortality"), ("clear_cell", "B2", None)]
    assert "R" not in by_sheet and any(s.startswith("R!C3:D4: a title in C2 would touch B2 ('Mortality') beside it, which is read by R!F1") for s in skipped), skipped
    assert not any(o[1] == "D2" for o in by_sheet.get("G", [])) and any(s.startswith("G!D3:E4: a title in D2 would touch C2 ('x') beside it, which is part of the grid") for s in skipped), skipped


def test_the_grid_namer_follows_the_same_rule(tmp_path):
    analysis = _analysis(tmp_path, {"R": RESULT_CF, "Top": {"A1": "X", "B1": "Y", "A2": 1, "B2": 2}})
    out = prep.plan_named_areas(analysis, [
        {"sheet": "R", "ref": "B3:C4", "name": "Result Cashflows", "flags": ["Input"]},
        {"sheet": "Top", "ref": "A1:B2", "name": "Top Grid", "flags": []},
    ])
    assert [(o["op"], o["sheet"], o["cell"], o["before"], o["after"]) for o in out["ops"]] == [("set_value", "R", "A1", "Running scenario", "#Result Cashflows /Input")]
    assert out["named"] == ["R!A2:C5"] and out["inserted_rows"] == {} and out["banded_rows"] == {}
    assert any(s.startswith("Top!A1:B2: starts on row 1") for s in out["skipped"]), out["skipped"]
    # selecting the caption cell alone names the table too
    out = prep.plan_named_areas(analysis, [{"sheet": "R", "ref": "A1", "name": "Result Cashflows", "flags": []}])
    assert [o["cell"] for o in out["ops"]] == ["A1"]
    # a formula caption is refused with the reason, for the user's name as for the automatic one
    analysis = _analysis(tmp_path, {"R": {**RESULT_CF, "A1": '="Running scenario ["&E1&"]"', "E1": "Base"}})
    out = prep.plan_named_areas(analysis, [{"sheet": "R", "ref": "B3", "name": "Result Cashflows", "flags": []}])
    assert not out["ops"] and any("holds a formula" in s for s in out["skipped"]), out["skipped"]


def test_an_old_plan_with_a_row_insert_is_still_ordered_and_resolved():
    old = [
        {"op": "insert_row", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "row": 1, "columns": "A:A"},
        {"op": "set_value", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "cell": "A1", "before": None, "after": "#T", "after_inserts": True, "insert_at": 1},
        {"op": "set_value", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "cell": "C1", "before": None, "after": "#U"},
    ]
    ordered = prep._ordered(old)
    assert [o["op"] for o in ordered] == ["set_value", "insert_row", "set_value"]
    assert ordered[0]["cell"] == "C1" and ordered[2]["cell"] == "A1"  # the banded insert at A does not move C1


@needs_excel
def test_excel_still_applies_an_old_plan_with_a_row_insert(tmp_path):
    path = _wb(tmp_path / "old.xlsx", {"S": {"A1": "k", "A2": "a", "A3": "b", "B2": 1, "B3": 2, "D1": "=A2&B2"}})
    old = [
        {"op": "insert_row", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "row": 1},
        {"op": "set_value", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "cell": "A1", "before": None, "after": "#T", "after_inserts": True, "insert_at": 1},
    ]
    out = prep.apply_operations(path, tmp_path / "prep", old)
    assert out["status"] == "APPLIED", out
    ws = openpyxl.load_workbook(out["output_path"])["S"]
    assert ws["A1"].value == "#T" and ws["A2"].value == "k" and ws["A3"].value == "a" and ws["B3"].value == 1
    assert ws["D2"].value == "=A3&B3"  # Excel followed the insert
