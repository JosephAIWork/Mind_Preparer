"""Mind instances (app.instance_builder): the macro's run loop as one copy of the model per key.

The plans are built on a small workbook invented here: a model that runs one row of a table at a
time. Its "Inputs" sheet holds the table with a text key column (a group) and, above it, a current
row of INDEX(table, driver, ...) formulas; "Calc" holds the driver cell and two results; "Runner"
holds the first and last row the macro reads from named cells, the block the macro copies out on
every pass, and the paste area it adds each pass's results to. The macro (read through a patched
`read_vba`, since an .xlsx carries none) sets the row number in the driver cell, looks the row's
group up with VLookup and pastes the block into the group's part of the paste area. Nothing here
comes from a client workbook."""
from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from app import model_map, prep
from app.excel_com import com_available, excel_text
from app.formula_utils import STORAGE_PREFIX_RE
from app.grids import detect_grids
from app.instance_builder import ACTION_ID, RULE_ID, build_instance_plan, describe, write_main_workbook, zip_for_mind
from app.inventory import build_analysis
from app.model_map import build_model_map
from app.prep import apply_operations

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")

MACRO = '''
Sub RunRows()
Dim RunStart As Long: RunStart = Range("RunStart").Value
Dim RunEnd As Long: RunEnd = Range("RunEnd").Value
Worksheets("Runner").Range("H20:M21").ClearContents
For rowNo = RunStart To RunEnd
    Range("POLICY_ROW").Value = rowNo
    Application.Calculate
    Temp1 = Application.WorksheetFunction.VLookup(rowNo, Range("Policies"), {k}, False)
    Temp2 = Application.WorksheetFunction.Match(Temp1, Range("GroupList"), 0)
    Range("ByGroupPaster").Offset(0, 2 * (Temp2 - 1)).Value2 = Range("AggregateOutput").Value2
Next rowNo
End Sub
'''

MACRO_WITHOUT_LOOKUP = '''
Sub RunRows()
Dim RunStart As Long: RunStart = Range("RunStart").Value
Dim RunEnd As Long: RunEnd = Range("RunEnd").Value
Worksheets("Runner").Range("H20:M21").ClearContents
For rowNo = RunStart To RunEnd
    Range("POLICY_ROW").Value = rowNo
    Application.Calculate
    Range("ByGroupPaster").Offset(0, 2 * (Temp2 - 1)).Value2 = Range("AggregateOutput").Value2
Next rowNo
End Sub
'''

MACRO_WITHOUT_COUNT_CELL = '''
Sub RunRows()
For rowNo = 1 To 8
    Range("POLICY_ROW").Value = rowNo
    Range("ByGroupPaster").Value2 = Range("AggregateOutput").Value2
Next rowNo
End Sub
'''

GROUPS = ["G1", "G1", "G1", "G2", "G2", "G3", "G3", "G3"]
REGIONS = ["North", "South"] * 4
DRIVER = '=MM_LOOPINSTANCE("POLICY_ROW",Runner!$E$5,Runner!$E$3)'


def _book(tmp_path: Path, *, key_first: bool = False, gap: int = 1, headers: list[str] | None = None, header: bool = True, groups: list | None = None,
          left_column: bool = False, table_ref: str | None = None, first: int = 1, last: int = 8, name: str = "model") -> Path:
    """The small model. `key_first`: the group is the table's first column. `gap`: empty rows between
    the current row and the table's header (1: a row has to be inserted for the title; 2: none).
    `header`: the table has a header row. `left_column`: column A holds a row number beside every
    row of the table (a column only the macro used). `table_ref`: the named range the macro indexes
    (default the data rows). `first` / `last`: what the macro reads as its first and last row."""
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Inputs"
    active = 1 if gap == 2 else 2
    data.cell(active, 1, "Current row")
    for i in range(4):
        data.cell(active, 2 + i, f"=INDEX(Policies,POLICY_ROW,{i + 1})")
    groups = list(groups if groups is not None else GROUPS)
    cols = [("Group", groups), ("Policy", list(range(1, 9))), ("Premium", [100 * (i + 1) for i in range(8)]), ("Region", REGIONS)] if key_first else \
           [("Policy", list(range(1, 9))), ("Group", groups), ("Premium", [100 * (i + 1) for i in range(8)]), ("Region", REGIONS)]
    for j, (head, values) in enumerate(cols):
        if header:
            data.cell(4, 2 + j, headers[j] if headers else head)
        for i, v in enumerate(values):
            if v is not None:
                data.cell(5 + i, 2 + j, v)
    if left_column:
        for i in range(8):
            data.cell(5 + i, 1, i + 1)
    calc = wb.create_sheet("Calc")
    calc["C2"], calc["D2"] = "Row to run", 7
    calc["B8"], calc["C8"] = f"=Inputs!D{active}*2", f"=Inputs!D{active}/2"
    run = wb.create_sheet("Runner")
    run["D2"], run["E2"], run["D3"], run["E3"] = "First row", first, "Last row", last
    run["B19"], run["C19"] = "Fee A", "Fee B"
    run["B20"], run["C20"] = "=Calc!B8+OFFSET(H20,0,0)", "=Calc!C8+OFFSET(I20,0,0)"
    run["B21"], run["C21"] = "=Calc!B8*2+OFFSET(H21,0,0)", "=Calc!C8*2+OFFSET(I21,0,0)"
    run["H20"], run["I21"], run["K20"] = 5, 6, 7  # left by an earlier run
    lists = wb.create_sheet("Lists")
    lists["B1"] = "Groups"
    for i, g in enumerate(["G1", "G2", "G3", "G4"]):
        lists.cell(2 + i, 2, g)
    for n, ref in (("RunStart", "Runner!$E$2"), ("RunEnd", "Runner!$E$3"), ("POLICY_ROW", "Calc!$D$2"), ("ByGroupPaster", "Runner!$H$20:$I$21"),
                   ("AggregateOutput", "Runner!$B$20:$C$21"), ("Policies", table_ref or "Inputs!$B$5:$E$12"), ("GroupList", "Lists!$B$2:$B$5")):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    p = tmp_path / f"{name}.xlsx"
    wb.save(p)
    wb.close()
    return p


def _plan_for(path: Path, tmp_path: Path, vba: str, monkeypatch, **kw) -> dict:
    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("Module1", vba)])
    analysis = build_analysis(path, tmp_path / "w", path.stem)
    return build_instance_plan(analysis, build_model_map(analysis), **kw)


def _by_cell(plan: dict) -> dict:
    return {(o["sheet"], o.get("cell") or o.get("ref")): o for o in plan["operations"] if o["op"] not in ("insert_row", "insert_column")}


def test_the_plan_for_a_model_that_runs_one_row_at_a_time(tmp_path, monkeypatch):
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["found"] is True and plan["status"] == "ready"
    # the loop: the counter, its count cell, the loop size cell found in a free row, the driver
    lp = plan["loop"]
    assert lp["name"] == "POLICY_ROW" and lp["kind"] == "counter" and lp["macro"] == "RunRows"
    assert lp["driver"] == "Calc!D2" and lp["count_cell"] == "Runner!E3" and lp["max_cell"] == "Runner!E5" and lp["key_cell"] == "Runner!G5"
    assert lp["formula"] == DRIVER and lp["count_formula"] == "=COUNTA(Inputs!$B$6:$B$13)" and lp["offset"] == 0
    # the table: the range the current row indexes through the driver, its header row above it
    t = plan["table"]
    assert (t["sheet"], t["ref"], t["header_row"], t["first_row"], t["last_row"], t["rows"], t["columns"], t["name"]) == ("Inputs", "B5:E12", 4, 5, 12, 8, 4, "Policies")
    # the key: the macro's own lookup column; the other text column is offered too
    k = plan["key"]
    assert (k["column"], k["header"], k["distinct"], k["chosen_by"], k["missing"]) == ("C", "Group", 3, "macro", 0)
    assert k["counts"] == {"G1": 3, "G2": 2, "G3": 3} and k["values"] == ["G1", "G2", "G3"]
    assert k["candidates"] == [{"column": "C", "header": "Group", "distinct": 3, "sampled": False}, {"column": "E", "header": "Region", "distinct": 2, "sampled": False}]
    # the split grid: one row inserted (the current row is two rows above the header), the key column inserted first
    sg = plan["split_grid"]
    assert sg == {"title_cell": "Inputs!B4", "title": "#Policies /InstanceSplit", "header_cell": "Inputs!B5", "inserted_rows": 1, "inserted_column": "B", "key_range": "Inputs!B6:B13", "reused_title": False}
    ops = plan["operations"]
    assert [o["op"] for o in ops[:5]] == ["insert_row", "insert_column", "set_value", "set_value", "set_range_values"]
    assert (ops[0]["sheet"], ops[0]["row"]) == ("Inputs", 4) and (ops[1]["sheet"], ops[1]["column"]) == ("Inputs", "B")
    assert (ops[2]["cell"], ops[2]["after"]) == ("B4", "#Policies /InstanceSplit") and (ops[3]["cell"], ops[3]["after"]) == ("B5", "Instance key")
    assert (ops[4]["ref"], ops[4]["cell"], ops[4]["values"]) == ("B6:B13", "B6:B13", [[g] for g in GROUPS])
    assert all(o.get("after_inserts") for o in ops[2:5]) and all(o["action_id"] == ACTION_ID and o["rule_id"] == RULE_ID for o in ops)
    by = _by_cell(plan)
    count = by[("Runner", "E3")]
    assert (count["op"], count["before"], count["after"], count.get("after_inserts")) == ("set_formula", 8, "=COUNTA(Inputs!$B$6:$B$13)", True)  # reads the table sheet: waits for the inserts
    assert (by[("Runner", "D5")]["after"], by[("Runner", "E5")]["after"]) == ("Loop size: rows of the biggest Group", 3)
    drv = by[("Calc", "D2")]
    assert (drv["before"], drv["after"], drv.get("after_inserts")) == (7, DRIVER, None)
    # the block the macro copies out on every pass: the sum over the copy's loop, cell by cell
    assert [(c, by[("Runner", c)]["after"]) for c in ("B20", "C20", "B21", "C21")] == [("B20", "=MM_RESULT(Calc!B8)"), ("C20", "=MM_RESULT(Calc!C8)"), ("B21", "=MM_RESULT(Calc!B8)"), ("C21", "=MM_RESULT(Calc!C8)")]
    assert by[("Runner", "B20")]["before"] == "=Calc!B8+OFFSET(H20,0,0)"
    out = plan["outputs"][0]
    assert (out["block"], out["source_sheet"], out["cells"], out["status"], out["example"]) == ("Runner!B20:C21", "Calc", 4, "ready", "Runner!B20 =MM_RESULT(Calc!B8)")
    # the paste area: what an earlier run left there goes
    cleared = sorted(c for (s, c), o in by.items() if o["op"] == "clear_cell")
    assert cleared == ["H20", "I21", "K20"] and plan["clears"] == [{"area": "Runner!H20:M21", "cells": 3, "status": "ready", "over": 0, "why": None}]
    # the copy's key, shown next to the loop size
    assert (by[("Runner", "F5")]["after"], by[("Runner", "G5")]["after"]) == ("This copy's Group", '=MM_INSTANCEKEY("G1")')
    assert plan["counts"] == {"key_values": 3, "result_cells": 4, "cleared": 3, "inserted_rows": 1, "inserted_columns": 1, "operations": len(ops)} and plan["operation_count"] == len(ops) == 18
    # the main workbook: the keys, the results by key, the whole block as plain links (the sum over the copies in Mind)
    m = plan["main"]
    assert (m["file"], m["model_file"], m["keys"], m["key_header"]) == ("model_Main.xlsx", "model_instances.xlsx", ["G1", "G2", "G3"], "Group")
    assert [(g["title"], g["sheet"], g["rows"], g["cols"]) for g in m["grids"]] == [
        ("#Group instances /InstanceKeys", "Instances", 3, 1), ("#Results by Group", "Summary", 3, 3), ("#All copies", "Summary", 4, 3)]
    assert "plain links" in m["grids"][2]["what"] and "sum over the 3 copies" in m["grids"][2]["what"]
    assert m["block"]["ref"] == "B20:C21" and m["block"]["headers"] == ["Fee A", "Fee B"] and m["block"]["row_labels"] == ["Row 1", "Row 2"]
    assert "one copy of model_instances.xlsx per Group (3 copies" in plan["how_it_runs"] and "RunRows is not needed" in plan["how_it_runs"] and "adds them up" in plan["how_it_runs"]
    assert plan["notes"] == []
    json.dumps(plan)
    text = describe(plan)
    assert "key column C ('Group', 3 values, chosen by macro)" in text and "operations: 18" in text


def test_every_key_of_the_list_the_macro_runs_through_can_be_asked_for(tmp_path, monkeypatch):
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch, keys="all")
    assert plan["status"] == "ready" and plan["main"]["keys"] == ["G1", "G2", "G3", "G4"] and plan["key"]["counts"] == {"G1": 3, "G2": 2, "G3": 3}
    assert plan["counts"]["key_values"] == 4 and plan["main"]["grids"][0]["rows"] == 4
    assert any("have no row in the table" in n and "G4" in n for n in plan["notes"])
    assert _by_cell(plan)[("Runner", "G5")]["after"] == '=MM_INSTANCEKEY("G1")'
    # the model's file stem can be chosen (it is what the keys grid's header must say)
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch, model_name="My model v2")
    assert plan["main"]["model_file"] == "My_model_v2.xlsx" and plan["main"]["model_stem"] == "My_model_v2"


def test_without_the_macros_lookup_a_header_word_picks_the_key_else_the_fewest_values(tmp_path, monkeypatch):
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO_WITHOUT_LOOKUP, monkeypatch)
    assert plan["status"] == "ready" and (plan["key"]["column"], plan["key"]["header"], plan["key"]["chosen_by"]) == ("C", "Group", "heuristic")
    assert any("list of keys the macro runs through was not found" in n for n in _plan_for(_book(tmp_path), tmp_path, MACRO_WITHOUT_LOOKUP, monkeypatch, keys="all")["notes"])
    # no header names a group: the text column with the fewest distinct values
    plan = _plan_for(_book(tmp_path, headers=["Policy", "Code", "Premium", "Zone"], name="coded"), tmp_path, MACRO_WITHOUT_LOOKUP, monkeypatch)
    assert (plan["key"]["column"], plan["key"]["header"], plan["key"]["distinct"], plan["key"]["chosen_by"]) == ("E", "Zone", 2, "heuristic")
    assert plan["key"]["counts"] == {"North": 4, "South": 4} and plan["split_grid"]["inserted_column"] == "B"
    assert _by_cell(plan)[("Runner", "D5")]["after"] == "Loop size: rows of the biggest Zone"


def test_the_users_key_wins_by_header_or_by_letter(tmp_path, monkeypatch):
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch, key="region")
    assert (plan["key"]["column"], plan["key"]["header"], plan["key"]["chosen_by"], plan["key"]["counts"]) == ("E", "Region", "user", {"North": 4, "South": 4})
    assert _by_cell(plan)[("Inputs", "B6:B13")]["values"] == [[r] for r in REGIONS]
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch, key="E")
    assert (plan["key"]["column"], plan["key"]["chosen_by"]) == ("E", "user")
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch, key="Nope")
    assert plan["status"] == "needs_decision" and plan["operations"] == [] and any("No column of the table is called 'Nope'" in n for n in plan["notes"])
    assert plan["key"]["candidates"] and plan["key"]["column"] is None and plan["table"]["sheet"] == "Inputs"


def test_a_workbook_without_a_loop_has_nothing_to_build(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    p = tmp_path / "plain.xlsx"
    wb.save(p)
    wb.close()
    analysis = build_analysis(p, tmp_path / "w", "plain")
    plan = build_instance_plan(analysis, build_model_map(analysis))
    assert plan["found"] is False and plan["status"] == "not_built" and plan["operations"] == [] and "No run loop" in plan["notes"][0]
    assert plan["main"] == {"file": "plain_Main.xlsx", "model_file": "plain_instances.xlsx", "keys": [], "grids": []}
    assert describe(plan).startswith("status: not_built")


def test_a_counter_without_a_count_cell_needs_a_decision_first(tmp_path, monkeypatch):
    plan = _plan_for(_book(tmp_path), tmp_path, MACRO_WITHOUT_COUNT_CELL, monkeypatch)
    assert plan["found"] is True and plan["status"] == "needs_decision" and plan["operations"] == []
    assert plan["loop"]["name"] == "POLICY_ROW" and any("no cell holds the count" in n and "Mind loop panel" in n for n in plan["notes"])


def test_a_range_that_starts_on_its_header_row_offsets_the_driver(tmp_path, monkeypatch):
    """The macro indexes a range that includes the header row and runs 2..n. MM_LOOPINSTANCE counts
    from 1, so the driver adds the header row: every pass reads a data row, and the count (COUNTA of
    the key column, data rows only) is still the copy's row count."""
    plan = _plan_for(_book(tmp_path, table_ref="Inputs!$B$4:$E$12", first=2, last=9), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["status"] == "ready" and (plan["loop"]["first"], plan["loop"]["last"], plan["loop"]["offset"]) == (2, 9, 1)
    assert plan["loop"]["formula"] == DRIVER + "+1" and plan["loop"]["count_formula"] == "=COUNTA(Inputs!$B$6:$B$13)"
    assert (plan["table"]["header_row"], plan["table"]["first_row"], plan["table"]["rows"]) == (4, 5, 8)
    drv = _by_cell(plan)[("Calc", "D2")]
    assert drv["after"] == DRIVER + "+1" and "in Excel it shows 2" in drv["note"]
    assert any("starts on the table's header row" in n and "adds 1" in n for n in plan["notes"]) and not any("holds 8" in n for n in plan["notes"])
    # the macro ran the header row too (1..9): the copies start at the first data row, said in a note
    plan = _plan_for(_book(tmp_path, table_ref="Inputs!$B$4:$E$12", first=1, last=9, name="fromone"), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["loop"]["formula"] == DRIVER + "+1" and any("read the table's header row" in n for n in plan["notes"])
    # a range of data rows with a macro that skips the first two: no offset (every row joins its copy), said in a note
    plan = _plan_for(_book(tmp_path, first=3, last=8, name="skip"), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["loop"]["formula"] == DRIVER and plan["loop"]["offset"] == 0
    assert any("2 data row(s) into the table" in n for n in plan["notes"]) and any("runs 6 row(s); the table holds 8" in n for n in plan["notes"])


def test_without_a_header_row_the_key_column_gets_one_header_only(tmp_path, monkeypatch):
    """No header row and the key column already first: the header row is made (a row inserted), the key
    column's header is "Instance key" -- written once -- and the other columns get their letters."""
    plan = _plan_for(_book(tmp_path, key_first=True, header=False), tmp_path, MACRO.format(k=1), monkeypatch)
    assert plan["status"] == "ready" and (plan["table"]["header_row"], plan["table"]["first_row"]) == (None, 5)
    sg = plan["split_grid"]
    assert (sg["title_cell"], sg["header_cell"], sg["inserted_rows"], sg["inserted_column"], sg["key_range"]) == ("Inputs!B4", "Inputs!B5", 1, None, "Inputs!B6:B13")
    headers = [(o["cell"], o["after"]) for o in plan["operations"] if o["op"] == "set_value" and o["sheet"] == "Inputs" and o["cell"].endswith("5")]
    assert headers == [("B5", "Instance key"), ("C5", "Column C"), ("D5", "Column D"), ("E5", "Column E")]
    assert not any(o["op"] == "set_range_values" for o in plan["operations"])  # the keys are already in the first column
    assert any("had no header row" in n for n in plan["notes"])


def test_a_filled_column_left_of_the_table_joins_the_grid(tmp_path, monkeypatch):
    """Column A holds a row number beside every row (a column only the macro used). Mind reads cells
    touching a grid as part of it, so the key column is inserted before column A, the title sits above
    it, and the old column A gets a text header (a grid's first row is its header)."""
    plan = _plan_for(_book(tmp_path, left_column=True), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["status"] == "ready" and plan["table"]["columns"] == 5 and plan["table"]["ref"] == "A5:E12"
    sg = plan["split_grid"]
    assert (sg["title_cell"], sg["header_cell"], sg["inserted_rows"], sg["inserted_column"], sg["key_range"]) == ("Inputs!A4", "Inputs!A5", 1, "A", "Inputs!A6:A13")
    assert plan["key"]["column"] == "C" and plan["loop"]["count_formula"] == "=COUNTA(Inputs!$A$6:$A$13)"
    by = _by_cell(plan)
    assert by[("Inputs", "A5")]["after"] == "Instance key" and by[("Inputs", "B5")]["after"] == "Column A" and by[("Inputs", "A6:A13")]["values"] == [[g] for g in GROUPS]
    assert any("1 filled column(s) left of the table (A) touch it" in n for n in plan["notes"])
    # a label two rows above the table does not count: only the header and data rows are looked at
    plan = _plan_for(_book(tmp_path, name="label"), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["split_grid"]["inserted_column"] == "B" and plan["table"]["columns"] == 4


def test_text_keys_that_look_like_numbers_stay_text(tmp_path, monkeypatch):
    """Keys like "001" and "1" are two keys: the plan keeps them as text, and excel_text (what every
    write through Excel goes through) prefixes what Excel would turn into a number, a date or a boolean."""
    codes = ["001", "001", "1", "1", "010", "010", "10", "10"]
    plan = _plan_for(_book(tmp_path, groups=codes, name="codes"), tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["status"] == "ready" and plan["key"]["counts"] == {"001": 2, "1": 2, "010": 2, "10": 2} and plan["main"]["keys"] == ["001", "1", "010", "10"]
    assert _by_cell(plan)[("Inputs", "B6:B13")]["values"] == [[c] for c in codes]
    assert [excel_text(v) for v in ("001", "1E3", "2024-01", "10:30", "-5", "-x", " TRUE", "$5", "=A1", "G1", "Fee A", "")] == \
           ["'001", "'1E3", "'2024-01", "'10:30", "'-5", "'-x", "' TRUE", "'$5", "'=A1", "G1", "Fee A", ""]
    assert excel_text(5) == 5 and excel_text(None) == "" and excel_text(2.5) == 2.5


def test_the_operations_apply_without_excel_when_nothing_has_to_be_inserted(tmp_path, monkeypatch):
    """The key column already first and two empty rows above the header: no insert, so openpyxl can
    write the plan (the inserts need Excel, which shifts every reference). The table sheet then shows
    the split grid, named from its title, its first column the key."""
    path = _book(tmp_path, key_first=True, gap=2)
    plan = _plan_for(path, tmp_path, MACRO.format(k=1), monkeypatch)
    assert plan["status"] == "ready" and (plan["key"]["column"], plan["key"]["chosen_by"]) == ("B", "macro")
    sg = plan["split_grid"]
    assert (sg["title_cell"], sg["inserted_rows"], sg["inserted_column"], sg["key_range"]) == ("Inputs!B3", 0, None, "Inputs!B5:B12")
    assert not any(o["op"] in ("insert_row", "insert_column", "set_range_values") or o.get("after_inserts") for o in plan["operations"])
    assert ("Inputs", "B4") not in _by_cell(plan)  # the key column's own header ("Group") is kept
    out = apply_operations(path, tmp_path / "apply", plan["operations"], prefer_excel=False)
    assert out["status"] == "APPLIED", out
    wb = openpyxl.load_workbook(out["output_path"])
    ws = wb["Inputs"]
    occupied = {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if c.value is not None}
    grids, _standalone = detect_grids("Inputs", occupied)
    g = next(x for x in grids if x["name"] == "Policies")
    assert g["flag_names"] == ["instancesplit"] and g["title_cell"] == "B3" and g["header_values"][0] == "Group" and (g["first_row"], g["last_row"]) == (4, 12)
    assert wb["Runner"]["E3"].value == "=COUNTA(Inputs!$B$5:$B$12)" and wb["Runner"]["E5"].value == 3
    assert wb["Calc"]["D2"].value == DRIVER
    assert wb["Runner"]["B20"].value == "=MM_RESULT(Calc!B8)" and wb["Runner"]["C21"].value == "=MM_RESULT(Calc!C8)"
    assert wb["Runner"]["H20"].value is None and wb["Runner"]["I21"].value is None and wb["Runner"]["K20"].value is None
    assert wb["Runner"]["G5"].value == '=MM_INSTANCEKEY("G1")'
    wb.close()


def test_set_range_values_is_written_by_openpyxl_and_logged_by_count(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.active["A1"] = "old"
    p = tmp_path / "range.xlsx"
    wb.save(p)
    wb.close()
    op = {"op": "set_range_values", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": "S", "ref": "A1:B3", "cell": "A1:B3", "values": [["k1", 1], ["k2", None], ["k3", 3]], "note": "test"}
    applied, failed = prep._apply_with_openpyxl(p, [op])
    assert failed == [] and len(applied) == 1 and "values" not in applied[0] and applied[0]["values_count"] == 3
    wb = openpyxl.load_workbook(p)
    assert [[c.value for c in row] for row in wb["S"]["A1:B3"]] == [["k1", 1], ["k2", None], ["k3", 3]]
    wb.close()
    # the shape is checked
    bad = {**op, "values": [["k1"], ["k2"], ["k3"]]}
    applied, failed = prep._apply_with_openpyxl(p, [bad])
    assert applied == [] and len(failed) == 1 and "one list per row" in failed[0]["error"]
    # the executor orders it with the other cell writes, before or after the inserts as marked
    ordered = prep._ordered([{**op, "after_inserts": True}, {"op": "insert_row", "sheet": "S", "row": 2}, op])
    assert [(o["op"], bool(o.get("after_inserts"))) for o in ordered] == [("set_range_values", False), ("insert_row", False), ("set_range_values", True)]
    assert prep._op_target({"op": "set_range_values", "sheet": "S", "ref": "A1:B3"}) == "S!A1:B3"
    # the change log counts the values instead of copying them (a fresh file: the first one already holds them)
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    p2 = tmp_path / "range2.xlsx"
    wb.save(p2)
    wb.close()
    res = apply_operations(p2, tmp_path / "log", [op], prefer_excel=False)
    assert res["status"] == "APPLIED" and res["change_log_entry"]["applied"][0]["values_count"] == 3 and "values" not in res["change_log_entry"]["applied"][0]


class _FakeCell:
    def __init__(self, store: dict, r: int, c: int):
        self.store, self.r, self.c = store, r, c

    @property
    def Value2(self):
        return self.store.get((self.r, self.c))


class _FakeRange:
    """Range.Value2 as Excel does it: '' assigned leaves an empty cell; `refuse` = a protected range
    that takes the assignment without an error and keeps nothing."""
    MergeCells = False

    def __init__(self, store: dict, refuse: bool = False):
        self.store, self.refuse = store, refuse

    def _set(self, v):
        if self.refuse:
            return
        for i, row in enumerate(v):
            for j, x in enumerate(row):
                self.store[(i + 1, j + 1)] = None if x == "" else x

    Value2 = property(lambda self: None, _set)

    def Cells(self, r: int, c: int) -> _FakeCell:
        return _FakeCell(self.store, r, c)


class _FakeSheet:
    def __init__(self, refuse: bool = False):
        self.store: dict = {}
        self.refuse = refuse

    def Range(self, _ref: str) -> _FakeRange:
        return _FakeRange(self.store, self.refuse)


def test_set_range_values_reads_back_the_last_planned_value_not_the_last_cell():
    """A key column whose last rows have no key ends in empty cells by design: the read-back after the
    write looks at the last cell that was planned to hold something, so the write is not refused."""
    def op(values):
        return {"op": "set_range_values", "action_id": ACTION_ID, "rule_id": RULE_ID, "sheet": "Inputs", "ref": f"B6:B{5 + len(values)}", "cell": f"B6:B{5 + len(values)}", "values": values}

    ws = _FakeSheet()
    assert prep._write_range_values(ws, op([["G1"], ["G2"], ["G3"]]))["values_count"] == 3 and ws.store[(3, 1)] == "G3"
    ws = _FakeSheet()
    assert prep._write_range_values(ws, op([["G1"], ["G2"], [None], [""]]))["values_count"] == 4 and ws.store[(2, 1)] == "G2" and ws.store.get((3, 1)) is None
    # text Excel would read as a number goes with the text prefix (the cell keeps "001", not 1)
    ws = _FakeSheet()
    prep._write_range_values(ws, op([["001"], ["1"]]))
    assert ws.store[(1, 1)] == "'001" and ws.store[(2, 1)] == "'1"
    # an all-empty plan has nothing to read back; a range that takes the write and keeps nothing is refused
    assert prep._write_range_values(_FakeSheet(), op([[None], [None]]))["values_count"] == 2
    with pytest.raises(ValueError, match="still empty"):
        prep._write_range_values(_FakeSheet(refuse=True), op([["G1"], ["G2"], [None]]))


@needs_excel
def test_the_split_grid_is_written_through_excel(tmp_path, monkeypatch):
    """The inserts and everything after them, through Excel: the row and the key column inserted, the
    title above the header, the keys in the new first column, every reference shifted by Excel."""
    path = _book(tmp_path)
    plan = _plan_for(path, tmp_path, MACRO.format(k=2), monkeypatch)
    out = apply_operations(path, tmp_path / "apply", plan["operations"], prefer_excel=True)
    assert out["status"] == "APPLIED" and out["method"] == "excel_com", out
    analysis = build_analysis(Path(out["output_path"]), tmp_path / "w2", "out")
    data = next(s for s in analysis["workbooks"][0]["sheets"] if s["name"] == "Inputs")
    g = next(x for x in data["grids"] if x["name"] == "Policies")
    assert g["flag_names"] == ["instancesplit"] and g["title_cell"] == "B4" and g["header_values"] == ["Instance key", "Policy", "Group", "Premium", "Region"]
    assert (g["first_row"], g["last_row"], g["first_col"], g["last_col"]) == (5, 13, 2, 6)
    wb = openpyxl.load_workbook(out["output_path"])
    assert [wb["Inputs"].cell(r, 2).value for r in range(6, 14)] == GROUPS
    assert wb["Inputs"]["C2"].value == "=INDEX(Policies,POLICY_ROW,1)" and wb.defined_names["Policies"].attr_text == "Inputs!$C$6:$F$13"  # shifted by Excel
    assert wb["Runner"]["E3"].value == "=COUNTA(Inputs!$B$6:$B$13)"
    assert STORAGE_PREFIX_RE.sub("", wb["Calc"]["D2"].value) == DRIVER
    assert STORAGE_PREFIX_RE.sub("", wb["Runner"]["B20"].value) == "=MM_RESULT(Calc!B8)" and wb["Runner"]["H20"].value is None
    wb.close()


@needs_excel
def test_a_table_with_a_filled_column_on_its_left_and_text_keys_through_excel(tmp_path, monkeypatch):
    """Through Excel: the key column goes before the row-number column, which gets a header, so the
    grid Mind sees starts with the key; keys that look like numbers stay text, and the last rows of
    the table without a key are written (and stay empty) without a refusal."""
    codes = ["001", "001", "1", "1", "010", "010", None, None]
    path = _book(tmp_path, left_column=True, groups=codes, name="codes")
    plan = _plan_for(path, tmp_path, MACRO.format(k=2), monkeypatch)
    assert plan["split_grid"]["inserted_column"] == "A" and _by_cell(plan)[("Inputs", "A6:A13")]["values"] == [[c] for c in codes]
    assert any("2 row(s) of the table have no Group" in n for n in plan["notes"])
    out = apply_operations(path, tmp_path / "apply", plan["operations"], prefer_excel=True)
    assert out["status"] == "APPLIED" and out["method"] == "excel_com", out
    analysis = build_analysis(Path(out["output_path"]), tmp_path / "w2", "out")
    data = next(s for s in analysis["workbooks"][0]["sheets"] if s["name"] == "Inputs")
    g = next(x for x in data["grids"] if x["name"] == "Policies")
    assert g["flag_names"] == ["instancesplit"] and g["title_cell"] == "A4" and g["header_values"] == ["Instance key", "Column A", "Policy", "Group", "Premium", "Region"]
    wb = openpyxl.load_workbook(out["output_path"])
    assert [wb["Inputs"].cell(r, 1).value for r in range(6, 14)] == codes and [wb["Inputs"].cell(r, 2).value for r in range(6, 14)] == list(range(1, 9))
    assert wb["Runner"]["E3"].value == "=COUNTA(Inputs!$A$6:$A$13)"
    wb.close()


@needs_excel
def test_the_main_workbook_and_the_zip(tmp_path, monkeypatch):
    path = _book(tmp_path)
    plan = _plan_for(path, tmp_path, MACRO.format(k=2), monkeypatch)
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    model_path = out_dir / plan["main"]["model_file"]
    shutil.copyfile(path, model_path)
    with pytest.raises(ValueError):
        write_main_workbook(plan, out_dir / "other.xlsx", out_dir / plan["main"]["file"])
    main_path = out_dir / plan["main"]["file"]
    done = write_main_workbook(plan, model_path, main_path)
    # the keys grid, the results by key, the whole block as plain links with a Total row
    assert done["grids"] == 3 and done["formulas"] == 3 * 2 + (2 * 2 + 2) and main_path.is_file()
    wb = openpyxl.load_workbook(main_path)
    inst = wb["Instances"]
    assert inst["B4"].value == "#Group instances /InstanceKeys" and inst["B5"].value == "model_instances" and [inst.cell(r, 2).value for r in (6, 7, 8)] == ["G1", "G2", "G3"]
    summ = wb["Summary"]
    cells = {c.value: c for row in summ.iter_rows() for c in row if isinstance(c.value, str)}
    results = cells["#Results by Group"]
    assert summ.cell(results.row + 1, 2).value == "Group" and summ.cell(results.row + 1, 3).value == "Row 1" and summ.cell(results.row + 2, 2).value == "G1"
    f = STORAGE_PREFIX_RE.sub("", summ.cell(results.row + 2, 3).value)
    assert f.startswith("=MM_INSTANCE([1]Runner!$C$20,$B") and f.endswith(")")  # [1] = the external link to the model file
    links = [getattr(getattr(link, "file_link", None), "Target", "") for link in wb._external_links]
    assert any(str(t).endswith("model_instances.xlsx") for t in links), links
    # the whole block as plain links to the model (in Mind: the sum over the copies), then a Total row
    all_copies = cells["#All copies"]
    assert all_copies.row > results.row and summ.cell(all_copies.row + 1, 3).value == "Fee A" and summ.cell(all_copies.row + 2, 2).value == "Row 1"
    assert STORAGE_PREFIX_RE.sub("", summ.cell(all_copies.row + 2, 3).value) == "=[1]Runner!B20"
    assert STORAGE_PREFIX_RE.sub("", summ.cell(all_copies.row + 3, 4).value) == "=[1]Runner!C21"
    total = cells["Total"]
    assert total.row == all_copies.row + 4 and summ.cell(total.row, 3).value.startswith("=SUM(")
    wb.close()
    zip_path = zip_for_mind(model_path, main_path, out_dir / "both.zip")
    with zipfile.ZipFile(zip_path) as z:
        assert sorted(z.namelist()) == sorted([model_path.name, main_path.name])


@needs_excel
def test_the_main_workbook_keeps_text_keys_as_text(tmp_path, monkeypatch):
    """"001" and "1" are two keys in the main workbook too (written as text, like the model's key column)."""
    codes = ["001", "001", "1", "1", "010", "010", "10", "10"]
    path = _book(tmp_path, groups=codes, name="codes")
    plan = _plan_for(path, tmp_path, MACRO.format(k=2), monkeypatch)
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    model_path = out_dir / plan["main"]["model_file"]
    shutil.copyfile(path, model_path)
    main_path = out_dir / plan["main"]["file"]
    write_main_workbook(plan, model_path, main_path)
    wb = openpyxl.load_workbook(main_path)
    assert [wb["Instances"].cell(r, 2).value for r in range(6, 10)] == ["001", "1", "010", "10"]
    summ = wb["Summary"]
    results = next(c for row in summ.iter_rows() for c in row if c.value == "#Results by Group")
    assert [summ.cell(results.row + 2 + i, 2).value for i in range(4)] == ["001", "1", "010", "10"]
    assert next(c for row in summ.iter_rows() for c in row if c.value == "#All copies") is not None
    wb.close()



# --- copies from loops: one copy per combination of the values of two loops -------------------------
SCENARIO_MACRO = """
Sub AllRuns()
Dim scenarioGrid As Range
Set scenarioGrid = ThisWorkbook.Names("ScenarioGrid").RefersToRange
Dim numScenario As Long
numScenario = scenarioGrid(1, 2).End(xlDown).Row - scenarioGrid(1, 2).Row + 1
Dim startY As Long: startY = Range("StartCohort").Value
Dim endY As Long: endY = Range("CalcYear").Value
For cohortYear = startY To endY
    Range("CohortRng").Value = cohortYear
    For i = 0 To numScenario - 1
        Range("Scenario").Value = i
        Calculate
    Next i
Next cohortYear
End Sub

Sub RollRun()
Dim startY As Long: startY = Range("StartCohort").Value
Dim endY As Long: endY = Range("CalcYear").Value
For y = startY To endY
    Range("RollYear").Value = y
    Calculate
Next y
End Sub
"""


def _scenario_book(tmp_path: Path, *, stale: bool = False, driver_loop: bool = False) -> Path:
    """A model run per cohort and scenario (the cohort a counter, the scenario a dropdown list) with a
    second macro that runs a roll-forward year: the shape of an IFRS 17 tool, invented here. "Result"
    holds the block the main workbook reads. `stale`: a cell still sums the Scenario loop (an earlier
    loop build). `driver_loop`: the cohort cell already holds the Mind loop."""
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ctrl = wb.active
    ctrl.title = "Ctrl"
    ctrl["B2"], ctrl["C2"] = "Calc year", 2024
    ctrl["B3"], ctrl["C3"] = "First cohort", 2022
    ctrl["B5"], ctrl["C5"] = "Cohort", '=MM_LOOP("CohortRng",Ctrl!C2,2022,Ctrl!C3,Ctrl!C2)' if driver_loop else 2024
    ctrl["B7"], ctrl["C7"] = "Scenario", 1
    ctrl["B9"] = "Scenarios"
    for i in range(3):
        ctrl.cell(9 + i, 3, i)
        ctrl.cell(9 + i, 4, ["Base", "Up", "Down"][i])
    dv = DataValidation(type="list", formula1="$C$9:$C$11")
    ctrl.add_data_validation(dv)
    dv.add("C7")
    res = wb.create_sheet("Result")
    res["A2"], res["B2"], res["C2"] = "Item", "Value", "Double"
    for r, (label, f) in enumerate([("Cohort", "=Ctrl!C5"), ("Scenario", "=Ctrl!C7"), ("Both", "=Ctrl!C5*10+Ctrl!C7")], start=3):
        res.cell(r, 1, label)
        res.cell(r, 2, f)
        res.cell(r, 3, f"=B{r}*2")
    roll = wb.create_sheet("Roll")
    roll["B2"], roll["C2"] = "Roll year", 2024
    roll["C3"] = "=C2+1"
    if stale:
        roll["E2"] = '=MM_RESULT(Result!B3,"Scenario",1)'
    for n, ref in (("CalcYear", "Ctrl!$C$2"), ("StartCohort", "Ctrl!$C$3"), ("CohortRng", "Ctrl!$C$5"), ("Scenario", "Ctrl!$C$7"), ("ScenarioGrid", "Ctrl!$B$9"),
                   ("RollYear", "Roll!$C$2")):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    p = tmp_path / "ifrs.xlsx"
    wb.save(p)
    wb.close()
    return p


KEYS9 = [f"CohortRng{y}_Scenario{s}" for y in (2022, 2023, 2024) for s in (0, 1, 2)]


def test_copies_from_two_loops_a_split_grid_of_their_values_and_the_other_macros_loop_is_left_out(tmp_path, monkeypatch):
    path = _scenario_book(tmp_path)
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch)
    # without loops asked for: the table way finds no table, and says which loops copies can come from
    assert plan["mode"] == "table" and plan["status"] != "ready"
    opts = {o["name"]: o for o in plan["loop_options"]}
    assert set(opts) == {"CohortRng", "Scenario", "RollYear"} and all(o["usable"] for o in opts.values())
    assert (opts["CohortRng"]["kind"], opts["CohortRng"]["count"], opts["Scenario"]["kind"], opts["Scenario"]["count"]) == ("counter", 3, "list", 3)
    assert any("Copies from loops" in n and "CohortRng" in n for n in plan["notes"])

    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["CohortRng", "Ctrl!$C$7"])
    assert plan["status"] == "ready" and plan["mode"] == "loops" and plan["found"] is True
    assert [(c["name"], c["driver"], c["values"]) for c in plan["copies_from"]] == [("CohortRng", "Ctrl!C5", [2022, 2023, 2024]), ("Scenario", "Ctrl!C7", [0, 1, 2])]
    # the keys: every combination, the first loop outside (as the macro nests them)
    assert plan["key"]["values"] == KEYS9 and plan["main"]["keys"] == KEYS9 and plan["counts"]["key_values"] == 9
    # the split grid: below the used range of the first loop's sheet, the key first, then each loop's value
    sg = plan["split_grid"]
    assert (sg["title"], sg["title_cell"], sg["header_cell"], sg["key_range"], sg["key_cell"]) == ("#Instance values /InstanceSplit", "Ctrl!B16", "Ctrl!B17", "Ctrl!B18:B26", "Ctrl!C14")
    by = _by_cell(plan)
    assert (by[("Ctrl", "B14")]["after"], by[("Ctrl", "C14")]["after"]) == ("This copy's key", '=MM_INSTANCEKEY("CohortRng2022_Scenario0")')
    assert by[("Ctrl", "B16")]["after"] == "#Instance values /InstanceSplit"
    grid = by[("Ctrl", "B17:D26")]
    assert grid["op"] == "set_range_values" and grid["values"][0] == ["Instance key", "CohortRng", "Scenario"]
    assert grid["values"][1:] == [[k, y, s] for k, (y, s) in zip(KEYS9, [(y, s) for y in (2022, 2023, 2024) for s in (0, 1, 2)])]
    # each chosen driver reads the copy's one row; the scenario's dropdown goes (Mind would drop the formula)
    assert (by[("Ctrl", "C5")]["before"], by[("Ctrl", "C5")]["after"]) == (2024, "=INDEX(Ctrl!$C$18:$C$26,1)")
    fx = {(o["sheet"], o.get("cell")): o for o in plan["operations"] if o["op"] == "set_formula"}
    assert (fx[("Ctrl", "C7")]["before"], fx[("Ctrl", "C7")]["after"]) == (1, "=INDEX(Ctrl!$D$18:$D$26,1)")
    assert [o["op"] for o in plan["operations"] if o["sheet"] == "Ctrl" and o.get("cell") == "C7"] == ["set_formula", "remove_validation"]
    # the other macro's loop is not written into the copies (review 2026-11-10): Excel runs RollRun on its own,
    # with the cohort and scenario fixed; inside every copy Mind would run it and add its passes up
    assert ("Roll", "C2") not in fx and not any(o["sheet"] == "Roll" for o in plan["operations"])
    assert plan["kept_loops"] == [] and any("The loops of the macro RollRun (RollYear) are not part of these copies" in n for n in plan["notes"])
    assert all(o["action_id"] == ACTION_ID for o in plan["operations"] if o.get("loop") is None)
    # the main workbook: the keys and every copy's results block (the output sheet's formula block), no sum over the copies
    m = plan["main"]
    assert (m["layout"], m["key_header"], m["model_file"]) == ("by_key_rows", "CohortRng and Scenario", "ifrs_instances.xlsx")
    assert m["block"]["ref"] == "B3:C5" and m["block"]["sheet"] == "Result" and m["block"]["headers"] == ["Value", "Double"] and m["block"]["row_labels"] == ["Row 3", "Row 4", "Row 5"]
    assert [(g["title"], g["rows"], g["cols"]) for g in m["grids"]] == [("#CohortRng and Scenario instances /InstanceKeys", 9, 1), ("#Results by CohortRng and Scenario", 27, 4)]
    assert plan["results"] == {"block": "Result!B3:C5", "rows": 3, "cols": 2, "chosen_by": "outputs"} and plan["counts"]["result_cells"] == 54
    assert any("add every copy up" in n for n in plan["notes"]) and any("Ctrl!C7 had a dropdown" in n for n in plan["notes"])
    assert "9 copies" in plan["how_it_runs"] and "stay" not in plan["how_it_runs"] and "AllRuns is not needed" in plan["how_it_runs"]
    json.dumps(plan)
    assert "status: ready" in describe(plan)


def test_copies_from_loops_can_be_written_without_excel(tmp_path, monkeypatch):
    """Nothing is inserted, so openpyxl can write the plan: Ctrl then shows the split grid, flagged, its first
    column the key; the drivers read its first row (the first copy, as Excel shows it)."""
    path = _scenario_book(tmp_path)
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["CohortRng", "Scenario"])
    out = apply_operations(path, tmp_path / "apply", plan["operations"], prefer_excel=False)
    assert out["status"] == "APPLIED", out
    wb = openpyxl.load_workbook(out["output_path"])
    ws = wb["Ctrl"]
    occupied = {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if c.value is not None}
    grids, _standalone = detect_grids("Ctrl", occupied)
    g = next(x for x in grids if x["name"] == "Instance values")
    assert g["flag_names"] == ["instancesplit"] and g["title_cell"] == "B16" and g["header_values"] == ["Instance key", "CohortRng", "Scenario"] and (g["first_row"], g["last_row"]) == (17, 26)
    assert [ws.cell(r, 2).value for r in (18, 26)] == ["CohortRng2022_Scenario0", "CohortRng2024_Scenario2"] and [ws.cell(26, c).value for c in (3, 4)] == [2024, 2]
    assert ws["C5"].value == "=INDEX(Ctrl!$C$18:$C$26,1)" and ws["C7"].value == "=INDEX(Ctrl!$D$18:$D$26,1)" and not ws.data_validations.dataValidation
    assert wb["Roll"]["C2"].value == 2024  # the other macro's loop driver is left as it is
    wb.close()


def test_copies_from_loops_choices_and_what_cannot_be_built(tmp_path, monkeypatch):
    path = _scenario_book(tmp_path, stale=True, driver_loop=True)
    # a results range can be named; the cohort cell's own Mind loop stays; a cell that still sums a chosen loop is named
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["Scenario"], results="Result!B5:C5", model_name="Tool")
    assert plan["status"] == "ready" and plan["key"]["values"] == ["Scenario0", "Scenario1", "Scenario2"]
    by = {(o["sheet"], o.get("cell")): o for o in plan["operations"] if o["op"] == "set_formula"}
    assert by[("Ctrl", "C7")]["after"] == "=INDEX(Ctrl!$C$18:$C$20,1)" and plan["split_grid"]["title_cell"] == "Ctrl!B16"
    assert ("Ctrl", "C5") not in by  # CohortRng stays a loop: its formula is already the plan's, nothing to write
    assert {k["name"]: k["already"] for k in plan["kept_loops"]} == {"CohortRng": True}  # RollYear is RollRun's: left out
    assert plan["results"] == {"block": "Result!B5:C5", "rows": 1, "cols": 2, "chosen_by": "user"} and plan["main"]["model_file"] == "Tool.xlsx"
    assert any("Roll!E2" in n and "MM_ function" in n for n in plan["notes"])
    # a loop that is not there, a results range that is not one, too many copies
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["Nope"])
    assert plan["status"] == "needs_decision" and "No loop is called 'Nope'" in plan["notes"][0] and plan["operations"] == []
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["Scenario"], results="Nowhere!A1:B2")
    assert plan["status"] == "ready" and plan["main"]["block"] is None and len(plan["main"]["grids"]) == 1 and any("not a range" in n for n in plan["notes"])
    import app.instance_builder as ib

    monkeypatch.setattr(ib, "MAX_COPIES", 5)
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["CohortRng", "Scenario"])
    assert plan["status"] == "needs_decision" and "3 x 3 = 9 copies, above the limit of 5" in plan["notes"][0]


def test_key_parts_name_the_loop_and_its_value():
    from app.instance_builder import _key_part

    assert _key_part("Scenario", 0, 1) == "Scenario0" and _key_part("CohortRng", 2021.0, 1) == "CohortRng2021" and _key_part("Group", "VFA-1", 2) == "GroupVFA-1"
    assert _key_part("Group", "変額", 3) == "Group3" and _key_part("Group", "a b", 4) == "Group4" and _key_part("Flag", True, 1) == "Flag1"


@needs_excel
def test_the_main_workbook_of_copies_from_loops_has_one_row_per_copy_and_block_row(tmp_path, monkeypatch):
    path = _scenario_book(tmp_path)
    plan = _plan_for(path, tmp_path, SCENARIO_MACRO, monkeypatch, loops=["CohortRng", "Scenario"])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    model_path = out_dir / plan["main"]["model_file"]
    shutil.copyfile(path, model_path)
    main_path = out_dir / plan["main"]["file"]
    done = write_main_workbook(plan, model_path, main_path)
    assert done["grids"] == 2 and done["formulas"] == 9 * 3 * 2
    wb = openpyxl.load_workbook(main_path)
    inst = wb["Instances"]
    assert inst["B4"].value == "#CohortRng and Scenario instances /InstanceKeys" and inst["B5"].value == "ifrs_instances" and [inst.cell(r, 2).value for r in range(6, 15)] == KEYS9
    summ = wb["Summary"]
    title = next(c for row in summ.iter_rows() for c in row if c.value == "#Results by CohortRng and Scenario")
    head = [summ.cell(title.row + 1, c).value for c in range(2, 6)]
    assert head == ["Key", "Row", "Value", "Double"]
    first = title.row + 2
    assert [summ.cell(first + i, 2).value for i in (0, 2, 3, 26)] == [KEYS9[0], KEYS9[0], KEYS9[1], KEYS9[8]]
    assert [summ.cell(first + i, 3).value for i in (0, 1, 2, 3)] == ["Row 3", "Row 4", "Row 5", "Row 3"]
    f = STORAGE_PREFIX_RE.sub("", summ.cell(first + 4, 5).value)
    assert f == f"=MM_INSTANCE([1]Result!$C$4,$B{first + 4})"
    assert not any(c.value == "#All copies" for row in summ.iter_rows() for c in row)
    wb.close()
