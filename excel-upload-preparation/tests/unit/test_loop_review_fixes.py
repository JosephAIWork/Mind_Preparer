"""Review of 10 November 2026 (Mind loop + copies from loops): the findings, each with a test.

- loops of two macros are built one macro at a time (Excel runs each macro on its own; in one model
  Mind would run them inside each other and add their results up);
- the loops are keyed by macro and loop variable (two macros both say `i`);
- a count cell that already holds this plan's count is fine; the count's label never lands on another
  loop's count, a cell the plan writes, or a cell a formula reads;
- a first value typed as a cell keeps the header pass, as the number does;
- copies from loops keep the clears of the macro's paste areas;
- copies' keys are always different;
- Prep's repairs are compared with the instance plan's cells at the addresses before its inserts.
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from app import model_map
from app.instance_builder import _key_parts, build_instance_plan
from app.inventory import build_analysis
from app.loop_builder import _add_repairs, _original_cells, build_loop_plan
from app.model_map import build_model_map

from test_loop_builder import ROW_LOOP_STYLE, _row_loop_book
from test_loop_choices import TWO_RUNS, _book
from test_instance_builder import SCENARIO_MACRO, _scenario_book


def _analyse(path: Path, tmp_path: Path, vba: str, monkeypatch) -> tuple[dict, dict]:
    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("Module1", vba)])
    analysis = build_analysis(path, tmp_path / "w", path.stem)
    return analysis, build_model_map(analysis)


def _loops(plan: dict) -> dict:
    return {lp["name"]: lp for lp in plan["loops"]}


# --- two macros, the same loop variable ---------------------------------------------------------------
SAME_VAR = '''
Sub RunA()
Dim na As Long: na = Range("CountA").Value
For i = 1 To na
    Range("ItemA").Value = i
    Calculate
    Worksheets("Out").Range("A" & i & ":B" & i) = Worksheets("Calc").Range("E10:F10").Value
Next i
End Sub

Sub RunB()
Dim nb As Long: nb = Range("CountB").Value
For i = 1 To nb
    Range("ItemB").Value = i
    Calculate
Next i
End Sub
'''


def _same_var_book(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    calc = wb.active
    calc.title = "Calc"
    calc["A1"], calc["B1"], calc["A2"], calc["B2"] = "Count A", 3, "Count B", 5
    calc["A4"], calc["B4"], calc["A5"], calc["B5"] = "Item A", 1, "Item B", 1
    calc["E10"], calc["F10"] = "=B4*10", "=B5*100"
    out = wb.create_sheet("Out")
    for r in range(1, 4):
        out.append([r, r])
    for n, ref in (("CountA", "Calc!$B$1"), ("CountB", "Calc!$B$2"), ("ItemA", "Calc!$B$4"), ("ItemB", "Calc!$B$5")):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    p = tmp_path / "same_var.xlsx"
    wb.save(p)
    wb.close()
    return p


def test_two_macros_with_the_same_loop_variable_keep_their_own_loop(tmp_path, monkeypatch):
    analysis, mp = _analyse(_same_var_book(tmp_path), tmp_path, SAME_VAR, monkeypatch)
    plan = build_loop_plan(analysis, mp)
    lp = _loops(plan)
    assert plan["macro"] == "RunA" and lp["ItemA"]["status"] == "ready" and lp["ItemB"]["status"] == "not_built" and lp["ItemB"]["other_macro"]
    # RunA's block is written by RunA's own `i` (3 values), not by RunB's (5)
    out = plan["outputs"][0]
    assert out["status"] == "ready" and out["loop"] == "ItemA" and out["items"] == 3 and out["cells"] == 6
    assert {o["after"] for o in plan["operations"] if o["sheet"] == "Out"} >= {'=MM_RESULT(Calc!$E$10,"ItemA",1)', '=MM_RESULT(Calc!$F$10,"ItemA",3)'}
    assert not any('"ItemB"' in str(o.get("after")) for o in plan["operations"])
    # RunB chosen: its loop alone; RunA's block is said to be RunA's, nothing is written into it
    b = build_loop_plan(analysis, mp, macro="RunB")
    assert _loops(b)["ItemB"]["status"] == "ready" and _loops(b)["ItemA"]["status"] == "not_built"
    assert b["outputs"][0]["status"] == "not_built" and "written by the macro RunA" in b["outputs"][0]["why"]
    assert not any(o["sheet"] == "Out" for o in b["operations"])
    # a macro that is not there: said, and the first macro's loops are shown
    nope = build_loop_plan(analysis, mp, macro="Nope")
    assert nope["macro"] == "RunA" and any("No loop runs in the macro 'Nope'" in n for n in nope["notes"])


# --- count cells and their labels ---------------------------------------------------------------------
@pytest.fixture()
def two_runs(tmp_path, monkeypatch):
    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("M1_Main", TWO_RUNS)])

    def make(edit=None):
        p = _book(tmp_path)
        if edit:
            wb = openpyxl.load_workbook(p)
            edit(wb["Ctrl_01"])
            wb.save(p)
            wb.close()
        analysis = build_analysis(p, tmp_path / f"w{len(list(tmp_path.iterdir()))}", "two_runs")
        return analysis, build_model_map(analysis)

    return make


SCEN_H27 = {"Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!H27"}}


def test_a_count_cell_holding_this_plans_own_count_can_be_planned_again(two_runs):
    def wrote_before(ws):  # what the first apply left: the count and its label
        ws["H27"], ws["G27"] = 4, "Scenario count"

    analysis, mp = two_runs(wrote_before)
    sc = _loops(build_loop_plan(analysis, mp, loops=SCEN_H27))["Scenario"]
    assert sc["status"] == "ready" and sc["cells"][0]["formula"] == '=MM_LOOP("Scenario",Ctrl_01!H27)-1'

    def other_number(ws):
        ws["H27"] = 5

    analysis, mp = two_runs(other_number)
    sc = _loops(build_loop_plan(analysis, mp, loops=SCEN_H27))["Scenario"]
    assert sc["status"] == "needs_decision" and "is not empty (it holds '5')" in sc["why"]


def test_the_count_label_never_overwrites_another_loops_count(two_runs):
    analysis, mp = two_runs()
    # CohortRng counts into Z5 (label Y5); Scenario counts into AA5, whose left cell is CohortRng's count
    plan = build_loop_plan(analysis, mp, loops={"Ctrl_01!C22": {"last": "2030", "count_cell": "Ctrl_01!Z5"}, "Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!AA5"}})
    lp = _loops(plan)
    assert lp["CohortRng"]["status"] == "ready" and lp["Scenario"]["status"] == "ready"
    sets = [(o["cell"], o["after"]) for o in plan["operations"] if o["op"] == "set_value" and o["sheet"] == "Ctrl_01"]
    assert sorted(sets) == [("AA5", 4), ("Y5", "CohortRng count"), ("Z5", 2030)]
    # the other way round: a count cell on the cell where another loop's label goes is refused
    plan = build_loop_plan(analysis, mp, loops={"Ctrl_01!C22": {"last": "2030", "count_cell": "Ctrl_01!AA5"}, "Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!Z5"}})
    sc = _loops(plan)["Scenario"]
    assert sc["status"] == "needs_decision" and "where the label of the count of the loop CohortRng goes" in sc["why"]


def test_the_count_label_never_goes_into_a_cell_a_formula_reads(two_runs):
    def reader(ws):
        ws["J40"] = "=COUNTA(Ctrl_01!G26:G28)"  # a text in G27 would change it

    analysis, mp = two_runs(reader)
    plan = build_loop_plan(analysis, mp, loops=SCEN_H27)
    assert _loops(plan)["Scenario"]["status"] == "ready"
    sets = {(o["cell"], o["after"]) for o in plan["operations"] if o["op"] == "set_value"}
    assert sets == {("H27", 4)}


# --- a first value typed as a cell --------------------------------------------------------------------
def test_a_first_value_typed_as_a_cell_keeps_the_header_pass_like_the_number(tmp_path, monkeypatch):
    analysis, mp = _analyse(_row_loop_book(tmp_path), tmp_path, ROW_LOOP_STYLE, monkeypatch)
    by_number = _loops(build_loop_plan(analysis, mp, first=1, last=4))["Policy_Data"]
    by_cell = _loops(build_loop_plan(analysis, mp, first="Calculation!T5", last=4))["Policy_Data"]  # T5 holds 1
    assert (by_number["first"], by_number["count"]) == (1, 4) and "header_skipped" not in by_number
    assert (by_cell["first"], by_cell["count"]) == (1, 4) and "header_skipped" not in by_cell and by_cell["first_from"]["chosen"] is True
    # not chosen: the header pass is skipped, and said
    auto = _loops(build_loop_plan(analysis, mp, last=4))["Policy_Data"]
    assert auto["first"] == 2 and auto["header_skipped"] == 1


# --- copies from loops ----------------------------------------------------------------------------------
def test_copies_from_loops_keep_the_clears_of_the_macros_paste_area(tmp_path, monkeypatch):
    vba = SCENARIO_MACRO.replace("For cohortYear", 'Worksheets("Result").Range("E2:E20").ClearContents\nFor cohortYear', 1)
    p = _scenario_book(tmp_path)
    wb = openpyxl.load_workbook(p)
    wb["Result"]["E3"] = "left from an earlier run"
    wb.save(p)
    wb.close()
    analysis, mp = _analyse(p, tmp_path, vba, monkeypatch)
    lp = build_loop_plan(analysis, mp, clear_old=True)
    clears = [o for o in lp["operations"] if o["op"] == "clear_cell"]
    assert [(o["sheet"], o["cell"]) for o in clears] == [("Result", "E3")] and "loop" not in clears[0]
    plan = build_instance_plan(analysis, mp, loop_plan=lp, loops=["CohortRng", "Scenario"])
    assert plan["status"] == "ready"
    kept = [o for o in plan["operations"] if o["op"] == "clear_cell"]
    assert [(o["sheet"], o["cell"]) for o in kept] == [("Result", "E3")] and plan["counts"]["cleared"] == 1
    assert any("the macro clears before a run are emptied" in n for n in plan["notes"])


def test_copies_keys_are_always_different():
    assert _key_parts("Scenario", [0, 1, 2]) == ["Scenario0", "Scenario1", "Scenario2"]
    # 'Base case' goes by its position (1), which meets the value '1': every value goes by its position
    assert _key_parts("Scenario", ["Base case", "1", "2"]) == ["Scenario1", "Scenario2", "Scenario3"]
    assert _key_parts("Scenario", ["A", "A"]) == ["Scenario1", "Scenario2"]


def test_copies_from_loops_of_two_macros_are_refused_and_one_macros_loops_pick_its_plan(tmp_path, monkeypatch):
    analysis, mp = _analyse(_scenario_book(tmp_path), tmp_path, SCENARIO_MACRO, monkeypatch)
    plan = build_instance_plan(analysis, mp, loops=["CohortRng", "RollYear"])
    assert plan["status"] == "needs_decision" and plan["operations"] == [] and "belong to different macros (AllRuns, RollRun)" in plan["notes"][0]
    # RollYear alone: RollRun's own plan, so its loop is usable and AllRuns' loops are left out
    plan = build_instance_plan(analysis, mp, loops=["RollYear"])
    assert plan["status"] == "ready" and [c["name"] for c in plan["copies_from"]] == ["RollYear"] and plan["kept_loops"] == []
    assert any("The loops of the macro AllRuns (CohortRng, Scenario)" in n for n in plan["notes"])
    # Scenario alone: the cohort, a loop of the same macro, stays a loop inside every copy
    plan = build_instance_plan(analysis, mp, loops=["Scenario"])
    assert plan["status"] == "ready" and [k["name"] for k in plan["kept_loops"]] == ["CohortRng"] and plan["kept_loops"][0]["written"] is True
    assert any(o["sheet"] == "Ctrl" and o.get("cell") == "C5" and str(o["after"]).startswith('=MM_LOOP("CohortRng"') for o in plan["operations"])


# --- Prep's repairs and the instance plan's inserts -----------------------------------------------------
def _repair(sheet: str, cell: str) -> dict:
    return {"op": "set_formula", "action_id": "freeze_environment_formulas", "sheet": sheet, "cell": cell, "before": '=CELL("filename",A1)', "after": '="Data"'}


def test_repairs_are_compared_with_the_cells_the_plan_writes_before_its_inserts():
    ops = [
        {"op": "insert_row", "sheet": "Data", "row": 5},
        {"op": "insert_column", "sheet": "Data", "column": "B"},
        {"op": "set_value", "sheet": "Data", "cell": "C7", "after": "x", "after_inserts": True},  # was B6
        {"op": "set_value", "sheet": "Data", "cell": "B7", "after": "Instance key", "after_inserts": True},  # the new column
        {"op": "set_value", "sheet": "Data", "cell": "D5", "after": "#T", "after_inserts": True},  # the new row
        {"op": "set_value", "sheet": "Data", "cell": "A2", "after": "y", "after_inserts": True},  # above and left: unmoved
        {"op": "set_range_values", "sheet": "Data", "cell": "E8:E9", "ref": "E8:E9", "values": [[1], [2]], "after_inserts": True},  # was D7:D8
    ]
    assert _original_cells(ops[2], ops) == [("Data", "B6")]
    assert _original_cells(ops[3], ops) == [] and _original_cells(ops[4], ops) == []
    assert _original_cells(ops[5], ops) == [("Data", "A2")]
    assert _original_cells(ops[6], ops) == [("Data", "D7"), ("Data", "D8")]
    prep = [{"id": "freeze_environment_formulas", "title": "Freeze", "default_on": True,
             "operations": [_repair("Data", "B6"), _repair("Data", "C7"), _repair("Data", "D8"), _repair("Data", "A2"), _repair("Data", "B7")]}]
    plan = {"notes": [], "counts": {}}
    out = list(ops)
    _add_repairs(plan, out, prep, True, what="the instances")
    written = [o["cell"] for o in out if o.get("action_id") == "freeze_environment_formulas"]
    # B6 (now C7), D8 (now E9) and A2 are the plan's; C7 and B7 are other cells that only share an address
    assert written == ["C7", "B7"] and plan["repairs"][0]["count"] == 2


def test_a_banded_row_insert_moves_only_its_columns():
    ops = [{"op": "insert_row", "sheet": "Data", "row": 3, "columns": "A:C"},
           {"op": "set_value", "sheet": "Data", "cell": "B5", "after": "x", "after_inserts": True},
           {"op": "set_value", "sheet": "Data", "cell": "F5", "after": "x", "after_inserts": True}]
    assert _original_cells(ops[1], ops) == [("Data", "B4")] and _original_cells(ops[2], ops) == [("Data", "F5")]


# --- the routes -------------------------------------------------------------------------------------------
def test_the_instances_route_picks_the_macro_of_the_loops_and_takes_the_loop_panels_choices(tmp_path, monkeypatch):
    import json

    from fastapi.testclient import TestClient

    from app.web import server

    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("Module1", SCENARIO_MACRO)])
    client = TestClient(server.app)
    path = _scenario_book(tmp_path)
    with path.open("rb") as f:
        sid = client.post("/api/sessions", files={"file": (path.name, f, "application/octet-stream")}, data={"mode": "plan"}).json()["sessionId"]
    plan = client.get(f"/api/sessions/{sid}/mind-instances-plan").json()["plan"]
    # every loop, each as its own macro's plan has it: all three can make copies
    assert {o["name"]: (o["macro"], o["usable"]) for o in plan["loop_options"]} == {"CohortRng": ("AllRuns", True), "Scenario": ("AllRuns", True), "RollYear": ("RollRun", True)}
    roll = client.get(f"/api/sessions/{sid}/mind-instances-plan", params={"loops": "RollYear"}).json()["plan"]
    assert roll["status"] == "ready" and [c["name"] for c in roll["copies_from"]] == ["RollYear"] and roll["kept_loops"] == []
    mixed = client.get(f"/api/sessions/{sid}/mind-instances-plan", params={"loops": "Scenario,RollYear"}).json()["plan"]
    assert mixed["status"] == "needs_decision" and "different macros" in mixed["notes"][0]
    # the loop panel's choices: the cohort runs 2023..2024 only, so 2 x 3 copies
    picks = {"Ctrl!C5": {"first": "2023"}}
    two = client.get(f"/api/sessions/{sid}/mind-instances-plan", params={"loops": "CohortRng,Scenario", "loop_choices": json.dumps(picks)}).json()["plan"]
    assert two["status"] == "ready" and [c["values"] for c in two["copies_from"]] == [[2023, 2024], [0, 1, 2]] and two["counts"]["key_values"] == 6
    # the macro of the loop panel reaches the plan by the table too (here: no table, the plan says so)
    assert client.get(f"/api/sessions/{sid}/mind-instances-plan", params={"macro": "RollRun"}).status_code == 200
    loop_plan = client.get(f"/api/sessions/{sid}/mind-loop-plan", params={"macro": "RollRun"}).json()["plan"]
    assert loop_plan["macro"] == "RollRun" and [le["status"] for le in loop_plan["loops"]] == ["not_built", "not_built", "ready"]
