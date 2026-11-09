"""1.12.5: the Mind loop goes to Mind with the formula repairs Prep still has for the version.

Found on the Tokyo IFRS 17 tool (2026-11-09): its four loops were written on a version Prep had not
repaired yet. BBA_01..BBA_05 and VFA_01 read their own sheet name from
=RIGHT(CELL("filename",B3),...FIND("]",CELL("filename",B3))) in B3; Mind answers CELL("filename") with
the workbook's name alone (no "]"), so B3 was NaN, the lookups keyed on it NaN, and every formula of
the six calculation sheets NaN: 14,737 differences in Mind's test. Probe workbooks proved it is not the
loops (the same NaN with and without an MM_LOOP). The loop step now writes those repairs with the loop.
"""
from __future__ import annotations

import pytest

from app import model_map
from app.inventory import build_analysis
from app.loop_builder import FORMULA_REPAIRS, build_loop_plan
from app.model_map import build_model_map
from app.prep import ACTIONS, plan_actions

from test_loop_choices import TWO_RUNS, _book  # the Tokyo-shaped workbook and its two macros

CELL_NAME = '=RIGHT(CELL("filename",B3),LEN(CELL("filename",B3))-FIND("]",CELL("filename",B3)))'
B45 = "=SUMPRODUCT(1-OFFSET($H$2,2,0,3,1),OFFSET($O$2,2,0,3,1))"


@pytest.fixture()
def tokyo(tmp_path, monkeypatch):
    import openpyxl

    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("M1_Main", TWO_RUNS)])
    p = _book(tmp_path)
    wb = openpyxl.load_workbook(p)
    ws = wb.create_sheet("BBA_01")
    ws["A3"], ws["B3"] = "Sheet", CELL_NAME
    ws["A5"], ws["B5"] = "Cohort", "=Ctrl_01!C22"
    ws["G3"], ws["H3"], ws["O3"] = "Month", "Rate", "Premium"
    for r in range(4, 10):
        ws.cell(r, 7, r - 4)
        ws.cell(r, 8, 1 - (r - 4) / 1000)
        ws.cell(r, 15, 100 + r)
    ws["A45"], ws["B45"] = "LF", B45
    wb.save(p)
    wb.close()
    analysis = build_analysis(p, tmp_path / "w", "tokyo")
    # what the user's own Excel showed in B3 (the server passes the original upload's values)
    analysis["original_environment_values"] = {"BBA_01!B3": "BBA_01"}
    return analysis, build_model_map(analysis)


def _prep(analysis):
    return plan_actions(analysis, {"findings": []})


def test_the_loop_is_written_with_the_formula_repairs_prep_still_has(tokyo):
    analysis, mp = tokyo
    plan = build_loop_plan(analysis, mp, prep_plan=_prep(analysis))
    ops = {(o["sheet"], o["cell"]): o for o in plan["operations"]}
    # the four loops, as before
    assert ops[("Ctrl_01", "C22")]["after"] == '=MM_LOOP("CohortRng",Ctrl_01!C18,2021,Ctrl_01!C24,Ctrl_01!C18)'
    assert ops[("RollForward", "B4")]["after"] == '=MM_LOOP("GroupName",MM_COLUMN(Ctrl_02!C20))'
    # CELL("filename") is written as what the user's Excel showed: Mind gave NaN for it
    assert ops[("BBA_01", "B3")]["after"] == '="BBA_01"' and ops[("BBA_01", "B3")]["action_id"] == "freeze_environment_formulas"
    # arithmetic on a range OFFSET returns: written as the range (Mind computed another number)
    assert ops[("BBA_01", "B45")]["after"] == "=SUMPRODUCT(1-$H$4:$H$6,OFFSET($O$2,2,0,3,1))"
    rep = {r["id"]: r for r in plan["repairs"]}
    assert rep["freeze_environment_formulas"]["count"] == 1 and rep["offset_ranges"]["count"] == 1 and all(r["written"] for r in rep.values())
    assert plan["counts"]["repairs"] == 2
    assert any(n.startswith("Written with the loop, so Mind computes what Excel does:") for n in plan["notes"])


def test_repairs_can_be_left_out_and_are_still_named(tokyo):
    analysis, mp = tokyo
    plan = build_loop_plan(analysis, mp, prep_plan=_prep(analysis), repairs=False)
    cells = {(o["sheet"], o["cell"]) for o in plan["operations"]}
    assert ("BBA_01", "B3") not in cells and ("BBA_01", "B45") not in cells
    assert {r["id"] for r in plan["repairs"]} == {"freeze_environment_formulas", "offset_ranges"} and not any(r["written"] for r in plan["repairs"])
    assert any(n.startswith("Not written with the loop:") for n in plan["notes"])


def test_a_repair_never_overwrites_a_cell_the_loop_writes_and_other_actions_stay_out(tokyo):
    analysis, mp = tokyo
    fake = [
        {"id": "freeze_environment_formulas", "title": "Freeze", "default_on": True, "operations": [
            {"op": "set_formula", "action_id": "freeze_environment_formulas", "sheet": "Ctrl_01", "cell": "C22", "before": "x", "after": "=1"}]},
        {"id": "create_grid_titles", "title": "Titles", "default_on": True, "operations": [
            {"op": "set_value", "action_id": "create_grid_titles", "sheet": "BBA_01", "cell": "A1", "before": None, "after": "#T"}]},
        {"id": "offset_ranges", "title": "Offsets", "default_on": False, "operations": [
            {"op": "set_formula", "action_id": "offset_ranges", "sheet": "BBA_01", "cell": "B45", "before": B45, "after": "=1"}]},
    ]
    plan = build_loop_plan(analysis, mp, prep_plan=fake)
    c22 = [o for o in plan["operations"] if (o["sheet"], o["cell"]) == ("Ctrl_01", "C22")]
    assert len(c22) == 1 and c22[0]["after"].startswith("=MM_LOOP(")  # the loop keeps its driver cell
    assert not any(o["sheet"] == "BBA_01" for o in plan["operations"])  # not a formula repair / switched off
    assert plan["repairs"] == []


def test_no_prep_plan_no_repairs(tokyo):
    analysis, mp = tokyo
    plan = build_loop_plan(analysis, mp)
    assert plan["repairs"] == [] and not any(o["sheet"] == "BBA_01" for o in plan["operations"])


def test_the_repairs_are_the_formula_actions_of_prep():
    ids = {a["id"] for a in ACTIONS}
    assert set(FORMULA_REPAIRS) <= ids
    assert "offset_ranges" in FORMULA_REPAIRS and "freeze_environment_formulas" in FORMULA_REPAIRS
    # structural changes (rows, titles) are Prep's, never the loop's
    assert not {"separate_merged_grids", "create_grid_titles", "title_lone_text_cells", "pivot_tables"} & set(FORMULA_REPAIRS)
