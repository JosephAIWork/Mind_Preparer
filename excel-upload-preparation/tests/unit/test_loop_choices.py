"""1.12.5: each loop has its own choices, and the loops a model like Tokyo's runs are built from the
workbook alone: a counter whose dropdown lists its values, an item of a named list counted to its
end, bounds reached through plain references, the workbook's own name over a sheet's copy of it.

The workbook is written here with the two macros' shapes (read through a patched `read_vba`)."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from app import model_map
from app.inventory import build_analysis
from app.loop_builder import _follow_refs, build_loop_plan
from app.model_map import build_model_map

TWO_RUNS = '''
Sub AllScenarioRun()
    Dim scenarioGrid As Range
    Set scenarioGrid = ThisWorkbook.Names("ScenarioGrid").RefersToRange
    Dim numScenario As Long
    numScenario = scenarioGrid(1, 2).End(xlDown).Row - scenarioGrid(1, 2).Row + 1
    Dim scenarioRng As Range
    Set scenarioRng = ThisWorkbook.Names("Scenario").RefersToRange
    Dim startcohortyearRng As Range
    Set startcohortyearRng = ThisWorkbook.Names("StartCohortYear").RefersToRange
    Dim startcohortyear As Long
    startcohortyear = startcohortyearRng.Value
    Dim calcyearRng As Range
    Set calcyearRng = ThisWorkbook.Names("CalcYear").RefersToRange
    Dim endcohortyear As Long
    endcohortyear = calcyearRng.Value
    Dim cohortRng As Range
    Set cohortRng = ThisWorkbook.Names("CohortRng").RefersToRange
    For cohortYear = startcohortyear To endcohortyear
        cohortRng.Value = cohortYear
        For i = 0 To numScenario - 1
            scenarioRng(1, 1).Value = i
            Calculate
        Next i
    Next cohortYear
End Sub

Sub EndemicCalculationRun()
    Dim cohortyearRng As Range
    Set cohortyearRng = ThisWorkbook.Names("StartYear").RefersToRange
    Dim startcohortyear As Long
    Dim endcohortyear As Long
    startcohortyear = cohortyearRng.Value
    endcohortyear = Sheets("Ctrl_02").Range("E12").Value
    Dim GroupGrid As Range
    Set GroupGrid = ThisWorkbook.Names("GroupGrid").RefersToRange
    Dim numGroup As Long
    numGroup = GroupGrid(1, 1).End(xlDown).Row - GroupGrid(2, 1).Row + 1
    Dim RunGroupNameRng As Range
    Set RunGroupNameRng = ThisWorkbook.Names("GroupName").RefersToRange
    Dim RunCohortYearRng As Range
    Set RunCohortYearRng = ThisWorkbook.Names("CohortYear").RefersToRange
    For i = 1 To numGroup
        RunGroupNameRng(1, 1).Value = GroupGrid(i + 1, 1).Value
        For cohortYear = startcohortyear To endcohortyear
            RunCohortYearRng(1, 1).Value = cohortYear
            Calculate
        Next cohortYear
    Next i
End Sub
'''

GROUPS = ["Death", "Medical", "Cancer", "Area1", "Area3", "Variable"]


def _book(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    c1 = wb.active
    c1.title = "Ctrl_01"
    c1["B18"], c1["C18"] = "Calc year", 2026
    c1["B22"], c1["C22"] = "Cohort", 2026
    c1["B24"], c1["C24"] = "Start cohort", 2021
    c1["C26"], c1["D26"] = "Scenario", "Run (Y/N)"
    c1["B27"], c1["C27"], c1["F27"] = "Scenario run", 3, "=C27*2"
    c1["B29"] = "Scenarios"
    for k in range(4):
        c1.cell(29 + k, 3, k)
        c1.cell(29 + k, 4, "Y")
    dv = DataValidation(type="list", formula1="$C$29:$C$32")
    dv.add("C27")
    c1.add_data_validation(dv)
    c2 = wb.create_sheet("Ctrl_02")
    c2["B10"], c2["C10"] = "Base year", "=CalcYear"
    c2["B12"], c2["C12"], c2["D12"], c2["E12"] = "Cohorts", "=Ctrl_01!C24", "~", "=C10"
    c2["B19"], c2["C19"], c2["D19"] = "Grouping", "Group", "Portfolio"
    for k, g in enumerate(GROUPS):
        c2.cell(20 + k, 3, g)
        c2.cell(20 + k, 4, g.lower())
    rf = wb.create_sheet("RollForward")
    rf["A4"], rf["B4"], rf["A6"], rf["B6"] = "Group", "Variable", "Cohort year", 2026
    mat = wb.create_sheet("Maturity")
    mat["A4"], mat["B4"], mat["A6"], mat["B6"] = "Group", "=RollForward!B4", "Cohort year", "=RollForward!B6"
    for n, ref in (("CalcYear", "Ctrl_01!$C$18"), ("CohortRng", "Ctrl_01!$C$22"), ("StartCohortYear", "Ctrl_01!$C$24"), ("Scenario", "Ctrl_01!$C$27"),
                   ("ScenarioGrid", "Ctrl_01!$B$29"), ("StartYear", "Ctrl_02!$C$12"), ("GroupGrid", "Ctrl_02!$C$19"),
                   ("GroupName", "RollForward!$B$4"), ("CohortYear", "RollForward!$B$6")):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    # the sheet's own copies of two names, pointing at the cells that follow the workbook's ones
    mat.defined_names["GroupName"] = DefinedName("GroupName", attr_text="Maturity!$B$4")
    mat.defined_names["CohortYear"] = DefinedName("CohortYear", attr_text="Maturity!$B$6")
    p = tmp_path / "two_runs.xlsx"
    wb.save(p)
    wb.close()
    return p


@pytest.fixture()
def built(tmp_path, monkeypatch):
    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("M1_Main", TWO_RUNS)])
    analysis = build_analysis(_book(tmp_path), tmp_path / "w", "two_runs")
    return analysis, build_model_map(analysis)


def _loops(plan: dict) -> dict:
    return {lp["name"]: lp for lp in plan["loops"]}


def test_the_four_loops_are_built_from_the_workbook_alone(built):
    analysis, mp = built
    drivers = [d["cell"] for d in mp["loop"]["drivers"]]
    # the macro writes the workbook's GroupName / CohortYear, not the sheet's copies that follow them
    assert drivers == ["Ctrl_01!C22", "Ctrl_01!C27", "RollForward!B4", "RollForward!B6"]
    plan = build_loop_plan(analysis, mp)
    lp = _loops(plan)
    assert all(x["status"] == "ready" for x in lp.values()), {k: v.get("why") for k, v in lp.items()}
    assert lp["CohortRng"]["cells"][0]["formula"] == '=MM_LOOP("CohortRng",Ctrl_01!C18,2021,Ctrl_01!C24,Ctrl_01!C18)'
    # the counter 0 .. numScenario - 1: the values its own dropdown lists, a fixed range (no header row above it)
    sc = lp["Scenario"]
    assert (sc["first"], sc["last"], sc["count"]) == (0, 3, 4) and sc["cells"][0]["formula"] == '=MM_LOOP("Scenario",Ctrl_01!C29:C32)'
    # item i of GroupGrid, counted to the end of the list as the macro does; the range follows the list
    gr = lp["GroupName"]
    assert (gr["first"], gr["last"]) == (1, 6) and gr["cells"][0]["formula"] == '=MM_LOOP("GroupName",MM_COLUMN(Ctrl_02!C20))'
    assert any("6 items counted in Ctrl_02!C20:C25" in n for n in gr["notes"])
    # E12 = C10 = CalcYear, C12 = Ctrl_01!C24: the loop reads the cells that hold the numbers, no count cell needed
    cy = lp["CohortYear"]
    assert (cy["first"], cy["last"]) == (2021, 2026) and cy["cells"][0]["formula"] == '=MM_LOOP("CohortYear",Ctrl_01!C18,2021,Ctrl_01!C24,Ctrl_01!C18)'
    ops = {(o["sheet"], o["cell"], o["op"]) for o in plan["operations"]}
    assert ("Ctrl_01", "C27", "remove_validation") in ops  # Mind drops a formula in a cell with a dropdown
    assert not any(sh == "Maturity" for sh, _, _ in ops)


def test_each_loop_takes_its_own_choices(built):
    analysis, mp = built
    plan = build_loop_plan(analysis, mp, loops={
        "Ctrl_01!C27": {"last": "2", "name": "Scen"},  # a number: the dropdown is no longer read
        "'RollForward'!$B$4": {"last": "3"},  # quotes and $ do not matter
        "RollForward!B6": {"first": "2023"},
    })
    lp = _loops(plan)
    assert lp["CohortRng"]["first"] == 2021  # untouched
    sc = lp["Scen"]
    assert sc["status"] == "needs_decision" and "count cell" in sc["why"] and sc["suggested_count_cell"]  # 0 .. 2 typed: counted in a cell
    assert sc["chosen"] == {"last": "2", "name": "Scen"}
    gr = lp["GroupName"]
    assert (gr["first"], gr["last"]) == (1, 3) and gr["cells"][0]["formula"] == '=MM_LOOP("GroupName",Ctrl_02!C20:C22)'  # chosen: a fixed stretch
    assert lp["CohortYear"]["cells"][0]["formula"] == '=MM_LOOP("CohortYear",Ctrl_01!C18,2023,2023,Ctrl_01!C18)'


def test_a_count_cell_starting_at_zero_shifts_the_loop_and_gets_a_label(built):
    analysis, mp = built
    free = build_loop_plan(analysis, mp, loops={"Ctrl_01!C27": {"last": "3"}})
    cell = _loops(free)["Scenario"]["suggested_count_cell"]
    assert cell == "Ctrl_01!F27" or cell.startswith("Ctrl_01!")
    plan = build_loop_plan(analysis, mp, loops={"Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!H27"}})
    sc = _loops(plan)["Scenario"]
    assert sc["status"] == "ready" and sc["cells"][0]["formula"] == '=MM_LOOP("Scenario",Ctrl_01!H27)-1'
    ops = {(o["sheet"], o["cell"]): o for o in plan["operations"] if o["op"] == "set_value"}
    assert ops[("Ctrl_01", "H27")]["after"] == 4 and ops[("Ctrl_01", "G27")]["after"] == "Scenario count"


def test_a_count_cell_is_never_a_filled_cell_a_driver_or_shared(built):
    analysis, mp = built
    filled = _loops(build_loop_plan(analysis, mp, loops={"Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!F27"}}))["Scenario"]
    assert filled["status"] == "needs_decision" and "is not empty" in filled["why"] and "=C27*2" in filled["why"]
    driver = _loops(build_loop_plan(analysis, mp, loops={"Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!C22"}}))["Scenario"]
    assert driver["status"] == "needs_decision" and "driver cell" in driver["why"]
    shared = build_loop_plan(analysis, mp, loops={"Ctrl_01!C27": {"last": "3", "count_cell": "Ctrl_01!H27"}, "RollForward!B6": {"last": "2030", "first": "-1", "count_cell": "Ctrl_01!H27"}})
    assert _loops(shared)["Scenario"]["status"] == "ready"
    assert _loops(shared)["CohortYear"]["status"] == "needs_decision" and "already holds the count of the loop Scenario" in _loops(shared)["CohortYear"]["why"]


def test_a_bound_can_be_a_cell_and_a_wrong_one_is_said(built):
    analysis, mp = built
    plan = build_loop_plan(analysis, mp, loops={"Ctrl_01!C22": {"first": "Ctrl_02!$C$12"}, "Ctrl_01!C27": {"last": "Ctrl_01!nope"}})
    lp = _loops(plan)
    assert lp["CohortRng"]["first"] == 2021 and lp["CohortRng"]["first_from"]["kind"] == "cell"
    assert lp["Scenario"]["status"] == "needs_decision" and "the last value 'Ctrl_01!nope' is not a number nor a cell holding one" in lp["Scenario"]["why"]


def test_reference_chains_end_on_the_number(built):
    analysis, _ = built
    assert _follow_refs(analysis, "Ctrl_02", "E12") == ("Ctrl_01", "C18")  # =C10 -> =CalcYear -> 2026
    assert _follow_refs(analysis, "Ctrl_02", "C12") == ("Ctrl_01", "C24")
    assert _follow_refs(analysis, "Ctrl_01", "F27") is None  # =C27*2 is not a plain reference
