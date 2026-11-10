"""1.14.0: a sheet the scan skipped at the user's request still exists for a reference.

Joseph, 10 November 2026, the Tokyo IFRS 17 tool: with the big calculation sheets ticked as "do not
check", the 480 INDIRECT formulas on Result_CF that build addresses on them were written as NA()
(INDEX($A$1,NA()) where a range must stand) instead of the plain reference the no-skip run wrote
(BBA_01!A:B, Mind = Excel). The decision was `exists = sheet in sheets and grids`, over the scanned
sheets only. Now a skipped sheet counts as existing (its name and tab index are known); what is not
known about it (its cells, its grids) is never guessed: a rule that needs its cells leaves the
formula as it is and says so.
"""
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from app.config import load_config
from app.excel_com import close_quietly, com_available, excel_session, open_for_write, save_in_place
from app.inventory import build_analysis, sheet_order
from app.prep import plan_actions, reference_index
from app.rules_engine import RulesEngine
from app.validators import run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
SKIP = ["Big"]


def _finding(analysis, rule_id):
    return run_rule(RulesEngine().get(rule_id), analysis, load_config())


def _plan(analysis):
    return {a["id"]: a for a in plan_actions(analysis, {"findings": []})}


def _model(path: Path) -> Path:
    """Inputs reads Big (the sheet to skip) through INDIRECT, a 3-D range, a defined name and a
    plain reference; Big sits between Inputs and Mid in tab order."""
    wb = openpyxl.Workbook()
    inp = wb.active
    inp.title = "Inputs"
    big = wb.create_sheet("Big")
    big.append(["Key", "Value"])
    big.append(["k1", 10])
    big.append(["k2", 20])
    big["C1"] = "Mid"      # a sheet chooser that lives on the skipped sheet
    big["D1"] = 1
    for name in ("Mid", "Tail"):
        ws = wb.create_sheet(name)
        ws.append(["Key", "Value"])
        ws.append(["k1", 5])
        ws["D1"] = 2
    inp["A1"], inp["A2"] = "Sheet", "Big"
    inp["B2"] = '=IFERROR(VLOOKUP("k2",INDIRECT($A2&"!A:B"),2,FALSE),"-")'   # the Tokyo pattern: a range on the skipped sheet
    inp["B3"] = '=INDIRECT("Nope!A1")+1'                                      # a sheet that truly does not exist
    inp["B4"] = "=INDIRECT(\"'\"&Big!C1&\"'!A1\")"                            # the text needs a cell of the skipped sheet
    inp["B5"] = "=SUM('Big:Tail'!D1)"                                         # a 3-D range over the skipped sheet
    inp["B6"] = '=INDIRECT("BigTotal")*2'                                     # a defined name pointing into the skipped sheet
    inp["B7"] = "=IF(Big!E5>0,1,0)"                                           # compares an empty cell of the skipped sheet (FRM-009)
    inp["B8"] = "=BigTotal+Big!B2"                                            # plain references into the skipped sheet
    wb.defined_names["BigTotal"] = DefinedName("BigTotal", attr_text="Big!$B$3")
    wb.save(path)
    wb.close()
    return path


def test_a_skipped_sheet_keeps_its_place_in_the_tab_order(tmp_path):
    analysis = build_analysis(_model(tmp_path / "m.xlsx"), tmp_path / "w", "m", ignore_sheets=SKIP)
    wb0 = analysis["workbooks"][0]
    assert [s["name"] for s in wb0["sheets"]] == ["Inputs", "Mid", "Tail"] and wb0["ignored_sheets"][0]["name"] == "Big"
    assert sheet_order(wb0) == ["Inputs", "Big", "Mid", "Tail"]


def test_indirect_and_3d_targets_on_a_skipped_sheet_resolve_as_without_the_skip(tmp_path):
    src = _model(tmp_path / "m.xlsx")
    with_skip = _plan(build_analysis(src, tmp_path / "skip", "s", ignore_sheets=SKIP))
    no_skip = _plan(build_analysis(src, tmp_path / "full", "f"))

    direct = {o["cell"]: o for o in with_skip["direct_references"]["operations"]}
    assert direct["B2"]["after"] == '=IFERROR(VLOOKUP("k2",Big!A:B,2,FALSE),"-")'       # the plain reference, never NA()
    assert "Big left out of the scan" in direct["B2"]["note"] and "not checked" in direct["B2"]["note"]
    assert direct["B3"]["after"] == "=NA()+1" and "Nope" in direct["B3"]["note"]         # a sheet that does not exist: NA() as before
    assert direct["B6"]["after"] == "=BigTotal*2"                                        # the name into the skipped sheet
    assert "B4" not in direct                                                            # the text needs a cell of Big: left for a person, and the note says why
    skipped = [s for s in with_skip["direct_references"]["skipped"] if s.startswith("Inputs!B4")]
    assert len(skipped) == 1 and "Big was left out of the scan" in skipped[0] and "cells were not read" in skipped[0]

    three_d = {o["cell"]: o for o in with_skip["expand_3d_references"]["operations"]}
    assert three_d["B5"]["after"] == "=SUM(Big!D1,Mid!D1,Tail!D1)"                      # Big in its own tab position, not left out
    assert "Big left out of the scan" in three_d["B5"]["note"]

    # the same formulas as the run without the skip (B4 is resolved there: Big!C1 can be read)
    full_direct = {o["cell"]: o["after"] for o in no_skip["direct_references"]["operations"]}
    assert {c: o["after"] for c, o in direct.items()} == {c: a for c, a in full_direct.items() if c != "B4"}
    assert full_direct["B4"] == "=Mid!A1"
    assert {o["cell"]: o["after"] for o in no_skip["expand_3d_references"]["operations"]} == {c: o["after"] for c, o in three_d.items()}


def test_the_finding_counts_the_formulas_that_read_a_skipped_sheet(tmp_path):
    analysis = build_analysis(_model(tmp_path / "m.xlsx"), tmp_path / "w", "m", ignore_sheets=SKIP)
    f = _finding(analysis, "FRM-007")
    sites = {s["cell"]: s for s in f["observed"]["sites"]}
    assert sites["B2"]["unscanned_sheets"] == ["Big"] and sites["B2"]["issue"] is None and sites["B2"]["calls"][0]["exists"] is True
    assert sites["B3"]["unscanned_sheets"] == [] and "Nope" in sites["B3"]["issue"]
    assert sites["B4"]["suggested_formula"] is None and "left out of the scan" in sites["B4"]["issue"]
    assert f["observed"]["reading_skipped_sheets"] == 1
    # the reference index still knows the sheets a 3-D range and an INDIRECT reach, skipped or not
    index = reference_index(analysis)
    assert "Big" in index["positional"] and "Big" in index["text_addressed"]


def test_a_sheet_that_does_not_exist_still_becomes_na_with_or_without_skips(tmp_path):
    src = _model(tmp_path / "m.xlsx")
    for label, ignore in (("skip", SKIP), ("full", None)):
        plan = _plan(build_analysis(src, tmp_path / label, label, ignore_sheets=ignore))
        assert {o["cell"]: o["after"] for o in plan["direct_references"]["operations"]}["B3"] == "=NA()+1"


def _calculated(path: Path) -> Path:
    def work(excel):
        wb = open_for_write(excel, path)
        try:
            excel.CalculateFullRebuild()
            save_in_place(wb, path)
        finally:
            close_quietly(wb)

    with excel_session() as excel:
        work(excel)
    return path


@needs_excel
def test_a_rule_that_needs_the_skipped_sheets_cells_leaves_the_formula_and_says_so(tmp_path):
    """FRM-009 reads the compared cell: without the skip Big!E5 is seen empty and B7 gets N();
    with the skip nothing is known about Big!E5, so B7 is left as it is and the finding says why."""
    src = _calculated(_model(tmp_path / "m.xlsx"))
    full = build_analysis(src, tmp_path / "full", "f")
    assert {o["cell"] for o in _plan(full)["empty_cells_as_zero"]["operations"]} == {"B7"}
    assert "left out of the scan" not in _finding(full, "FRM-009")["message"]

    skip = build_analysis(src, tmp_path / "skip", "s", ignore_sheets=SKIP)
    assert _plan(skip)["empty_cells_as_zero"]["count"] == 0
    for rule_id in ("FRM-008", "FRM-009", "FRM-010", "FRM-013"):
        f = _finding(skip, rule_id)
        assert "read a sheet left out of the scan at your request (Big)" in f["message"] and "left as they are" in f["message"], rule_id
        assert "Big" in f["observed"]["skipped_sheets_note"]


@needs_excel
def test_two_skipped_sheets_do_not_break_the_mind_value_rules(tmp_path):
    """The rules' cache key sorted the skipped sheets as dicts: with two or more, TypeError, and
    FRM-008/009/010/013 came back NOT_SUPPORTED (so the Prep actions planned nothing) -- exactly
    the case the skipped-sheet handling was built for, the big calculation sheets ticked together."""
    analysis = build_analysis(_calculated(_model(tmp_path / "m.xlsx")), tmp_path / "w", "m", ignore_sheets=["Big", "Tail"])
    assert [s["name"] for s in analysis["workbooks"][0]["ignored_sheets"]] == ["Big", "Tail"]
    for rule_id in ("FRM-008", "FRM-009", "FRM-010", "FRM-013"):
        f = _finding(analysis, rule_id)
        assert f["status"] != "NOT_SUPPORTED" and "TypeError" not in (f["message"] or ""), (rule_id, f["message"])
        assert "read a sheet left out of the scan at your request (Big, Tail)" in f["message"] and "left as they are" in f["message"], (rule_id, f["message"])
