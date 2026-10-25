"""1.8.6: external links frozen to the values Excel keeps (app/external_links.py, REF-002).

UC-4: 20,934 formulas read a workbook on a colleague's machine; Excel shows the copies
it kept (the link's cache), Mind cannot follow a link at all -- "20,000 deterministic
errors" and "external links" in the converter."""
import os
from pathlib import Path

import pytest

from app.config import load_config
from app.excel_com import close_quietly, com_available, excel_session
from app.external_links import external_references, freeze_formula, read_external_links
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")


def test_external_references_are_found_and_frozen_from_the_cache():
    links = [{"index": 1, "cache": {("General Assumptions", "D13"): 0.44, ("Plan by Line", "F8"): "Group A", ("Plan by Line", "F9"): "", ("Flags", "B2"): True}}]
    assert [r["ref"] for r in external_references("='[1]General Assumptions'!$D$13*H21+[1]Flags!B2")] == ["$D$13", "B2"]
    assert freeze_formula("='[1]General Assumptions'!$D$13*H21", links)[0] == "=0.44*H21"
    assert freeze_formula("=IFERROR('[1]Plan by Line'!F8,0)&[1]Flags!B2", links)[0] == '=IFERROR("Group A",0)&TRUE'
    assert freeze_formula("='[1]Plan by Line'!F9", links)[0] == '=""'
    fixed, notes, issue = freeze_formula("=IFERROR('[1]Plan by Line'!Z99,0)", links)          # no copy kept: #REF! in Excel
    assert fixed == "=IFERROR(NA(),0)" and "no copy kept" in notes[0]
    assert freeze_formula("=SUM('[1]Plan by Line'!F8:F20)", links)[0] is None                   # a range: by hand
    assert freeze_formula("=A1+1", links) == (None, [], None)


def _dead_link_model(folder: Path) -> Path:
    """Book B reads book A; A is then deleted, so B's link cannot be reached and Excel shows its cache."""
    a, b = folder / "A.xlsx", folder / "B.xlsx"

    def make(excel):
        wa = excel.Workbooks.Add()
        wa.Worksheets(1).Name = "Inputs"
        wa.Worksheets(1).Range("A1").Value = 0.44
        wa.Worksheets(1).Range("A2").Value = "Group A"
        wa.SaveAs(str(a), FileFormat=51)
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Calc"
        ws.Range("B1").Value = 100
        ws.Range("C1").Formula = "='[A.xlsx]Inputs'!A1*B1"
        ws.Range("C2").Formula = "=IFERROR('[A.xlsx]Inputs'!A2,0)"
        ws.Range("C3").Formula = "=IFERROR('[A.xlsx]Inputs'!A9,0)"      # A9 is empty in A
        ws.Range("C4").Formula = "=C1*2"
        excel.CalculateFullRebuild()
        wb.SaveAs(str(b), FileFormat=51)
        close_quietly(wb)
        close_quietly(wa)

    with excel_session() as excel:
        make(excel)
    os.remove(a)
    return b


@needs_excel
def test_a_dead_link_is_frozen_and_broken_and_the_numbers_stay(tmp_path):
    src = _dead_link_model(tmp_path)
    links = read_external_links(src)
    assert len(links) == 1 and links[0]["file_name"] == "A.xlsx" and links[0]["reachable"] is False
    assert links[0]["cache"][("Inputs", "A1")] == 0.44 and links[0]["cache"][("Inputs", "A2")] == "Group A"
    analysis = build_analysis(src, tmp_path / "w", "b")
    f = run_rule(RulesEngine().get("REF-002"), analysis, load_config())
    assert f["status"] == "ERROR" and f["observed"]["sites_count"] == 3 and "cannot follow" in f["message"]
    plan = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}
    a = plan["freeze_external_links"]
    assert a["level"] == "blocking" and not a["skipped"]
    by = {o.get("cell"): o for o in a["operations"]}
    assert by["C1"]["after"] == "=0.44*B1" and by["C2"]["after"] == '=IFERROR("Group A",0)'
    assert by["C3"]["after"] in ("=IFERROR(NA(),0)", '=IFERROR("",0)', "=IFERROR(0,0)")   # what Excel kept for the empty cell
    assert [o["op"] for o in a["operations"]].count("break_external_link") == 1
    out = apply_operations(src, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"], out["failed"]
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["clean"] is True and res["counts"]["good_to_other"] == 0
    after = build_analysis(Path(out["output_path"]), tmp_path / "after", "a")
    assert read_external_links(Path(out["output_path"])) == []
    assert run_rule(RulesEngine().get("REF-002"), after, load_config())["status"] == "PASS"
