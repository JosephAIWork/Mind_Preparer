"""1.7.2: formulas Excel refuses to store -- a value where a function needs a
cell range. An assistant fix rewrote SUMIFS('BASE Polices'!#REF!, ...) to
SUMIFS(NA(), ...) in 17 cells and Excel refused every one (COM 0x800A03EC,
reported to the user as a raw tuple). The ACCEPTED / REFUSED lists below were
checked one by one in Excel on 2026-09-24."""
import pytest

from app.config import load_config
from app.excel_com import com_available
from app.formula_utils import collapse_na_reference_calls, reference_arg_problems
from app.modes import plan_mode
from app.prep import _com_message, apply_operations, plan_fix_broken_refs, plan_fix_broken_refs_whole, validate_proposal
from app.rules_engine import RulesEngine

needs_excel = pytest.mark.skipif(not com_available(), reason="requires pywin32 + an installed Excel")

REFUSED = [
    "=SUMIF(#N/A,1)", "=SUMIF(1,1)", '=SUMIF("a",1)', "=SUMIF(TRUE,1)", "=SUMIF({1,2},1)", "=SUMIF(A1:A5+0,1)",
    "=SUMIF(IFERROR(A:A,B:B),1)", "=SUMIF(FILTER(A:A,A:A>0),1)", "=SUMIF(A:A,1,NA())", "=COUNTIF(NA(),1)",
    "=AVERAGEIF(NA(),1)", "=MAXIFS(NA(),A:A,1)", "=MINIFS(A:A,NA(),1)", "=AVERAGEIFS(A:A,A:A,1,NA(),2)",
    "=COUNTIFS(A:A,1,NA(),2)", "=COUNTBLANK(NA())", "=OFFSET(NA(),1,1)", "=ROW(NA())", "=COLUMN(NA())",
    "=SUBTOTAL(9,NA())", "=AREAS(NA())", '=CELL("row",NA())', "=+SUMIF(NA(),YEAR(CALCUL!$N$6),NA())",
    "=+_xlfn.MAXIFS(NA(),A:A,1)",
]
ACCEPTED = [
    "=SUMIF(INDEX(A:A,0),1)", "=SUMIF(OFFSET(A1,0,0,5),1)", "=SUMIF(IF(1,A:A,B:B),1)", "=SUMIF(CHOOSE(1,A:A,B:B),1)",
    '=SUMIF(INDIRECT("A:A"),1)', "=SUMIF((A:A),1)", "=SUMIF(_xlfn.XLOOKUP(1,A:A,B:B),1)", "=SUMIF(A:A,NA(),B:B)",
    "=SUMIF(Sheet1!A:A,1)", "=SUMIF(A1:A5 A3:A9,1)", "=SUMIF(LET(x,A:A,x),1)", "=SUMIF(SWITCH(1,1,A:A),1)",
    "=SUMIFS(A:A,B:B,1,C:C,NA())", "=ROWS(NA())", "=SUMPRODUCT(NA())", "=VLOOKUP(1,NA(),2,0)", "=MATCH(1,NA(),0)",
    "=INDEX(NA(),1)", "=SUMIF(missing_name,1)", "=+SUMIF('BASE Polices'!#REF!,YEAR(CALCUL!$N$6),'BASE Polices'!#REF!)",
    "=SUMIF('P&L, x'!A:A,\"a,b\",'P&L, x'!B:B)", "=ROW()", "=SUM(NA())", "=SUMIF(#REF!,1)", '=SUMIF(A:A,"<>"&B1,C:C)',
]


@pytest.mark.parametrize("formula", REFUSED)
def test_formulas_excel_refuses_are_flagged(formula):
    assert reference_arg_problems(formula), formula


@pytest.mark.parametrize("formula", ACCEPTED)
def test_formulas_excel_accepts_are_not_flagged(formula):
    assert reference_arg_problems(formula) == [], formula


def test_problem_names_the_function_position_and_argument():
    (p,) = reference_arg_problems('=+SUMIFS(NA(),OUTPUT!J:J,"<>COURTAGE")')
    assert (p["function"], p["position"], p["arg"]) == ("SUMIFS", 1, "NA()") and p["call"].startswith("SUMIFS(")


def test_collapse_turns_only_na_holding_calls_into_na():
    assert collapse_na_reference_calls('=+SUMIFS(NA(),OUTPUT!J:J,"<>COURTAGE")') == "=+NA()"
    assert collapse_na_reference_calls("=IFERROR(SUMIF(NA(),1),0)+SUM(NA())") == "=IFERROR(NA(),0)+SUM(NA())"
    assert collapse_na_reference_calls("=SUMIF(A:A,1,OFFSET(NA(),0,0))*2") == "=NA()*2"  # innermost first
    assert collapse_na_reference_calls("=SUMIF(5,1)") == "=SUMIF(5,1)"  # not an error: left for a person
    assert collapse_na_reference_calls("=SUMIFS(A:A,B:B,1,C:C,NA())") == "=SUMIFS(A:A,B:B,1,C:C,NA())"  # criteria may be NA()


def test_broken_ref_fix_handles_sheet_qualified_refs_and_range_arguments():
    """REF-001's NA() rewrite used to skip 'Sheet'!#REF! and would have
    produced SUMIFS(NA(),...) -- now every one becomes a formula Excel stores."""
    formulas = [
        ("C20", "B7", "=+SUMIF('BASE Polices'!#REF!,YEAR(CALCUL!$N$6),'BASE Polices'!#REF!)", "=+NA()"),
        ("C20", "B8", "=+SUM('BASE Polices'!#REF!)", "=+SUM(NA())"),
        ("C20", "C22", "=+COUNTIFS('BASE Polices'!#REF!,1,OUTPUT!J:J,\"<>COURTAGE\")", "=+NA()"),
        ("C20", "I37", "=SUMIFS('BASE Polices'!#REF!,OUTPUT!V:V,'C20'!G36)", "=NA()"),
        ("C20", "K1", "=IFERROR(SUMIF(Data!#REF!,1),0)", "=IFERROR(NA(),0)"),
        ("OUTPUT", "S3", "=#REF!-Q3", "=NA()-Q3"),
    ]
    analysis = {"workbooks": [{"defined_names": [], "formulas": [{"sheet": s, "cell": c, "formula": f} for s, c, f, _ in formulas]}]}
    # 1.7.4: a call that needed the range goes to the separate "whole formula"
    # action (its own checkbox); the plain NA() swaps stay in the piecewise one
    piecewise, skipped = plan_fix_broken_refs(analysis, {})
    whole, _ = plan_fix_broken_refs_whole(analysis, {})
    assert {o["cell"] for o in piecewise} == {"B8", "S3"} and len(skipped) == 4
    ops = piecewise + whole
    assert {(o["sheet"], o["cell"]): o["after"] for o in ops} == {(s, c): want for s, c, _, want in formulas}
    assert all(not reference_arg_problems(o["after"]) for o in ops)
    # a range endpoint still cannot be rewritten: left for review, as before
    ops, skipped = plan_fix_broken_refs({"workbooks": [{"defined_names": [], "formulas": [{"sheet": "S", "cell": "A1", "formula": "=SUM(A1:#REF!)"}]}]}, {})
    assert ops == [] and len(skipped) == 1


def test_assistant_proposal_with_a_refused_formula_is_sent_back_with_the_repair(tmp_path, plain_grid_xlsx):
    result = plan_mode.run(plain_grid_xlsx, tmp_path / "work", load_config(), engine=RulesEngine())
    analysis = result["workbook_analysis"]
    ops, errors = validate_proposal({"operations": [
        {"op": "set_formula", "sheet": "Data", "cell": "D7", "formula": '=+SUMIFS(NA(),A:A,"<>COURTAGE")'},
        {"op": "set_formula", "sheet": "Data", "cell": "D8", "formula": "=SUMIF(5,1)"},
        {"op": "set_array_formula", "sheet": "Data", "range": "E7:E8", "formula": "=ROW(NA())"},
        {"op": "set_formula", "sheet": "Data", "cell": "D9", "formula": "=+SUM(NA())"},
    ]}, analysis)
    assert [o["cell"] for o in ops] == ["D9"]
    assert len(errors) == 3
    assert "SUMIFS argument 1 must be a cell range, but it is NA()" in errors[0] and "Use =+NA() instead" in errors[0]
    assert "argument 1 must be a cell range, but it is 5" in errors[1] and "replace the whole formula with =NA()" in errors[1]
    assert "ROW argument 1" in errors[2]


def test_excel_refusal_is_explained_in_plain_english():
    class FakeComError(Exception):
        excepinfo = (0, None, None, None, 0, -2146827284)

    msg = _com_message(FakeComError(), {"after": '=+SUMIFS(NA(),OUTPUT!J:J,"<>COURTAGE")'})
    assert msg.startswith("Excel refused this formula: SUMIFS argument 1 must be a cell range, but it is NA()")
    assert "not a formula Excel can store" in _com_message(FakeComError(), {"after": "=SUM(1"})
    assert "0x800A03EC" in _com_message(FakeComError(), {"after": 5})

    class WithText(Exception):
        excepinfo = (0, "Excel", "Cannot change part of an array.", None, 0, -2146827284)

    assert _com_message(WithText()) == "Cannot change part of an array."


@needs_excel
def test_excel_really_refuses_and_accepts_as_the_checker_says(tmp_path, plain_grid_xlsx):
    base = {"action_id": "assistant", "rule_id": "CHAT", "sheet": "Data", "before": None}
    ops = [
        {**base, "op": "set_formula", "cell": "D7", "after": '=+SUMIFS(NA(),A:A,"<>COURTAGE")'},  # bypasses validate_proposal on purpose
        {**base, "op": "set_formula", "cell": "D8", "after": collapse_na_reference_calls('=+SUMIFS(NA(),A:A,"<>COURTAGE")')},
    ]
    res = apply_operations(plain_grid_xlsx, tmp_path / "apply", ops)
    assert [o["cell"] for o in res["applied"]] == ["D8"]
    assert [o["cell"] for o in res["failed"]] == ["D7"]
    assert res["failed"][0]["error"].startswith("Excel refused this formula: SUMIFS argument 1 must be a cell range")
