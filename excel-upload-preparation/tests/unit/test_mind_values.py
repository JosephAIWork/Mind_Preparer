"""1.8.1: two places where Mind computes another value than Excel (app/validators/mind_values.py).

Found by uploading a model that recalculates clean in Excel (UC-1) and
confirmed in Mind with a 49-cell test workbook:

  FRM-008  =IFERROR(1+INT((D192-1)/12),"") with D192 = "": blank in Excel, a number in Mind
  FRM-009  =IF(D35<=$H$9,0,1) with H9 = a VLOOKUP that lands on an empty cell: 1 in Excel, 0 in Mind
  FRM-010  =1+(...)*IFERROR(VLOOKUP(F192,Criteria!$N$11:$O$33,2,0),...) with F192 = "": Excel finds nothing,
           Mind matches the empty header cell N11 and returns the header text -- NaN
"""
from pathlib import Path

import openpyxl
import pytest

from app.config import load_config
from app.excel_com import close_quietly, com_available, excel_session, open_for_write, save_in_place
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import mind_values, run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")


def _finding(analysis, rule_id):
    return run_rule(RulesEngine().get(rule_id), analysis, load_config())


def _model(path: Path) -> Path:
    wb = openpyxl.Workbook()
    loans = wb.active
    loans.title = "Loans"
    loans.append(["Loan", "Capital", "Deferral"])
    loans.append([1, 100, 12])
    loans.append([2, 200, None])                                # C3 stays empty
    c = wb.create_sheet("Calc")
    c["A1"], c["B1"] = "Loan", 2
    c["A2"], c["B2"] = "Deferral", "=VLOOKUP(B1,Loans!A:C,3,0)"  # lands on Loans!C3: 0 in Excel
    c["A3"], c["B3"] = "Same, by reference", "=+Loans!C3"
    c["D4"], c["E4"], c["F4"] = "Step", "Year", "Live"
    c["D5"] = 1
    for r in range(5, 11):
        if r > 5:
            c[f"D{r}"] = f'=IF(OR(D{r - 1}="",D{r - 1}>=3),"",D{r - 1}+1)'      # 1, 2, 3, "", "", ""
        c[f"E{r}"] = f'=IFERROR(1+INT((D{r}-1)/12),"")'                          # blank in Excel on the last three rows
        c[f"F{r}"] = f"=IF(D{r}<=$B$2,0,1)*10"                                   # reads B2 into a comparison
    c["G5"] = "=0=B3"                                           # reads the plain reference into a comparison
    c["H5"] = "=IF(1<=Loans!C3,0,1)"                            # the empty cell itself, compared
    c["H6"] = '=IF(Loans!C3="","",IF(Loans!C3>0,1,2))'          # asks first whether it is empty: same branch in both
    c["H7"] = '=IF(D8="","",D8-1)'                              # no IFERROR, tested first
    c["H8"] = '=IFERROR(IF(D8="","",D8-1),"")'                  # tested first, inside the IFERROR
    c["H9"] = "=1+IFERROR(D8*2,5)"                              # an IFERROR inside a larger formula
    c["H10"] = "=Loans!C3+1"                                    # arithmetic on an empty cell: 0 in both
    c["I5"] = "=IF(Loans!C2>5,1,0)"                             # a cell that holds a number
    c["I6"] = '=IF(Loans!C3="x",1,0)'                           # compared with text
    c["M5"], c["M6"], c["M7"] = 3, 4, 5
    c["M8"], c["M9"], c["M10"] = '=IF(B1=2,"",0)', "=1/0", 6                     # "" on row 8, an error on row 9
    for r in range(5, 11):
        c[f"L{r}"] = f"=IFERROR(M{r}*2,0)"                                     # one block: row 8 needs the test, row 9 must be left alone
    rates = wb.create_sheet("Rates")                            # a table whose first header cell is empty, as in UC-1
    rates["B1"] = "Loading"
    for i, (age, loading) in enumerate(((50, 1.08), (51, 1.09), (52, 1.1), (53, 1.11), (54, 1.12), (55, 1.15)), start=2):
        rates[f"A{i}"], rates[f"B{i}"] = age, loading
    for r in range(5, 11):
        c[f"J{r}"] = f"=1+($B$1=9)*IFERROR(VLOOKUP(E{r},Rates!$A$1:$B$7,2,0),IF(E{r}<50,Rates!$B$2,Rates!$B$7))"   # as in UC-1: the key E is "" on rows 8-10
        c[f"K{r}"] = f"=IFERROR(VLOOKUP(E{r},Loans!A:C,2,0),0)"                                                  # the same key, but no empty cell in Loans!A to match
    wb.save(path)
    wb.close()
    return path


def _calculated(path: Path) -> Path:
    """Excel calculates and saves: the file now holds the values the rules read."""
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


def test_a_workbook_excel_never_calculated_is_not_checked(tmp_path):
    analysis = build_analysis(_model(tmp_path / "raw.xlsx"), tmp_path / "w", "raw")
    for rule_id in ("FRM-008", "FRM-009", "FRM-010"):
        f = _finding(analysis, rule_id)
        assert f["status"] == "NOT_SUPPORTED" and "no stored values" in f["message"]
    plan = {a["id"]: a for a in plan_actions(analysis, {"findings": []})}
    assert plan["guard_blank_arithmetic"]["count"] == 0 and plan["empty_cells_as_zero"]["count"] == 0


@needs_excel
def test_empty_text_in_arithmetic_and_empty_cells_in_comparisons_are_found_and_rewritten(tmp_path):
    src = _calculated(_model(tmp_path / "model.xlsx"))
    analysis = build_analysis(src, tmp_path / "w", "m")

    blank = _finding(analysis, "FRM-008")
    assert blank["status"] == "ERROR" and blank["priority"] == "REQUIRED"
    sites = {s["cell"]: s for s in blank["observed"]["sites"]}
    # the three rows where D is "", the rest of that filled-down block, and the IFERROR inside a larger formula
    assert set(sites) == {"E5", "E6", "E7", "E8", "E9", "E10", "H9", "L5", "L6", "L7", "L8", "L9", "L10"}   # not the lookups: those are FRM-010's
    assert [c for c, s in sites.items() if s["blank_now"]] == ["E8", "E9", "E10", "H9", "L8"]
    # row 9 of the L block reads an error: IF(#DIV/0!="",...) would be #DIV/0! where IFERROR gave 0 -- left alone (UC-4, 19 values changed)
    assert sites["L9"]["suggested_formula"] is None and "holds an error" in sites["L9"]["issue"]
    assert sites["L8"]["suggested_formula"] == '=IF(M8="",0,IFERROR(M8*2,0))' and sites["L10"]["suggested_formula"] == '=IF(M10="",0,IFERROR(M10*2,0))'
    assert sites["E8"]["suggested_formula"] == '=IF(D8="","",IFERROR(1+INT((D8-1)/12),""))'
    assert sites["E5"]["suggested_formula"] == '=IF(D5="","",IFERROR(1+INT((D5-1)/12),""))'
    assert sites["H9"]["suggested_formula"] == '=1+IF(D8="",5,IFERROR(D8*2,5))'
    assert blank["observed"]["blank_now"] == 5 and blank["observed"]["with_proposal"] == 12 and blank["observed"]["without_proposal"] == 1
    assert "Mind counts" in blank["message"] and "Test for empty text before the arithmetic" in blank["message"]

    lookup = _finding(analysis, "FRM-010")
    assert lookup["status"] == "ERROR" and lookup["priority"] == "REQUIRED"
    sites = {s["cell"]: s for s in lookup["observed"]["sites"]}
    assert set(sites) == {"J5", "J6", "J7", "J8", "J9", "J10"}                                 # K: Loans!A1:A3 has no empty cell, Excel and Mind both find nothing
    assert [c for c, s in sites.items() if s["blank_now"]] == ["J8", "J9", "J10"]
    assert sites["J8"]["suggested_formula"] == '=1+($B$1=9)*IF(E8="",IF(E8<50,Rates!$B$2,Rates!$B$7),IFERROR(VLOOKUP(E8,Rates!$A$1:$B$7,2,0),IF(E8<50,Rates!$B$2,Rates!$B$7)))'
    assert sites["J8"]["lookup_key"] and not sites["J8"]["arithmetic"]
    assert "header cell of a table" in lookup["message"]

    empty = _finding(analysis, "FRM-009")
    assert empty["status"] == "ERROR"
    sites = {(s["kind"], s["cell"]): s for s in empty["observed"]["sites"]}
    assert set(sites) == {("lands_on_empty", "B2"), ("lands_on_empty", "B3"), ("reads_empty", "H5")}
    assert sites[("lands_on_empty", "B2")]["lands_on"] == "Loans!C3" and sites[("lands_on_empty", "B2")]["readers"] == 3   # F5:F7 -- on F8:F10 the other side is ""
    assert sites[("lands_on_empty", "B2")]["suggested_formula"] == "=N(VLOOKUP(B1,Loans!A:C,3,0))"
    assert sites[("lands_on_empty", "B3")]["suggested_formula"] == "=N(Loans!C3)"
    assert sites[("reads_empty", "H5")]["suggested_formula"] == "=IF(1<=N(Loans!C3),0,1)"
    assert "Read empty cells as 0 where they are compared" in empty["message"]

    plan = {a["id"]: a for a in plan_actions(analysis, {"findings": [blank, lookup, empty]})}
    assert plan["guard_blank_arithmetic"]["level"] == "blocking" and plan["guard_blank_arithmetic"]["count"] == 18
    assert len(plan["guard_blank_arithmetic"]["skipped"]) == 1 and "L9" in plan["guard_blank_arithmetic"]["skipped"][0]
    assert {o["rule_id"] for o in plan["guard_blank_arithmetic"]["operations"]} == {"FRM-008", "FRM-010"}
    assert plan["empty_cells_as_zero"]["count"] == 3 and not plan["empty_cells_as_zero"]["skipped"]
    ops = plan["guard_blank_arithmetic"]["operations"] + plan["empty_cells_as_zero"]["operations"]
    out = apply_operations(src, tmp_path / "prep", ops)
    assert out["status"] == "APPLIED" and not out["failed"]

    # Excel computes exactly what it did
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["ran"] and res["clean"] is True and res["counts"]["same"] > 40 and res["counts"]["error_to_error"] == 1   # M9 = 1/0 before and after

    after = build_analysis(Path(out["output_path"]), tmp_path / "after", "a")
    assert {_finding(after, rid)["status"] for rid in ("FRM-008", "FRM-009", "FRM-010")} == {"PASS"}


def test_which_side_of_a_comparison_a_reference_is():
    def other(formula, ref):
        s = formula.index(ref)
        span = mind_values._other_side(formula, s, s + len(ref))
        return formula[span[0]:span[1]] if span else None

    assert other("=IF(D35<=$H$9,0,1)", "$H$9") == "D35" and other("=IF(D35<=$H$9,0,1)", "D35") == "$H$9"
    assert other("=+C3+C7-C11=C15", "C15") == "+C3+C7-C11" and other("=+C3+C7-C11=C15", "C11") is None   # C11 is part of a sum
    assert other("=0=B3", "B3") == "0" and other("=B3<>SUM(A1:A3)*2", "B3") == "SUM(A1:A3)*2"
    assert other("=A1", "A1") is None and other("=SUM(A1,B3)", "B3") is None and other("=IF(A1>B3+1,1,0)", "B3") is None


def test_the_test_for_empty_text_goes_in_front_of_each_iferror():
    import re

    from app.formula_utils import REF_RE

    formula = '=IFERROR(C191+F192-F191,0)+IFERROR(A192*2,"")'
    refs = list(REF_RE.finditer(formula))
    calls = mind_values._iferror_calls(formula)
    assert mind_values._guarded(formula, refs, calls, {0: {1, 2}, 1: {3}}) == '=IF(OR(F192="",F191=""),0,IFERROR(C191+F192-F191,0))+IF(A192="","",IFERROR(A192*2,""))'
    # a filled-down block has one relative form, whatever the row
    a = mind_values._relative_form("=IFERROR($C$6+A35-YEAR($H$6),\"\")", "=IFERROR($C$6+A35-YEAR($H$6),\"\")", list(REF_RE.finditer("=IFERROR($C$6+A35-YEAR($H$6),\"\")")), 35, 6)
    b = mind_values._relative_form("=IFERROR($C$6+A812-YEAR($H$6),\"\")", "=IFERROR($C$6+A812-YEAR($H$6),\"\")", list(REF_RE.finditer("=IFERROR($C$6+A812-YEAR($H$6),\"\")")), 812, 6)
    assert a == b and re.search(r"C\[-5\]R\[0\]", a)


def test_a_line_break_inside_a_string_literal_is_written_as_char10():
    """1.8.5 -- FRM-011 (UC-3: another text in Mind; with CHAR(10) the same)."""
    assert mind_values.line_breaks_as_char('=BB62&"\n/ Beneficiary - "&D8') == '=BB62&"&CHAR(10)&"/ Beneficiary - "&D8'.replace('"&CHAR(10)&"/', 'CHAR(10)&"/') or True
    assert mind_values.line_breaks_as_char('=BB62&"\n/ Beneficiary - "&D8') == '=BB62&CHAR(10)&"/ Beneficiary - "&D8'
    assert mind_values.line_breaks_as_char('="a\nb\r\nc"') == '="a"&CHAR(10)&"b"&CHAR(13)&CHAR(10)&"c"'
    assert mind_values.line_breaks_as_char('=IF(A1="\n",1,"x\ny")') == '=IF(A1=CHAR(10),1,"x"&CHAR(10)&"y")'
    assert mind_values.line_breaks_as_char('="say ""hi""\nnow"') == '="say ""hi"""&CHAR(10)&"now"'
    assert mind_values.line_breaks_as_char('=A1&"no break"') is None and mind_values.line_breaks_as_char("=A1+\n1") is None  # a break outside the quotes is not touched


@needs_excel
def test_the_line_break_rewrite_gives_the_same_text_in_excel(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wave"
    ws["A1"], ws["B1"] = "HBP Category", "Combined All"
    ws["C1"] = '=A1&"\n/ Beneficiary - "&B1'
    ws["C2"] = "=LEN(C1)"
    p = tmp_path / "wave.xlsx"
    wb.save(p)
    wb.close()
    analysis = build_analysis(p, tmp_path / "w", "v")
    f = _finding(analysis, "FRM-011")
    assert f["status"] == "ERROR" and f["observed"]["sites"][0]["cell"] == "C1"
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["line_breaks_as_char"]
    assert a["count"] == 1 and a["operations"][0]["after"] == '=A1&CHAR(10)&"/ Beneficiary - "&B1'
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED"
    res = compare_workbooks(p, Path(out["output_path"]), Moves())
    assert res["clean"] is True and res["counts"]["same"] == 2
    assert _finding(build_analysis(Path(out["output_path"]), tmp_path / "after", "a"), "FRM-011")["status"] == "PASS"


@needs_excel
def test_a_formula_that_asks_excel_about_itself_is_frozen_to_what_it_shows(tmp_path):
    """1.8.7 -- FRM-012 (UC-2: CELL("filename") stopped Mind's run)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Saved paths"
    from openpyxl.worksheet.formula import ArrayFormula

    ws["B2"], ws["C2"] = "Directory", ArrayFormula("C2", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)')  # a one-cell array, as in UC-2 (1.8.12)
    ws["B5"], ws["C5"] = "Spreadsheet", '=MID(CELL("filename",A1),FIND("[",CELL("filename",A1))+1,FIND("]",CELL("filename",A1))-FIND("[",CELL("filename",A1))-1)'
    ws["C6"] = "=C2&C5"
    ws["C7"] = "=LEN(C5)"
    p = tmp_path / "env.xlsx"
    wb.save(p)
    wb.close()
    src = _calculated(p)
    analysis = build_analysis(src, tmp_path / "w", "e")
    f = _finding(analysis, "FRM-012")
    assert f["status"] == "ERROR" and {s["cell"] for s in f["observed"]["sites"]} == {"C2", "C5"}
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["freeze_environment_formulas"]
    by = {o["cell"]: o for o in a["operations"]}
    assert by["C5"]["after"] == '="env.xlsx"' and by["C2"]["after"].startswith('="') and by["C2"]["after"].endswith('\\"')
    out = apply_operations(src, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["clean"] is True and res["counts"]["same"] == 4
    assert _finding(build_analysis(Path(out["output_path"]), tmp_path / "after", "a"), "FRM-012")["status"] == "PASS"


def _environment_workbook(folder: Path) -> Path:
    """CELL("filename") cells as in UC-2: the folder (C2) and the file name (C5),
    calculated by Excel where the file lies."""
    folder.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Saved paths"
    ws["C2"] = '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)'
    ws["C5"] = '=MID(CELL("filename",A1),FIND("[",CELL("filename",A1))+1,FIND("]",CELL("filename",A1))-FIND("[",CELL("filename",A1))-1)'
    ws["C6"] = "=C2&C5"
    p = folder / "env.xlsx"
    wb.save(p)
    wb.close()
    return _calculated(p)


@needs_excel
def test_frm012_never_writes_the_apps_working_folder_and_prefers_the_originals_values(tmp_path):
    """1.11.10's leftover: 'Saved paths'!C2 was frozen to the app's temporary folder (with the Windows
    user's name in it), because the app's Excel had saved the version there."""
    src = _environment_workbook(tmp_path / "mind_ready_abc" / "apply")  # where the app's Excel saves versions
    analysis = build_analysis(src, tmp_path / "w", "e")
    f = _finding(analysis, "FRM-012")
    assert "working folder" in f["message"] and "mind_ready_abc" not in f["message"]
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["freeze_environment_formulas"]
    assert [o["cell"] for o in a["operations"]] == ["C5"]
    assert any(x.startswith("Saved paths!C2") and "working folder" in x for x in a["skipped"])
    # what the original upload shows (the user's own Excel saved it) wins
    analysis["original_environment_values"] = {"Saved paths!C2": "L:\\Models\\", "Saved paths!C5": "model v1.xlsx"}
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["freeze_environment_formulas"]
    assert {o["cell"]: o["after"] for o in a["operations"]} == {"C2": '="L:\\Models\\"', "C5": '="model v1.xlsx"'}


@needs_excel
def test_the_server_freezes_cell_filename_to_what_the_original_upload_shows(tmp_path):
    """The original's values follow their cells into later versions (a row inserted above them),
    whatever the version the app's Excel saved in its own folder shows."""
    from fastapi.testclient import TestClient

    from app.web import server

    src = _environment_workbook(tmp_path / "user folder")
    shown = str(tmp_path / "user folder") + "\\"
    client = TestClient(server.app)
    with open(src, "rb") as fh:
        body = client.post("/api/sessions", files={"file": (src.name, fh, "application/octet-stream")}, data={"mode": "plan"}).json()
    plan = {a["id"]: a for a in body["plan"]}["freeze_environment_formulas"]
    assert {o["cell"]: o["after"] for o in plan["operations"]} == {"C2": f'="{shown}"', "C5": '="env.xlsx"'}
    out = client.post(f"/api/sessions/{body['sessionId']}/apply", json={"operations": [{"op": "insert_row", "sheet": "Saved paths", "row": 1}], "reanalyze": True}).json()
    assert out["result"]["status"] == "APPLIED"
    plan = {a["id"]: a for a in out["plan"]}["freeze_environment_formulas"]
    assert {o["cell"]: o["after"] for o in plan["operations"]} == {"C3": f'="{shown}"', "C6": '="env.xlsx"'}
    assert all("original upload" in o["note"] for o in plan["operations"])


@needs_excel
def test_an_error_cell_kept_on_purpose_stays_an_error_in_mind(tmp_path):
    """1.8.9 -- FRM-013 (UC-4, 'Line 1'!AB78 = AB77-AB84: #VALUE! in Excel, a number in Mind)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["A1"], ws["B1"] = 10, '=IF(1=1,"",0)'             # B1 holds ""
    ws["C1"] = "=A1-B1"                                    # #VALUE! in Excel, kept on purpose
    ws["D1"] = '=IFERROR(C1/2,"No GEP")'                   # what reads it
    ws["C2"] = "=IFERROR(A1-B1,0)"                         # guarded: FRM-008's, not FRM-013's
    ws["C3"] = "=A1-A1"                                    # no "" involved
    p = tmp_path / "kept.xlsx"
    wb.save(p)
    wb.close()
    src = _calculated(p)
    analysis = build_analysis(src, tmp_path / "w", "k")
    f = _finding(analysis, "FRM-013")
    assert f["status"] == "ERROR" and [s["cell"] for s in f["observed"]["sites"]] == ["C1"]
    assert f["observed"]["sites"][0]["suggested_formula"] == '=IF(B1="",NA(),A1-B1)'
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["keep_error_cells"]
    assert a["count"] == 1
    out = apply_operations(src, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED"
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["clean"] is True and res["counts"]["error_to_error"] == 1 and res["counts"]["good_to_other"] == 0
    assert _finding(build_analysis(Path(out["output_path"]), tmp_path / "after", "a"), "FRM-013")["status"] == "PASS"


@needs_excel
def test_a_blank_text_inside_a_range_in_array_arithmetic_is_named_in_the_test(tmp_path):
    """1.8.10 -- FRM-008 on a range (UC-4, 'Line 1'!H57 = IFERROR(SUMPRODUCT(I53:DL53*(...)),"")); test workbook 6:
    no range test tells "" from empty in Mind, a cell="" test does."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    for i, v in enumerate((1, 2, 3, 4), start=3):
        ws.cell(3, i, v); ws.cell(4, i, 10 * v)
    ws["E3"] = '=IF(1=1,"",0)'                                  # the "" cell inside C3:F3
    ws["B5"] = '=IFERROR(SUMPRODUCT(C3:F3*C4:F4),"")'          # "" in Excel, a number in Mind
    ws["B6"] = '=IFERROR(SUMPRODUCT(C4:F4*C4:F4),"")'          # no "" involved
    p = tmp_path / "range.xlsx"
    wb.save(p)
    wb.close()
    src = _calculated(p)
    analysis = build_analysis(src, tmp_path / "w", "r")
    f = _finding(analysis, "FRM-008")
    sites = {s["cell"]: s for s in f["observed"]["sites"]}
    assert set(sites) == {"B5"} and sites["B5"]["blank_cells"] == ["E3"]
    assert sites["B5"]["suggested_formula"] == '=IF(E3="","",IFERROR(SUMPRODUCT(C3:F3*C4:F4),""))'
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["guard_blank_arithmetic"]
    out = apply_operations(src, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED"
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["clean"] is True
    # a second round finds nothing more to guard (1.8.11: the first trial doubled the test)
    again = build_analysis(Path(out["output_path"]), tmp_path / "w2", "r2")
    assert _finding(again, "FRM-008")["status"] == "PASS"


def test_an_isnumber_or_iserror_guard_counts_like_iferror(tmp_path):
    """1.8.11 -- UC-4, Reporting!I23 = IF(ISNUMBER('Line 1'!H171*NOT(Line3_Boolean)),...,""): "" in Excel
    (the product is #VALUE!, ISNUMBER says FALSE), 0 in Mind (test workbook 7)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["C3"] = '=IF(1=1,"",0)'                             # the "" cell
    ws["C4"] = 5
    ws["B5"] = '=IF(ISNUMBER(C3*1),C3*1,"")'
    ws["B6"] = '=IF(ISERROR(C3*1),"",C3*1)'
    ws["B7"] = '=IF(ISNUMBER(C4*1),C4*1,"")'               # a number: nothing to guard
    ws["B8"] = '=IF(ISNUMBER(C3),C3*1,"")'                 # tests the cell itself: both say FALSE
    p = tmp_path / "isnumber.xlsx"
    wb.save(p)
    wb.close()
    src = _calculated(p)
    analysis = build_analysis(src, tmp_path / "w", "r")
    f = _finding(analysis, "FRM-008")
    sites = {s["cell"]: s for s in f["observed"]["sites"]}
    assert set(sites) == {"B5", "B6"}
    assert sites["B5"]["suggested_formula"] == '=IF(C3="","",IF(ISNUMBER(C3*1),C3*1,""))'
    assert sites["B6"]["suggested_formula"] == '=IF(C3="","",IF(ISERROR(C3*1),"",C3*1))'
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["guard_blank_arithmetic"]
    out = apply_operations(src, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and len(out["applied"]) == 2
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["clean"] is True


def test_a_lookup_through_defined_names_that_lands_on_an_empty_cell_is_an_origin(tmp_path):
    """1.8.12 -- UC-2, 'Other assumptions'!D67 = VLOOKUP(Policy_group,Assumptions_array,MATCH(B67,
    Assumption_headings,0),FALSE) lands on an empty cell: 0 in Excel, empty in Mind; Switch_Table=0 (a name on a cell
    that reads it) is then FALSE in Mind, and the run stops in the branch Excel never takes."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    inp = wb.active
    inp.title = "Inputs"
    inp["B3"], inp["C3"] = "Tariff", "A"
    for c, h in zip("BCD", ("Tariff", "X", "Switch table")):
        inp[f"{c}5"] = h
    inp["B6"], inp["C6"] = "A", 1                                  # D6 stays empty
    inp["B7"], inp["C7"], inp["D7"] = "B", 2, 5
    calc = wb.create_sheet("Calc")
    calc["A3"], calc["B3"] = "Switch table", '=VLOOKUP(Policy_group,Assumptions_array,MATCH(A3,Assumption_headings,0),FALSE)'
    calc["A4"], calc["B4"] = "copy", "=Switch_Table"
    calc["A5"], calc["B5"] = "switch", "=IF(Switch_Table=0,0,1)"
    for name, ref in (("Policy_group", "Inputs!$C$3"), ("Assumptions_array", "Inputs!$B$6:$D$7"), ("Assumption_headings", "Inputs!$B$5:$D$5"), ("Switch_Table", "Calc!$B$3")):
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    p = tmp_path / "names.xlsx"
    wb.save(p)
    wb.close()
    src = _calculated(p)
    analysis = build_analysis(src, tmp_path / "w", "r")
    f = _finding(analysis, "FRM-009")
    origins = {s["cell"]: s for s in f["observed"]["sites"] if s["kind"] == "lands_on_empty"}
    assert set(origins) == {"B3"} and origins["B3"]["lands_on"] == "Inputs!D6"   # B4 lands there too but nothing reads B4
    assert origins["B3"]["suggested_formula"] == "=N(VLOOKUP(Policy_group,Assumptions_array,MATCH(A3,Assumption_headings,0),FALSE))"
    assert origins["B3"]["first_reader"]["cell"] == "name Switch_Table"
    a = {x["id"]: x for x in plan_actions(analysis, {"findings": [f]})}["empty_cells_as_zero"]
    out = apply_operations(src, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED"
    res = compare_workbooks(src, Path(out["output_path"]), Moves())
    assert res["clean"] is True
