"""2026-11-16: the lone-text-cell rule (STR-008) is fixed in ONE Prep pass.

Model D (an IFRS 17 tool), sheet 'ロールフォワード': under the grid U21:AA23 the texts V24, W25, X26, Y27 stand one row
and one column further each time (a staircase). Every cell of those rows is read by a formula on its own row
(MATCH(X25, ...)), so no title may go above a step, and only the step right under the grid could join it:
one cell a pass (v5 V24, v6 W25, v7 X26, v8 Y27), and "Do everything" ran out of passes and stopped for a person.
The planner now sees the sheet as its own earlier operations will leave it, so the whole staircase goes in one pass.
"""
from pathlib import Path

import openpyxl
import pytest

from app.config import load_config
from app.excel_com import com_available
from app.grids import all_grids
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
ACTION = "title_lone_text_cells"


def _finding(analysis):
    return run_rule(RulesEngine().get("STR-008"), analysis, load_config())


def _action(analysis):
    return {a["id"]: a for a in plan_actions(analysis, {"findings": []})}[ACTION]


def _sites(analysis):
    return [s["cell"] for s in _finding(analysis)["observed"]["sites"]]


def _recalculated(path: Path) -> Path:
    """Excel computes the workbook once, so the numbers check has the values the original shows."""
    from app.excel_com import close_quietly, excel_session, open_for_write, save_in_place

    with excel_session() as excel:
        book = open_for_write(excel, path)
        try:
            excel.CalculateFullRebuild()
            save_in_place(book, path)
        finally:
            close_quietly(book)
    return path


def _staircase(path: Path, steps=("V24", "W25", "X26", "Y27"), extra: dict | None = None) -> Path:
    """The Model D shape: field names in U21:Z23 (the labels in U), then one name a row, a column further each
    time. Columns D:H hold one formula per cell of V:Z on rows 21-27 (as the real MATCH(X25, ...)): the empty
    cells above each step are read, the label column U is read on the grid's own rows only."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Roll"
    for r in (21, 22, 23):
        for c in "UVWXYZ":
            ws[f"{c}{r}"] = f"LI_{c}{r}"
        ws[f"C{r}"] = f'=LEN(U{r})'
    for i, cell in enumerate(steps):
        ws[cell] = f"plan_{i + 1}"
    for r in range(21, 28):
        for j, c in enumerate("VWXYZ"):
            ws.cell(r, 4 + j).value = f'=IF({c}{r}="",0,LEN({c}{r}))'
    for ref, value in (extra or {}).items():
        ws[ref] = value
    wb.save(path)
    wb.close()
    return path


def test_a_staircase_under_a_grid_is_reported_whole_and_fixed_in_one_pass(tmp_path):
    p = _staircase(tmp_path / "roll.xlsx")
    analysis = build_analysis(p, tmp_path / "w", "s")
    f = _finding(analysis)
    assert f["status"] == "ERROR" and f["observed"]["sites_count"] == 4 and _sites(analysis) == ["V24", "W25", "X26", "Y27"]
    a = _action(analysis)
    assert not a["skipped"], a["skipped"]
    assert [(o["op"], o["cell"], o["after"]) for o in a["operations"]] == [
        ("set_value", "U24", "Value in V24"), ("set_value", "U25", "Value in W25"),
        ("set_value", "U26", "Value in X26"), ("set_value", "U27", "Value in Y27")]
    # each step joins the grid the step before it made one row longer
    assert ["the grid U21:Z23" in a["operations"][0]["note"], "the grid U21:Z24" in a["operations"][1]["note"],
            "the grid U21:Z25" in a["operations"][2]["note"], "the grid U21:Z26" in a["operations"][3]["note"]] == [True] * 4
    out = apply_operations(p, tmp_path / "prep", a["operations"], prefer_excel=False)
    assert out["status"] == "APPLIED"
    after = build_analysis(Path(out["output_path"]), tmp_path / "after", "a")
    assert _finding(after)["status"] == "PASS"                       # nothing left for a second pass
    assert not _action(after)["operations"]
    grids = [g for g in all_grids(after) if g["sheet"] == "Roll"]
    assert any(g["ref"] == "U21:Z27" for g in grids) and not any(g["inner_title_cells"] for g in grids)


def test_the_plan_of_one_pass_is_what_the_passes_one_by_one_would_have_done(tmp_path):
    """Pass by pass the old way (only the first operation of each plan applied), against the one pass: the same cells."""
    p = _staircase(tmp_path / "roll.xlsx")
    one_pass = [(o["cell"], o["after"]) for o in _action(build_analysis(p, tmp_path / "w", "s"))["operations"]]
    path, by_hand = p, []
    for n in range(6):
        ops = _action(build_analysis(path, tmp_path / f"w{n}", f"s{n}"))["operations"][:1]
        if not ops:
            break
        by_hand += [(o["cell"], o["after"]) for o in ops]
        path = Path(apply_operations(path, tmp_path / f"p{n}", ops, prefer_excel=False)["output_path"])
    assert one_pass == by_hand and len(one_pass) == 4


def test_a_step_whose_label_cell_a_lookup_reads_stops_the_staircase_there(tmp_path):
    """The protection stays: U26 lies inside a range a MATCH reads, so no label goes there, and no title goes into
    a cell a formula reads. The two steps before are still fixed; X26, now right under the grid, moves to a column
    of its own as such a cell always did (only single cells read it). Y27 (no title, no label, no move: only G27 reads
    it, through one reference) gets its text written into G27 in place of the reference (17 Nov 2026; it was left for
    a person before)."""
    p = _staircase(tmp_path / "roll.xlsx", extra={"A1": '=IFERROR(MATCH("x",$U$26:$U$30,0),0)'})
    a = _action(build_analysis(p, tmp_path / "w", "s"))
    assert [(o["op"], o["cell"], o.get("to") or o["after"]) for o in a["operations"]] == [
        ("set_value", "U24", "Value in V24"), ("set_value", "U25", "Value in W25"), ("move_cell", "X26", "AB26"), ("set_value", "AB25", "#Roll X26"),
        ("set_formula", "G27", '=IF("plan_4"="",0,LEN("plan_4"))')]
    assert a["skipped"] == []
    assert not any(o["cell"] in ("U26", "U27", "X25", "Y26") for o in a["operations"])


def test_a_step_that_a_range_reads_and_cannot_join_is_left_for_a_person(tmp_path):
    """The same, with X26 read through a range as well (COUNTIF needs its text): it cannot move nor take its text into
    the formula, and says so. Y27 takes its text into G27 (17 Nov 2026)."""
    p = _staircase(tmp_path / "roll.xlsx", extra={"A1": '=IFERROR(MATCH("x",$U$26:$U$30,0),0)', "A2": '=COUNTIF(X26:X27,"plan_3")'})
    a = _action(build_analysis(p, tmp_path / "w", "s"))
    assert [(o["cell"], o["after"]) for o in a["operations"]] == [("U24", "Value in V24"), ("U25", "Value in W25"), ("G27", '=IF("plan_4"="",0,LEN("plan_4"))')]
    assert len(a["skipped"]) == 1 and a["skipped"][0].endswith("-- by hand")
    assert a["skipped"][0].startswith("Roll!X26: the cell above (X25) is taken, and a range or a name over several cells reads it (it cannot move)")


def test_two_lone_cells_on_the_row_under_a_grid_take_one_label(tmp_path):
    """W24 and Y24 both under the grid: one label takes the row, and the second cell asks for nothing more
    (it used to plan the same label again, dropped with "run Prep again for this one")."""
    p = _staircase(tmp_path / "roll.xlsx", steps=("W24", "Y24"))
    analysis = build_analysis(p, tmp_path / "w", "s")
    assert _sites(analysis) == ["W24", "Y24"]
    a = _action(analysis)
    assert [(o["cell"], o["after"]) for o in a["operations"]] == [("U24", "Value in W24")] and not a["skipped"]
    out = apply_operations(p, tmp_path / "prep", a["operations"], prefer_excel=False)
    assert _finding(build_analysis(Path(out["output_path"]), tmp_path / "after", "a"))["status"] == "PASS"


@needs_excel
def test_the_staircase_fixed_in_one_pass_changes_no_value_in_excel(tmp_path):
    p = _recalculated(_staircase(tmp_path / "roll.xlsx"))
    a = _action(build_analysis(p, tmp_path / "w", "s"))
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"] and len(out["applied"]) == 4
    done = Path(out["output_path"])
    check = compare_workbooks(p, done, Moves([a["operations"]]))
    assert check["clean"] is True and check["counts"]["compared"] == 38 and check["counts"]["same"] == 38
    assert _finding(build_analysis(done, tmp_path / "after", "a"))["status"] == "PASS"


def _diagonal(path: Path, middle_read_through_a_range: bool = False) -> Path:
    """Three texts on a diagonal with nothing around them, each read by a formula; the cells above them are free."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Diag"
    ws["F10"], ws["G11"], ws["H12"] = "a", "b", "c"
    ws["A1"], ws["A3"] = "=LEN(F10)", "=LEN(H12)"
    ws["A2"] = '=COUNTIF(G11:G13,"b")' if middle_read_through_a_range else "=LEN(G11)"
    wb.save(path)
    wb.close()
    return path


def test_titles_on_a_diagonal_are_not_planned_where_they_would_trap_each_other(tmp_path):
    """Before: three titles F9, G10, H11 in one pass. G10 sits right of F10, so the two were read as one grid
    (title F9, the '#' title G10 trapped in it), G11 and H11 as another, and H12 stayed alone: a second pass,
    and two trapped titles nobody could clear. Now the middle cell moves to a column of its own instead."""
    p = _diagonal(tmp_path / "d.xlsx")
    analysis = build_analysis(p, tmp_path / "w", "d")
    assert _sites(analysis) == ["F10", "G11", "H12"]
    a = _action(analysis)
    assert not a["skipped"], a["skipped"]
    assert [(o["op"], o["cell"], o.get("to") or o["after"]) for o in a["operations"]] == [
        ("set_value", "F9", "#Diag F10"), ("move_cell", "G11", "J11"), ("set_value", "J10", "#Diag G11"), ("set_value", "H11", "#Diag H12")]


def test_a_cell_on_a_diagonal_that_cannot_move_is_left_for_a_person_with_the_reason(tmp_path):
    p = _diagonal(tmp_path / "d.xlsx", middle_read_through_a_range=True)
    a = _action(build_analysis(p, tmp_path / "w", "d"))
    cells = [(o["op"], o["cell"]) for o in a["operations"]]
    assert ("set_value", "F9") in cells and ("set_value", "G10") not in cells and ("set_value", "H11") not in cells
    assert len(a["skipped"]) == 1 and "Diag!G11" in a["skipped"][0] and "one grid with the text beside it (F10)" in a["skipped"][0] and "by hand" in a["skipped"][0]
    assert ("move_cell", "H12") in cells     # G11 stays where it is, so a title in H11 would join it: H12 moves instead


@needs_excel
def test_the_diagonal_is_fixed_in_one_pass_in_excel_and_the_numbers_stay(tmp_path):
    p = _recalculated(_diagonal(tmp_path / "d.xlsx"))
    a = _action(build_analysis(p, tmp_path / "w", "d"))
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    done = Path(out["output_path"])
    book = openpyxl.load_workbook(done)
    assert book["Diag"]["A2"].value == "=LEN(J11)" and book["Diag"]["J11"].value == "b" and book["Diag"]["G11"].value is None
    assert compare_workbooks(p, done, Moves([a["operations"]]))["clean"] is True
    after = build_analysis(done, tmp_path / "after", "a")
    assert _finding(after)["status"] == "PASS"
    assert not any(g["inner_title_cells"] for g in all_grids(after))


def _beside(path: Path, extra: dict | None = None, lookup: str = "$C$36:$C$45") -> Path:
    """The IFRS risk model's shape: a caption over a column of names (the grid B34:B45, one column wide) and, beside
    it, values with holes in them. An INDEX reads the value column, so no title may go into its empty cells."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Parameters"
    ws["B34"], ws["B35"], ws["C35"] = "Input Translation Key", "Name", "Used?"
    for r in range(36, 46):
        ws[f"B{r}"] = f"field {r}"
    for r in (38, 39, 41, 44):                       # C38:C39 a small grid; C41 and C44 stand alone
        ws[f"C{r}"] = "Not Used"
    for ref, value in (extra or {}).items():
        ws[ref] = value
    calc = wb.create_sheet("Calc")
    for i, r in enumerate((36, 38, 41, 44)):
        calc[f"A{i + 1}"] = f'=INDEX(Parameters!{lookup},MATCH("field {r}",Parameters!$B$36:$B$45,0))'
    wb.save(path)
    wb.close()
    return path


def test_lone_cells_beside_a_one_column_grid_join_it_with_one_header(tmp_path):
    """Ten 'Not Used' cells of the real model were left "by hand": the cell above each lies inside the INDEX's
    range, and the range keeps them from moving. One header beside the caption makes the grid two columns wide."""
    p = _beside(tmp_path / "p.xlsx")
    analysis = build_analysis(p, tmp_path / "w", "p")
    assert _sites(analysis) == ["C41", "C44"]
    a = _action(analysis)
    assert [(o["op"], o["cell"], o["after"]) for o in a["operations"]] == [("set_value", "C34", "Values in column C")] and not a["skipped"]
    assert "the grid B34:B45" in a["operations"][0]["note"]
    out = apply_operations(p, tmp_path / "prep", a["operations"], prefer_excel=False)
    after = build_analysis(Path(out["output_path"]), tmp_path / "after", "a")
    assert _finding(after)["status"] == "PASS"
    grids = [g for g in all_grids(after) if g["sheet"] == "Parameters"]
    assert [g["ref"] for g in grids] == ["B34:C45"] and not grids[0]["inner_title_cells"]


@pytest.mark.parametrize("why, extra, lookup", [
    ("the header's cell lies inside the range the INDEX reads", {}, "$C$34:$C$45"),
    ("a '#' title in the column (of C38:C39) would be trapped in the grid", {"C37": "#Rates"}, "$C$36:$C$45"),
    ("a grid two columns wide starts in the column: it would be cut", {"D38": "x", "D39": "y"}, "$C$36:$C$45"),
    ("something stands right of the header's cell: the grid would take more than this column", {"D34": "note"}, "$C$36:$C$45"),
    ("something stands above the header's cell: it would take the header", {"C33": "note"}, "$C$36:$C$45"),
    ("the grid carries a flag (an input keeps its columns)", {"B33": "#Keys /Input"}, "$C$36:$C$45"),
])
def test_no_header_beside_a_grid_when_it_would_change_more_than_the_column(tmp_path, why, extra, lookup):
    p = _beside(tmp_path / "p.xlsx", extra=extra, lookup=lookup)
    a = _action(build_analysis(p, tmp_path / "w", "p"))
    assert not any(o["cell"] == "C34" for o in a["operations"]), why
    assert any("Parameters!C41" in s and "by hand" in s for s in a["skipped"]), (why, a["skipped"])


@needs_excel
def test_the_header_beside_a_grid_changes_no_value_in_excel(tmp_path):
    p = _recalculated(_beside(tmp_path / "p.xlsx"))
    a = _action(build_analysis(p, tmp_path / "w", "p"))
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    done = Path(out["output_path"])
    check = compare_workbooks(p, done, Moves([a["operations"]]))
    assert check["clean"] is True and check["counts"]["compared"] == 4
    assert openpyxl.load_workbook(done)["Parameters"]["C34"].value == "Values in column C"
