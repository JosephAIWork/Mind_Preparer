"""20 Nov 2026, Joseph's answers of the morning.

- CNHI: 2,320 formulas =INDIRECT("Cluster"&n), n built from ROW, COLUMN, ROUNDUP and COUNTA, over the names Cluster1..
  Cluster20 = 'Validation list'!B7..B26. "A: INDEX": each becomes INDEX('Validation list'!$B$7:$B$26,n) -- the same cell
  for every whole n, and Mind does not follow INDIRECT. Only a whole n is accepted; a number worked out from the cell's
  position alone must lie between 1 and the number of names.
- "App tries, never stops": a lone text inside a range a formula reads, where no title, label, header, move or text in
  the formula fits, gets a label in the empty cell beside it (left first, else right): the two touching texts are a
  small grid Mind imports. Seen in Mind on a made-up workbook the same day (COUNTA and INDEX over the range agree).
"""
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from app.config import load_config
from app.excel_com import com_available, close_quietly, excel_session, open_for_write, save_in_place
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule
from app.validators.references import _position_value, _whole_number

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
LIST = "'Validation list'!$B$2:$B$5"


def _actions(analysis):
    return {a["id"]: a for a in plan_actions(analysis, {"findings": []})}


def _recalc(p: Path) -> None:
    with excel_session() as excel:
        book = open_for_write(excel, p)
        try:
            excel.CalculateFullRebuild()
            save_in_place(book, p)
        finally:
            close_quietly(book)


# --- the number: whole by construction ------------------------------------------------------------------------
@pytest.mark.parametrize("expr", [
    "(ROW($A15)-ROW($A$14))", "ROUNDUP((ROW($A15)-ROW($A$14))/18,0)", "ROUNDUP((ROW($A15)-ROW($A$14))/COUNTA($B$11:$B$15),0)",
    "COLUMN(E$1)-COLUMN($D$1)", "INT(B4/3)+1", "2*ROWS(A1:A3)-1", "-(-3)", "TRUNC(B2)", "ROUNDDOWN(B2 / 2, 0)",
])
def test_a_number_built_from_whole_parts_is_whole(expr):
    assert _whole_number(expr)


@pytest.mark.parametrize("expr", [
    "B4", "(ROW($A15)-ROW($A$14))/18", "ROUNDUP(B4/18,1)", "ROUND(B4,2)", "MOD(B4,3)", "1.5", "2^3", "SUM(A1:A3)", "ROW(A1)&\"1\"", "",
])
def test_a_number_that_may_have_a_fraction_is_not(expr):
    assert not _whole_number(expr)


def test_a_number_from_the_cells_position_alone_is_worked_out():
    assert _position_value("(COLUMN(V$1)-COLUMN($B$1)-3)/17") == 1
    assert _position_value("(COLUMN(W$1)-COLUMN($B$1)-3)/17").denominator == 17
    assert _position_value("ROW(A5)*2-COLUMN(C1)") == 7
    assert _position_value("COUNTA(A1:A3)") is None and _position_value("B4/2") is None and _position_value("ROW(A:A)") is None


# --- INDIRECT("Cluster"&n) becomes INDEX over the list -------------------------------------------------------------
def _clusters(path: Path, formulas: dict[str, str], names: dict[str, str] | None = None) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["A1"], ws["B1"] = "Step", "Count"
    for r in range(2, 8):
        ws[f"A{r}"], ws[f"B{r}"] = r - 1, "x" if r < 4 else None
    for cell, f in formulas.items():
        ws[cell] = f
    lst = wb.create_sheet("Validation list")
    lst["B1"] = "Clusters"
    for k, name in enumerate(("North", "South", "East", "West"), start=1):
        lst[f"B{k + 1}"] = name
    for k in range(1, 5):
        wb.defined_names[f"Cluster{k}"] = DefinedName(f"Cluster{k}", attr_text=f"'Validation list'!$B${k + 1}")
    for n, ref in (names or {}).items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    wb.save(path)
    wb.close()
    return path


FORMULAS = {
    "D2": '=INDIRECT("Cluster"&(ROW($A3)-ROW($A$2)))',
    "D3": '=INDIRECT("Cluster"&ROUNDUP((ROW($A5)-ROW($A$2))/COUNTA($B$2:$B$7),0))',
    "D4": '=INDIRECT("Cluster"&((COLUMN(H$1)-COLUMN($B$1)-2)/2))',
    "D5": '="Region "&INDIRECT("Cluster"&(COLUMN(E$1)-COLUMN($D$1)))',
}


def test_indirect_over_numbered_names_becomes_index_over_their_cells(tmp_path):
    a = _actions(build_analysis(_clusters(tmp_path / "c.xlsx", FORMULAS), tmp_path / "w", "c"))["direct_references"]
    after = {o["cell"]: o["after"] for o in a["operations"]}
    assert after == {
        "D2": f"=INDEX({LIST},(ROW($A3)-ROW($A$2)))",
        "D3": f"=INDEX({LIST},ROUNDUP((ROW($A5)-ROW($A$2))/COUNTA($B$2:$B$7),0))",
        "D4": f"=INDEX({LIST},((COLUMN(H$1)-COLUMN($B$1)-2)/2))",
        "D5": f'="Region "&INDEX({LIST},(COLUMN(E$1)-COLUMN($D$1)))',
    }
    assert not a["skipped"]
    assert all("Cluster1..Cluster4" in o["note"] for o in a["operations"])


@pytest.mark.parametrize("formula, why", [
    ('=INDIRECT("Cluster"&MOD(ROW($A5),3))', "MOD is not taken as a whole number"),
    ('=INDIRECT("Cluster"&((COLUMN(F$1)-COLUMN($B$1))/3))', "the position works out to 4/3"),
    ('=INDIRECT("Cluster"&(COLUMN(H$1)-COLUMN($B$1)))', "the position works out to 6, past the 4 names"),
    ('=INDIRECT("Cluster"&(ROW($A2)-ROW($A$2)))', "the position works out to 0"),
    ('=INDIRECT("Region"&(ROW($A3)-ROW($A$2)))', "no names Region1.."),
])
def test_indirect_is_left_when_the_number_or_the_names_do_not_fit(tmp_path, formula, why):
    a = _actions(build_analysis(_clusters(tmp_path / "c.xlsx", {"D2": formula}), tmp_path / "w", "c"))["direct_references"]
    assert a["operations"] == [], why
    assert len(a["skipped"]) == 1


@pytest.mark.parametrize("extra", [
    {"Cluster0": "'Validation list'!$B$1"},     # INDIRECT("Cluster"&0) finds it, INDEX(..., 0) does not
    {"Cluster6": "'Validation list'!$B$7"},     # a gap: Cluster5 is missing
])
def test_indirect_is_left_when_the_names_are_not_one_unbroken_list(tmp_path, extra):
    a = _actions(build_analysis(_clusters(tmp_path / "c.xlsx", {"D2": FORMULAS["D2"]}, extra), tmp_path / "w", "c"))["direct_references"]
    assert a["operations"] == []


def test_indirect_is_left_when_the_names_are_not_neighbours(tmp_path):
    p = _clusters(tmp_path / "c.xlsx", {"D2": FORMULAS["D2"]})
    wb = openpyxl.load_workbook(p)
    wb.defined_names["Cluster3"] = DefinedName("Cluster3", attr_text="'Validation list'!$C$4")
    wb.save(p)
    a = _actions(build_analysis(p, tmp_path / "w", "c"))["direct_references"]
    assert a["operations"] == []


@needs_excel
def test_the_index_rewrite_changes_no_value_in_excel(tmp_path):
    p = _clusters(tmp_path / "c.xlsx", FORMULAS)
    _recalc(p)
    analysis = build_analysis(p, tmp_path / "w", "c")
    a = _actions(analysis)["direct_references"]
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    done = Path(out["output_path"])
    assert compare_workbooks(p, done, Moves([a["operations"]]))["clean"] is True
    book = openpyxl.load_workbook(done, data_only=True)
    assert [book["Calc"][c].value for c in ("D2", "D3", "D4", "D5")] == ["North", "South", "South", "Region North"]
    after = build_analysis(done, tmp_path / "after", "a")
    assert run_rule(RulesEngine().get("FRM-007"), after, load_config())["status"] == "PASS"


# --- a label beside a lone text that sits inside a range a formula reads --------------------------------------------
def _columns(path: Path, col: int, extra: dict[str, object] | None = None) -> Path:
    """Numbers in two blocks of column `col` (and the next), a text alone between them, read by COUNTA and INDEX."""
    wb = openpyxl.Workbook()
    probe = wb.active
    probe.title = "Probe"
    letter = openpyxl.utils.get_column_letter(col)
    probe["A1"] = f"=COUNTA(Col!${letter}$7:${letter}$23)"
    probe["A2"] = f"=INDEX(Col!${letter}$7:${letter}$23,9)"
    ws = wb.create_sheet("Col")
    for i, (a, b) in enumerate(zip((3, 5, 2, 8, 6, 4), (2, 2, 2, 2, 2, 2))):
        ws.cell(7 + i, col, a)
        ws.cell(7 + i, col + 1, b)
    ws.cell(15, col, "Second part")
    for i, (a, b) in enumerate(zip((7, 1, 9, 2, 5, 3), (3, 3, 3, 3, 3, 3))):
        ws.cell(18 + i, col, a)
        ws.cell(18 + i, col + 1, b)
    for ref, v in (extra or {}).items():
        sheet, cell = ref.split("!")
        wb[sheet][cell] = v
    wb.save(path)
    wb.close()
    return path


@pytest.mark.parametrize("col, label", [(3, "B15"), (1, "B15")])
def test_a_lone_text_inside_a_read_range_gets_a_label_beside_it(tmp_path, col, label):
    a = _actions(build_analysis(_columns(tmp_path / "c.xlsx", col), tmp_path / "w", "c"))["title_lone_text_cells"]
    cell = f"{openpyxl.utils.get_column_letter(col)}15"
    assert [(o["op"], o["cell"], o["after"]) for o in a["operations"]] == [("set_value", label, f"Text in {cell}")]
    assert ("left" if col == 3 else "right") in a["operations"][0]["note"] and a["skipped"] == []


@pytest.mark.parametrize("extra, why", [
    ({"Probe!A3": "=SUM(Col!$B$7:$B$23)"}, "a formula reads the label's cell"),
    ({"Col!B14": "x"}, "a cell touching the label would join the pair to something else"),
    ({"Col!C16": 1}, "a cell touching the label at a corner would join the pair to something else"),
])
def test_no_label_where_it_could_be_read_or_join_something(tmp_path, extra, why):
    a = _actions(build_analysis(_columns(tmp_path / "c.xlsx", 1, extra), tmp_path / "w", "c"))["title_lone_text_cells"]
    assert not any(o["cell"] == "B15" for o in a["operations"]), why
    assert any(s.startswith("Col!A15:") for s in a["skipped"]), why


@needs_excel
def test_the_label_changes_no_value_and_the_cell_no_longer_stands_alone(tmp_path):
    p = _columns(tmp_path / "c.xlsx", 1)
    _recalc(p)
    a = _actions(build_analysis(p, tmp_path / "w", "c"))["title_lone_text_cells"]
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    done = Path(out["output_path"])
    assert compare_workbooks(p, done, Moves([a["operations"]]))["clean"] is True
    after = build_analysis(done, tmp_path / "after", "a")
    assert run_rule(RulesEngine().get("STR-008"), after, load_config())["observed"]["sites"] == []


# --- after the Haiku 5.5 refuter (20 Nov 2026) ----------------------------------------------------------------------
@pytest.mark.parametrize("formula, why", [
    ('=INDIRECT("Cluster"&(ROW($A5)-ROW($A$5)+COUNTA($B$2:$B$7)))', "0 when column B is empty: INDEX(..., 0) is the whole list"),
    ('=INDIRECT("Cluster"&(COUNTA($B$2:$B$7)-2))', "may be below 1"),
    ('=INDIRECT("Cluster"&INT(ROW($A5)/2-COUNTA($B$2:$B$7)))', "a count taken away: may be 0 or less"),
])
def test_indirect_is_left_when_the_number_may_be_below_one(tmp_path, formula, why):
    a = _actions(build_analysis(_clusters(tmp_path / "c.xlsx", {"D2": formula}), tmp_path / "w", "c"))["direct_references"]
    assert a["operations"] == [], why


def test_a_count_divided_and_rounded_up_is_never_below_one():
    from app.validators.references import _whole_bounds

    assert _whole_bounds("ROUNDUP((ROW($A15)-ROW($A$14))/COUNTA($B$11:$B$15),0)")[0] == 1
    assert _whole_bounds("ROUNDUP((ROW($A14)-ROW($A$14))/COUNTA($B$11:$B$15),0)")[0] == 0
    assert _whole_bounds("(ROW($A15)-ROW($A$14))/18") is None and _whole_bounds('"a"&1') is None


def test_a_short_prefix_that_spells_a_cell_is_left(tmp_path):
    p = tmp_path / "c.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["D2"] = '=INDIRECT("CN"&(ROW($A3)-ROW($A$2)))'
    lst = wb.create_sheet("List")
    lst["B1"] = "Names"
    for k in range(1, 4):
        lst[f"B{k + 1}"] = f"n{k}"
        wb.defined_names[f"CN{k}"] = DefinedName(f"CN{k}", attr_text=f"List!$B${k + 1}")
    wb.save(p)
    a = _actions(build_analysis(p, tmp_path / "w", "c"))["direct_references"]
    assert a["operations"] == []  # "CN"&21 would be the cell CN21


@pytest.mark.parametrize("reader", ["=OFFSET(Col!$D$15,0,-2)", "=SUM(INDEX(Col!$A$1:$A$30,15):Col!$A$16)"])
def test_no_label_on_a_sheet_offset_or_an_index_range_reads(tmp_path, reader):
    a = _actions(build_analysis(_columns(tmp_path / "c.xlsx", 1, {"Probe!A3": reader}), tmp_path / "w", "c"))["title_lone_text_cells"]
    assert not any(o["cell"] == "B15" for o in a["operations"])
