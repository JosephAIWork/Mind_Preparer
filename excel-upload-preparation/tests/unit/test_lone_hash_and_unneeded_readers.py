"""2026-11-17: Phoenix TVOG Calc stopped Start on "a formula reads a lone text cell, in 6 cells; the first is E62 on sheet CFs".

- CFs!B29 holds a bare '#', the first header of the table B29:L110 (a column of row numbers). It was read as a grid
  title, so the table fell into pieces and four lookup keys inside it (F66, F78, F86, F109) stood alone. A bare '#'
  with headers right of it is now a header, not a title.
- F109 still stood alone: the row numbers in column B skip B101:B103, and Mind reads a grid down its first column.
  Prep writes a label in each of those empty cells, and the grid runs on to row 110.
- CFs!E62 and Calc!A259 were false alarms: every formula that reads them gets the same result with the cell empty
  (VLOOKUP(key, C30:G110, 5, FALSE) reads only C and G; MAX($A$7:A304) skips text). They are not counted.
- A formula that does need the text, where no title, label or move fits, gets the text itself in place of the
  reference (same value in Excel, and Mind no longer needs the cell).
"""
from pathlib import Path

import openpyxl
import pytest

from app.config import load_config
from app.excel_com import com_available
from app.grids import all_grids, detect_grids
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
ACTION = "title_lone_text_cells"


def _finding(analysis):
    return run_rule(RulesEngine().get("STR-008"), analysis, load_config())


def _sites(analysis):
    return [s["cell"] for s in _finding(analysis)["observed"]["sites"]]


def _action(analysis):
    return {a["id"]: a for a in plan_actions(analysis, {"findings": []})}[ACTION]


def _occupied(cells: dict[str, object]) -> dict[tuple[int, int], object]:
    out = {}
    for ref, v in cells.items():
        c = openpyxl.utils.cell.coordinate_from_string(ref)
        out[(c[1], openpyxl.utils.column_index_from_string(c[0]))] = v
    return out


# --- the bare '#' header --------------------------------------------------------------------------------------
def _table(first_header: str) -> dict[tuple[int, int], object]:
    cells = {"B2": first_header, "C2": "Key", "D2": "Value"}
    for r in range(3, 8):
        cells.update({f"B{r}": r - 2, f"C{r}": f"k{r}", f"D{r}": r * 10})
    return _occupied(cells)


def test_a_bare_hash_heading_a_row_of_headers_is_a_header_not_a_title():
    grids, lone = detect_grids("S", _table("#"))
    assert [(g["ref"], g["title_cell"]) for g in grids] == [("B2:D7", None)]
    assert grids[0]["header_values"] == ["#", "Key", "Value"] and not grids[0]["inner_title_cells"] and not lone


@pytest.mark.parametrize("title", ["#Rates", "# /Input", "#Rates /Resize.A"])
def test_a_named_or_flagged_title_is_still_a_title(title):
    grids, _ = detect_grids("S", _table(title))
    assert any(g["title_cell"] == "B2" and g["ref"] == "B3:D7" for g in grids)


def test_a_bare_hash_with_nothing_right_of_it_is_still_read_as_before():
    grids, _ = detect_grids("S", _occupied({"B2": "#", "B3": 1, "B4": 2}))
    assert [(g["ref"], g["title_cell"]) for g in grids] == [("B3:B4", "B2")]


def test_a_bare_hash_among_the_headers_is_no_trapped_title():
    cells = {"B2": "Key", "C2": "#", "D2": "Value"}
    for r in range(3, 6):
        cells.update({f"B{r}": f"k{r}", f"C{r}": r, f"D{r}": r * 10})
    grids, _ = detect_grids("S", _occupied(cells))
    assert [(g["ref"], g["inner_title_cells"]) for g in grids] == [("B2:D5", [])]


# --- the Phoenix shape -----------------------------------------------------------------------------------------
def _phoenix(path: Path, gap_read_by_a_range: bool = False) -> Path:
    """CFs: a table B29:L50 headed '#' (row numbers in B, keys in C and F, a note column E, the factor in G), its row
    numbers skipping B44:B46; F48 a key on a row past the gap; E40 a note. Calc: VLOOKUP(key, CFs!$C$30:$G$50, 5) as in
    the model, VLOOKUP(key, tab, 2) over the name tab = CFs!$F$30:$G$50 (keys in F), and a counter MAX($A$7:A..)+1
    over a section label A20 beside a small grid."""
    wb = openpyxl.Workbook()
    cfs = wb.active
    cfs.title = "CFs"
    for c, h in zip("BCDEFGHIJKL", ["#", "CF_Item", "Hebrew", "Desc", "Source", "Factor", "Timing", "Profit", "PV Gross", "PV Reins", "CF Item"]):
        cfs[f"{c}29"] = h
    for r in range(30, 51):
        if not 44 <= r <= 46:
            cfs[f"B{r}"] = f"=B{r - 1 if r != 47 else 43}+1" if r > 30 else 1
        cfs[f"C{r}"] = f"item{r}"
        if r not in (44, 45, 46):
            cfs[f"D{r}"] = f"he{r}"
        cfs[f"F{r}"] = f"SRC{r}" if (r % 3 == 0 and r < 44) or r == 48 else None
        cfs[f"G{r}"] = r % 2 or -1
        for c in "HIJK":
            cfs[f"{c}{r}"] = 0
        cfs[f"L{r}"] = f"=C{r}"
    cfs["E40"] = "Original Name: CFCOMM"
    if gap_read_by_a_range:
        cfs["A60"] = "=COUNTA(B44:B46)"
    calc = wb.create_sheet("Calc")
    wb.defined_names["tab"] = openpyxl.workbook.defined_name.DefinedName("tab", attr_text="CFs!$F$30:$G$50")
    for i, r in enumerate(range(3, 12)):
        key_row = 30 + i * 2
        calc[f"C{r}"] = f"item{key_row}"
        calc[f"D{r}"] = f"=IFERROR(VLOOKUP(C{r},CFs!$C$30:$G$50,5,FALSE),0)"
        calc[f"E{r}"] = f'=IFERROR(VLOOKUP("SRC48",tab,2,FALSE),0)'
    calc["A20"] = "LH2_calc"
    for r in range(18, 24):
        calc[f"B{r}"], calc[f"C{r}"] = f"label{r}", r
    calc["A30"] = "=MAX($A$7:A29)+1"
    calc["A31"] = "=MAX($A$7:A30)+1"
    wb.save(path)
    wb.close()
    return path


def test_phoenix_only_the_real_key_past_the_gap_is_left_and_labels_close_the_gap(tmp_path):
    p = _phoenix(tmp_path / "phoenix.xlsx")
    analysis = build_analysis(p, tmp_path / "w", "s")
    cfs = [g for g in all_grids(analysis) if g["sheet"] == "CFs"]
    assert any(g["ref"] == "B29:L43" and g["title_cell"] is None for g in cfs)  # the '#' heads the table, E40 inside it
    # Calc!A20 (beside a grid, read only by MAX($A$7:A29)+1, which skips text) is no longer reported
    assert _sites(analysis) == ["F48"]
    a = _action(analysis)
    assert not a["skipped"], a["skipped"]
    assert [(o["op"], o["cell"], o["after"]) for o in a["operations"]] == [
        ("set_value", "B44", "Row 44"), ("set_value", "B45", "Row 45"), ("set_value", "B46", "Row 46")]
    assert "runs on to row 50" in a["operations"][0]["note"]
    out = apply_operations(p, tmp_path / "prep", a["operations"], prefer_excel=False)
    assert out["status"] == "APPLIED"
    after = build_analysis(Path(out["output_path"]), tmp_path / "after", "a")
    assert _finding(after)["status"] == "PASS"
    assert any(g["ref"] == "B29:L50" for g in all_grids(after) if g["sheet"] == "CFs")
    assert not _action(after)["operations"]


def test_no_label_goes_into_a_gap_cell_a_formula_reads(tmp_path):
    p = _phoenix(tmp_path / "phoenix.xlsx", gap_read_by_a_range=True)
    a = _action(build_analysis(p, tmp_path / "w", "s"))
    assert not any(o["cell"] in ("B44", "B45", "B46") for o in a["operations"])
    assert any(s.startswith("CFs!F48:") and s.endswith("-- by hand") for s in a["skipped"])


# --- which readers need the text --------------------------------------------------------------------------------
def _one_reader(path: Path, formula: str, lone: str = "D5", text: str = "abc") -> Path:
    """A text alone in D5 with a grid A1:B3 far from it, read by one formula in H1 (the cell above, D4, is read too,
    so no title can go there)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["B1"], ws["A2"], ws["B2"], ws["A3"], ws["B3"] = "k", "v", "x", 1, "y", 2
    ws[lone] = text
    ws["H1"] = formula
    ws["H3"] = "=D4*2"
    wb.save(path)
    wb.close()
    return path


@pytest.mark.parametrize("formula", ["=MAX($D$1:D9)+1", "=SUM(D1:D9)", "=COUNT(D2:D9)", "=AVERAGE(D2:D9,B2)", "=IFERROR(VLOOKUP(\"k\",C2:F9,4,FALSE),0)"])
def test_a_reader_that_gets_the_same_result_from_an_empty_cell_does_not_count(tmp_path, formula):
    assert _sites(build_analysis(_one_reader(tmp_path / "s.xlsx", formula), tmp_path / "w", "s")) == []


@pytest.mark.parametrize("formula", ["=COUNTA(D1:D9)", "=VLOOKUP(\"abc\",D2:F9,2,FALSE)", "=VLOOKUP(\"k\",B2:D9,3,FALSE)", "=SUMPRODUCT((D1:D9=\"abc\")*1)",
                                     "=INDEX(D1:D9,5)", "=VLOOKUP(\"k\",B2:F9,A1,FALSE)"])
def test_a_reader_that_needs_the_text_still_counts(tmp_path, formula):
    assert _sites(build_analysis(_one_reader(tmp_path / "s.xlsx", formula), tmp_path / "w", "s")) == ["D5"]


def test_a_name_some_formula_uses_still_counts(tmp_path):
    p = _one_reader(tmp_path / "s.xlsx", "=MAX($D$1:D9)+1")
    wb = openpyxl.load_workbook(p)
    wb.defined_names["keys"] = openpyxl.workbook.defined_name.DefinedName("keys", attr_text="S!$D$2:$E$9")
    wb["S"]["H5"] = '=VLOOKUP("abc",keys,2,FALSE)'
    wb.save(p)
    assert _sites(build_analysis(p, tmp_path / "w", "s")) == ["D5"]


# --- the text into the formula -----------------------------------------------------------------------------------
def test_a_formula_that_needs_the_text_gets_it_in_place_of_the_reference(tmp_path):
    p = _one_reader(tmp_path / "s.xlsx", '=IF(D5="abc","yes "&D5,LEN(D5))')
    wb = openpyxl.load_workbook(p)
    wb["S"]["H7"] = "=MAX(D1:D9)+D5&\"!\""  # the range is left alone, the single reference takes the text
    wb["S"]["J5"] = 'say "hi"'
    wb["S"]["H9"] = "=J5"
    wb["S"]["H10"] = "=J4+J5"  # J4 is read: no title above J5 either
    wb.save(p)
    analysis = build_analysis(p, tmp_path / "w", "s")
    assert _sites(analysis) == ["D5", "J5"]
    a = _action(analysis)
    assert not a["skipped"], a["skipped"]
    got = {(o["cell"], o["after"]) for o in a["operations"] if o["op"] == "set_formula"}
    assert ("H1", '=IF("abc"="abc","yes "&"abc",LEN("abc"))') in got
    assert ("H7", '=MAX(D1:D9)+"abc"&"!"') in got
    assert ("H9", '="say ""hi"""') in got and ("H10", '=J4+"say ""hi"""') in got
    out = apply_operations(p, tmp_path / "prep", a["operations"], prefer_excel=False)
    assert _finding(build_analysis(Path(out["output_path"]), tmp_path / "after", "a"))["status"] == "PASS"


@pytest.mark.parametrize("formula", ["=ROW(D5)", "=CELL(\"address\",D5)", "=SUM(D5:INDEX(D1:D9,7))", "=MM_INSTANCE(D5)", "=OFFSET(D5,1,0)"])
def test_no_text_goes_where_excel_wants_a_reference(tmp_path, formula):
    p = _one_reader(tmp_path / "s.xlsx", formula)
    wb = openpyxl.load_workbook(p)
    wb["S"]["H10"] = '=D5&"x"'
    wb.save(p)
    a = _action(build_analysis(p, tmp_path / "w", "s"))
    assert not any(o["op"] == "set_formula" for o in a["operations"])


def test_two_lone_cells_in_one_formula_are_rewritten_one_pass_at_a_time(tmp_path):
    p = _one_reader(tmp_path / "s.xlsx", '=D5&J5')
    wb = openpyxl.load_workbook(p)
    wb["S"]["J5"] = "def"
    wb["S"]["H11"] = "=J4+1"  # J4 is read: no title above J5
    wb.save(p)
    a = _action(build_analysis(p, tmp_path / "w", "s"))
    rewrites = [(o["cell"], o["after"]) for o in a["operations"] if o["op"] == "set_formula"]
    assert rewrites == [("H1", '="abc"&J5')]
    # 20 Nov 2026: J5 no longer waits for the next pass -- its formula is taken this pass, so it gets a label beside it
    labels = [o["cell"] for o in a["operations"] if o["op"] == "set_value" and o["after"] == "Text in J5"]
    assert len(labels) == 1 and labels[0] in ("I5", "K5") and not any(s.startswith("S!J5:") for s in a["skipped"])


@needs_excel
def test_phoenix_shape_and_the_text_rewrite_change_no_value_in_excel(tmp_path):
    from app.excel_com import close_quietly, excel_session, open_for_write, save_in_place

    p = _phoenix(tmp_path / "phoenix.xlsx")
    wb = openpyxl.load_workbook(p)
    wb["Calc"]["H1"] = '=IF(Calc!A20="LH2_calc",1,2)&Calc!A20'
    wb["Calc"]["H2"] = "=A19+1"  # A19 is read: no title above A20
    wb.save(p)
    with excel_session() as excel:
        book = open_for_write(excel, p)
        try:
            excel.CalculateFullRebuild()
            save_in_place(book, p)
        finally:
            close_quietly(book)
    analysis = build_analysis(p, tmp_path / "w", "s")
    assert sorted(_sites(analysis)) == ["A20", "F48"]
    a = _action(analysis)
    assert not a["skipped"], a["skipped"]
    out = apply_operations(p, tmp_path / "prep", a["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    done = Path(out["output_path"])
    check = compare_workbooks(p, done, Moves([a["operations"]]))
    assert check["clean"] is True, check
    assert _finding(build_analysis(done, tmp_path / "after", "a"))["status"] == "PASS"


# --- repair round (17 Nov 2026, refuter findings) -----------------------------------------------------------------
def _phoenix_with(tmp_path: Path, cell: str, formula: str) -> Path:
    p = _phoenix(tmp_path / "phoenix.xlsx")
    wb = openpyxl.load_workbook(p)
    sheet, ref = cell.split("!")
    wb[sheet][ref] = formula
    wb.save(p)
    return p


@pytest.mark.parametrize("cell, formula", [
    ("Calc!Z1", '=IFERROR(INDIRECT("CFs!B"&44),0)'),  # a label cell named by a text address
    ("Calc!Z1", "=SUM(CFs:Calc!Z5)"),                  # a 3-D reference reads the sheet by position
    ("Calc!Z1", '=COUNTIF(CFs!B:B,"")'),               # a whole column read by a function that does not look for text
    ("Calc!Z1", '=SUMPRODUCT(--(CFs!B:B=""))'),
])
def test_no_label_closes_a_gap_that_a_formula_could_see(tmp_path, cell, formula):
    a = _action(build_analysis(_phoenix_with(tmp_path, cell, formula), tmp_path / "w", "s"))
    assert not any(o["cell"] in ("B44", "B45", "B46") for o in a["operations"]), a["operations"]


def test_the_formula_table_never_serves_another_workbook(tmp_path, monkeypatch):
    """The server runs for days: an id() can be given to a new analysis once the old one is gone."""
    import app.prep as prep

    monkeypatch.setattr(prep, "id", lambda _o: 1, raising=False)  # every analysis gets the same id, as after a reuse
    first = {"workbooks": [{"formulas": [{"sheet": "S", "cell": "A1", "formula": "=1"}]}]}
    second = {"workbooks": [{"formulas": [{"sheet": "S", "cell": "A1", "formula": "=2"}]}]}
    assert prep._formula_record(first, "S!A1")["formula"] == "=1"
    assert prep._formula_record(second, "S!A1")["formula"] == "=2"


def test_a_bare_hash_beside_a_named_title_stays_a_title():
    cells = {"B2": "#", "C2": "#Rates"}
    for r in range(3, 6):
        cells.update({f"B{r}": r, f"C{r}": r * 10})
    grids, _ = detect_grids("S", _occupied(cells))
    assert any(g["title_cell"] == "B2" for g in grids), [(g["ref"], g["title_cell"]) for g in grids]
