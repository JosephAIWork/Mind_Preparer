"""2026-11-19: Model B stopped Start on "a formula reads a lone text cell" for 337 cells; 332 were one column heading,
"Per Capita Cost", repeated over the yearly tables of 'RAC Baseline' (rows 245, 362, 479, 596).

Eight formulas cover those tables with INDEX('RAC Baseline'!$X$237:$MI$705, 1+(active_year-proj_year-1)*117, MATCH(...)).
The row is always 1 + a whole number x 117 -- the 'Specialised Services' line of each year -- so INDEX never returns a
heading, and Mind leaving the headings empty changes no result. Such a cell is no longer listed.

The row number is followed only when it is plain whole-number arithmetic (+ - * and brackets over whole constants and
single cells, or names of one cell, that hold a typed-in whole number, or the whole-number values of a Mind loop).
Anything else, and the cell is listed as before.
"""
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from app.config import load_config
from app.excel_com import com_available
from app.inventory import build_analysis
from app.prep import _row_steps, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
HEADINGS = ["D7", "D13", "D19"]
STEPPED = "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2)"


def _years(path: Path, reader: str, yr: object = 2027, loop_values: tuple = (), second_yr: bool = False, yr_of_one_sheet: bool = False,
           pick: str | None = None) -> Path:
    """Base: three yearly blocks of 6 rows in C5:E22. Each block: the line the model reads (rows 5, 11, 17), an empty
    row, a heading alone in D (rows 7, 13, 19), an empty row, two rows of data. Inputs: the year in B1 (name yr) and
    the first year in B2 (name first_yr). Out!A1: the reader."""
    wb = openpyxl.Workbook()
    base = wb.active
    base.title = "Base"
    for block, top in enumerate((5, 11, 17)):
        for c, col in enumerate("CDE"):
            base[f"{col}{top}"] = 100 * (block + 1) + c
            base[f"{col}{top + 4}"] = 10 * (block + 1) + c
            base[f"{col}{top + 5}"] = 20 * (block + 1) + c
        base[f"D{top + 2}"] = "Per Capita Cost"
    inputs = wb.create_sheet("Inputs")
    inputs["A1"], inputs["B1"] = "Year", yr
    inputs["A2"], inputs["B2"] = "First year", 2026
    for i, v in enumerate(loop_values):
        inputs[f"D{i + 1}"] = v
    out = wb.create_sheet("Out")
    out["A1"] = reader
    out["B9"] = 2
    if yr_of_one_sheet:
        out.defined_names["yr"] = DefinedName("yr", attr_text="Inputs!$B$1")
    else:
        wb.defined_names["yr"] = DefinedName("yr", attr_text="Inputs!$B$1")
    wb.defined_names["first_yr"] = DefinedName("first_yr", attr_text="Inputs!$B$2")
    if second_yr:
        out.defined_names["yr"] = DefinedName("yr", attr_text="Inputs!$B$2")
    if pick:
        wb.defined_names["pick"] = DefinedName("pick", attr_text=pick)
    wb.save(path)
    wb.close()
    return path


def _analysis(tmp_path: Path, reader: str, **how):
    return build_analysis(_years(tmp_path / "years.xlsx", reader, **how), tmp_path / "w", "s")


def _sites(analysis) -> list[str]:
    return [s["cell"] for s in run_rule(RulesEngine().get("STR-008"), analysis, load_config())["observed"]["sites"]]


# --- the Model B shape -------------------------------------------------------------------------------------------
def test_a_heading_the_stepped_index_cannot_return_is_not_listed(tmp_path):
    analysis = _analysis(tmp_path, STEPPED)
    finding = run_rule(RulesEngine().get("STR-008"), analysis, load_config())
    assert finding["status"] == "PASS" and finding["observed"]["sites"] == []
    action = {a["id"]: a for a in plan_actions(analysis, {"findings": []})}["title_lone_text_cells"]
    assert action["operations"] == [] and not action.get("skipped")
    # the same workbook with a row number that reaches the headings: the rule and the Prep action do see them
    other = build_analysis(_years(tmp_path / "reached.xlsx", "=INDEX(Base!$C$5:$E$22,3+(yr-first_yr-1)*6,2)"), tmp_path / "w2", "s2")
    assert _sites(other) == HEADINGS


@pytest.mark.parametrize("reader", [
    "=INDEX(Base!$C$5:$E$22,1+(Inputs!$B$1-Inputs!B2-1)*6,2)",  # the cells themselves in place of their names
    "=INDEX(Base!$C$5:$E$22,1+(B9-1)*6,2)",  # a cell of the formula's own sheet
    "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,MATCH(101,Base!$C$5:$E$5,0))",  # the column worked out, as in the model
    "=INDEX(Base!$C$5:$E$22, 1 + ( yr - first_yr - 1 ) * 6 , 2)",
    "=IF(yr=first_yr,0,INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2))+INDEX(Base!$D$5:$D$22,1+(yr-first_yr)*6,1)",
])
def test_the_same_through_cells_spaces_a_worked_out_column_and_two_calls(tmp_path, reader):
    assert _sites(_analysis(tmp_path, reader)) == []


def test_a_fixed_row_reads_that_row_only(tmp_path):
    assert _sites(_analysis(tmp_path, "=INDEX(Base!$C$5:$E$22,9,2)")) == ["D13"]


# --- where the row can reach a heading, the cell is still listed ---------------------------------------------
@pytest.mark.parametrize("reader", [
    "=INDEX(Base!$C$5:$E$22,3+(yr-first_yr-1)*6,2)",  # rows 3, 9, 15 of the range: the headings themselves
    "=INDEX(Base!$C$5:$E$22,(yr-first_yr)*6,2)",  # the row can be 0: INDEX then returns the whole column
    "=INDEX(Base!$C$5:$E$22,yr-first_yr,2)",  # any row
    "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr)/1*6,2)",  # a division: no longer known to be whole
    "=INDEX(Base!$C$5:$E$22,1+ROUND(yr-first_yr,0)*6,2)",  # a function: not followed
    "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6.5,2)",
    "=INDEX(Base!$C$5:$E$22,1+(yr-nobody-1)*6,2)",  # a name the workbook does not have
    "=INDEX(Base!$C$5:$E$22,1+(Inputs!B1:B2-1)*6,2)",  # a range in the row number
    "=INDEX(Base!$C$5:$E$22,,2)",
])
def test_a_row_that_can_reach_a_heading_or_cannot_be_followed_still_lists_them(tmp_path, reader):
    assert _sites(_analysis(tmp_path, reader)) == HEADINGS


@pytest.mark.parametrize("reader", [
    "=OFFSET(INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2),2,0)",  # moved two rows down: the heading itself
    "=COUNTA(INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2):Base!$D$22)",  # stretched into a range
    "=COUNTA(Base!$D$5:INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2))",
    "=COUNTA(IF(yr>0,INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2),Base!C5):Base!$D$22)",  # stretched through IF
    # 19 Nov, the refuter: stretched through a name LET gives it (as typed, and as Excel stores it)
    "=LET(x,INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2),COUNTA(x:Base!$D$22))",
    "=_xlfn.LET(_xlpm.x,INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2),COUNTA(Base!$D$5:_xlpm.x))",
    "=LET(x,INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2),x)",  # LET at all: not followed
    "=MyShift(INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2))",  # a macro's function, or a LAMBDA name
    "=_xludf.MyShift(INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2))",
    "=COUNTA(_xlfn.ANCHORARRAY(INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2)))",  # INDEX(...)#
    "=SUM(MM_COLUMN(INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2)))",  # Mind reads the column on from there
    "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2)+COUNTA(yr:first_yr)",  # a ':' between two names, anywhere
])
def test_a_reference_that_is_moved_or_stretched_after_index_still_lists_them(tmp_path, reader):
    assert _sites(_analysis(tmp_path, reader)) == HEADINGS


def test_a_range_stretched_from_a_name_that_holds_the_index_still_lists_them(tmp_path):
    reader, pick = "=COUNTA(pick:Base!$D$22)", "INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2)"
    assert _sites(_analysis(tmp_path, reader, pick=pick)) == HEADINGS


def test_an_index_along_one_row_still_lists_the_cell(tmp_path):
    # INDEX(one row, n): n runs along the row, so a "row number" of 3 is the third cell
    assert _sites(_analysis(tmp_path, "=INDEX(Base!$C$7:$G$7,1+(yr-first_yr)*2)")) == ["D7"]


# --- the year must be a whole number today ------------------------------------------------------------------
@pytest.mark.parametrize("yr", [2027.5, "=first_yr+1", None, "2027", True])
def test_a_year_that_is_not_a_typed_in_whole_number_still_lists_them(tmp_path, yr):
    assert _sites(_analysis(tmp_path, STEPPED, yr=yr)) == HEADINGS


def test_a_whole_number_stored_as_a_float_counts(tmp_path):
    assert _sites(_analysis(tmp_path, STEPPED, yr=2027.0)) == []


def test_two_names_of_the_same_spelling_are_not_followed(tmp_path):
    assert _sites(_analysis(tmp_path, STEPPED, second_yr=True)) == HEADINGS


def test_a_name_of_one_sheet_only_is_not_followed(tmp_path):
    assert _sites(_analysis(tmp_path, STEPPED, yr_of_one_sheet=True)) == HEADINGS


# --- the year as a Mind loop (what Build the Mind loop writes in the driver cell) ------------------------------
def test_a_mind_loop_over_typed_in_whole_numbers_counts(tmp_path):
    assert _sites(_analysis(tmp_path, STEPPED, yr='=MM_LOOP("years",Inputs!$D$1:$D$3)', loop_values=(2027, 2028, 2029))) == []


@pytest.mark.parametrize("yr, values", [
    ('=MM_LOOP("years",Inputs!$D$1:$D$3)', (2027, 2028.5, 2029)),
    ('=MM_LOOP("years",Inputs!$D$1:$D$3)', (2027, None, 2029)),
    ('=MM_LOOP("years",Inputs!$D$1:$D$3)', (2027, "=D1+1", 2029)),
    ('=MM_LOOP("years",Inputs!$D$1:$D$3)+0.5', (2027, 2028, 2029)),
    ('=MM_LOOP("years",MM_COLUMN(Inputs!$D$1))', (2027, 2028, 2029)),
    ('=MM_LOOP("years",Inputs!$D$1:Inputs!$D$3)', (2027, 2028, 2029)),
    ('=MM_LOOPINSTANCE("years",Inputs!$D$1,Inputs!$D$2)', (3, 3, 3)),
    ('=MM_LOOP("years")', (2027, 2028, 2029)),
    ('=MM_LOOP("years",Inputs!$D$1:$D$3,2)', (2027, 2028, 2029)),  # a third argument: what it does is not known here
])
def test_a_mind_loop_whose_values_are_not_known_to_be_whole_still_lists_them(tmp_path, yr, values):
    assert _sites(_analysis(tmp_path, STEPPED, yr=yr, loop_values=values)) == HEADINGS


# --- the arithmetic ------------------------------------------------------------------------------------------
@pytest.mark.parametrize("text, steps", [
    ("1+(yr-first_yr-1)*117", (1, 117)),
    ("1+(yr-first_yr)*117", (1, 117)),
    ("3+6*yr", (3, 6)),
    ("2*(yr+1)+4*first_yr", (2, 2)),
    ("-(yr)*5+2", (2, 5)),
    ("(yr*2+1)*(first_yr*2+1)", (1, 2)),  # odd times odd is odd
    ("yr*first_yr", (0, 1)),
    ("7", (7, 0)),
    ("7.0", (7, 0)),
    ("9007199254740993", (9007199254740993, 0)),  # past what a float holds exactly
    ("--7", (7, 0)),
    ("(1+2)*3", (9, 0)),
    ("yr*0+4", (4, 0)),
])
def test_the_rows_a_row_number_can_take(tmp_path, text, steps):
    first, every = _row_steps(_analysis(tmp_path, STEPPED), text, "Out")
    assert every == steps[1] and (first == steps[0] if every == 0 else first % every == steps[0] % every)


@pytest.mark.parametrize("text", ["1.5", "yr/2", "yr^2", "1+", "(yr", "yr)", "", "yr yr", "yr%", '1+"2"', "1e3", "TRUE", "Base!C5:C6", "yr&1", "A:A", "7.50", "Out!yr", "yr#", "@yr"])
def test_a_row_number_that_is_not_plain_whole_arithmetic_is_not_followed(tmp_path, text):
    assert _row_steps(_analysis(tmp_path, STEPPED), text, "Out") is None


# --- in real Excel: the headings change nothing ---------------------------------------------------------------
@needs_excel
def test_in_excel_the_reader_gives_the_same_with_the_headings_empty(tmp_path):
    from app.excel_com import close_quietly, excel_session, open_for_write

    p = _years(tmp_path / "years.xlsx", STEPPED)

    def results(empty_the_headings: bool) -> list:
        got = []
        with excel_session(mm_addin=False) as excel:
            book = open_for_write(excel, p)
            try:
                if empty_the_headings:
                    for cell in HEADINGS:
                        book.Worksheets("Base").Range(cell).ClearContents()
                for year in (2027, 2028, 2029):
                    book.Worksheets("Inputs").Range("B1").Value = year
                    excel.Calculate()
                    got.append(book.Worksheets("Out").Range("A1").Value)
                book.Saved = True
            finally:
                close_quietly(book)
        return got

    assert results(False) == results(True) == [101.0, 201.0, 301.0]
