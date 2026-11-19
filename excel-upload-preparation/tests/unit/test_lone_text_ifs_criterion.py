"""2026-11-19: CNHI's last lone text cell, 'Cost data'!AF10 = "Deliveries and inpatient services".

Twenty formulas read it as the criterion of SUMIFS -- SUMIFS($H:$H,$C:$C,$AF10,...) -- and no title could go above it
(AF32 is =AF9). Prep already knows how to put the text itself into a formula in place of the reference, but every
argument of SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS and MINIFS counted as "must stay a reference". Only their ranges
must: a criterion is a value, and the cell's own text does the same there.
"""
from pathlib import Path

import openpyxl
import pytest

from app.config import load_config
from app.excel_com import com_available
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
ACTION = "title_lone_text_cells"
ROWS = [("north", 10), ("south", 20), ("north", 5), ("east", 1), ("n*", 7), (">3", 4), ("5", 3), ('say "hi"', 2)]
AMOUNTS, KEYS = "$B$2:$B$9", "$A$2:$A$9"


def _book(path: Path, readers: dict[str, str], text: str = "north") -> Path:
    """Keys in A2:A9 and amounts in B2:B9 under their headers. The text alone in D5, far from the table. D4, the cell
    above it, is read by H9, so no title can go there. `readers`: the formulas that read D5."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["B1"] = "Region", "Amount"
    for i, (key, amount) in enumerate(ROWS):
        ws[f"A{i + 2}"], ws[f"B{i + 2}"] = key, amount
    ws["D5"] = text
    ws["H9"] = "=D4*2"
    for cell, formula in readers.items():
        ws[cell] = formula
    wb.save(path)
    wb.close()
    return path


def _finding(analysis):
    return run_rule(RulesEngine().get("STR-008"), analysis, load_config())


def _action(analysis):
    return {a["id"]: a for a in plan_actions(analysis, {"findings": []})}[ACTION]


def _rewrites(action) -> dict[str, str]:
    return {o["cell"]: o["after"] for o in action["operations"] if o["op"] == "set_formula"}


@pytest.mark.parametrize("formula, after", [
    (f"=SUMIFS({AMOUNTS},{KEYS},D5)", f'=SUMIFS({AMOUNTS},{KEYS},"north")'),
    (f"=SUMIFS({AMOUNTS},{KEYS},$D$5)", f'=SUMIFS({AMOUNTS},{KEYS},"north")'),
    (f'=SUMIFS({AMOUNTS},{KEYS},"<>south",{KEYS},D5)', f'=SUMIFS({AMOUNTS},{KEYS},"<>south",{KEYS},"north")'),  # the second criterion
    (f'=SUMIFS({AMOUNTS},{KEYS},"<>"&D5)', f'=SUMIFS({AMOUNTS},{KEYS},"<>"&"north")'),  # inside the criterion
    (f"=COUNTIFS({KEYS},D5)", f'=COUNTIFS({KEYS},"north")'),
    (f'=COUNTIFS({KEYS},"<>south",{KEYS},D5)', f'=COUNTIFS({KEYS},"<>south",{KEYS},"north")'),
    (f"=AVERAGEIFS({AMOUNTS},{KEYS},D5)", f'=AVERAGEIFS({AMOUNTS},{KEYS},"north")'),
    (f"=MAXIFS({AMOUNTS},{KEYS},D5)", f'=MAXIFS({AMOUNTS},{KEYS},"north")'),
    (f"=MINIFS({AMOUNTS},{KEYS},D5)", f'=MINIFS({AMOUNTS},{KEYS},"north")'),
    (f"=(SUMIFS({AMOUNTS},{KEYS},$D5)+IF(B2>0,SUMIFS({AMOUNTS},{KEYS},$D5),0))*2",
     f'=(SUMIFS({AMOUNTS},{KEYS},"north")+IF(B2>0,SUMIFS({AMOUNTS},{KEYS},"north"),0))*2'),  # CNHI's own shape
])
def test_the_text_goes_into_a_criterion(tmp_path, formula, after):
    analysis = build_analysis(_book(tmp_path / "s.xlsx", {"H1": formula}), tmp_path / "w", "s")
    assert [s["cell"] for s in _finding(analysis)["observed"]["sites"]] == ["D5"]
    action = _action(analysis)
    assert _rewrites(action) == {"H1": after} and not action["skipped"], action["skipped"]
    out = apply_operations(tmp_path / "s.xlsx", tmp_path / "prep", action["operations"], prefer_excel=False)
    assert _finding(build_analysis(Path(out["output_path"]), tmp_path / "after", "a"))["status"] == "PASS"


@pytest.mark.parametrize("formula", [
    f"=SUMIFS(D5,$A$5,\"north\")",  # the values to add: a reference
    f"=SUMIFS($B$5,D5,\"north\")",  # a criterion's range: a reference
    f'=SUMIFS({AMOUNTS},{KEYS},"north",D5,"north")',  # the second criterion's range
    f"=COUNTIFS(D5,\"north\")",
    f'=COUNTIFS({KEYS},"north",D5,"north")',
    f"=AVERAGEIFS(D5,$A$5,\"north\")",
    f"=MAXIFS($B$5,D5,\"north\")",
    # 19 Nov, the refuter: Excel refuses a text beside the intersection operator (a space) or as one end of a range
    f"=SUMIFS({AMOUNTS},{KEYS},D5 $D:$D)",
    f"=SUMIFS({AMOUNTS},{KEYS},$D:$D D5)",
    f"=SUMIFS({AMOUNTS},{KEYS},(D5):D6)",
    f"=SUMIFS({AMOUNTS},{KEYS},D3:(D5))",
    f"=SUMIFS({AMOUNTS},{KEYS},(D5) $D:$D)",
])
def test_no_text_goes_where_an_ifs_function_wants_a_range(tmp_path, formula):
    action = _action(build_analysis(_book(tmp_path / "s.xlsx", {"H1": formula}), tmp_path / "w", "s"))
    assert _rewrites(action) == {}


@pytest.mark.parametrize("formula, after", [
    (f"=SUMIFS({AMOUNTS}, {KEYS}, D5)", f'=SUMIFS({AMOUNTS}, {KEYS}, "north")'),  # a space after a comma is only a space
    (f"=SUMIFS({AMOUNTS},{KEYS},D5 )", f'=SUMIFS({AMOUNTS},{KEYS},"north" )'),
    (f"=SUMIFS({AMOUNTS},{KEYS},(D5))", f'=SUMIFS({AMOUNTS},{KEYS},("north"))'),
    ("=IF(D5 = \"north\", 1, 2)", '=IF("north" = "north", 1, 2)'),
])
def test_a_space_or_a_bracket_that_is_no_operator_does_not_stop_the_text(tmp_path, formula, after):
    action = _action(build_analysis(_book(tmp_path / "s.xlsx", {"H1": formula}), tmp_path / "w", "s"))
    assert _rewrites(action) == {"H1": after} and not action["skipped"], action["skipped"]


def test_a_text_too_long_for_a_formula_as_excel_counts_it_is_not_written(tmp_path):
    long_text = "😀" * 130  # 130 characters, 260 as Excel counts them
    action = _action(build_analysis(_book(tmp_path / "s.xlsx", {"H1": f"=SUMIFS({AMOUNTS},{KEYS},D5)"}, long_text), tmp_path / "w", "s"))
    assert _rewrites(action) == {}


def test_every_reader_of_the_cell_gets_the_text_in_one_pass(tmp_path):
    readers = {f"{col}1": f"=SUMIFS({AMOUNTS},{KEYS},$D5)*{i + 1}" for i, col in enumerate("HIJKL")}
    readers["H3"] = "=D5"  # CNHI's AF33: the cell read as it is
    action = _action(build_analysis(_book(tmp_path / "s.xlsx", readers), tmp_path / "w", "s"))
    assert _rewrites(action) == {**{f"{col}1": f'=SUMIFS({AMOUNTS},{KEYS},"north")*{i + 1}' for i, col in enumerate("HIJKL")}, "H3": '="north"'}
    assert not action["skipped"]


def test_one_reader_that_wants_a_range_keeps_the_cell_by_hand(tmp_path):
    readers = {"H1": f"=SUMIFS({AMOUNTS},{KEYS},D5)", "I1": "=COUNTIFS(D5,\"north\")"}
    action = _action(build_analysis(_book(tmp_path / "s.xlsx", readers), tmp_path / "w", "s"))
    assert _rewrites(action) == {}  # all its readers, or none: one left reading the cell and Mind still loses it


@needs_excel
@pytest.mark.parametrize("text", ["north", "n*", ">3", "5", 'say "hi"'])
def test_in_excel_the_text_in_the_criterion_gives_the_same_values(tmp_path, text):
    from app.excel_com import close_quietly, excel_session, open_for_write, save_in_place

    readers = {
        "H1": f"=SUMIFS({AMOUNTS},{KEYS},D5)", "I1": f"=COUNTIFS({KEYS},D5)", "J1": f"=IFERROR(AVERAGEIFS({AMOUNTS},{KEYS},D5),-1)",
        "K1": f"=MAXIFS({AMOUNTS},{KEYS},D5)", "L1": f"=MINIFS({AMOUNTS},{KEYS},D5)", "M1": f'=SUMIFS({AMOUNTS},{KEYS},"<>"&D5)', "H3": "=D5",
    }
    p = _book(tmp_path / "s.xlsx", readers, text)
    with excel_session(mm_addin=False) as excel:
        book = open_for_write(excel, p)
        try:
            excel.CalculateFullRebuild()
            save_in_place(book, p)
        finally:
            close_quietly(book)
    action = _action(build_analysis(p, tmp_path / "w", "s"))
    assert set(_rewrites(action)) == set(readers) and not action["skipped"], action["skipped"]
    out = apply_operations(p, tmp_path / "prep", action["operations"])
    assert out["status"] == "APPLIED" and not out["failed"]
    done = Path(out["output_path"])
    check = compare_workbooks(p, done, Moves([action["operations"]]))
    assert check["clean"] is True and check["counts"]["compared"] >= len(readers), check
    assert _finding(build_analysis(done, tmp_path / "after", "a"))["status"] == "PASS"
