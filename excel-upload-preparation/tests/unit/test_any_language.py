"""Sheet and name text in any language (the Tokyo IFRS 17 tool, 2026-11-07): Japanese sheet names were
read as no sheet at all (`=ロールフォワード!B10` in 仕訳!B10 looked like =B10, a cell reading itself),
so the scan reported dozens of false circular references; CELL("filename", C6) in C6 counted as one too;
and text typed as '=... was read as a formula."""
import openpyxl

from app.formula_utils import REF_RE, cell_refs_in_formula
from app.inventory import build_analysis
from app.validators.references import circular_references

SHEETS = ["ロールフォワード", "גיליון", "Été", "Лист", "表1"]


def test_unquoted_sheet_names_in_any_language_are_read_as_sheets():
    for name in SHEETS:
        refs = [m.group(0) for m in REF_RE.finditer(f"={name}!B10+C2")]
        assert refs == [f"{name}!B10", "C2"], name
    assert [m.group(0) for m in REF_RE.finditer("=合計A1")] == []  # a defined name, not the cell A1


def _book(tmp_path, build):
    wb = openpyxl.Workbook()
    build(wb)
    p = tmp_path / "book.xlsx"
    wb.save(p)
    return build_analysis(p, tmp_path / "work", "t")


def test_a_reference_to_the_same_cell_on_a_japanese_sheet_is_not_circular(tmp_path):
    def build(wb):
        a = wb.active
        a.title = "仕訳"
        b = wb.create_sheet("ロールフォワード")
        b["B10"] = 5
        a["B10"] = "=ロールフォワード!B10"
        a["B11"] = "=B10*2"

    res = circular_references(None, _book(tmp_path, build), {})
    assert res["status"] == "PASS", res["message"]


def test_cell_filename_of_its_own_cell_is_not_circular(tmp_path):
    def build(wb):
        a = wb.active
        a.title = "Ctrl"
        a["D1"] = "data"
        a["C6"] = '=IFERROR(LEFT(CELL("filename",C6),FIND("\\"&$D$1,CELL("filename",C6))-1),"")'

    res = circular_references(None, _book(tmp_path, build), {})
    assert res["status"] == "PASS", res["message"]


def test_text_typed_with_a_leading_equals_sign_is_text_not_a_formula(tmp_path):
    def build(wb):
        a = wb.active
        a.title = "仕訳"
        a["B2"] = "=B2"  # a real self reference
        a["B3"].value = "=B3"
        a["B3"].data_type = "s"  # typed as '=B3: text

    analysis = _book(tmp_path, build)
    cells = {f["cell"] for f in analysis["workbooks"][0]["formulas"]}
    assert "B2" in cells and "B3" not in cells
    assert cell_refs_in_formula("=表1!A1")[0]["sheet"] == "表1"
