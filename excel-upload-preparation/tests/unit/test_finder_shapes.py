"""Formula shapes and repetition patterns for the instance finder (app.finder_shapes,
app.finder_patterns): the A1 -> R1C1 normaliser, loose keys, the shape book, bands and fills,
sheet copies, how copies differ and the edits that unify them, chains, and what a repetition steps
through (relative references and computed indices).

Every workbook here is invented with openpyxl (SYNTHETIC_TEST_FIXTURE in A1 of the first sheet);
most are read as in-memory analysis dicts shaped like app.inventory.build_analysis's output, one
goes through build_analysis itself. No Excel, no numpy, nothing from a client workbook."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

from app import finder_patterns as fp
from app.autofix import _EP, _REF_RE, _STRING_RE, r1c1_to_a1
from app.finder_shapes import A1_REF_RE, Budget, a1_to_r1c1, build_book, loose_keys, number_literals
from app.inventory import build_analysis
from app.model_map import _Names

LABEL = "SYNTHETIC_TEST_FIXTURE"


def _analysis_of(wb: openpyxl.Workbook) -> dict:
    """An in-memory analysis of an openpyxl workbook, shaped like build_analysis's output."""
    formulas, sheets = [], []
    for ws in wb.worksheets:
        n = 0
        for (_r, _c), cell in sorted(ws._cells.items()):
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                formulas.append({"sheet": ws.title, "cell": cell.coordinate, "formula": v})
                n += 1
            elif isinstance(v, ArrayFormula):
                formulas.append({"sheet": ws.title, "cell": cell.coordinate, "formula": v.text, "array_ref": v.ref})
                n += 1
        sheets.append({"name": ws.title, "state": "visible", "dimensions": ws.dimensions, "formula_count": n})
    names = [{"name": n, "value": d.attr_text, "local_sheet_id": None} for n, d in wb.defined_names.items()]
    return {"workbooks": [{"formulas": formulas, "sheets": sheets, "defined_names": names, "copy_path": None, "file_name": "invented.xlsx"}]}


def _book_of(wb: openpyxl.Workbook):
    analysis = _analysis_of(wb)
    return build_book(analysis, Budget(120)), analysis


def _new() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.active.title = "Inputs"
    wb.active["A1"] = LABEL
    return wb


def _scen_blocks(ws, rows: int = 10) -> None:
    """Three scenario blocks of 4 columns (year, rate read from Inputs!B$2/C$2/D$2, fund, a
    constant), one empty column between them: 20 formulas per block."""
    for k in range(3):
        c0 = 1 + k * 5
        a, b, c, d = (L(c0 + i) for i in range(4))
        for t in range(1, rows + 1):
            r = t + 2
            ws[f"{a}{r}"] = t
            ws[f"{b}{r}"] = f"=Inputs!{L(2 + k)}$2"
            ws[f"{c}{r}"] = f"={b}{r}*100"
            ws[f"{d}{r}"] = 5


# --- the normaliser --------------------------------------------------------------------------------
_EP_RE = re.compile(rf"^(?:{_EP})(?::(?:{_EP}))?$")

INVERSE_CASES = [
    ("=B9*$B$3+SUM(F:F)+$5:$5", 10, 7, "=R[-1]C[-5]*R3C2+SUM(C[-1])+R5"),
    ("=A1:A1", 3, 3, "=R[-2]C[-2]:R[-2]C[-2]"),
    ("=$A:$A+1:1", 5, 5, "=C1+R[-4]"),
    ("=A$1+$A1", 2, 2, "=R1C[-1]+R[-1]C1"),
    ("=SUM(A:C)+SUM(2:4)+Data!5:5", 9, 9, "=SUM(C[-8]:C[-6])+SUM(R[-7]:R[-5])+Data!R[-4]"),
    ("=SUM(Data!B2:D4)*'My Sheet'!$C$3", 7, 2, None),
    ("=IF(A1>0,\"A1 \"&B1,C1)", 1, 1, "=IF(RC>0,\"A1 \"&RC[1],RC[2])"),
    ("=LOG10(A1)+ATAN2(B1,C1)+DAYS360(D1,E1)", 1, 1, None),
    ("=Rate2*Q1Sales+1E5+2.5E-3", 1, 1, "=Rate2*Q1Sales+1E5+2.5E-3"),
    ("=OFFSET(Pick, COLUMN()-3, (ROW()-1)*2)", 2, 3, None),
    ("=VLOOKUP($A5,Rates!$A:$D,3,FALSE)", 5, 4, None),
    ("=INDEX($B$5:$AI$27,ROW_SELECT,COUNTA($B$1:B1))", 2, 2, "=INDEX(R5C2:R27C35,ROW_SELECT,COUNTA(R1C2:R[-1]C))"),
    ("=SUMPRODUCT((A2:A100=\"x\")*(B2:B100))", 1, 5, None),
    ("='Sheet A1'!B2+'It''s'!C3", 1, 1, None),
    ("=A1+ A2 + A3", 4, 1, "=R[-3]C+ R[-2]C + R[-1]C"),
]


def test_a1_to_r1c1_is_the_inverse_of_r1c1_to_a1():
    cells = [(1, 1), (10, 7), (100, 26), (3, 3)]
    n = 0
    for formula, row, col, expected in INVERSE_CASES:
        rc = a1_to_r1c1(formula, row, col)
        if expected is not None:
            assert rc == expected, (formula, rc)
        for r, c in [(row, col)] + cells:
            rc = a1_to_r1c1(formula, r, c)
            assert r1c1_to_a1(rc, r, c) == formula, (formula, r, c, rc)
            masked = _STRING_RE.sub('""', rc)
            refs = list(_REF_RE.finditer(masked))
            a1_refs = list(A1_REF_RE.finditer(_STRING_RE.sub('""', formula)))
            assert len(refs) == len(a1_refs), (formula, rc)
            for m in refs:  # every converted reference is R1C1 as autofix's grammar reads it
                assert _EP_RE.match(m.group("a") + (":" + m.group("b") if m.group("b") else "")), m.group(0)
            n += 1
    assert n >= 30


def test_normaliser_keeps_strings_names_sheets_externals_and_tables():
    f = "=\"A1 \"&'My Sheet'!$B$2&[1]Sheet1!A1&Jan:Dec!A1&Rate&Table1[Col]"
    rc = a1_to_r1c1(f, 5, 5)
    assert rc == "=\"A1 \"&'My Sheet'!R2C2&[1]Sheet1!R[-4]C[-4]&Jan:Dec!R[-4]C[-4]&Rate&Table1[Col]"
    assert r1c1_to_a1(rc, 5, 5) == f
    for t in ("=Table1[Q1]*2", "=Table1[[#This Row],[Q1]]", "=SUM(DeptSales[[#Headers],[Region]:[Q1]])"):
        assert a1_to_r1c1(t, 2, 2) == t
        assert a1_to_r1c1(t, 2, 2) == a1_to_r1c1(t, 9, 4)  # the same shape id on any row
    assert a1_to_r1c1("=Table1[[#This Row],[Q1]]+B2", 6, 6) == "=Table1[[#This Row],[Q1]]+R[-4]C[-4]"
    assert a1_to_r1c1("=#REF!+NA()+{1,2;3}+TRUE", 1, 1) == "=#REF!+NA()+{1,2;3}+TRUE"


def test_normaliser_leaves_a_name_that_looks_like_a_cell_past_excels_last_row_or_column():
    """ABC1234567 or XFE1 is no cell (Excel stops at XFD1048576), so Excel allows it as a defined name: the
    normaliser copies it as it is, and a formula using it keeps one shape wherever it is filled down."""
    assert a1_to_r1c1("=ABC1234567+A1+XFE1+XFD1048576", 10, 7) == "=ABC1234567+R[-9]C[-6]+XFE1+R[1048566]C[16377]"
    assert a1_to_r1c1("=SUM($A:$ZZZ)+SUM(1:9999999)+SUM(A:B)", 2, 2) == "=SUM($A:$ZZZ)+SUM(1:9999999)+SUM(C[-1]:C)"
    assert a1_to_r1c1("=ABC1234567*2", 3, 3) == a1_to_r1c1("=ABC1234567*2", 9, 9) == "=ABC1234567*2"


def test_normaliser_strips_storage_prefixes_and_implicit_intersection():
    rc = a1_to_r1c1("=_xlfn.XLOOKUP(@C3,Data!C:C,Data!E:E)", 3, 5)
    assert rc == "=XLOOKUP(RC[-2],Data!C[-2],Data!C)"
    assert a1_to_r1c1("=_xll.MM_LOOP(\"Scenario\",A2:A6)", 1, 2) == "=MM_LOOP(\"Scenario\",R[1]C[-1]:R[5]C[-1])"
    # the analysis keeps the prefix, COM's FormulaR1C1 has none: both give the same shape
    assert a1_to_r1c1("=_xlfn.IFS(A1>0,1,TRUE,0)", 1, 2) == a1_to_r1c1("=IFS(A1>0,1,TRUE,0)", 1, 2)


def test_loose_keys_mask_refs_then_sheets_then_numbers():
    k1, k2, k3 = loose_keys("=Data!R[-1]C*0.12+SUM(C[-1])+1")
    assert (k1, k2, k3) == ("=Data!@*0.12+SUM(@)+1", "=@*0.12+SUM(@)+1", "=@*#+SUM(@)+#")
    # the two copies of the use-case's kind: same shape but for the input read / the typed number
    assert loose_keys("=Inputs!R2C[-1]*RC[-2]")[1] == loose_keys("=Inputs!R2C[-6]*RC[-2]")[1]
    assert loose_keys("=R[-1]C*0.1")[2] == loose_keys("=R[-1]C*0.12")[2]
    assert loose_keys("=R[-1]C*0.1")[1] != loose_keys("=R[-1]C*0.12")[1]
    assert loose_keys('=IF(RC[-1]="2024",1,0)')[2] == '=IF(@="2024",#,#)'  # the string literal stays
    assert number_literals("=R[-1]C*0.12+1") == ["0.12", "1"]
    assert number_literals('=R3C2*1E5+"12"&LOG10(R[-1]C)') == ["1E5"]


def test_shape_book_gives_one_id_to_a_filled_down_formula_and_to_array_members(tmp_path: Path):
    wb = _new()
    ws = wb.active
    for r in range(2, 52):
        ws[f"A{r}"] = r
        ws[f"C{r}"] = f"=A{r}*2"
    ws["B1"] = ArrayFormula("B1:B3", "=A2:A4*3")
    ws["D1"] = "text"
    ws["D2"] = 42
    path = tmp_path / "shapes.xlsx"
    wb.save(path)
    wb.close()
    analysis = build_analysis(path, tmp_path / "w", path.stem)
    book = build_book(analysis, Budget(120))
    sh = book.sheets["Inputs"]
    ids = {sh.cells[(r, 3)] for r in range(2, 52)}
    assert len(ids) == 1 and min(ids) >= 10
    assert len({sh.cells[(r, 2)] for r in (1, 2, 3)}) == 1  # the array members share the anchor's id
    assert sh.cells[(1, 2)] != sh.cells[(2, 3)]
    assert (1, 4) not in sh.cells and (2, 4) not in sh.cells  # constants are not stored
    assert sh.n == 53 and (sh.r_lo, sh.r_hi, sh.c_lo, sh.c_hi) == (1, 51, 2, 3)
    assert sh.cols[3][10] == sh.cells[(10, 3)] and sh.rows[10][3] == sh.cells[(10, 3)]
    assert book.a1_text("Inputs", 10, 3) == "=A10*2" and book.a1_text("Inputs", 3, 2) == "=A2:A4*3"
    assert book.texts[sh.cells[(10, 3)]] == "=RC[-2]*2"
    assert len(book.loose) == 3 and book.loose[1][sh.cells[(10, 3)]] >= 10
    assert book.workbook() is not None  # the analysed copy, through the app's cache


# --- bands and fills -------------------------------------------------------------------------------
def test_bands_find_three_blocks_and_skip_a_time_axis():
    wb = _new()
    wb.active["B2"], wb.active["C2"], wb.active["D2"] = 0.02, 0.03, 0.01
    sc = wb.create_sheet("Scen")
    _scen_blocks(sc)
    sc["Q20"] = 1000  # a 60-column recurrent projection: each month reads the one before
    for m in range(1, 60):
        sc[f"{L(17 + m)}20"] = f"={L(16 + m)}20*0.98"
    book, _ = _book_of(wb)
    found = fp.sheet_findings(book, book.sheets["Scen"], Budget(120))
    blocks = [f for f in found if not f["fill"]]
    fills = [f for f in found if f["fill"]]
    assert len(blocks) == 1 and len(fills) == 1
    b = blocks[0]
    # positions start at the first FORMULA column (B: column A holds the years), period 5
    assert (b["axis"], b["where"], b["copies"], b["width"], b["per_copy"]) == ("columns", "B:P", 3, 5, 20)
    assert b["positions"][1] == [7, 8, 9, 10, 11] and len(b["cells"][2]) == 20
    assert b["chain_share"] == 0.0 and b["diff"]["reads a different input"] == 10 and b["diff"]["identical"] == 10
    assert b["inner_axis"]["axis"] == "rows" and b["inner_axis"]["copies"] == 10  # the years inside each block
    assert b["first_cell"] == (3, 2) and b["first_formula_index"] == 0
    f = fills[0]
    assert (f["axis"], f["copies"], f["chain_share"], f["time_axis"]) == ("columns", 59, 1.0, True)
    assert f["where"] == "R:BX" and f["axes"] is None


def test_a_uniform_area_is_one_fill_with_two_axes():
    wb = _new()
    u = wb.create_sheet("Out")
    for r in range(22, 48):
        for c in range(14, 142):
            u.cell(r, c, f"=$A{r}*{L(c)}$1")
    book, _ = _book_of(wb)
    found = fp.sheet_findings(book, book.sheets["Out"], Budget(120))
    assert len(found) == 1
    f = found[0]
    assert f["fill"] is True and f["axes"] == {"rows": 26, "columns": 128}
    assert f["n_cells"] == 26 * 128 and f["diff"]["identical"] == f["per_copy"]
    assert set(f["chain_shares"]) == {"rows", "columns"} and f["time_axis"] is False
    assert f["n_cells"] >= fp.BIG_FILL  # handed to the merge step, never a dimension on its own


def test_chain_test_ignores_absolute_reads_of_a_parameter_in_the_first_block():
    wb = _new()
    ws = wb.create_sheet("Blocks")
    ws["B2"] = 0.05  # a parameter inside block 1's columns, read by every block
    for k in range(3):
        c0 = 1 + k * 4
        a, b = L(c0), L(c0 + 1)
        for r in range(4, 16):
            ws[f"{a}{r}"] = f"=$B$2*ROW()"
            ws[f"{b}{r}"] = f"=$B$2+{a}{r}"
    book, _ = _book_of(wb)
    found = [f for f in fp.sheet_findings(book, book.sheets["Blocks"], Budget(120)) if not f["fill"]]
    assert len(found) == 1 and found[0]["copies"] == 3
    assert found[0]["chain_share"] == 0.0 and fp.chain_share(book, found[0]) == 0.0
    # the real thing: each block reads the block before it
    wb2 = _new()
    ws2 = wb2.create_sheet("Chain")
    ws2["B2"] = 1
    for k in range(3):
        c0 = 1 + k * 4
        a, b = L(c0), L(c0 + 1)
        for r in range(4, 16):
            ws2[f"{a}{r}"] = "=$B$2*1.1" if k == 0 else f"={L(c0 - 4)}{r}*1.1"  # blocks 2 and 3 read the block before
            ws2[f"{b}{r}"] = f"={a}{r}+1"
    book2, _ = _book_of(wb2)
    chain = [f for f in fp.sheet_findings(book2, book2.sheets["Chain"], Budget(120)) if not f["fill"]][0]
    assert chain["chain_share"] >= fp.CHAIN_SHARE


# --- sheet copies ----------------------------------------------------------------------------------
def _prod_sheets(wb, names=("Prod_A", "Prod_B", "Prod_C"), extra=("*3", "+1", "-1")) -> None:
    """Three product sheets with 80 formula cells each at the same addresses, 60 of them equal."""
    for p, x in zip(names, extra):
        ws = wb.create_sheet(p)
        for r in range(1, 61):
            ws[f"B{r}"] = f"=A{r}*2"
        for r in range(61, 81):
            ws[f"B{r}"] = f"=A{r}{x}"


def test_sheet_copies_group_at_the_threshold_and_report_min_similarity():
    wb = _new()
    _prod_sheets(wb)
    tot = wb.create_sheet("Total")
    for r in range(1, 21):
        tot[f"B{r}"] = f"=Prod_A!B{r}+Prod_B!B{r}+Prod_C!B{r}"
    for i, yr in enumerate((2024, 2025, 2026)):  # chained year sheets: each opens with the previous closing
        ws = wb.create_sheet(f"Y{yr}")
        ws["B1"] = 500 if i == 0 else f"=Y{yr - 1}!B4"
        ws["B2"], ws["B3"], ws["B4"] = 120, 80, "=B1+B2-B3"
        for m in range(1, 13):
            c = L(m + 2)
            ws[f"{c}1"] = "=B1" if m == 1 else f"={L(m + 1)}4"
            ws[f"{c}2"], ws[f"{c}3"], ws[f"{c}4"] = "=$B$2/12", "=$B$3/12", f"={c}1+{c}2-{c}3"
    book, _ = _book_of(wb)
    groups = fp.sheet_copies(book, Budget(120))
    assert [g["members"] for g in groups] == [["Prod_A", "Prod_B", "Prod_C"], ["Y2024", "Y2025", "Y2026"]]
    prod = groups[0]
    assert prod["min_similarity"] == 0.75 and prod["per_copy"] == 80 and prod["copies"] == 3
    assert prod["diff"] == {"identical": 60, "reads a different input": 0, "differs only in a typed-in number": 0, "different logic": 20}
    assert prod["chain"] is False and prod["cross_reads"] == []
    assert (prod["adding_cells"], prod["adding_example"]) == (20, "Total!B1")
    assert prod["cells"]["Prod_B"] is book.sheets["Prod_B"].cells
    years = groups[1]
    assert years["chain"] is True and years["cross_reads"] == []
    # out of order the same reads are scattered cross reads (listed in the members' order), not a chain
    assert fp.consecutive_reads(book, {"members": ["Y2024", "Y2026", "Y2025"]}) == (False, [{"from": "Y2026", "to": "Y2025", "cells": 1}, {"from": "Y2025", "to": "Y2024", "cells": 1}])
    # two sheets that only share a few shapes never reach the aligned compare
    assert fp._overlap(fp.Counter({1: 60, 2: 20}), fp.Counter({1: 60, 3: 20})) == 0.6


# --- how copies differ, and the edits -------------------------------------------------------------
def test_diff_kinds_sort_identical_input_number_logic():
    wb = _new()
    for name, inp, num, logic in (("S1", "B", "0.1", "+1"), ("S2", "C", "0.12", "*2"), ("S3", "D", "0.15", "-1")):
        ws = wb.create_sheet(name)
        ws["B1"] = "=A1*2"  # identical
        ws["B2"] = f"=Inputs!{inp}$2"  # reads a different input
        ws["B3"] = f"=A3*{num}"  # differs only in a typed-in number
        ws["B4"] = f"=A4{logic}"  # different logic
        if name != "S3":
            ws["B5"] = "=A5*5"  # a formula in some copies only
    book, _ = _book_of(wb)
    stacks = fp.stacks_of(book, {"members": ["S1", "S2", "S3"]})
    assert fp.diff_kinds(book, stacks) == {"identical": 1, "reads a different input": 1, "differs only in a typed-in number": 1, "different logic": 2}
    assert stacks.address(1, (2, 2)) == ("S2", 2, 2)
    assert fp.formula_numbers_differ(book, stacks) == [{"cell": "S1!B3", "literals": [["0.1"], ["0.12"], ["0.15"]]}]
    logic = fp.logic_cells(book, stacks)
    assert [(x["cell"], x["action"], x["copies_with_cell"]) for x in logic] == [("S1!B4", fp.IF_ACTION, 3), ("S1!B5", fp.SIZE_DIFFERS, 2)]
    assert logic[0]["formulas"] == ["=A4+1", "=A4*2", "=A4-1"] and logic[1]["formulas"] == ["=A5*5", "=A5*5", None]


def _near_copies(wb) -> None:
    """Four product sheets: own numbers in B1:B3, row 8 *0.1 or *0.12 (a typed number inside the
    formula), row 9 +1 in one copy (logic, 2 forms), row 10 four different forms, row 11 missing in
    one copy, rows 12-25 identical."""
    specs = {"Prod_A": (100, 10, 0.05, "*0.1", "", "*1"), "Prod_B": (200, 6, 0.04, "*0.12", "", "+1"),
             "Prod_C": (50, 16, 0.07, "*0.1", "", "-1"), "Prod_D": (300, 8, 0.06, "*0.12", "+1", "/2")}
    for p, (pol, prem, lapse, rate, plus, form) in specs.items():
        ws = wb.create_sheet(p)
        ws["A1"], ws["B1"], ws["A2"], ws["B2"], ws["A3"], ws["B3"] = "Policies", pol, "Premium", prem, "Lapse", lapse
        for y in range(1, 21):
            c, prev = L(y + 1), L(y)
            ws[f"{c}5"] = y
            ws[f"{c}6"] = "=$B$1" if y == 1 else f"={prev}6*(1-$B$3)"
            ws[f"{c}7"] = f"={c}6*$B$2"
            ws[f"{c}8"] = f"={c}7{rate}"
            ws[f"{c}9"] = f"={c}7-{c}8{plus}"
            ws[f"{c}10"] = f"={c}6{form}"
            if p != "Prod_D":
                ws[f"{c}11"] = f"={c}9*2"
            for r in range(12, 26):
                ws[f"{c}{r}"] = f"={c}{r - 1}+{c}$5"


def test_edits_list_formula_numbers_and_decide_if_or_override():
    wb = _new()
    _near_copies(wb)
    book, _ = _book_of(wb)
    groups = fp.sheet_copies(book, Budget(120))
    assert len(groups) == 1 and groups[0]["members"] == ["Prod_A", "Prod_B", "Prod_C", "Prod_D"]
    e = fp.edits_for(book, groups[0], wb)
    assert (e["numbers_count"], e["formula_numbers_count"], e["logic_count"], e["size_count"]) == (3, 20, 40, 20)
    assert e["numbers"][0] == {"cell": "Prod_A!B1", "values_per_copy": [100, 200, 50, 300], "action": fp.NUMBERS_ACTION}
    fn = e["formula_numbers"][0]
    assert fn["cell"] == "Prod_A!B8" and fn["literals_per_copy"] == [["0.1"], ["0.12"], ["0.1"], ["0.12"]] and fn["action"] == fp.FORMULA_NUMBERS_ACTION
    by_cell = {x["cell"]: x for x in e["logic"]}
    assert by_cell["Prod_A!B9"]["action"] == fp.IF_ACTION  # 2 forms, present everywhere
    assert by_cell["Prod_A!B9"]["formulas_per_copy"] == ["=B7-B8", "=B7-B8", "=B7-B8", "=B7-B8+1"]
    assert by_cell["Prod_A!B10"]["action"] == fp.OVERRIDE_ACTION  # 4 distinct forms
    assert e["size"][0] == {"cell": "Prod_A!B11", "copies_with_cell": 3, "action": fp.SIZE_ACTION}
    assert "MM_LOOPINSTANCE" in fp.SIZE_ACTION and "pad the shorter copies" in fp.SIZE_ACTION
    assert len(e["logic"]) <= fp.MAX_EDITS and len(e["formula_numbers"]) <= fp.MAX_EDITS
    # without the workbook the typed-in numbers are not read, everything else is
    e2 = fp.edits_for(book, groups[0])
    assert e2["numbers"] == [] and e2["formula_numbers_count"] == 20


# --- what the copies step through -----------------------------------------------------------------
def test_step_target_sees_consecutive_rows_and_column_slots():
    wb = _new()
    wb.active["B2"], wb.active["C2"], wb.active["D2"] = 0.02, 0.03, 0.01
    sc = wb.create_sheet("Scen")
    _scen_blocks(sc)
    rows = wb.create_sheet("Rows")  # three blocks of two rows, each reading one 10-column slot of Paste
    for j in range(3):
        r0 = 1 + j * 3
        for c in range(2, 8):
            rows.cell(r0, c, f"=Paste!{L(c + 10 * j)}$1")
            rows.cell(r0 + 1, c, f"={L(c)}{r0}*2")
    book, analysis = _book_of(wb)
    names = _Names(analysis["workbooks"][0]["defined_names"])
    band = [f for f in fp.sheet_findings(book, book.sheets["Scen"], Budget(120)) if not f["fill"]][0]
    t = fp.step_target(book, band, book.a1_text("Scen", *band["first_cell"]), names)
    assert t == {"sheet": "Inputs", "r1": 2, "r2": 2, "c1": 2, "c2": 4, "stride": 1, "axis": "columns", "computed": None, "cells": 10}
    band = [f for f in fp.sheet_findings(book, book.sheets["Rows"], Budget(120)) if not f["fill"]][0]
    assert (band["axis"], band["copies"], band["width"]) == ("rows", 3, 3)
    t = fp.step_target(book, band, book.a1_text("Rows", *band["first_cell"]), names)
    assert (t["sheet"], t["axis"], t["stride"], t["computed"], t["cells"]) == ("Paste", "columns", 10, None, 6)
    assert (t["r1"], t["r2"], t["c1"], t["c2"]) == (1, 1, 2, 31)  # B1 .. the end of the third slot


def test_step_target_reads_computed_indices():
    wb = _new()
    out = wb.create_sheet("Out")
    for c in range(2, 40):
        out.cell(2, c, c)
    wb.defined_names["Pick"] = DefinedName("Pick", attr_text="Out!$B$2")
    one = wb.create_sheet("One")  # (1) rows step through 2-column slots, columns through the attributes
    for r in range(2, 10):
        for c in range(3, 7):
            one.cell(r, c, "=OFFSET(Pick, COLUMN()-3, (ROW()-1)*2)")
    two = wb.create_sheet("Two")  # (2) the expanding COUNTA range anchored at the first row (12 rows)
    for r in range(2, 14):
        two.cell(r, 1, f"item {r}")
        two.cell(r, 2, f"=OFFSET(Out!B$2, 0, 2*(COUNTA($A$2:$A{r})-1))")
    three = wb.create_sheet("Three")  # (3) the keyed slot: MATCH of a key in a list
    three["A1"] = "k3"
    for r in range(3, 13):
        three.cell(r, 3, f"=X{r}+OFFSET(Out!B{r - 1}, 0, 2*(MATCH($A$1, List!$A$1:$A$8, 0)-1))")
    lst = wb.create_sheet("List")
    for r in range(1, 9):
        lst.cell(r, 1, f"k{r}")
    plain = wb.create_sheet("Plain")  # a fill with no computed index steps through nothing
    for r in range(2, 14):
        plain.cell(r, 2, f"=A{r}*2")
    wb.defined_names["Pick2"] = DefinedName("Pick2", attr_text="Out!$U$13")
    four = wb.create_sheet("Four")  # (4) 12 rows reading the LAST column of 10-column slots: row 2 reads U13 + 10 = AE13
    for r in range(2, 14):
        four.cell(r, 2, "=OFFSET(Pick2,0,(ROW()-1)*10)")
    book, analysis = _book_of(wb)
    names = _Names(analysis["workbooks"][0]["defined_names"])

    def target(sheet):
        found = fp.sheet_findings(book, book.sheets[sheet], Budget(120))
        assert len(found) == 1 and found[0]["fill"] and found[0]["diff"]["identical"] == found[0]["per_copy"]
        return found[0], fp.step_target(book, found[0], book.a1_text(sheet, *found[0]["first_cell"]), names)

    f, t = target("One")
    assert f["axes"] == {"rows": 8, "columns": 4}
    c = t["computed"]
    assert (c["form"], c["w"], c["slot_axis"], c["step_axis"], c["anchor"], c["positions"]) == ("row", 2, "columns", "rows", "Out!$B$2", 8)
    # the hull runs from the first read cell (D, slot 1) to the last (R, slot 8): 15 columns, not 8 whole slots
    assert (t["sheet"], t["stride"], t["axis"], t["c1"], t["c2"], t["r1"], t["r2"]) == ("Out", 2, "columns", 4, 18, 2, 5)
    assert c["others"][0]["form"] == "column" and c["others"][0]["w"] == 1  # the attribute axis
    f, t = target("Two")
    c = t["computed"]
    assert (c["form"], c["w"], c["slot_axis"], c["step_axis"], c["anchor"]) == ("counta", 2, "columns", "rows", "Out!$B$2")
    assert (t["c1"], t["c2"], t["stride"], c["positions"]) == (2, 24, 2, 12)  # 12 rows, 2-column slots from B: B .. the 12th read cell (X)
    f, t = target("Three")
    c = t["computed"]
    assert (c["form"], c["w"], c["list"], c["key"], c["list_ref"], c["list_len"]) == ("match", 2, "List!$A$1:$A$8", "$A$1", "List!A1:A8", 8)
    assert (t["sheet"], t["c1"], t["c2"], t["axis"]) == ("Out", 2, 16, "columns")  # anchor .. anchor + w*(len(list)-1)
    f, t = target("Four")  # the index lands on the LAST column of each 10-column slot: the hull ends at the area's end, not 9 past it
    assert (t["sheet"], t["r1"], t["r2"], t["c1"], t["c2"], t["stride"], t["computed"]["positions"]) == ("Out", 13, 13, 31, 141, 10, 12)  # AE13 .. EK13, 12 slots
    found = fp.sheet_findings(book, book.sheets["Plain"], Budget(120))
    assert len(found) == 1 and found[0]["fill"]
    assert fp.step_target(book, found[0], book.a1_text("Plain", *found[0]["first_cell"]), names) is None


def test_the_computed_index_idioms_have_one_parser_with_balanced_parentheses():
    """app.finder_grouping.displacement_form reads through finder_patterns._computed_arg, so the readers pass
    and the step target see the same forms. ROW()-21*10 is ROW()-210 in Excel, not an idiom; a pair of
    parentheses around the whole text is dropped; a whole-column or two-dimensional COUNTA is no expanding range."""
    from app.finder_grouping import displacement_form

    cases = {
        "(ROW()-21)*10": ("row", 10, 21, "rows"), "10*(ROW()-21)": ("row", 10, 21, "rows"), "ROW()-21": ("row", 1, 21, "rows"),
        "(COLUMN()-14)": ("column", 1, 14, "columns"), "@COLUMN()-14": ("column", 1, 14, "columns"), "10 * ( ROW() - 21 )": ("row", 10, 21, "rows"),
        "((ROW()-21)*10)": ("row", 10, 21, "rows"),
        "10*(COUNTA($B$9:$B9)-1)": ("counta", 10, None, "rows"), "(COUNTA($B$9:B9)-1)*10": ("counta", 10, None, "rows"), "2*(COUNTA($B$1:C$1)-1)": ("counta", 2, None, "columns"),
        "10*(MATCH($C$7,'M 1'!$B$9:$B$34,0)-1)": ("match", 10, None, None), "(MATCH(A1,List,0)-1)": ("match", 1, None, None), "MATCH(TRIM(A1),List,0)-1": ("match", 1, None, None),
        "ROW()-21*10": None, "(ROW()-21)*10+1": None, "ROW(A1)-1": None, "ROW()": None, "(ROW()-21)/10": None, "10*(ROW()-21)+COLUMN()": None,
        "COUNTA($A:$A)-1": None, "10*(COUNTA($A$1:B2)-1)": None, "(ROWS($A$2:A2)-1)*10": None, "0": None, "Slot": None,
    }
    for text, want in cases.items():
        a, b = fp._computed_arg(text), displacement_form(text)
        got = (a["form"], a["w"], a.get("n"), a.get("step_axis")) if a else None
        assert got == want, (text, got, want)
        assert ((b["form"], b["w"], b["n"], b["step_axis"]) if b else None) == got, text
    assert displacement_form("10 * (ROW()-21)")["text"] == "10*(ROW()-21)"  # the readers pass keeps its tight text
