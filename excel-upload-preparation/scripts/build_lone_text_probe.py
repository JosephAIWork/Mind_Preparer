"""A  made-up workbook  for Mind (invented numbers, no client data): which formulas need a lone text cell?

Mind leaves out a text cell that stands alone (docs/MIND_VS_EXCEL.md, 8 Oct 2026). Since 17 Nov 2026 the  app lists
such a cell only when a formula that reads it  would get another result from an  empty cell (prep.readers_needing).
Each formula on the sheet  'Probe' tests one of those judgements:
- a reader that skips text  (MAX, SUM ... over a range; VLOOKUP  / HLOOKUP in a column it neither searches nor
  returns; N(); an INDEX whose stepped row never lands on the cell): the app does not list the cell;
- a reader that needs the text: the app lists the cell, and Prep fixes it where it can (a title above, or the text
  itself written into the formula).
Column C is the formula, column D  the value Excel computed (typed in), column F what the app expects of Mind on the
workbook AS WRITTEN: "same" or "differs". Rows H1 to H5 read another probe cell and put brackets around it (or add
5): Mind's difference list does not name a cell where Excel shows a text and Mind returns nothing, and these do.
The sheets Nw and Lone hold a label and a number beside the lone text: a sheet with nothing but lone text cells is
not imported at  all, and Mind stops the  conversion ("Spreadsheet not found", rule FRM-006).

    python scripts/build_lone_text_probe.py  <output folder>      (run from  excel-upload-preparation)

Writes lone_text_probe.xlsx,  lone_text_probe.tests.json (formula and Excel's value of every test) and
lone_text_probe.app_view.json (the text cells standing alone, and the ones rule STR-008 lists). Upload it with
scripts/mind_upload.py --keep. Seen in Mind on 20 Nov 2026:  docs/MIND_VS_EXCEL.md.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl  # noqa: E402
from openpyxl.workbook.defined_name import DefinedName  # noqa: E402

OUT = Path(sys.argv[1]).resolve()
SRC = OUT / "lone_text_probe_formulas_only.xlsx"
BOOK = OUT / "lone_text_probe.xlsx"

TESTS = [
    # ---  A: a  lone text inside a range that a text-blind function reads (Model C: Calc!A259 under MAX($A$7:A304)+1)
    ("A1", "MAX over a column with a lone text inside it (Col!A15)", "=MAX(Col!$A$7:$A$23)+1", "same"),
    ("A2", "MIN over the same column", "=MIN(Col!$A$7:$A$23)", "same"),
    ("A3", "SUM over the same column", "=SUM(Col!$A$7:$A$23)", "same"),
    ("A4", "AVERAGE over the same column", "=AVERAGE(Col!$A$7:$A$23)", "same"),
    ("A5", "COUNT over the same column", "=COUNT(Col!$A$7:$A$23)", "same"),
    ("A6", "PRODUCT over the same column", "=PRODUCT(Col!$A$7:$A$23)", "same"),
    ("A7", "SUMPRODUCT of that column and the one beside it", "=SUMPRODUCT(Col!$A$7:$A$23,Col!$B$7:$B$23)", "same"),
    ("A8", "the other way: COUNTA does count the lone text (ColN!A15)", "=COUNTA(ColN!$A$7:$A$23)", "differs"),
    ("A9", "the other way: INDEX returns the lone text itself (ColN!A15)", "=INDEX(ColN!$A$7:$A$23,9)", "differs"),
    # --- B: a lookup table with a lone text in a column that is neither searched nor  returned (Model C: CFs!E62)
    ("B1", "VLOOKUP: a lone text in a column neither searched nor returned (Look!E35)", "=VLOOKUP(105,Look!$C$30:$G$40,5,FALSE)", "same"),
    ("B2", "the other way: VLOOKUP returns the lone text's column (Look2!E35)", "=VLOOKUP(105,Look2!$C$30:$G$40,3,FALSE)", "differs"),
    ("B3", "HLOOKUP: a lone text in a row neither searched nor returned (Hl!G5)", "=HLOOKUP(104,Hl!$C$3:$M$7,5,FALSE)", "same"),
    # --- C: N() of a lone text
    ("C1", "N() of a lone text (Nw!C3): 0 for a text and for an empty cell", "=N(Nw!$C$3)", "same"),
    # --- D:  the stepped INDEX (Model B: 332  headings 'Per Capita Cost' inside INDEX(range, 1+(year-first  year-1)*117, ...))
    ("D1", "INDEX whose row is 1 + n x 6: it never lands on the headings (Base!D7, D13, D19)", "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,2)", "same"),
    ("D2", "the same with the column from MATCH, as in the model", "=INDEX(Base!$C$5:$E$22,1+(yr-first_yr-1)*6,MATCH(102,Base!$C$5:$E$5,0))", "same"),
    ("D3", "the other way: an INDEX that does land on a heading (Base2!D13)", "=INDEX(Base2!$C$5:$E$22,3+(yr-first_yr-1)*6,2)", "differs"),
    # --- E: a lone text as the criterion of SUMIFS  (Model B: 'Cost data'!AF10)
    ("E1", "a lone text as the criterion of SUMIFS (Crit!D5)", "=SUMIFS(Crit!$B$2:$B$9,Crit!$A$2:$A$9,Crit!$D$5)", "differs"),
    ("E2", "the same in COUNTIFS (Crit!D5)", "=COUNTIFS(Crit!$A$2:$A$9,Crit!$D$5)", "differs"),
    ("E3", "the same in AVERAGEIFS (Crit!D5)", "=AVERAGEIFS(Crit!$B$2:$B$9,Crit!$A$2:$A$9,Crit!$D$5)", "differs"),
    ("E4", "SUMIFS criterion with no room for a title above (Crit2!D5): Prep writes the text into the formula", "=SUMIFS(Crit2!$B$2:$B$9,Crit2!$A$2:$A$9,Crit2!$D$5)", "differs"),
    ("E5", "COUNTIFS criterion with no room for a title above (Crit2!D5)", "=COUNTIFS(Crit2!$A$2:$A$9,Crit2!$D$5)", "differs"),
    ("E6", "the criterion typed into the formula, the way Prep writes it", '=SUMIFS(Crit!$B$2:$B$9,Crit!$A$2:$A$9,"north")', "same"),
    # --- F: the plain  case, seen in Mind on 8 Oct 2026
    ("F1", "a lone text joined to a text (Lone!C3)", '="x#"&Lone!$C$3', "differs"),
    # --- G: a bare # as the first header of a table (Model C: CFs!B29)
    ("G1", "a table whose first header is a bare #: VLOOKUP by the row number", "=VLOOKUP(3,Hash!$B$4:$D$8,3,FALSE)", "same"),
    ("G2", "the same table: a name out of it", "=INDEX(Hash!$C$4:$C$8,2)", "same"),
    # --- H: one step further, to SEE what came back: a text that comes back empty is not named in Mind's list
    ("H1", "what A9 returned, in brackets (Probe!C10)", '="["&C10&"]"', "differs"),
    ("H2", "what B2 returned, in brackets (Probe!C12)", '="["&C12&"]"', "differs"),
    ("H3", "what D3 returned, in brackets (Probe!C17)", '="["&C17&"]"', "differs"),
    ("H4", "what C1 returned, plus 5 (Probe!C14)", "=C14+5", "same"),
    ("H5", "what G2 returned, in brackets (Probe!C26)", '="["&C26&"]"', "same"),
]


def formulas_only() -> None:
    wb = openpyxl.Workbook()
    probe = wb.active
    probe.title = "Probe"
    for c, head in enumerate(("Test", "What it tests", "Result", "Excel's value", "Formula", "The app expects Mind"), start=1):
        probe.cell(1, c, head)
    for r, (tid, what, formula, expects) in enumerate(TESTS, start=2):
        probe.cell(r, 1, tid)
        probe.cell(r, 2, what)
        probe.cell(r, 3, formula)
        probe.cell(r, 5, formula[1:])
        probe.cell(r, 5).number_format = "@"
        probe.cell(r, 6, expects)

    inputs = wb.create_sheet("Inputs")
    inputs["A1"], inputs["B1"] = "Year", 2028
    inputs["A2"], inputs["B2"] = "First year", 2026
    wb.defined_names["yr"] = DefinedName("yr", attr_text="Inputs!$B$1")
    wb.defined_names["first_yr"] = DefinedName("first_yr", attr_text="Inputs!$B$2")

    for name in ("Col", "ColN"):  # two  blocks of  numbers in A and B, a text alone between them in A15
        ws = wb.create_sheet(name)
        for i, (a, b) in enumerate(zip((3, 5, 2, 8, 6, 4), (2, 2, 2, 2, 2, 2))):
            ws.cell(7 + i, 1, a)
            ws.cell(7 + i, 2, b)
        ws["A15"] = "Second part"
        for i, (a, b) in enumerate(zip((7, 1, 9, 2, 5, 3), (3, 3, 3, 3, 3, 3))):
            ws.cell(18 + i, 1, a)
            ws.cell(18 + i, 2, b)

    for name in ("Look", "Look2"):  # keys in C, values in G, a text alone in E35
        ws = wb.create_sheet(name)
        for i in range(11):
            ws.cell(30 + i, 3, 100 + i)
            ws.cell(30 + i, 7, 1000 + 7 * i)
        ws["E35"] = "note"

    ws = wb.create_sheet("Hl")  # keys in row 3, values in row 7, a text  alone in G5
    for i in range(11):
        ws.cell(3, 3 + i, 100 + i)
        ws.cell(7, 3 + i, 2000 + 3 * i)
    ws["G5"] = "note"

    ws = wb.create_sheet("Nw")
    ws["C3"] = "Line 1"
    ws["A20"], ws["B20"] = "Rate", 0.5  # a grid, so that Mind imports the sheet at  all

    for name in ("Base", "Base2"):  # three yearly blocks of 6 rows  in C5:E22, a heading alone in D  of each block
        ws = wb.create_sheet(name)
        for block, top in enumerate((5, 11, 17)):
            for c, col in enumerate("CDE"):
                ws[f"{col}{top}"] = 100 * (block + 1) + c
                ws[f"{col}{top + 4}"] = 10 * (block + 1) + c
                ws[f"{col}{top + 5}"] = 20 * (block + 1) + c
            ws[f"D{top + 2}"] = "Per Capita Cost"

    rows = [("north", 10), ("south", 20), ("north", 5), ("east", 1), ("west", 7), ("south", 4), ("north", 3), ("east", 2)]
    for name in ("Crit", "Crit2"):  # a table of regions and amounts, the region asked  for alone in D5
        ws = wb.create_sheet(name)
        ws["A1"], ws["B1"] = "Region", "Amount"
        for i, (key, amount) in enumerate(rows):
            ws[f"A{i + 2}"], ws[f"B{i + 2}"] = key, amount
        ws["D5"] = "north"
    wb["Crit2"]["H9"] = "=D4*2"  # reads the cell above  D5: no title can go there

    ws = wb.create_sheet("Lone")
    ws["C3"] = "Line 1"
    ws["A20"], ws["B20"] = "Rate", 0.5

    ws = wb.create_sheet("Hash")  # a table whose first header is a bare '#'
    ws["B3"], ws["C3"], ws["D3"] = "#", "Name", "Amount"
    for i, nm in enumerate(("alpha", "beta", "gamma", "delta", "epsilon")):
        ws[f"B{4 + i}"], ws[f"C{4 + i}"], ws[f"D{4 + i}"] = i + 1, nm, 10 * (i + 1)

    wb.save(SRC)
    wb.close()


def through_excel() -> list[dict]:
    """Excel  computes every  formula; its value is typed beside the formula; Excel saves the workbook."""
    from app.excel_com import excel_session, open_workbook

    out = []
    with excel_session(mm_addin=False) as excel:
        wb = open_workbook(excel, SRC, read_only=False)
        excel.CalculateFullRebuild()
        probe = wb.Worksheets("Probe")
        for r, (tid, what, formula, expects) in enumerate(TESTS, start=2):
            v = probe.Range(f"C{r}").Value
            text = probe.Range(f"C{r}").Text
            if v not in (None, ""):
                probe.Range(f"D{r}").Value = v
            out.append({"test": tid, "cell": f"Probe!C{r}", "what": what, "formula": probe.Range(f"C{r}").Formula, "excel": text, "app_expects": expects})
        for ws in wb.Worksheets:
            ws.Columns.AutoFit()
        probe.Columns(5).ColumnWidth = 60
        wb.SaveAs(str(BOOK), FileFormat=51)
        wb.Close(SaveChanges=False)
    return out


def app_view() -> dict:
    """What  the app's  own scan says of the workbook: every text cell standing alone, and the ones it lists."""
    from app.config import load_config
    from app.inventory import build_analysis
    from app.prep import standalone_labels
    from app.rules_engine import RulesEngine
    from app.validators import run_rule

    analysis = build_analysis(BOOK, OUT / "scan_work", "probe")
    lone = [f"{lab['sheet']}!{lab['cell']}" for lab in standalone_labels(analysis)]
    finding = run_rule(RulesEngine().get("STR-008"), analysis, load_config())
    listed = [f"{s['sheet']}!{s['cell']}" for s in finding["observed"]["sites"]]
    return {"standing_alone": lone, "listed_by_STR-008": listed, "not_listed": [c for c in lone if c not in listed], "status": finding["status"]}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for p in (SRC, BOOK):
        if p.exists():
            p.unlink()
    formulas_only()
    tests = through_excel()
    (OUT / "lone_text_probe.tests.json").write_text(json.dumps(tests, indent=1, ensure_ascii=False), encoding="utf-8")
    for t in tests:
        print(f"{t['test']}  {t['cell']:10s} {t['excel']!r:18s} {t['app_expects']:8s} {t['what']}")
    view = app_view()
    (OUT / "lone_text_probe.app_view.json").write_text(json.dumps(view, indent=1, ensure_ascii=False), encoding="utf-8")
    print("standing alone:", view["standing_alone"])
    print("listed by the app (STR-008):", view["listed_by_STR-008"])
    print("not listed:", view["not_listed"])
    print("saved", BOOK, BOOK.stat().st_size, "bytes")


main()
