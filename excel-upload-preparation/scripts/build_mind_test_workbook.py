"""A small test workbook for Mind (no client data): the three constructs that came back different
on Horizon, each next to candidate rewrites. Column C is the formula, column D the value Excel
computed (typed in), so the difference list Mind gives says which construct behaves differently.

    python build_probe.py <output.xlsx>      (run from excel-upload-preparation)
"""
import datetime
import json
import sys
from pathlib import Path

from app.excel_com import excel_session

OUT = Path(sys.argv[1]).resolve()


def serial(y, m, d):
    return (datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days


LOANS = [  # loan id, purpose, date, capital, payment -- invented
    (1, "Home", (2022, 7, 1), 100000, 500), (1, "Home", (2023, 1, 1), 95000, 500), (1, "Home", (2024, 1, 1), 90000, 500),
    (1, "Home", (2025, 1, 1), 85000, 500), (1, "Home", (2026, 1, 1), 80000, 500), (1, "Home", (2027, 1, 1), 75000, 500),
    (2, "Rental", (2022, 7, 1), 60000, 300), (2, "Rental", (2023, 1, 1), 55000, 300), (2, "Rental", (2023, 7, 1), 5000, 300),
    (2, "Rental", (2024, 1, 1), 50000, 300), (2, "Rental", (2025, 1, 1), 45000, 300), (2, "Rental", (2026, 1, 1), 40000, 300),
]
N = len(LOANS) + 1  # last row of the table

ID, YEAR_, S0, S1, D0, D1, BLANK, EMPTY, NUM = ("Params!$B$%d" % i for i in range(1, 10))
ORIG = 'SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!C:C,">="&DATE(%s,1,1),Loans!C:C,"<"&DATE(%s,12,31))' % (ID, YEAR_, YEAR_)

TESTS = [
    # --- A: SUMIFS with a comparison criterion built from a date
    ("A00", "Horizon's formula: whole columns, criterion \">=\"&DATE()", "=" + ORIG),
    ("A01", "the same on the table's rows only", '=SUMIFS(Loans!$D$2:$D$%d,Loans!$A$2:$A$%d,%s,Loans!$C$2:$C$%d,">="&DATE(%s,1,1),Loans!$C$2:$C$%d,"<"&DATE(%s,12,31))' % (N, N, ID, N, YEAR_, N, YEAR_)),
    ("A02", "control: loan id only", "=SUMIFS(Loans!D:D,Loans!A:A,%s)" % ID),
    ("A03", "date criterion alone: \">=\"&DATE()", '=SUMIFS(Loans!D:D,Loans!C:C,">="&DATE(2023,1,1))'),
    ("A04", "date criterion typed as a number: \">=44927\"", '=SUMIFS(Loans!D:D,Loans!C:C,">=44927")'),
    ("A05", "date criterion with equals: DATE()", "=SUMIFS(Loans!D:D,Loans!C:C,DATE(2023,1,1))"),
    ("A06", "the text the criterion becomes", '=">="&DATE(%s,1,1)' % YEAR_),
    ("A07", "comparison criterion on a column of plain numbers", '=SUMIFS(Loans!D:D,Loans!F:F,">="&2023)'),
    ("A08", "\">=\"&DATE() against the dates held as plain numbers (column G)", '=SUMIFS(Loans!D:D,Loans!G:G,">="&DATE(%s,1,1))' % YEAR_),
    ("A10", "candidate 1: helper year column, equals criteria only", "=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!F:F,%s)" % (ID, YEAR_)),
    ("A11", "candidate 2: SUMPRODUCT on the table's rows", "=SUMPRODUCT((Loans!$A$2:$A$%d=%s)*(Loans!$C$2:$C$%d>=DATE(%s,1,1))*(Loans!$C$2:$C$%d<DATE(%s,12,31))*Loans!$D$2:$D$%d)" % (N, ID, N, YEAR_, N, YEAR_, N)),
    ("A12", "candidate 3: the bounds typed as numbers in cells", '=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!C:C,">="&%s,Loans!C:C,"<"&%s)' % (ID, S0, S1)),
    ("A13", "candidate 3b: the bounds in cells holding =DATE(), shown as numbers", '=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!C:C,">="&%s,Loans!C:C,"<"&%s)' % (ID, D0, D1)),
    ("A14", "candidate 4: \">=\"&(DATE()+0)", '=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!C:C,">="&(DATE(%s,1,1)+0),Loans!C:C,"<"&(DATE(%s,12,31)+0))' % (ID, YEAR_, YEAR_)),
    ("A15", "candidate 5: \">=\"&INT(DATE())", '=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!C:C,">="&INT(DATE(%s,1,1)),Loans!C:C,"<"&INT(DATE(%s,12,31)))' % (ID, YEAR_, YEAR_)),
    ("A16", "candidate 1 on a row after the end (the year is \"\")", "=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!F:F,%s)" % (ID, BLANK)),
    ("A17", "candidate 1 on a row after the end (\"\"+1)", "=SUMIFS(Loans!D:D,Loans!A:A,%s,Loans!F:F,%s+1)" % (ID, BLANK)),
    # --- B: arithmetic and functions on a cell whose formula returns ""
    ("B00", "Horizon column E: IFERROR(1+INT((\"\"-1)/12),\"\")", '=IFERROR(1+INT((%s-1)/12),"")' % BLANK),
    ("B01", "Horizon column F: IFERROR(51+\"\"-2022,\"\")", '=IFERROR(51+%s-2022,"")' % BLANK),
    ("B02", "candidate: IF(x=\"\",\"\",...) in front", '=IF(%s="","",IFERROR(1+INT((%s-1)/12),""))' % (BLANK, BLANK)),
    ("B03", "N(\"\")", "=N(%s)" % BLANK),
    ("B04", "MIN(90,\"\")", "=MIN(90,%s)" % BLANK),
    ("B05", "Horizon column C: 14+N(\"\")-N(64)", "=14+N(%s)-N(%s)" % (BLANK, NUM)),
    ("B06", "candidate for column C: IF(x=\"\",0,x)", '=14+IF(%s="",0,%s)-IF(%s="",0,%s)' % (BLANK, BLANK, NUM, NUM)),
    ("B07", "candidate for column O: IF(x=\"\",90,MIN(90,x))", '=IF(%s="",90,MIN(90,%s))' % (BLANK, BLANK)),
    ("B08", "\"\">=156 inside IF", "=IF(%s>=156,0,1)" % BLANK),
    ("B09", "\"\"<=0 inside IF", "=IF(%s<=0,0,1)" % BLANK),
    ("B10", "(\"\"<=156)*1", "=(%s<=156)*1" % BLANK),
    ("B11", "x=\"\"", '=%s=""' % BLANK),
    ("B12", "ISNUMBER(\"\")", "=ISNUMBER(%s)" % BLANK),
    # --- C: a comparison with a reference to an empty cell
    ("C00", "Horizon C20!C15: =+(an empty cell)", "=+%s" % EMPTY),
    ("C01", "Horizon C20!C18: 2311+0-2311 = that cell", "=+2311+0-2311={C00}"),
    ("C02", "0 = (an empty cell), directly", "=0=%s" % EMPTY),
    ("C03", "candidate a: N(empty cell)", "=N(%s)" % EMPTY),
    ("C04", "the check against candidate a", "=+2311+0-2311={C03}"),
    ("C05", "candidate b: empty cell + 0", "=%s+0" % EMPTY),
    ("C06", "the check against candidate b", "=+2311+0-2311={C05}"),
    ("C07", "candidate c: IF(cell=\"\",0,cell)", '=IF(%s="",0,%s)' % (EMPTY, EMPTY)),
    ("C08", "the check against candidate c", "=+2311+0-2311={C07}"),
    ("C09", "(empty cell)=\"\"", '=%s=""' % EMPTY),
    ("C10", "candidate d: N() on the check's side", "=+2311+0-2311=N({C00})"),
    # --- D (round 2): a lookup that lands on an empty cell, then compared (Horizon CALCUL!H9 and column I)
    ("D00", "Horizon H9: VLOOKUP that lands on an empty cell", "=VLOOKUP(%s,Loans!A:H,8,0)" % ID),
    ("D01", "Horizon column I: IF(1<=that cell,0,1)*60000", "=IF(1<={D00},0,1)*60000"),
    ("D02", "candidate: N(VLOOKUP(...))", "=N(VLOOKUP(%s,Loans!A:H,8,0))" % ID),
    ("D03", "the same IF against the N() cell", "=IF(1<={D02},0,1)*60000"),
    ("D04", "candidate: VLOOKUP(...)+0", "=VLOOKUP(%s,Loans!A:H,8,0)+0" % ID),
    ("D05", "the same IF against the +0 cell", "=IF(1<={D04},0,1)*60000"),
    ("D06", "IF(1<=(an empty cell),0,1), directly", "=IF(1<=%s,0,1)" % EMPTY),
    ("D07", "IF(1<=N(an empty cell),0,1)", "=IF(1<=N(%s),0,1)" % EMPTY),
]


def build(excel):
    wb = excel.Workbooks.Add()
    while wb.Worksheets.Count < 3:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    loans, params, probe = wb.Worksheets(1), wb.Worksheets(2), wb.Worksheets(3)
    loans.Name, params.Name, probe.Name = "Loans", "Params", "Probe"

    loans.Range("A1:H1").Value = (("Loan id", "Purpose", "Date", "Capital", "Payment", "Year", "Date as number", "Deferral"),)
    for i, (lid, purpose, (y, m, d), capital, payment) in enumerate(LOANS, start=2):
        if lid == 1:
            loans.Range(f"H{i}").Value = 12        # loan 2 has none: the cell stays empty, as in Horizon
        loans.Range(f"A{i}:E{i}").Value = ((lid, purpose, serial(y, m, d), capital, payment),)
        loans.Range(f"F{i}").Formula = f"=IF(AND(MONTH(C{i})=12,DAY(C{i})=31),0,YEAR(C{i}))"
        loans.Range(f"G{i}").Formula = f"=C{i}+0"
    loans.Range(f"C2:C{N}").NumberFormat = "m/d/yyyy"
    loans.Range(f"G2:G{N}").NumberFormat = "General"

    rows = [
        ("Loan id", 2), ("Year", 2023), ("Start of the year, typed as a number", serial(2023, 1, 1)), ("31 December, typed as a number", serial(2023, 12, 31)),
        ("Start of the year, =DATE()", "=DATE(B2,1,1)"), ("31 December, =DATE()", "=DATE(B2,12,31)"),
        ("A formula that returns \"\"", '=IF(B1=0,1,"")'), ("An empty cell", None), ("A number", 64),
    ]
    params.Range("A1").Value = "Parameter"
    params.Range("A1").ClearContents()
    for i, (label, value) in enumerate(rows, start=1):
        params.Range(f"A{i}").Value = label
        if isinstance(value, str) and value.startswith("="):
            params.Range(f"B{i}").Formula = value
        elif value is not None:
            params.Range(f"B{i}").Value = value
    params.Range("B5:B6").NumberFormat = "General"

    probe.Range("A1:E1").Value = (("Test", "What it tests", "Result", "Excel's value", "Formula"),)
    where = {tid: f"C{r}" for r, (tid, _, _) in enumerate(TESTS, start=2)}
    for r, (tid, what, formula) in enumerate(TESTS, start=2):
        for k, v in where.items():
            formula = formula.replace("{" + k + "}", v)
        probe.Range(f"A{r}").Value = tid
        probe.Range(f"B{r}").Value = what
        probe.Range(f"C{r}").Formula = formula
        probe.Range(f"E{r}").NumberFormat = "@"
        probe.Range(f"E{r}").Value = formula[1:]
    excel.CalculateFullRebuild()
    out = []
    for r, (tid, what, formula) in enumerate(TESTS, start=2):
        v = probe.Range(f"C{r}").Value
        text = probe.Range(f"C{r}").Text
        if v not in (None, ""):
            probe.Range(f"D{r}").Value = v
        out.append({"test": tid, "cell": f"Probe!C{r}", "what": what, "formula": probe.Range(f"C{r}").Formula, "excel": text})
    for ws in (loans, params, probe):
        ws.Columns.AutoFit()
    probe.Columns(5).ColumnWidth = 60
    wb.SaveAs(str(OUT), FileFormat=51)
    wb.Close(SaveChanges=False)
    return out


def main():
    if OUT.exists():
        OUT.unlink()
    with excel_session() as excel:
        out = build(excel)
    OUT.with_suffix(".tests.json").write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8")
    for t in out:
        print(f"{t['test']}  {t['cell']:10s} {t['excel']!r:14s} {t['what']}")
    print("saved", OUT, OUT.stat().st_size, "bytes")


main()
