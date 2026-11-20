# -*- coding: utf-8 -*-
"""20  Nov 2026,  a made-up workbook (invented numbers, no client data): does a label beside a lone text
make Mind import it, when the text sits inside a range COUNTA and INDEX read? Joseph: "App  tries, never stops".

ColR!A15: a lone text in column A, nothing left  of it -> the app's label  goes RIGHT (B15), not yet seen in  Mind.
ColL!C15: a lone text in column C, B empty            -> the app's  label goes LEFT (B15), the shape Mind
                                                          imported  on 8 Oct 2026 (Lone!F3 'Name' beside G3).
Builds label_probe.xlsx (Excel computes it), runs the app's Prep action for lone text cells on it, applies it
through Excel, runs the numbers check, and writes label_probe_prepared.xlsx and label_probe.json.
Usage: python scripts/build_label_probe.py <out folder>   (run from excel-upload-preparation)
Then: python scripts/mind_upload.py <out folder>/label_probe_prepared.xlsx --project zzz_mindready_... --keep"""
import json
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.config import load_config
from app.excel_com import excel_session, open_workbook
from app.inventory import build_analysis
from app.numbers_check import Moves, compare_workbooks
from app.prep import apply_operations, plan_actions
from app.rules_engine import RulesEngine
from app.validators import run_rule

OUT = Path(sys.argv[1]).resolve()
OUT.mkdir(parents=True, exist_ok=True)
SRC = OUT / "label_probe_src.xlsx"
BOOK = OUT / "label_probe.xlsx"

TESTS = [
    ("R1", "COUNTA over a column with a lone text inside it (ColR!A15, label goes right)", "=COUNTA(ColR!$A$7:$A$23)"),
    ("R2", "INDEX that returns that lone text", "=INDEX(ColR!$A$7:$A$23,9)"),
    ("R3", "what R2 returned, in brackets", '="["&C3&"]"'),
    ("L1", "COUNTA over a column with a lone text inside it (ColL!C15, label goes left)", "=COUNTA(ColL!$C$7:$C$23)"),
    ("L2", "INDEX that returns that lone text", "=INDEX(ColL!$C$7:$C$23,9)"),
    ("L3", "what L2 returned, in brackets", '="["&C6&"]"'),
]

wb = openpyxl.Workbook()
probe = wb.active
probe.title = "Probe"
for c, head in enumerate(("Test", "What it tests", "Result", "Excel's value"), start=1):
    probe.cell(1, c, head)
for r, (tid, what, formula) in enumerate(TESTS, start=2):
    probe.cell(r, 1, tid)
    probe.cell(r, 2, what)
    probe.cell(r, 3, formula)
for name, col in (("ColR", 1), ("ColL", 3)):
    ws = wb.create_sheet(name)
    for i, (a, b) in enumerate(zip((3, 5, 2, 8, 6, 4), (2, 2, 2, 2, 2, 2))):
        ws.cell(7 + i, col, a)
        ws.cell(7 + i, col + 1, b)
    ws.cell(15, col, "Second part")
    for i, (a, b) in enumerate(zip((7, 1, 9, 2, 5, 3), (3, 3, 3, 3, 3, 3))):
        ws.cell(18 + i, col, a)
        ws.cell(18 + i, col + 1, b)
wb.save(SRC)
wb.close()

tests = []
with excel_session(mm_addin=False) as excel:
    book = open_workbook(excel, SRC, read_only=False)
    excel.CalculateFullRebuild()
    sh = book.Worksheets("Probe")
    for r, (tid, what, formula) in enumerate(TESTS, start=2):
        v = sh.Range(f"C{r}").Value
        sh.Range(f"D{r}").Value = v
        tests.append({"test": tid, "cell": f"Probe!C{r}", "what": what, "formula": formula, "excel": sh.Range(f"C{r}").Text})
    book.SaveAs(str(BOOK), FileFormat=51)
    book.Close(SaveChanges=False)
SRC.unlink()

analysis = build_analysis(BOOK, OUT / "scan_before", "label_probe")
listed = [f"{s['sheet']}!{s['cell']}" for s in run_rule(RulesEngine().get("STR-008"), analysis, load_config())["observed"]["sites"]]
action = {a["id"]: a for a in plan_actions(analysis, {"findings": []})}["title_lone_text_cells"]
ops = action["operations"]
out = apply_operations(BOOK, OUT / "prep", ops)
done = Path(out["output_path"])
prepared = OUT / "label_probe_prepared.xlsx"
prepared.write_bytes(done.read_bytes())
check = compare_workbooks(BOOK, prepared, Moves([ops]))
after = build_analysis(prepared, OUT / "scan_after", "label_probe_prepared")
left = [f"{s['sheet']}!{s['cell']}" for s in run_rule(RulesEngine().get("STR-008"), after, load_config())["observed"]["sites"]]
result = {
    "tests": tests, "listed_before": listed,
    "operations": [{"op": o["op"], "sheet": o["sheet"], "cell": o.get("cell"), "after": o.get("after"), "note": o.get("note")} for o in ops],
    "skipped": action["skipped"], "apply_status": out["status"], "failed": out["failed"],
    "numbers_check_clean": check["clean"], "listed_after": left,
}
(OUT / "label_probe.json").write_text(json.dumps(result, indent=1, ensure_ascii=False, default=str), encoding="utf-8")
print(json.dumps(result, indent=1, ensure_ascii=False, default=str)[:3000])
