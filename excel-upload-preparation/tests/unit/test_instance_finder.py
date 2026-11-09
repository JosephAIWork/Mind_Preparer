"""The instance finder (app.instance_finder, app.finder_grouping): where Mind instances or a loop would help.

Two invented workbooks. The first is the instance builder's small model (a table with a text key column
read one row at a time by a macro that pastes each pass's block into the key's slot), with optional extra
sheets: rows that read a slot by a computed index, a block that reads its own key's slot, cells that add
every slot up, lookups keyed by the current row. The second mirrors the structure-check prototype's
control workbook: product sheets with their own assumptions, scenario blocks side by side, chained year
sheets, a plain time projection. The macro is read through a patched `read_vba` (an .xlsx carries none).
Nothing here comes from a client workbook."""
from __future__ import annotations

import json
import re
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

from app import instance_finder, model_map
from app.finder_grouping import _offset_slots, branch_key_name, branches_text, displacement_form, distinct_literals, macro_chain
from app.instance_finder import EMPTY_ASSISTANT, _switches, describe, find_instances, main, make_finder_advisor, merge_assistant, parse_assistant
from app.inventory import build_analysis
from app.loop_builder import build_loop_plan
from app.model_map import build_model_map
from local_usecases import usecase_or_skip

MACRO = '''
Sub RunRows()
Dim RunStart As Long: RunStart = Range("RunStart").Value
Dim RunEnd As Long: RunEnd = Range("RunEnd").Value
Worksheets("Runner").Range("H20:M21").ClearContents
For rowNo = RunStart To RunEnd
    Range("POLICY_ROW").Value = rowNo
    Application.Calculate
    grp = Application.WorksheetFunction.VLookup(rowNo, Range("Policies"), {k}, False)
    slotNo = Application.WorksheetFunction.Match(grp, Range("GroupList"), 0)
    ' Range("P").Offset(rowNo).Value = Range("Out1").Value2
    Range("ByGroupPaster").Offset(0, 2 * (slotNo - 1)).Value2 = Range("AggregateOutput").Value2
Next rowNo
End Sub
'''

MACRO_WITHOUT_LOOKUP = '''
Sub RunRows()
Dim RunStart As Long: RunStart = Range("RunStart").Value
Dim RunEnd As Long: RunEnd = Range("RunEnd").Value
Worksheets("Runner").Range("H20:M21").ClearContents
For rowNo = RunStart To RunEnd
    Range("POLICY_ROW").Value = rowNo
    Application.Calculate
    Range("ByGroupPaster").Offset(0, 2 * (slotNo - 1)).Value2 = Range("AggregateOutput").Value2
Next rowNo
End Sub
'''

MACRO_BY_COUNTER = '''
Sub RunRows()
Dim RunStart As Long: RunStart = Range("RunStart").Value
Dim RunEnd As Long: RunEnd = Range("RunEnd").Value
Worksheets("Runner").Range("H20:M21").ClearContents
For rowNo = RunStart To RunEnd
    Range("POLICY_ROW").Value = rowNo
    Application.Calculate
    Range("ByGroupPaster").Offset(0, 2 * (rowNo - 1)).Value2 = Range("AggregateOutput").Value2
Next rowNo
End Sub
'''

GROUPS = ["G1", "G1", "G1", "G2", "G2", "G3", "G3", "G3"]
REGIONS = ["North", "South"] * 4
LABEL = "SYNTHETIC_TEST_FIXTURE"


def _book(tmp_path: Path, *, groups: list | None = None, group_formula: str | None = None, stepping: bool = False, keyed: bool = False, adders: bool = False,
          adders_fill: bool = False, lookups: int = 0, mapped: bool = False, name: str = "model") -> Path:
    """The small model that runs one row of a table at a time. `group_formula`: the key column holds this
    formula (with {r} for the row) instead of texts. `stepping`: a sheet of identical OFFSET(Pick,
    COLUMN()-3, (ROW()-1)*2) cells, one row per slot. `keyed`: a block reading its own key's slot through
    MATCH on a list that holds the group texts. `adders`: cells adding every slot up (`adders_fill`: an 8-row block of them, big enough to be a fill). `lookups`: that many
    VLOOKUPs keyed by the current row's group. `mapped`: an alias of the current row, a key mapped through
    a table, 12 lookups by the mapped key and 8 IFs branching on it."""
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Inputs"
    data["A1"] = LABEL
    data.cell(2, 1, "Current row")
    for i in range(4):
        data.cell(2, 2 + i, f"=INDEX(Policies,POLICY_ROW,{i + 1})")
    groups = list(groups if groups is not None else GROUPS)
    cols = [("Policy", list(range(1, 9))), ("Group", groups), ("Premium", [100 * (i + 1) for i in range(8)]), ("Region", REGIONS)]
    for j, (head, values) in enumerate(cols):
        data.cell(4, 2 + j, head)
        for i, v in enumerate(values):
            if head == "Group" and group_formula:
                data.cell(5 + i, 2 + j, group_formula.format(r=5 + i))
            elif v is not None:
                data.cell(5 + i, 2 + j, v)
    calc = wb.create_sheet("Calc")
    calc["C2"], calc["D2"] = "Row to run", 7
    calc["B8"], calc["C8"] = "=Inputs!D2*2", "=Inputs!D2/2"
    run = wb.create_sheet("Runner")
    run["D2"], run["E2"], run["D3"], run["E3"] = "First row", 1, "Last row", 8
    run["B19"], run["C19"] = "Fee A", "Fee B"
    run["B20"], run["C20"] = "=Calc!B8+OFFSET(H20,0,0)", "=Calc!C8+OFFSET(I20,0,0)"
    run["B21"], run["C21"] = "=Calc!B8*2+OFFSET(H21,0,0)", "=Calc!C8*2+OFFSET(I21,0,0)"
    run["H20"], run["I21"], run["K20"] = 5, 6, 7  # left by an earlier run
    lists = wb.create_sheet("Lists")
    lists["B1"], lists["D1"], lists["F1"], lists["G1"], lists["H1"], lists["I1"] = "Groups", "Other list", "Group", "Rate", "Group", "Code"
    for i, g in enumerate(["G1", "G2", "G3", "G4"]):
        lists.cell(2 + i, 2, g)
        lists.cell(2 + i, 4, g)
        lists.cell(2 + i, 6, g)
        lists.cell(2 + i, 7, 0.01 * (i + 1))
        lists.cell(2 + i, 8, g)
        lists.cell(2 + i, 9, "ABCD"[i])
    names = {"RunStart": "Runner!$E$2", "RunEnd": "Runner!$E$3", "POLICY_ROW": "Calc!$D$2", "ByGroupPaster": "Runner!$H$20:$I$21", "AggregateOutput": "Runner!$B$20:$C$21",
             "Policies": "Inputs!$B$5:$E$12", "GroupList": "Lists!$B$2:$B$5", "OtherList": "Lists!$D$2:$D$5", "Rates": "Lists!$F$2:$G$5", "Map": "Lists!$H$2:$I$5",
             "Pick": "Runner!$G$20", "Paste": "Runner!$H$20"}
    if stepping:
        pick = wb.create_sheet("Pick")
        pick["A1"] = "Slot per row"
        for r in range(2, 5):
            for c in (3, 4):
                pick.cell(r, c, "=OFFSET(Pick,COLUMN()-3,(ROW()-1)*2)")
    if keyed:
        ky = wb.create_sheet("Keyed")
        ky["A1"] = "Own slot"
        for r in (2, 3):
            for c in (2, 3):
                ky.cell(r, c, "=Calc!B8+OFFSET(Paste,0,2*(MATCH(Inputs!$C$2,OtherList,0)-1))")
    if adders or adders_fill:
        ad = wb.create_sheet("Adders")
        ad["A1"] = "All slots"
        for r in range(2, 10) if adders_fill else (2, 3):
            for c in (2, 3):
                ad.cell(r, c, "=Runner!$H$20+Runner!$J$20+Runner!$L$20" if adders_fill else f"=Runner!H{18 + r}+Runner!J{18 + r}+Runner!L{18 + r}")
    if lookups:
        for i in range(lookups):
            calc.cell(12 + i, 2, "=VLOOKUP(Inputs!C2,Rates,2,0)")
    if mapped:
        calc["B3"] = "=Inputs!C2"  # an alias of the current row's group
        calc["B4"] = "=VLOOKUP(B3,Map,2,0)"  # the group mapped to a code
        names["Code"] = "Calc!$B$4"
        for i in range(12):
            calc.cell(30 + i, 2, "=VLOOKUP(Code,Rates,2,0)")
        for i in range(8):
            calc.cell(30 + i, 4, '=IF(OR(Code="A",Code="B"),0.1,0.2)')
    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    p = tmp_path / f"{name}.xlsx"
    wb.save(p)
    wb.close()
    return p


def _control(tmp_path: Path, *, scen2: bool = False, plain: bool = False, groups: int = 0, name: str = "control") -> Path:
    """The control workbook: three product sheets with their own assumptions (one formula differs by a
    typed-in number, one by its logic) and a Total that adds them; three scenario blocks side by side (and
    a second sheet of them with `scen2`); chained year sheets; a plain time projection. `plain`: a big
    filled column that joins nothing and a small fill. `groups`: that many pairs of copied sheets."""
    wb = openpyxl.Workbook()
    inp = wb.active
    inp.title = "Inputs"
    inp["A1"], inp["B1"], inp["C1"], inp["D1"] = LABEL, "S1", "S2", "S3"
    inp["A2"], inp["B2"], inp["C2"], inp["D2"] = "Rate", 0.02, 0.03, 0.01
    assump = {"A": (100, 10, 0.05), "B": (200, 6, 0.04), "C": (50, 16, 0.07)}
    for p, (pol, prem, lapse) in assump.items():
        ws = wb.create_sheet(f"Prod_{p}")
        ws["A1"], ws["B1"] = "Policies", pol
        ws["A2"], ws["B2"] = "Premium", prem
        ws["A3"], ws["B3"] = "Lapse", lapse
        ws["A5"] = "Year"
        ws["A6"], ws["A7"], ws["A8"], ws["A9"] = "In force", "Premium", "Expense", "Profit"
        for y in range(1, 21):
            c, prev = L(y + 1), L(y)
            ws[f"{c}5"] = y
            ws[f"{c}6"] = "=$B$1" if y == 1 else f"={prev}6*(1-$B$3)"
            ws[f"{c}7"] = f"={c}6*$B$2"
            ws[f"{c}8"] = f"={c}7*0.1"
            ws[f"{c}9"] = f"={c}7-{c}8"
    wb["Prod_B"]["B8"] = "=B7*0.12"  # a typed-in number differs in one copy
    wb["Prod_C"]["C8"] = "=C7*0.12+1"  # the logic differs in one copy
    tot = wb.create_sheet("Total")
    tot["A1"] = "Profit"
    for y in range(1, 21):
        c = L(y + 1)
        tot[f"{c}1"] = f"=Prod_A!{c}9+Prod_B!{c}9+Prod_C!{c}9"
    sc = wb.create_sheet("Scen")
    for k in range(3):
        c0 = 1 + k * 5
        A, B, C, D = (L(c0 + i) for i in range(4))
        sc[f"{A}1"] = f"Scenario {k + 1}"
        sc[f"{A}2"], sc[f"{B}2"], sc[f"{C}2"], sc[f"{D}2"] = "Year", "Rate", "Fund", "Interest"
        for t in range(1, 11):
            r = t + 2
            sc[f"{A}{r}"] = t
            sc[f"{B}{r}"] = f"=Inputs!{L(2 + k)}$2"
            sc[f"{C}{r}"] = "=1000" if t == 1 else f"={C}{r - 1}+{D}{r - 1}"
            sc[f"{D}{r}"] = f"={C}{r}*{B}{r}"
    sc["T2"] = "Total fund"  # well right of the blocks: a column next to them would cut the third copy off
    for t in range(1, 11):
        sc[f"T{t + 2}"] = f"=C{t + 2}+H{t + 2}+M{t + 2}"
    if scen2:  # another sheet, another layout, the same three scenario rates read block by block
        s2 = wb.create_sheet("Scen2")
        for k in range(3):
            c0 = 2 + k * 4
            B, C, D = (L(c0 + i) for i in range(3))
            s2[f"{B}1"], s2[f"{C}1"], s2[f"{D}1"] = f"Rate {k + 1}", "Fund", "Gain"
            for t in range(1, 9):
                r = t + 1
                s2[f"{B}{r}"] = f"=Inputs!{L(2 + k)}$2*100"
                s2[f"{C}{r}"] = "=1000" if t == 1 else f"={C}{r - 1}*(1+{B}{r}/100)"
                s2[f"{D}{r}"] = f"={C}{r}-1000"
    for i, yr in enumerate((2024, 2025, 2026)):
        ws = wb.create_sheet(f"Y{yr}")
        ws["A1"], ws["A2"], ws["A3"], ws["A4"] = "Opening", "Premium", "Claims", "Closing"
        ws["B1"] = 500 if i == 0 else f"=Y{yr - 1}!B4"
        ws["B2"], ws["B3"] = 120 + 10 * i, 80 + 5 * i
        ws["B4"] = "=B1+B2-B3"
        for m in range(1, 13):
            c = L(m + 2)
            ws[f"{c}1"] = "=B1" if m == 1 else f"={L(m + 1)}4"
            ws[f"{c}2"] = "=$B$2/12"
            ws[f"{c}3"] = "=$B$3/12"
            ws[f"{c}4"] = f"={c}1+{c}2-{c}3"
    pj = wb.create_sheet("Proj")
    pj["A1"], pj["A2"], pj["A3"] = "Month", "Reserve", "Release"
    for m in range(1, 61):
        c = L(m + 1)
        pj[f"{c}1"] = m
        pj[f"{c}2"] = "=1000" if m == 1 else f"={L(m)}2-{L(m)}3"
        pj[f"{c}3"] = f"={c}2*0.02"
    if plain:
        pl = wb.create_sheet("Plain")
        pl["A1"], pl["B1"] = "Item", "Length"
        for r in range(2, 602):
            pl[f"A{r}"] = f"Item {r}"
            pl[f"B{r}"] = f"=LEN(A{r})*2"
        sm = wb.create_sheet("Small")
        for c in range(2, 10):
            sm.cell(1, c, f"Item {c}")
            for r in (2, 3, 4):
                sm.cell(r, c, f"={L(c)}1&\"-\"&ROW()")
    for g in range(groups):
        for part in ("A", "B"):
            ws = wb.create_sheet(f"Grp{g + 1}_{part}")
            ws["A1"] = f"Group {g + 1}"
            for r in range(2, 14):
                ws[f"A{r}"] = r
                ws[f"B{r}"] = f"=A{r}+{g + 1}"
    p = tmp_path / f"{name}.xlsx"
    wb.save(p)
    wb.close()
    return p


def _add_cached_values(path: Path, sheet_index: int, values: dict[str, str]) -> None:
    """Give formula cells of a saved xlsx the cached text value Excel would have stored (openpyxl writes
    none): the app's fast reader then serves them as the sheet's `mind_ready_cached`."""
    tmp = path.with_suffix(".tmp")
    part = f"xl/worksheets/sheet{sheet_index}.xml"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == part:
                xml = data.decode("utf-8")
                for cell, value in values.items():
                    xml, n = re.subn(rf'<c r="{cell}"([^>]*)><f>(.*?)</f><v></v></c>', lambda m: f'<c r="{cell}"{m.group(1)} t="str"><f>{m.group(2)}</f><v>{value}</v></c>', xml, count=1)
                    assert n == 1, cell
                data = xml.encode("utf-8")
            dst.writestr(item, data)
    tmp.replace(path)


def _run(path: Path, tmp_path: Path, vba: str | None, monkeypatch, **kw) -> dict:
    """The finder on a workbook, the macro read through the patched read_vba (None: no macro)."""
    if vba is not None:
        monkeypatch.setattr(model_map, "read_vba", lambda _p: [("Module1", vba)])
    else:
        monkeypatch.setattr(model_map, "read_vba", lambda _p: [])
    analysis = build_analysis(path, tmp_path / f"w_{path.stem}", path.stem)
    mp = build_model_map(analysis)
    return find_instances(analysis, mp, loop_plan=build_loop_plan(analysis, mp), **kw)


def _top(result: dict) -> dict:
    assert result["dimensions"], describe(result)
    return result["dimensions"][0]


def _lines(result: dict, kind: str) -> list[dict]:
    return [e for d in result["dimensions"] for e in d["evidence"] if e["kind"] == kind]


def _strip(d: dict) -> dict:
    return {k: v for k, v in d.items() if not k.startswith("assistant_")}


def test_the_table_key_column_is_found_blind_from_the_macro(tmp_path, monkeypatch):
    res = _run(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch)
    top = _top(res)
    assert (top["key"]["column"], top["key"]["header"], top["key"]["distinct"], top["key"]["rows"], top["key"]["type"]) == ("C", "Group", 3, 8, "text")
    assert {"macro", "list"} <= set(top["kinds"]) and (top["copies"], top["biggest_copy"]) == (3, 3) and top["label"] == "Group" and top["label_from"] == "header"
    assert top["headline"] == "Column C (Group) splits the 8 rows into 3 groups; the macro already groups its output by it."
    macro = next(e for e in top["evidence"] if e["kind"] == "macro")
    assert "column 2 of Policies (VLookup)" in macro["text"] and "GroupList (Match)" in macro["text"] and "slot n of ByGroupPaster (Offset(0, 2*(n-1)))" in macro["text"]
    assert any("slotNo" in v for v in macro["vba"]) and not any(v.lstrip().startswith("'") for v in macro["vba"])
    assert [e["kind"] for e in top["evidence"]] == ["macro", "list", "counts"]  # the counts line is the weakest: last, cut first
    lst = next(e for e in top["evidence"] if e["kind"] == "list")
    assert "GroupList" in lst["text"] and lst["count"] == 4
    assert res["builder"] == {"column": "C", "header": "Group", "chosen_by": "macro", "candidates": ["C", "E"], "matches_top": True, "source": "rule"}
    region = next(d for d in res["dimensions"] if d["key"] and d["key"]["column"] == "E")
    assert region["rank"] > top["rank"] and "has 2 values" in region["headline"] and region["fit"]["instances"]["fits"] == "partly"
    assert "two copies: an IF or a loop of 2 does the same" in region["fit"]["instances"]["reasons"]
    assert res["counts"]["dimensions"] == 2 and res["counts"]["formulas"] > 0 and res["schema"] == "instance-finder/1" and res["skipped"] == []
    # the fixed-slot readers of the paste area (the block the macro adds into) are classified, nothing is "already the sum"
    assert top["readers"]["fixed_slot"]["count"] == 4 and top["readers"]["all_slots"]["count"] == 0 and "already the sum" not in json.dumps(res)


def test_the_macro_offset_is_parsed_with_balanced_parentheses():
    vba = '''Sub X()
grp = Application.WorksheetFunction.VLookup(rowNo, Range("Policies"), 2, False)
slotNo = Application.WorksheetFunction.Match(grp, Range("GroupList"), 0)
Range("ByGroupPaster").Offset(0, 2 * (slotNo - 1)).Value2 = Range("AggregateOutput").Value2
Range("P").Offset(0, 10 * (slotNo - 1)).Value2 = Range("Q").Value2
Worksheets("S").Range("R").Offset((slotNo - 1) * 10, 0).Value2 = Range("Q").Value2
' Range("P").Offset(rowNo).Value = Range("Out1").Value2
Range("OutTotal").Offset(rowNo).Value = Range("Out1").Value2
If Application.WorksheetFunction.VLookup(grp, Range("RunFlags"), 2, False) <> "Yes" Then GoTo Skip
End Sub'''
    slots = _offset_slots(vba)
    assert [(s["anchor"], s["w"], s["axis"], s["var"]) for s in slots] == [("ByGroupPaster", 2, "columns", "slotNo"), ("P", 10, "columns", "slotNo"), ("R", 10, "rows", "slotNo")]
    assert slots[2]["sheet"] == "S" and _offset_slots(vba, {"slotno"}) == slots and _offset_slots(vba, {"other"}) == []
    assert _offset_slots("' Range(\"P\").Offset(0, 10 * (slotNo - 1)).Value2 = 1") == []
    names = model_map._Names([{"name": n, "value": v, "local_sheet_id": None} for n, v in (("Policies", "Inputs!$B$5:$E$12"), ("GroupList", "Lists!$B$2:$B$5"),
                                                                                           ("ByGroupPaster", "Runner!$H$20:$I$21"), ("RunFlags", "Lists!$F$2:$G$5"))])
    table = {"sheet": "Inputs", "r1": 5, "r2": 12, "c1": 2, "c2": 5}
    chain = macro_chain(vba, "rowNo", table, names)
    assert (chain["col_k"], chain["col"], chain["rng_name"], chain["list"], chain["list_name"], chain["var1"], chain["var2"]) == (2, 3, "Policies", "Lists!B2:B5", "GroupList", "grp", "slotNo")
    assert (chain["anchor"], chain["anchor_name"], chain["w"], chain["axis"]) == ("Runner!H20:I21", "ByGroupPaster", 2, "columns")
    assert chain["keyed_tables"] == [{"name": "RunFlags", "ref": "Lists!F2:G5", "column": 2}] and len(chain["lines"]) == 4 and all(len(x) <= 160 for x in chain["lines"])
    assert macro_chain("Sub Y()\n' nothing here\nEnd Sub", "rowNo", table, names) is None
    # the displacement idioms the finder reads inside OFFSET / INDEX
    assert displacement_form("(ROW()-21)*10") == {"form": "row", "w": 10, "n": 21, "rng": None, "step_axis": "rows", "key": None, "list": None, "text": "(ROW()-21)*10"}
    assert displacement_form("COLUMN()-14")["w"] == 1 and displacement_form("10*(COUNTA($B$9:$B9)-1)")["step_axis"] == "rows"
    m = displacement_form("2*(MATCH($A$1, List!$A$1:$A$8, 0)-1)")
    assert (m["form"], m["w"], m["key"], m["list"]) == ("match", 2, "$A$1", "List!$A$1:$A$8") and displacement_form("0") is None and displacement_form("Slot") is None
    assert displacement_form("ROW()-21*10") is None and displacement_form("(COLUMN()-14)")["w"] == 1  # balanced parentheses: one parser with finder_patterns


def test_without_the_macro_lookup_the_column_is_found_through_keyed_lookups(tmp_path, monkeypatch):
    res = _run(_book(tmp_path, lookups=6), tmp_path, MACRO_WITHOUT_LOOKUP, monkeypatch)
    top = _top(res)
    assert top["key"]["column"] == "C" and "lookups" in top["kinds"] and "macro" not in top["kinds"] and top["key"]["keyed_formulas"] == 6 and "Rates" in top["key"]["keyed_tables"]
    assert "6 formulas look a value up by the current row's Group in Rates" in next(e for e in top["evidence"] if e["kind"] == "lookups")["text"]
    assert res["builder"]["chosen_by"] == "heuristic" and res["builder"]["matches_top"] is True
    with_macro = _run(_book(tmp_path, lookups=6, name="withmacro"), tmp_path, MACRO.format(k=2), monkeypatch)
    assert with_macro["dimensions"][0]["score"] > top["score"] and "macro" in with_macro["dimensions"][0]["kinds"]


def test_lookups_through_a_mapped_key_and_branches_count_as_evidence(tmp_path, monkeypatch):
    res = _run(_book(tmp_path, mapped=True), tmp_path, MACRO.format(k=2), monkeypatch)
    top = _top(res)
    k = top["key"]
    assert k["column"] == "C" and k["mapped_through"] == [{"cell": "Calc!B4", "name": "Code", "table": "Map"}]
    assert k["keyed_formulas"] >= 12 and k["branch_formulas"] == 8 and k["branch_literals_count"] == 2 and k["literals_are_values"] == 2
    assert "through the mapping table Map" in next(e for e in top["evidence"] if e["kind"] == "lookups")["text"]
    branches = next(e for e in top["evidence"] if e["kind"] == "branches")
    assert branches["count"] == 8 and 'IF(Code="' in branches["text"] and "on 2 of its values" in branches["text"] and len(branches["cells"]) == 6
    assert {"macro", "lookups", "branches", "list"} <= set(top["kinds"])
    assert top["edits"]["branches"] == {"count": 8, "action": "move the values chosen by these IFs into a table with one row per copy"} and top["cells_changed"] >= 8


def test_the_branches_line_names_the_compared_key_and_counts_literals_case_blind(tmp_path, monkeypatch):
    """Excel's = on text is case-blind: "a" and "A" are one literal and "g1" is the value G1. The line names
    the key most formulas compare (a name, never an address) and says how many compared texts are not values."""
    p = _book(tmp_path, mapped=True)
    wb = openpyxl.load_workbook(p)
    calc = wb["Calc"]
    for i in range(6):
        calc.cell(50 + i, 2, '=IF(Inputs!$C$2="g1",1,0)')  # the current-row cell by address, a value in another case
    for i in range(5):
        calc.cell(50 + i, 4, '=IF(Code="a",1,0)')  # the literal "A" in another case
    for i in range(4):
        calc.cell(50 + i, 6, '=IF(Code="Z",1,0)')  # a text that is no value of the key or of its mapping
    wb.save(p)
    wb.close()
    top = _top(_run(p, tmp_path, MACRO.format(k=2), monkeypatch))
    k = top["key"]
    assert (k["branch_formulas"], k["branch_literals_count"], k["literals_are_values"]) == (8 + 6 + 5 + 4, 4, 3)
    branches = next(e for e in top["evidence"] if e["kind"] == "branches")
    assert branches["text"] == '23 formulas choose a value with IF(Code="…") on 3 of its values; 1 other text compared is not among its values'
    # the helpers on their own
    assert distinct_literals(Counter({"Open": 3, "open": 2, "Closed": 1, "CLOSED": 1})) == ["Open", "Closed"]
    assert branch_key_name(Counter({"Inputs!$C$2": 6, "Code": 3}), [], "Group") == "Code"
    assert branch_key_name(Counter({"Inputs!$C$2": 6, "'My sheet'!C7": 2}), [{"name": "Code"}], "Group") == "Code"
    assert branch_key_name(Counter({"$C$2": 6}), [{"name": None}], "Group") == "Group"
    assert branches_text(5, "Flag", 2, 0) == '5 formulas choose a value with IF(Flag="…") comparing 2 texts that are not among its values'
    assert branches_text(5, "Flag", 1, 1) == '5 formulas choose a value with IF(Flag="…") on 1 of its values'
    assert branches_text(9, "Sex", 3, 1) == '9 formulas choose a value with IF(Sex="…") on 1 of its values; 2 other texts compared are not among its values'


def test_a_key_column_filled_by_formulas_uses_the_cached_values(tmp_path, monkeypatch):
    formula = '=IF(D{r}>400,"X","Y")'
    res = _run(_book(tmp_path, group_formula=formula), tmp_path, MACRO.format(k=2), monkeypatch)
    assert not any(d["key"] and d["key"]["column"] == "C" for d in res["dimensions"])
    assert any("column C" in s and "formulas without saved values" in s for s in res["skipped"])
    path = _book(tmp_path, group_formula=formula, name="cached")
    _add_cached_values(path, 1, {f"C{r}": ("X" if 100 * (r - 4) > 400 else "Y") for r in range(5, 13)})
    res = _run(path, tmp_path, MACRO.format(k=2), monkeypatch)
    top = _top(res)
    assert (top["key"]["column"], top["key"]["distinct"], top["key"]["counts"], top["key"]["type"]) == ("C", 2, {"X": 4, "Y": 4}, "text") and res["skipped"] == []
    assert "macro" in top["kinds"] and res["builder"]["matches_top"] is True


def test_blocks_stepping_through_one_input_row_merge_into_one_dimension(tmp_path, monkeypatch):
    res = _run(_control(tmp_path, scen2=True), tmp_path, None, monkeypatch)
    scen = [d for d in res["dimensions"] if any(b["sheet"] in ("Scen", "Scen2") for b in d["blocks"])]
    assert len(scen) == 1, [d["label"] for d in res["dimensions"]]
    d = scen[0]
    assert sorted((b["sheet"], b["copies"], b["per_copy"], b["axis"]) for b in d["blocks"]) == [("Scen", 3, 30, "columns"), ("Scen2", 3, 24, "columns")]
    assert d["copies"] == 3 and d["formulas_saved"] == 2 * 30 + 2 * 24 and "bands" in d["kinds"] and "sheets" not in d["kinds"]
    assert d["fit"]["loop"]["fits"] == "yes" and d["fit"]["leave"]["fits"] == "yes" and d["fit"]["instances"]["fits"] == "partly"
    assert all("Inputs!B2:D2" in e["text"] for e in d["evidence"] if e["kind"] == "bands") and d["edits"]["logic_count"] == 0
    assert d["headline"].startswith("Sheet Scen: 3 blocks of 5 columns compute the same")  # the block is 4 columns and the gap: 5 positions per copy


def test_rows_stepping_by_computed_index_join_the_macro_dimension(tmp_path, monkeypatch):
    res = _run(_book(tmp_path, stepping=True, keyed=True, adders=True), tmp_path, MACRO.format(k=2), monkeypatch)
    top = _top(res)
    assert top["key"]["column"] == "C" and {"macro", "stepping", "keyed", "list"} <= set(top["kinds"]) and res["counts"]["dimensions"] == 2
    rd = top["readers"]
    assert {k: v["count"] for k, v in rd.items()} == {"all_slots": 4, "fixed_slot": 4, "computed_slot": 6, "keyed_slot": 4, "follow_adders": 0}
    assert rd["keyed_slot"]["text"] == "4 cells read the current key's slot and add each pass to it: in a copy it is the copy's own block, and MM_RESULT does the adding; nothing to write"
    lookups = next(e for e in top["evidence"] if e["kind"] == "lookups")  # the keyed block's MATCH on the key is a keyed lookup by the rule: said, not counted twice
    assert lookups["text"] == "4 formulas look a value up by the current row's Group in OtherList; 4 of them are the block that reads its own key's slot" and lookups["count"] == 4
    assert [e["kind"] for e in top["evidence"]] == ["macro", "lookups", "stepping", "keyed", "list", "counts"]
    assert rd["all_slots"]["example"] == "Adders!B2" and "move the adding row to the main workbook" in rd["all_slots"]["text"] and "already the sum" not in json.dumps(res)
    assert rd["computed_slot"]["text"] == "6 cells in 3 rows each read one slot: one row per copy in the main workbook (a plain link per copy, or MM_INSTANCE)"
    step = next(e for e in top["evidence"] if e["kind"] == "stepping")
    assert step["text"] == "Pick rows 2-4: 3 rows each read one 2-column slot of Runner!H20:M21 (OFFSET with (ROW()-1)*2)" and (step["count"], step["per_copy"]) == (3, 2)
    keyed = next(e for e in top["evidence"] if e["kind"] == "keyed")
    assert keyed["text"].startswith("Keyed!B2:C3 reads the current key's slot and adds the pass to it") and keyed["count"] == 4
    lst = next(e for e in top["evidence"] if e["kind"] == "list")
    assert "GroupList" in lst["text"] and "Lists!D2:D5" in lst["text"] and "holds the same values" in lst["text"]
    assert top["formulas_saved"] == 6 and top["cells_changed"] == 8 and top["cross_reads"] == [] and top["blocks"] == []
    assert any("4 adding cells move to the main workbook" in r for r in top["fit"]["instances"]["reasons"]) and top["fit"]["loop"]["fits"] == "partly"
    assert res["ignored"]["filled_areas"] == 0


def test_a_paste_area_cut_by_the_loop_counter_joins_no_dimension_and_is_noted(tmp_path, monkeypatch):
    """The macro pastes slot n by the loop counter, not by a key (no VLookup, no Match): the cells that read the
    paste area join no column, list or stepping block. With an adders block big enough to be a fill the component
    is the slots area and that fill alone -- no dimension of its own (an IndexError before) -- and a note says what
    reads the area; with a small adders block the same note: nothing is dropped in silence."""
    for name, fill in (("fill", True), ("small", False)):
        res = _run(_book(tmp_path, adders=True, adders_fill=fill, name=name), tmp_path, MACRO_BY_COUNTER, monkeypatch)
        json.dumps(res)
        assert res["dimensions"] and all(d["readers"] is None for d in res["dimensions"]) and not any("Adders" in d["label"] for d in res["dimensions"])
        n_add = 16 if fill else 4
        assert [x for x in res["skipped"] if "paste area" in x] == [f"{n_add + 4} cells read the macro's paste area but no key column, list or stepping block joins them "
                                                                      f"({n_add} add every slot up; 4 read one slot): they are in this note only"]
        assert res["macro_loops"][0]["handled_by"] == "the Mind loop step"


def test_switch_readers_are_counted_by_whole_address_or_name_on_the_right_sheet():
    """A switch at B3 is read by $B$3, B$3 or B3 on its own sheet, by Sheet!B3 from anywhere and by its defined
    name -- never by B30, AB3 or another sheet's B3 (a bare substring count gave hundreds of false readers)."""
    mm = {"run_settings": {"switches": [{"cell": "Control!B3", "name": "VAL"}, {"cell": "Control!$C$4", "name": None}]}}
    wb0 = {"formulas": [{"sheet": "Control", "cell": "D1", "formula": "=$B$3*2"}, {"sheet": "Control", "cell": "D2", "formula": "=B30+1"},
                        {"sheet": "Control", "cell": "D3", "formula": "=AB3"}, {"sheet": "Other", "cell": "D4", "formula": "=B3"},
                        {"sheet": "Other", "cell": "D5", "formula": "=Control!B3"}, {"sheet": "Other", "cell": "D6", "formula": "='Control'!$B$3"},
                        {"sheet": "Other", "cell": "D7", "formula": "=VAL*2"}, {"sheet": "Other", "cell": "D8", "formula": "=VALUE(1)"},
                        {"sheet": "Control", "cell": "D9", "formula": "=B$3+C4"}, {"sheet": "Other", "cell": "E1", "formula": "=Other!B3"}]}
    assert _switches(mm, wb0) == [{"cell": "Control!B3", "name": "VAL", "readers": 5}, {"cell": "Control!$C$4", "name": None, "readers": 1}]
    assert _switches({"run_settings": {"switches": []}}, wb0) == [] and _switches({}, {"formulas": []}) == []


def test_sheet_copies_with_own_assumptions_list_the_edits(tmp_path, monkeypatch):
    res = _run(_control(tmp_path), tmp_path, None, monkeypatch)
    prod = next(d for d in res["dimensions"] if "sheets" in d["kinds"])
    assert prod["blocks"] == [{"sheet": "Prod_A", "axis": "sheets", "where": "Prod_A, Prod_B, Prod_C", "copies": 3, "per_copy": 80, "computed": None, "diff": prod["blocks"][0]["diff"], "chain_share": 0.0}]
    assert prod["blocks"][0]["diff"] == {"identical": 78, "reads a different input": 0, "differs only in a typed-in number": 1, "different logic": 1}
    ed = prod["edits"]
    assert ed["numbers_count"] == 3 and [x["cell"] for x in ed["numbers"]] == ["Prod_A!B1", "Prod_A!B2", "Prod_A!B3"] and ed["numbers"][0]["values_per_copy"] == [100, 200, 50]
    assert ed["numbers"][0]["action"] == "move into a table row per copy"
    assert ed["formula_numbers_count"] == 1 and ed["formula_numbers"][0]["cell"] == "Prod_A!B8" and ed["formula_numbers"][0]["literals_per_copy"] == [["0.1"], ["0.12"], ["0.1"]]
    assert ed["logic_count"] == 1 and ed["logic"][0] == {"cell": "Prod_A!C8", "formulas_per_copy": ["=C7*0.1", "=C7*0.1", "=C7*0.12+1"], "action": "IF on the copy's key"}
    assert ed["size_count"] == 0 and ed["branches"] is None
    sheets = next(e for e in prod["evidence"] if e["kind"] == "sheets")
    assert "20 cells on other sheets add them up (Total!B1)" in sheets["text"] and prod["copies"] == 3 and prod["formulas_saved"] == 160 and prod["cells_changed"] == 5
    assert prod["fit"]["instances"]["fits"] == "partly" and prod["fit"]["loop"]["fits"] == "partly" and "IF(MM_DIMINDEX" in " ".join(prod["fit"]["loop"]["reasons"])
    assert prod["headline"] == "Sheets Prod_A, Prod_B, Prod_C are copies (98% the same): one sheet and 3 instances, 1 cell kept per copy."


def test_year_sheets_are_a_chain_and_not_a_dimension(tmp_path, monkeypatch):
    res = _run(_control(tmp_path), tmp_path, None, monkeypatch)
    chain = next(c for c in res["chains"] if c["label"] == "Y2024, Y2025, Y2026")
    assert (chain["copies"], chain["why"], chain["cells"], chain["chain_share"]) == (3, "each sheet reads the sheet before it", ["Y2025!B1"], 1.0)
    assert not any("Y2024" in json.dumps(d) for d in res["dimensions"]) and res["counts"]["chains"] == len(res["chains"]) >= 1
    assert "Set aside (chains): Y2024, Y2025, Y2026 (3 copies: each sheet reads the sheet before it)" in describe(res)


def test_time_axes_and_filled_areas_are_counted_not_reported(tmp_path, monkeypatch):
    res = _run(_control(tmp_path, plain=True), tmp_path, None, monkeypatch)
    ig = res["ignored"]
    assert ig["time_axes"] >= 1 and ig["filled_areas"] >= 1 and ig["small_fills"] >= 1 and res["counts"]["time_axes"] == ig["time_axes"]
    assert res["ignored_detail"]["filled_areas"][0] == {"sheet": "Plain", "where": "rows 2-601", "cells": 600}
    assert not any(b["axis"] == "fill" for d in res["dimensions"] for b in d["blocks"]) and not any("Plain" in d["label"] or "Proj" in d["label"] for d in res["dimensions"])
    assert "Filled areas that join nothing: Plain!rows 2-601 (600 cells)" in describe(res)


def test_macro_loops_are_listed_as_handled_by_the_loop_step(tmp_path, monkeypatch):
    res = _run(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch)
    assert res["macro_loops"] == [{"name": "POLICY_ROW", "driver": "Calc!D2", "kind": "counter", "status": "ready", "first": 1, "last": 8, "handled_by": "the Mind loop step"}]
    assert "Macro loops the Mind loop step handles: POLICY_ROW (counter, ready)" in describe(res)
    assert res["switches"] == [] and res["file"] == "model.xlsx" and res["version"]


def test_result_is_json_safe_bounded_and_the_same_twice(tmp_path, monkeypatch):
    path = _control(tmp_path, groups=14)
    res = _run(path, tmp_path, None, monkeypatch)
    json.dumps(res)  # no default= needed
    again = _run(path, tmp_path, None, monkeypatch)
    assert {k: v for k, v in res.items() if k != "counts"} == {k: v for k, v in again.items() if k != "counts"}
    assert len(res["dimensions"]) == 12 and res["counts"]["dimensions"] >= 16 and res["truncated"]["dimensions"] is True
    assert [d["rank"] for d in res["dimensions"]] == list(range(1, 13)) and all(d["id"] == f"dim-{d['rank']}" for d in res["dimensions"])
    scores = [d["score"] for d in res["dimensions"]]
    assert scores == sorted(scores, reverse=True)
    assert all(len(d["evidence"]) <= 8 and all(len(e.get("cells") or []) <= 6 for e in d["evidence"]) and len(d["key"]["counts"]) <= 13 if d["key"] else True for d in res["dimensions"])
    assert set(res) == {"schema", "version", "file", "counts", "dimensions", "chains", "macro_loops", "switches", "builder", "ignored", "ignored_detail", "skipped", "truncated", "assistant"}
    assert set(res["dimensions"][0]) == {"id", "rank", "score", "label", "label_from", "headline", "kinds", "key", "copies", "biggest_copy", "formulas_saved", "cells_changed", "evidence",
                                         "blocks", "readers", "cross_reads", "edits", "fit", "assistant_rank", "assistant_text", "assistant_edits"}
    assert res["assistant"] == EMPTY_ASSISTANT and res["builder"]["column"] is None and res["builder"]["matches_top"] is None and res["macro_loops"] == []


def test_describe_says_whether_the_builder_agrees(tmp_path, monkeypatch):
    res = _run(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch)
    text = describe(res)
    assert text.startswith("#1 Group: Column C (Group) splits the 8 rows into 3 groups; the macro already groups its output by it.")
    assert "Builder's key: C Group (macro) (the builder's rule) -- the same as the top dimension" in text and "   Instances: yes" in text and "   A loop: " in text
    other = json.loads(json.dumps(res))
    other["builder"].update(column="E", header="Region", chosen_by="heuristic", matches_top=False, source="plan")
    assert "Builder's key: E Region (heuristic) -- DIFFERENT: the top dimension is column C (Group)" in describe(other)
    empty = {**other, "dimensions": [], "chains": [], "macro_loops": [], "builder": None, "skipped": ["a note"], "ignored_detail": {"filled_areas": []}}
    assert describe(empty).startswith("Nothing repeats enough to split: leave the model as it is.") and "Note: a note" in describe(empty)


def test_assistant_pass_merges_ranking_explanations_edits_and_flags_extras(tmp_path, monkeypatch):
    plain = _run(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch)
    answer = ('Here is my view.\n{"ranking": [2, 1], "explanations": {"1": "The group column splits the rows.", "2": "Two regions only."}, '
              '"edits": {"2": ["move the two regions into a table"]}, "extra": [{"headline": "Sheet X repeats", "why": "looks alike", "cells": ["Inputs!B2", "nonsense"]}]}\nThanks.')
    calls: list = []

    def fake(messages, system, **kw):
        calls.append((messages, system, kw))
        return {"available": True, "text": answer, "message": None}

    res = _run(_book(tmp_path, name="asked"), tmp_path, MACRO.format(k=2), monkeypatch, assistant=True, complete=fake)
    a = res["assistant"]
    assert a["asked"] is True and a["available"] is True and a["ranking"] == [2, 1] and a["message"] is None
    assert a["explanations"] == {"1": "The group column splits the rows.", "2": "Two regions only."} and a["edits"] == {"2": ["move the two regions into a table"]}
    assert a["extra"] == [{"headline": "Sheet X repeats", "why": "looks alike", "cells": ["Inputs!B2"], "flag": "suggested by the assistant, not checked"}]
    assert a["text"] == "The group column splits the rows. Two regions only."
    d1, d2 = res["dimensions"]
    assert (d1["assistant_rank"], d2["assistant_rank"]) == (2, 1) and d1["assistant_text"] == "The group column splits the rows." and d2["assistant_edits"] == ["move the two regions into a table"]
    assert [_strip(d) for d in res["dimensions"]] == [_strip(d) for d in plain["dimensions"]]  # rank, score, evidence, fit untouched
    assert len(calls) == 1 and calls[0][1] == instance_finder.SYSTEM_PROMPT and "### Dimension 1" in calls[0][0][0]["content"] and "### Builder's choice" in calls[0][0][0]["content"]
    assert calls[0][2]["max_tokens"] == 2000 and calls[0][2]["timeout"] == 90 and len(calls[0][0][0]["content"]) <= 12_000
    # the merge alone, on a hand-made block: a ranking naming an unknown dimension is dropped
    block = parse_assistant('{"ranking": [1, 9], "explanations": {"3": "x"}, "edits": {"1": ["a", 5]}}', plain)
    assert block["ranking"] is None and block["explanations"] == {} and block["edits"] == {"1": ["a"]} and block["asked"] is True
    merged = merge_assistant(plain, block)
    assert merged["assistant"]["edits"] == {"1": ["a"]} and merged["dimensions"][0]["assistant_edits"] == ["a"] and merged["dimensions"][0]["assistant_rank"] is None and plain["assistant"] == EMPTY_ASSISTANT


def test_assistant_failures_leave_the_deterministic_result(tmp_path, monkeypatch):
    plain = _run(_book(tmp_path), tmp_path, MACRO.format(k=2), monkeypatch)

    def raising(*_a, **_k):
        raise RuntimeError("boom")

    no_key = "No AI key is saved yet: paste one under Advanced > Assistant."
    cases = [(raising, True, "Request failed: RuntimeError"), (lambda *a, **k: {"available": True, "text": "No JSON here, just prose.", "message": None}, True, "The assistant's answer was not the JSON asked for"),
             (lambda *a, **k: {"available": False, "text": None, "message": no_key}, False, no_key), (lambda *a, **k: "a string", True, "The assistant gave no usable answer"),
             (lambda *a, **k: None, True, "The assistant gave no usable answer"),
             (lambda *a, **k: {"available": True, "text": '{"notes": "nothing useful"}', "message": None}, True, "The assistant gave no usable answer")]
    for i, (fake, available, message) in enumerate(cases):
        res = _run(_book(tmp_path, name=f"m{i}"), tmp_path, MACRO.format(k=2), monkeypatch, assistant=True, complete=fake)
        a = res["assistant"]
        assert (a["asked"], a["available"], a["message"], a["ranking"], a["extra"], a["text"]) == (True, available, message, None, [], None), i
        assert all(d["assistant_rank"] is None and d["assistant_text"] is None and d["assistant_edits"] == [] for d in res["dimensions"])
        assert [_strip(d) for d in res["dimensions"]] == [_strip(d) for d in plain["dimensions"]]
    # no key: the advisor is None (patched on the finder module, where the name is bound), the result says so
    monkeypatch.setattr(instance_finder, "llm_available", lambda: False)
    assert make_finder_advisor() is None
    res = _run(_book(tmp_path, name="nokey"), tmp_path, MACRO.format(k=2), monkeypatch, assistant=True)
    assert res["assistant"]["asked"] is True and res["assistant"]["available"] is False and res["assistant"]["message"] == no_key
    # a key but a gateway that fails: chat_completion patched on the finder module
    monkeypatch.setattr(instance_finder, "llm_available", lambda: True)
    monkeypatch.setattr(instance_finder, "chat_completion", lambda *a, **k: {"available": True, "text": None, "message": "Request failed: timeout"})
    res = _run(_book(tmp_path, name="timeout"), tmp_path, MACRO.format(k=2), monkeypatch, assistant=True)
    assert res["assistant"]["available"] is True and res["assistant"]["message"] == "Request failed: timeout" and res["assistant"]["text"] is None
    assert [_strip(d) for d in res["dimensions"]] == [_strip(d) for d in plain["dimensions"]]


def test_cli_prints_text_and_json(tmp_path, monkeypatch, capsys):
    path = _book(tmp_path)
    monkeypatch.setattr(model_map, "read_vba", lambda _p: [("Module1", MACRO.format(k=2))])
    assert main([str(path), "--work", str(tmp_path / "cli")]) == 0
    out = capsys.readouterr().out
    assert out.startswith("#1 Group: Column C (Group) splits the 8 rows into 3 groups") and "Builder's key: C Group (macro)" in out
    assert main([str(path), "--json", "--work", str(tmp_path / "cli2")]) == 0
    data = json.loads(capsys.readouterr().out)
    assert data["schema"] == "instance-finder/1" and data["dimensions"][0]["key"]["column"] == "C" and data["assistant"]["asked"] is False


def test_the_table_range_ends_at_the_last_filled_row_not_the_used_range(tmp_path, monkeypatch):
    """A table name that runs past the data, on a sheet whose used range runs further still: the finder reports
    the table to its last filled row, as the builder does, and counts only the filled rows."""
    path = _book(tmp_path, name="tail")
    wb = openpyxl.load_workbook(path)
    wb.defined_names["Policies"] = DefinedName("Policies", attr_text="Inputs!$B$5:$E$16")  # four rows past the data
    wb["Inputs"].cell(16, 3).number_format = "0.00"  # a styled empty cell: the used range reaches row 16
    wb.save(path)
    res = _run(path, tmp_path, MACRO.format(k=2), monkeypatch)
    top = _top(res)
    assert top["key"]["column"] == "C" and top["key"]["rows"] == 8 and top["key"]["table"].endswith(":E12"), top["key"]


def test_uc2_blind_top_dimension_matches_the_builders_key(tmp_path):
    """The real 2c use case, when its copy is here: the blind top dimension is the builder's key, with the
    kinds and reader counts the local expectations hold (`finder` in expected.json: top_key_column, top_kinds,
    builder_matches, reader_counts). Nothing of the model is written here."""
    path, expected = usecase_or_skip("UC-2")
    facts = expected.get("finder")
    if not facts:
        pytest.skip("UC-2: no finder expectations in the local expected.json")
    analysis = build_analysis(path, tmp_path / "uc2", "UC-2")
    mp = build_model_map(analysis)
    res = find_instances(analysis, mp, loop_plan=build_loop_plan(analysis, mp))
    top = res["dimensions"][0]
    assert top["key"]["column"] == facts["top_key_column"] and set(facts["top_kinds"]) <= set(top["kinds"]) and res["builder"]["matches_top"] is facts["builder_matches"]
    for k, v in (facts.get("reader_counts") or {}).items():
        assert top["readers"][k]["count"] == v, k
