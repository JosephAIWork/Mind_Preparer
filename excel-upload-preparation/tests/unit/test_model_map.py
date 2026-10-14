"""Model map (1.9.0): where a workbook's inputs, loop driver, run settings and outputs are.

The macro parser is tested on VBA written here in the shapes the four client models use
(a row-per-record loop with sheet bounds; a named driver with named bounds and a paste per
item; nested scenario / year / cluster loops over named lists; a sheet-building loop). The
four models themselves are checked when their local copies are present (never committed)."""
from __future__ import annotations

import os
from pathlib import Path

import openpyxl
import pytest

from app.inventory import build_analysis
from app.model_map import build_model_map, describe, parse_macros, sheet_role

USECASES = Path(os.environ.get("LOCALAPPDATA", "")) / "MindReady" / "autofix" / "usecases"


def _names(*pairs: tuple[str, str]) -> list[dict]:
    return [{"name": n, "value": v, "local_sheet_id": None} for n, v in pairs]


HORIZON_STYLE = '''
Sub Bouton2_Cliquer()
num_debut = Worksheets("CALCUL").Cells(5, 20).Value
num_fin = Worksheets("CALCUL").Cells(6, 20).Value
'Nettoyage de l'output
Worksheets("OUTPUT").Range("A" & num_debut & ":I100000").ClearContents
For ins = num_debut To (num_fin)
    Application.StatusBar = "MP : " & ins
    Worksheets("CALCUL").Range("A2:NI2") = Worksheets("BASE Polices").Range("A" & ins & ":NI" & ins).Value
    Worksheets("OUTPUT").Range("A" & ins & ":J" & ins) = Worksheets("CALCUL").Range("AY35:BH35").Value
Next
End Sub
'''

PALERMO_STYLE = '''
Sub RunPolicies()
Dim MacroStart As Long: MacroStart = Range("MacroStart").Value     'first pol
Dim MacroEnd As Long: MacroEnd = Range("MacroEnd").Value
If Range("DeleteOld").Value = "Yes" Then
    Range("ByTariffOutput").ClearContents
End If
For polProcess = MacroStart To MacroEnd
    Range("MP_ROW_SELECT").Value = polProcess
    If Range("Full_Run?").Value = "No" Then
        GoTo SkipPolicy
    End If
    Application.Calculate
    Range("ByTariffPaster").Offset(0, 10 * (Temp2 - 1)).Value2 = Range("AggregateOutput").Value2
SkipPolicy:
Next polProcess
Call QuickPaste
End Sub
Sub SaveTotal()
Dim SaveAs_txt As String:  SaveAs_txt = Range("SaveNameTotal")
    Sheets(Array("Summary Output Pasted")).Copy
    ActiveWorkbook.SaveAs Filename:=SaveAs_txt, FileFormat:=52
End Sub
Sub QuickPaste()
    Range("SummaryOutput").Copy
    Sheets("Summary Output Pasted").Select
    Range("PasteOutput").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
'''

CNHI_STYLE = '''
Sub baseline_projections()
Dim paste_start_cell As Range
For i_scenario = Range("start_1").Value2 To Range("end_1").Value2
    Range("active_scenario").Value2 = Range("scenario_start").Offset(0, i_scenario - 1).Value2
    For i_year = 1 To 5
        Range("active_year").Value2 = Range("proj_year").Value2 + i_year - 1
        For i_cluster = 1 To 20
            Range("active_cluster").Value2 = Range("Cluster1").Offset(i_cluster - 1, 0).Value2
            Range("CatA_paste_start").Offset((i_year - 1) * 5, (i_cluster - 1) * 3).Value2 = Range("copyA_final").Value2
            Set paste_start_cell = Range("paste_output").Cells(1, 1).Offset(236 * (tracker - 1), 0)
            paste_start_cell.Resize(Range("copy_output").Rows.Count, Range("copy_output").Columns.Count).Value2 = Range("copy_output").Value2
        Next i_cluster
    Next i_year
Next i_scenario
End Sub
Sub combined_macro()
Call baseline_projections
End Sub
'''

BUILDER_STYLE = '''
Sub Macro1()
For i = 11 To 40
    Sheets("LoB 1").Copy Before:=Sheets(">>")
    Sheets("LoB 1 (2)").Name = "LoB " & i
    Range("G2").Value = "LoB " & i
Next i
End Sub
'''


def test_a_row_per_record_loop_with_sheet_bounds_is_read():
    v = parse_macros([("Module1", HORIZON_STYLE)], [])
    assert len(v["loops"]) == 1
    loop = v["loops"][0]
    assert loop["var"] == "ins" and loop["from"]["cell"] == "CALCUL!T5" and loop["to"]["cell"] == "CALCUL!T6"
    d = loop["drivers"][0]
    assert d["cell"] == "CALCUL!A2:NI2" and d["set_from"] == "'BASE Polices'!A<ins>:NI<ins>"
    o = loop["outputs"][0]
    assert o["to"] == "OUTPUT!A<ins>:J<ins>" and o["from"] == "CALCUL!AY35:BH35" and o["per_item"]
    assert v["cleared"][0]["cell"] == "OUTPUT!A<num_debut>:I100000"
    assert v["entry_points"] == ["Bouton2_Cliquer"]


def test_a_named_driver_with_named_bounds_switches_and_a_paste_per_item():
    names = _names(("MacroStart", "Macro!$B$1"), ("MacroEnd", "Macro!$B$2"), ("MP_ROW_SELECT", "Projection!$C$3"), ("DeleteOld", "Macro!$B$3"), ("Full_Run?", "Macro!$B$4"),
                   ("ByTariffPaster", "Macro!$V$13:$AE$140"), ("AggregateOutput", "Macro!$A$11:$J$138"), ("ByTariffOutput", "Macro!$V$13:$JU$140"),
                   ("SummaryOutput", "'Summary Output'!$M$9:$EK$48"), ("PasteOutput", "'Summary Output Pasted'!$A$1:$DY$40"), ("SaveNameTotal", "'Saved down'!$C$9"))
    v = parse_macros([("Module1", PALERMO_STYLE)], names)
    loop = v["loops"][0]
    assert loop["from"] == {"text": "MacroStart = Range(\"MacroStart\").Value", "kind": "cell", "cell": "Macro!B1", "name": "MacroStart"}
    assert loop["to"]["cell"] == "Macro!B2"
    assert loop["drivers"][0]["cell"] == "Projection!C3" and loop["drivers"][0]["name"] == "MP_ROW_SELECT" and loop["drivers"][0]["how"] == "the loop value itself"
    assert loop["outputs"][0]["to"] == "Macro!V13:AE140" and loop["outputs"][0]["from"] == "Macro!A11:J138" and loop["outputs"][0]["per_item"] is False
    sw = {s["name"]: s for s in v["switches"]}
    assert sw["Full_Run?"]["cell"] == "Macro!B4" and sw["Full_Run?"]["compared_with"] == ["No"] and sw["DeleteOld"]["compared_with"] == ["Yes"]
    assert v["cleared"][0]["name"] == "ByTariffOutput"
    assert v["pastes"] == [{"from_name": "SummaryOutput", "from": "'Summary Output'!M9:EK48", "in": "QuickPaste", "to_name": "PasteOutput", "to": "'Summary Output Pasted'!A1:DY40"}]
    assert v["entry_points"] == ["RunPolicies"]
    assert v["saves"] == [{"in": "SaveTotal", "path_from": "'Saved down'!C9", "name": "SaveNameTotal"}]


def test_nested_loops_over_named_lists_and_a_block_per_pass():
    names = _names(("start_1", "Output!$G$2"), ("end_1", "Output!$G$3"), ("active_scenario", "'Model Inputs'!$C$8"), ("scenario_start", "'Model Inputs'!$E$8"),
                   ("active_year", "'Model Inputs'!$C$173"), ("proj_year", "'Assumption Setting'!$D$8"), ("active_cluster", "'Model Inputs'!$C$174"), ("Cluster1", "'Validation list'!$B$7"),
                   ("CatA_paste_start", "'RAC Baseline'!$V$12"), ("copyA_final", "'RAC Calculation'!$DI$13:$DK$118"), ("paste_output", "Output!$B$9:$U$244"), ("copy_output", "'Output working'!$AB$13:$AU$248"))
    v = parse_macros([("Module1", CNHI_STYLE)], names)
    loops = {l["var"]: l for l in v["loops"]}
    assert loops["i_scenario"]["from"]["cell"] == "Output!G2" and loops["i_scenario"]["to"]["cell"] == "Output!G3"
    assert loops["i_year"]["from"] == {"text": "1", "kind": "literal", "value": 1} and loops["i_cluster"]["to"]["value"] == 20
    assert loops["i_scenario"]["drivers"][0]["cell"] == "'Model Inputs'!C8" and loops["i_scenario"]["drivers"][0]["set_from_name"] == "scenario_start"
    assert loops["i_year"]["drivers"][0]["cell"] == "'Model Inputs'!C173"
    assert loops["i_cluster"]["drivers"][0]["cell"] == "'Model Inputs'!C174" and loops["i_cluster"]["drivers"][0]["set_from_name"] == "Cluster1"
    outs = {(o["to_name"], o["from_name"]) for l in v["loops"] for o in l["outputs"]}
    assert ("CatA_paste_start", "copyA_final") in outs and ("paste_output", "copy_output") in outs
    assert v["entry_points"] == ["combined_macro"]


def test_a_loop_that_builds_sheets_is_not_a_run_loop():
    v = parse_macros([("Module2", BUILDER_STYLE)], [])
    loop = v["loops"][0]
    assert loop["builds_sheets"] and loop["copies_sheet"] == "LoB 1"
    assert loop["drivers"][0]["cell"] == "'LoB 1'!G2"
    assert v["entry_points"] == []


def test_sheet_roles_by_name():
    assert sheet_role("Inputs_PVFP") == "input" and sheet_role("OUTPUT") == "output" and sheet_role("CALCUL") == "calculation"
    assert sheet_role(">> Inputs", "A1:A1") == "divider" and sheet_role("Data inputs-->") == "divider"
    assert sheet_role("Non-economic assumptions") == "input" and sheet_role("Reporting_LoB Total") == "output" and sheet_role("Saved down") == "output"


def test_a_workbook_without_macros_gets_dropdown_switches_named_settings_and_key_cells(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    inp = wb.active
    inp.title = "Inputs"
    inp["B3"], inp["C3"] = "Basis", "SCR"
    dv = DataValidation(type="list", formula1='"SCR,MSR"')
    inp.add_data_validation(dv)
    dv.add("C3")
    inp["B4"], inp["C4"] = "Last year", 2016
    inp["B6"], inp["C6"] = "Threshold", 0.001
    wb.defined_names["Last_Year"] = DefinedName("Last_Year", attr_text="Inputs!$C$4")
    calc = wb.create_sheet("Calc")
    calc["G2"] = "LoB 1"
    for r in range(4, 40):
        calc[f"B{r}"] = f'="x#"&$G$2&{r}'
        calc[f"C{r}"] = f"=Last_Year+{r}"
    rep = wb.create_sheet("Reporting")
    rep["B2"] = "=SUM(Calc!C4:C39)"
    p = tmp_path / "nomacro.xlsx"
    wb.save(p)
    wb.close()
    analysis = build_analysis(p, tmp_path / "w", "m")
    mp = build_model_map(analysis)
    assert mp["has_macros"] is False and mp["loop"]["found"] is False
    assert mp["loop"]["key_cells"][0]["cell"] == "Calc!G2" and mp["loop"]["key_cells"][0]["readers"] == 36
    sw = mp["run_settings"]["switches"]
    assert sw[0]["cell"] == "Inputs!C3" and sw[0]["options"] == ["SCR", "MSR"] and sw[0]["label"] == "Basis" and sw[0]["value"] == "SCR"
    named = {n["name"]: n for n in mp["run_settings"]["named_settings"]}
    assert named["Last_Year"]["cell"] == "Inputs!C4" and named["Last_Year"]["readers"] == 36
    lab = {x["label"]: x for x in mp["run_settings"]["labelled"]}
    assert lab["Threshold"]["cell"] == "Inputs!C6" and lab["Threshold"]["value"] == 0.001
    assert [s["sheet"] for s in mp["inputs"]["sheets"]] == ["Inputs"] and [s["sheet"] for s in mp["outputs"]["sheets"]] == ["Reporting"]
    text = describe(mp)
    assert "key cell  Calc!G2" in text and "switch    Inputs!C3" in text


# --- the four client models, when their local copies are here ---------------------------------------
def _map_of(name: str, tmp_path: Path) -> dict:
    p = USECASES / f"{name}.xlsm"
    if not p.is_file():
        pytest.skip(f"{p} not here")
    return build_model_map(build_analysis(p, tmp_path / "w", name))


def test_horizon_map(tmp_path):
    mp = _map_of("horizon", tmp_path)
    lp = mp["loop"]
    assert lp["found"] and lp["drivers"][0]["cell"] == "CALCUL!A2:NI2" and lp["drivers"][0]["set_from"] == "'BASE Polices'!A<ins>:NI<ins>"
    bounds = {b["role"]: b for b in lp["bounds"]}
    assert bounds["first"]["cell"] == "CALCUL!T5" and bounds["first"]["label"] == "Ligne début" and bounds["last"]["cell"] == "CALCUL!T6" and bounds["last"]["formula"] == "=N5"
    assert lp["entry_points"] == ["Bouton2_Cliquer"]
    assert mp["inputs"]["per_item"][0]["cell"] == "'BASE Polices'!A<ins>:NI<ins>"
    assert {s["sheet"] for s in mp["inputs"]["sheets"]} >= {"BASE Polices", "Parametres", "Base Pret", "Criteres", "Tarifs CRD", "Tarifs CI"}
    outs = {o["cell"]: o for o in mp["outputs"]["cells"]}
    assert outs["OUTPUT!A<ins>:J<ins>"]["from"] == "CALCUL!AY35:BH35"
    assert any(x["sheet"] == "Parametres" for x in mp["run_settings"]["labelled"])


def test_pvfp_map(tmp_path):
    mp = _map_of("pvfp", tmp_path)
    lp = mp["loop"]
    assert lp["found"] is False
    assert {b["cell"] for b in lp["item_cells"]} == {"'LoB 1'!G2", "'Reporting_BNP_LoB 1'!E2"}
    sw = {s["name"]: s for s in mp["run_settings"]["switches"] if s.get("name")}
    assert sw["DAC_Model"]["cell"] == "Inputs_PVFP!H71" and sw["DAC_Model"]["options"] == ["Yes", "No"]
    assert sw["LoB3_Boolean"]["cell"] == "Inputs_PVFP!G53" and sw["LoB3_Boolean"]["options"] == ["VRAI", "FAUX"]
    assert any(s["cell"] == "CoC!F2" and s["options"] == ["SCR", "MSR"] for s in mp["run_settings"]["switches"])
    named = {n["name"]: n for n in mp["run_settings"]["named_settings"]}
    assert named["Last_Year_Triangle"]["cell"] == "Inputs_PVFP!G69"
    assert {s["sheet"] for s in mp["inputs"]["sheets"]} == {"Yield_Curve", "Inputs_CoC", "Inputs_PVFP"}
    assert {s["sheet"] for s in mp["outputs"]["sheets"]} == {"Reporting_LoB 1", "Reporting_LoB Total", "Reporting_General"}


def test_palermo_map(tmp_path):
    mp = _map_of("palermo", tmp_path)
    lp = mp["loop"]
    d = {x["name"]: x for x in lp["drivers"]}
    assert d["MP_ROW_SELECT"]["cell"] == "Projection!C3"
    bounds = {b["role"]: b for b in lp["bounds"]}
    assert bounds["first"]["name"] == "MacroStart_UDC2" and bounds["first"]["cell"] == "Macro!B1" and bounds["last"]["cell"] == "Macro!B2"
    assert lp["entry_points"] == ["UndiscountedCashflow"]
    sw = {s["name"]: s for s in mp["run_settings"]["switches"] if s.get("name")}
    assert sw["Full_Run?"]["cell"] == "Macro!B4" and sw["Del_range2"]["cell"] == "Macro!B3" and sw["SendEmails"]["cell"] == "Macro!B5"
    assert sw["Full_Run?"]["how"].startswith("a dropdown in the cell") and sw["Full_Run?"]["options"] == ["Yes", "No"]
    outs = {o["name"]: o for o in mp["outputs"]["cells"] if o.get("name")}
    assert outs["ByTariffPaster"]["from_name"] == "AggregateOutputByTariff" and outs["PasteOutput"]["from_name"] == "SummaryOutput"
    assert "ByTariffOutput" in outs
    tables = {t["name"] for t in mp["inputs"]["tables"]}
    assert "TarRange" in tables or any(s["sheet"] == "Data" for s in mp["inputs"]["sheets"])
    assert any(sv["name"] == "SaveNameTotal" for sv in mp["outputs"]["saves"])


def test_cnhi_map(tmp_path):
    mp = _map_of("cnhi", tmp_path)
    lp = mp["loop"]
    d = {x["name"]: x for x in lp["drivers"]}
    assert d["active_scenario"]["cell"] == "'Model Inputs'!C8" and d["active_year"]["cell"] == "'Model Inputs'!C173" and d["active_cluster"]["cell"] == "'Model Inputs'!C174"
    assert d["active_cluster"]["set_from_name"] == "Cluster1" and d["active_scenario"]["set_from_name"] == "scenario_start"
    bounds = {(b["role"], b["name"]) for b in lp["bounds"]}
    assert ("first", "start_1") in bounds and ("last", "end_1") in bounds
    assert "combined_macro" in lp["entry_points"]
    outs = {o["name"]: o for o in mp["outputs"]["cells"] if o.get("name")}
    assert outs["paste_output"]["from_name"] == "copy_output" and outs["CatA_benef_start"]["from_name"] == "CatA_benef_copy"
    sw = {s["cell"]: s for s in mp["run_settings"]["switches"]}
    assert "'Model Inputs'!C174" in sw and "'Model Inputs'!C8" in sw
    assert {s["sheet"] for s in mp["outputs"]["sheets"]} >= {"Output", "Exhibits", "Inbound Output", "SS Output"}
