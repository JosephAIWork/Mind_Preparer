"""Outside input files (app/outside_inputs.py): a workbook whose macro opens other files and pastes
their data in -- the Tokyo IFRS 17 tool's calcCF reads one .xlsm per row of Ctrl_01!H35:H40 into
BBA_01..05 / VFA_01.

Everything but the last test runs without Excel, on text written here in the shapes of the Tokyo
tool (its VBA, its path formula, its readers, the event log Excel gave on the real run). The last
test runs the whole chain on the anonymized Tokyo copy with invented input files, when that fake
data is on this machine (it is never committed).

Tests marked xfail show a defect of the module: the reason says what goes wrong."""
from __future__ import annotations

import os
import random
import shutil
import threading
import time
from pathlib import Path

import openpyxl
import pytest

from app import outside_inputs as oi
from app.excel_com import com_available
from app.formula_utils import parse_ref
from app.outside_inputs import (
    INPUT_FILE_EXTENSIONS,
    analyse_events,
    csv_number,
    describe_match,
    file_name_tail,
    find_import_table,
    import_table_from_analysis,
    match_files,
    normalize_name,
    plan_mind_inputs,
    plan_rewrites,
    redirect_formula,
    repair_dialog_text,
    scan_vba,
    source_of_message,
    split_partial_sums,
    write_input_csv,
)

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")

# --- the Tokyo tool's VBA (shortened; the Japanese text as the tool shows it) -----------------------
M1_MAIN = '''Option Explicit

'選択したシナリオで現在推計の計算を実行
Sub OneScenarioRun()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Call calcCF
    MsgBox "計算が完了しました"
    Application.Calculation = xlCalculationAutomatic
End Sub

'全てのシナリオで計算を実行し、Outputをcsvに出力
Sub AllScenarioRun()
    Dim res As Variant
    res = MsgBox("実行しますか？", vbYesNo)
    If res = vbYes Then
        Call calcCF
        Call OutputResult(0, "2026Q1")
        MsgBox "計算が完了しました"
    End If
End Sub

'各種Inputを実行
Sub calcCF()
    Const MAX_CF_MONTH = 1200
    Dim readCFRng As Range
    Set readCFRng = ThisWorkbook.Names("ReadCF").RefersToRange
    For i = 1 To numCFfile
        filePath = CFfolderPath(i, 1)
        fileName = CFfileName(i, 1)
        'インプットエリアのリセット
        Set inputRng = ThisWorkbook.Sheets(fileName).Cells(2, findCol(fileName, "csv file", 1) + 1)
        inputColumn = inputRng.End(xlToRight).Column - inputRng.Column + 1
        inputRng(2, 2).Resize(MAX_CF_MONTH, inputColumn - 1).ClearContents
        If fso.FileExists(filePath) Then
            Call InputCFFile(fileName, filePath, MAX_CF_MONTH, inputRng, inputColumn)
            readCFRng(i + 1, 5).Value = Format(Now(), "YYYY/MM/DD HH:MM:SS")
        Else
            readCFRng(i + 1, 5).Value = "ファイルが存在しません"
        End If
    Next i
End Sub

'IFRS17固有の計算を実行し、Outputをcsvに出力
Sub EndemicCalculationRun()
    Dim res As Variant
    res = MsgBox("実行しますか？", vbYesNo)
    If res = vbYes Then
        csvFile = filePath & "\\Result_" & yyyyQx & ".csv"
        Open csvFile For Output As #1
        Print #1, outputSummry.Cells(i, j).Value & ",";
        Close #1
        MsgBox "計算が完了しました"
    End If
End Sub
'''

M2_INPUT_CF_FILE = '''Option Explicit

Sub InputCFFile(ByVal fileName As String, ByVal filePath As String, ByVal MAX_CF_MONTH As Long, _
                ByVal inputRng As Range, ByVal inputColumn As Long)
    Workbooks.Open filePath
    fileName2 = ActiveWorkbook.Name
    'BBAのinputファイルとVFAのinputファイルは異なる
    If fileName2 Like "*BBA*" Then
        Set myRng = Cells.Find("Description").Cells(5, 1)
    ElseIf fileName2 Like "*VFA*" Then
        Set myRng = Cells.Find("Accrued interest at end of month").Cells(1, 1)
    End If
    fileRow = myRng.End(xlDown).Row - myRng.Row + 1
    If inputColumn < fileRow + 1 Then
        MsgBox (fileName & "のinput領域には" & inputColumn - 1 & "項目まで入力可能であるのに対し、inputファイルには" & fileRow & "項目あります")
        End
    End If
    inputRng(2, 2).Resize(MaxCol, fileRow) = InputArray_mod
    Workbooks(fileName2).Close
End Sub
'''

M3_OUTPUT = '''Option Explicit

Sub OutputResult(ByVal iScenario As Long, ByVal yyyyQx As String)
    csvFile = filePath & "\\Output " & yyyyQx & " - Run" & iScenario & ".csv"
    Open csvFile For Output As #1
    Print #1, Sheets(OutputSheet).Cells(i, j).Value & ",";
    Close #1
End Sub
'''

M4_OTHERS = '''Option Explicit

Function CountCSVRow(ByVal filePath As String)
    fileNum = FreeFile
    Open filePath For Input As #fileNum
        Do Until EOF(1)
            Line Input #fileNum, buf
            cnt = cnt + 1
        Loop
    Close #fileNum
    CountCSVRow = cnt - 1
End Function

Function findCol(ByVal sheetName As String, ByVal x As String, ByVal r As Long, Optional startCol As Long = 1, Optional endCol As Long = 0) As Long
    For i = 1 To endCol
        If ThisWorkbook.Sheets(sheetName).Cells(r, i).Value = x Then
            findCol = i
            Exit For
        End If
    Next i
End Function
'''

TOKYO_VBA = [("M1_Main", M1_MAIN), ("M2_InputCFFile", M2_INPUT_CF_FILE), ("M3_Output", M3_OUTPUT), ("M4_Others", M4_OTHERS)]

# the path formula of Ctrl_01!H35 (H36..H40 read D36..D40)
TOKYO_H35 = (r'=VLOOKUP($G$27,$B$6:$J$12,2,FALSE)&"\"&IF(RIGHT($G$27,4)="当初認識","NB_","")'
             r'&IF(RIGHT($G$27,3)="前期末",$C$20&$D$20,$C$18&$D$18)&"_"&$C$22&"_"&D35&".xlsm"')

TOKYO_SHEETS = ("BBA_01", "BBA_02", "BBA_03", "BBA_04", "BBA_05", "VFA_01")
BOOK = "SBI_IFRS17_Anonymized_original.xlsm"
# per sheet: the range calcCF clears (1,200 months) and the block InputCFFile pastes (1,201 months)
TOKYO_RUN = [(s, "AB3:DP1202", "AB3:DP1203") for s in TOKYO_SHEETS[:5]] + [("VFA_01", "AE3:ABS1202", "AE3:ABS1203")]


def _area(name: str, sheet: str, ref: str, file: str | None = None) -> dict:
    r = parse_ref(ref)
    return {"name": name, "sheet": sheet, "ref": ref, "r1": r["r1"], "c1": r["c1"], "r2": r["r2"], "c2": r["c2"],
            "rows": r["r2"] - r["r1"] + 1, "cols": r["c2"] - r["c1"] + 1, "file": file}


BBA_04 = _area("IN_BBA_04", "BBA_04", "AB3:DP1203")
VFA_01 = _area("IN_VFA_01", "VFA_01", "AE3:ABS1203")


def _table_row(cell: str, keys=(), expected: str | None = None, sheet: str = "Ctrl_01") -> dict:
    r = parse_ref(cell)
    return {"sheet": sheet, "cell": cell, "row": r["r1"], "col": r["c1"], "formula": None, "cached": None,
            "expected": expected, "tail": None, "keys": list(keys), "target_sheet": None}


# --- 1. VBA: which macro runs the import ------------------------------------------------------------
def test_tokyo_vba_import_macro_is_calccf_and_its_chain_reaches_the_workbooks_open():
    s = scan_vba(TOKYO_VBA)
    assert s["import_macro"] == "calcCF"
    assert s["chain"] == ["calcCF", "InputCFFile"]
    assert {(i["procedure"], i["module"], i["call"]) for i in s["imports"]} == {
        ("InputCFFile", "M2_InputCFFile", "Workbooks.Open"),
        ("CountCSVRow", "M4_Others", "Open ... For Input"),
    }
    procs = s["procedures"]
    assert procs["OneScenarioRun"]["calls"] == ["calcCF"]
    assert procs["AllScenarioRun"]["calls"] == ["OutputResult", "calcCF"]
    assert procs["calcCF"]["calls"] == ["InputCFFile", "findCol"]
    assert procs["calcCF"]["params"] == "" and procs["calcCF"]["kind"] == "sub"
    # the signature continued with " _" is one line: its parameters are read
    assert procs["InputCFFile"]["params"].startswith("ByVal fileName As String") and procs["InputCFFile"]["params"].endswith("ByVal inputColumn As Long")
    assert (procs["AllScenarioRun"]["msgbox"], procs["AllScenarioRun"]["yes_no"]) == (True, True)
    assert (procs["calcCF"]["msgbox"], procs["calcCF"]["yes_no"]) == (False, False)


def test_open_for_output_writes_a_file_and_is_not_an_import():
    s = scan_vba(TOKYO_VBA)
    writers = {"EndemicCalculationRun", "OutputResult"}
    assert not writers & {i["procedure"] for i in s["imports"]}
    assert s["procedures"]["OutputResult"]["opens"] == [] and s["procedures"]["EndemicCalculationRun"]["opens"] == []
    s = scan_vba([("M", "Sub W()\n    Open p For Output As #1\n    Open p For Append As #2\nEnd Sub\n")])
    assert s["imports"] == [] and s["import_macro"] is None and s["chain"] == []


def test_every_kind_of_opening_call_is_found():
    code = '''Sub R()
    Open p For Binary Access Read As #3
End Sub
Sub Q()
    ActiveSheet.QueryTables.Add Connection:="TEXT;" & p, Destination:=Range("A1")
    Set o = GetObject(p)
    Workbooks.OpenText Filename:=p
End Sub
'''
    s = scan_vba([("M", code)])
    assert [(i["procedure"], i["call"]) for i in s["imports"]] == [
        ("R", "Open ... For Input"), ("Q", "QueryTables.Add"), ("Q", "GetObject"), ("Q", "Workbooks.OpenText")]
    assert s["procedures"]["Q"]["opens"] == ["GetObject", "QueryTables.Add", "Workbooks.OpenText"]


def test_a_sub_with_a_japanese_name_is_found_and_can_be_the_import_macro():
    code = '''Sub 一括実行()
    Call 取込実行
End Sub
Sub 取込実行()
    Dim パス As String
    パス = Range("H35").Value
    Workbooks.Open パス
End Sub
'''
    s = scan_vba([("標準モジュール1", code)])
    assert [(i["procedure"], i["module"]) for i in s["imports"]] == [("取込実行", "標準モジュール1")]
    assert s["procedures"]["一括実行"]["calls"] == ["取込実行"]
    assert s["import_macro"] == "取込実行" and s["chain"] == ["取込実行"]  # the opener itself is the nearest


def test_workbooks_open_in_comments_and_strings_does_not_count():
    code = '''Sub Notes()
    ' Workbooks.Open is called in another tool
    Rem Workbooks.Open here too
    MsgBox "Use Workbooks.Open to load the file"
    x = 1 ' Workbooks.Open trailing comment
End Sub
'''
    s = scan_vba([("M", code)])
    assert s["imports"] == [] and s["import_macro"] is None
    # an apostrophe inside a string does not start a comment; a call continued with " _" is still a call
    s = scan_vba([("M", 'Sub Load()\n    MsgBox "Don\'t stop": Workbooks.Open Range("A1").Value\nEnd Sub\n'),
                  ("N", 'Sub Load2()\n    Set wb = Workbooks _\n        .Open(Range("A1").Value)\nEnd Sub\n')])
    assert [(i["procedure"], i["call"]) for i in s["imports"]] == [("Load", "Workbooks.Open"), ("Load2", "Workbooks.Open")]


def test_a_procedure_with_parameters_or_a_function_is_never_the_import_macro():
    s = scan_vba([("M", 'Sub Load(ByVal p As String)\n    Workbooks.Open p\nEnd Sub\nFunction LoadAll()\n    Call Load("x")\nEnd Function\n')])
    assert [i["procedure"] for i in s["imports"]] == ["Load"]
    assert s["import_macro"] is None and s["chain"] == []
    # with a caller without parameters, the caller is chosen, never the opener with parameters
    s = scan_vba([("M", 'Sub Go()\n    Call Load("x")\nEnd Sub\nSub Load(ByVal p As String)\n    Workbooks.Open p\nEnd Sub\n')])
    assert s["import_macro"] == "Go" and s["chain"] == ["Go", "Load"]


def test_a_procedure_that_asks_yes_no_loses_a_tie_and_one_with_a_message_box_too():
    code = '''Sub A_Ask()
    If MsgBox("実行しますか？", vbYesNo) = vbYes Then Call Load("x")
End Sub
Sub B_Info()
    Call Load("x")
    MsgBox "Done"
End Sub
Sub C_Plain()
    Call Load("x")
End Sub
Sub Load(ByVal p As String)
    Workbooks.Open p
End Sub
'''
    s = scan_vba([("M", code)])
    assert s["import_macro"] == "C_Plain" and s["chain"] == ["C_Plain", "Load"]  # A_Ask comes first by name, still loses
    without_plain = code.replace('Sub C_Plain()\n    Call Load("x")\nEnd Sub\n', "")
    assert scan_vba([("M", without_plain)])["import_macro"] == "B_Info"  # an OK box beats a question


def test_no_vba_finds_nothing():
    assert scan_vba([]) == {"imports": [], "import_macro": None, "chain": [], "procedures": {}, "messages": []}


def test_a_module_qualified_call_still_links_the_chain():
    opener = 'Sub InputCFFile(ByVal fileName As String, ByVal filePath As String)\n    Workbooks.Open filePath\nEnd Sub\n'
    s = scan_vba([("M1_Main", 'Sub calcCF()\n    Call M2_InputCFFile.InputCFFile("BBA_01", "C:\\x.xlsm")\nEnd Sub\n'),
                  ("M2_InputCFFile", opener)])
    assert s["import_macro"] == "calcCF" and s["chain"] == ["calcCF", "InputCFFile"]


def test_an_opening_call_in_a_second_module_with_the_same_procedure_name_is_listed():
    s = scan_vba([("Sheet1", 'Private Sub CommandButton1_Click()\n    MsgBox "hi"\nEnd Sub\n'),
                  ("Sheet2", 'Private Sub CommandButton1_Click()\n    Workbooks.Open Range("H35").Value\nEnd Sub\n')])
    assert [(i["module"], i["call"]) for i in s["imports"]] == [("Sheet2", "Workbooks.Open")]


# --- message boxes: what the VBA says, and what a garbled box said -----------------------------------
def test_message_boxes_are_listed_and_a_shown_message_is_traced_to_its_line():
    s = scan_vba(TOKYO_VBA)
    assert [m["procedure"] for m in s["messages"]] == [
        "OneScenarioRun", "AllScenarioRun", "AllScenarioRun", "EndemicCalculationRun", "EndemicCalculationRun", "InputCFFile"]
    # the box InputCFFile shows when an input file has more items than the sheet holds (fake_inputs/errors/too_many_items)
    shown = "BBA_01のinput領域には93項目まで入力可能であるのに対し、inputファイルには95項目あります"
    src = source_of_message(shown, s["messages"])
    assert src["procedure"] == "InputCFFile" and "項目あります" in src["line"]
    assert "計算が完了しました" in source_of_message("計算が完了しました", s["messages"])["line"]
    assert source_of_message("Something else entirely", s["messages"]) is None
    assert source_of_message(shown, []) is None


def _system_code_page(monkeypatch, cp: int) -> None:
    import ctypes

    if not hasattr(ctypes, "windll"):
        pytest.skip("Windows only")
    monkeypatch.setattr(ctypes.windll.kernel32, "GetACP", lambda: cp)


def test_repair_dialog_text_leaves_ascii_alone_and_undoes_japanese_garbled_by_a_western_code_page(monkeypatch):
    assert repair_dialog_text("File not found") == "File not found"
    assert repair_dialog_text("") == ""
    _system_code_page(monkeypatch, 1252)
    garbled = "計算が完了しました".encode("cp932").decode("cp1252")  # what a cp1252 Windows shows
    assert garbled != "計算が完了しました"
    assert repair_dialog_text(garbled) == "計算が完了しました"


@pytest.mark.parametrize("cp, text", [(1255, "הקובץ לא נמצא"), (1252, "Fichier introuvable : vérifiez le nom")])
def test_repair_dialog_text_leaves_a_correct_message_in_the_system_language_alone(monkeypatch, cp, text):
    _system_code_page(monkeypatch, cp)  # 1255 is this machine's own code page (a Hebrew Windows)
    assert repair_dialog_text(text) == text


@pytest.fixture
def fake_dialog(monkeypatch):
    """One message box of Excel process 4242, as the Win32 calls of _dialog_watch see it: no real window."""
    win32api = pytest.importorskip("win32api")
    win32con = pytest.importorskip("win32con")
    win32gui = pytest.importorskip("win32gui")
    win32process = pytest.importorskip("win32process")
    pid = 4242
    box = {"children": []}  # (hwnd, class, text, control id)
    posted: list[tuple[int, int]] = []
    stop = threading.Event()

    def child(h):
        return next((c for c in box["children"] if c[0] == h), None)

    monkeypatch.setattr(win32gui, "EnumWindows", lambda cb, extra: [cb(100, extra)] and None)
    monkeypatch.setattr(win32gui, "EnumChildWindows", lambda h, cb, extra: [cb(c[0], extra) for c in box["children"]] and None if h == 100 else None)
    monkeypatch.setattr(win32gui, "GetClassName", lambda h: "#32770" if h == 100 else (child(h) or (0, ""))[1])
    monkeypatch.setattr(win32gui, "GetWindowText", lambda h: "Microsoft Excel" if h == 100 else (child(h) or (0, "", ""))[2])
    monkeypatch.setattr(win32gui, "IsWindowVisible", lambda h: True)
    monkeypatch.setattr(win32gui, "GetDlgCtrlID", lambda h: (child(h) or (0, "", "", 0))[3], raising=False)
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda h: (1, pid))

    def post(h, msg, wparam, lparam):
        posted.append((h, msg))
        stop.set()
        return True

    monkeypatch.setattr(win32api, "PostMessage", post)

    def show(message: str, buttons: list[tuple[str, int]]) -> tuple[list[dict], list[tuple[int, int]]]:
        box["children"] = [(201, "Static", message, 0xFFFF)] + [(202 + i, "Button", label, cid) for i, (label, cid) in enumerate(buttons)]
        found: list[dict] = []
        oi._dialog_watch(pid, stop, found, time.time() + 5, lambda: stop.set())
        return found, posted

    show.BM_CLICK, show.WM_CLOSE = win32con.BM_CLICK, win32con.WM_CLOSE
    return show


@pytest.mark.parametrize("buttons, pressed", [
    ([("はい(&Y)", 6), ("いいえ(&N)", 7)], "いいえ(&N)"),  # a question on a Japanese Windows: No
    ([("&Yes", 6), ("&No", 7)], "&No"),
    ([("OK", 1)], "OK"),
])
def test_dialog_watch_answers_ok_or_no_and_records_the_box(fake_dialog, buttons, pressed):
    found, posted = fake_dialog("上書きしますか？", buttons)
    hwnd = 202 + [label for label, _ in buttons].index(pressed)
    assert posted == [(hwnd, fake_dialog.BM_CLICK)]
    pressed_id = dict(buttons)[pressed]
    assert found == [{"title": "Microsoft Excel", "message": "上書きしますか？", "buttons": [b for b, _ in buttons],
                      "pressed": pressed, "pressed_id": pressed_id}]


def test_dialog_watch_answers_no_on_a_hebrew_windows(fake_dialog):
    found, posted = fake_dialog("להחליף את הקובץ?", [("כן", 6), ("לא", 7)])
    assert posted == [(203, fake_dialog.BM_CLICK)]


# --- 2. the import table -----------------------------------------------------------------------------
def test_file_name_tail_of_the_tokyo_path_formula():
    assert file_name_tail(TOKYO_H35) == (r'IF(RIGHT($G$27,4)="当初認識","NB_","")&IF(RIGHT($G$27,3)="前期末",$C$20&$D$20,$C$18&$D$18)'
                                         r'&"_"&$C$22&"_"&D35&".xlsm"')


@pytest.mark.parametrize("formula, tail", [
    (r'="C:\data\file.csv"', '"file.csv"'),
    (r'="C:\data\"&"a.csv"', '"a.csv"'),
    (r'=$B$1&"/"&C5&".csv"', 'C5&".csv"'),
    (r'=$B$1&"\BBA_"&C5&".xlsm"', '"BBA_"&C5&".xlsm"'),  # the literal holding the last "\" also starts the name
    (r'=$B$1&"\sub\"&C5&"_x.csv"', 'C5&"_x.csv"'),
    ('="file.csv"', None),  # no folder separator to cut at
    ('=A1&".xlsm"', None),
])
def test_file_name_tail_simple_cases(formula, tail):
    assert file_name_tail(formula) == tail


def _cells_and_values():
    cells = [
        {"sheet": "Ctrl_01", "cell": "H35", "formula": r'=$B$1&"\"&$C$22&"_"&D35&".xlsm"'},
        {"sheet": "Ctrl_01", "cell": "H36", "formula": r'=$B$1&"\"&$C$22&"_"&D36&".xlsm"'},
        {"sheet": "Ctrl_01", "cell": "H37", "formula": r'=$B$1&"\"&D37&"_"&E37&".csv"'},
        {"sheet": "Ctrl_01", "cell": "H38", "formula": "=D38*2"},  # no file name in it
        {"sheet": "Ctrl_01", "cell": "J35", "value": r"C:\data\rates.csv"},  # a typed path
        {"sheet": "Ctrl_01", "cell": "J36", "value": "rates.csv"},  # a name without a folder: not a path
        {"sheet": "Ctrl_01", "cell": "J37", "value": "C:\\data\\"},  # a folder without a file
        {"sheet": "入力", "cell": "B2", "value": " D:/共有/金利.xlsx "},
        {"sheet": "Ctrl_01", "cell": "XYZ", "formula": '="a\\b.csv"'},  # not a cell address
    ]
    values = {
        ("Ctrl_01", 35, 8): r"C:\in\2026Q1_2026_BBA_01.xlsm",  # the path Excel last computed
        ("Ctrl_01", 36, 8): "#VALUE!",  # what the anonymized Tokyo copy holds in H35:H40
        ("Ctrl_01", 35, 4): "BBA_01",
        ("Ctrl_01", 36, 4): "ＢＢＡ＿０２",  # full-width: still the sheet BBA_02
        ("Ctrl_01", 37, 4): "保険料",  # a Japanese sheet name
        ("Ctrl_01", 37, 5): 2026,
        ("Ctrl_01", 22, 3): "2026",  # $C$22 is not on the row: not a key
    }
    return cells, values


def test_find_import_table_reads_keys_target_sheet_and_expected_name():
    cells, values = _cells_and_values()
    asked: list[tuple] = []

    def value_of(sheet, row, col):
        asked.append((sheet, row, col))
        return values.get((sheet, row, col))

    rows = find_import_table(cells, value_of, ["Ctrl_01", "BBA_01", "BBA_02", "保険料", "入力"])
    by_cell = {r["cell"]: r for r in rows}
    assert [r["cell"] for r in rows] == ["H35", "H36", "H37", "J35", "B2"]  # by sheet, column, row
    h35 = by_cell["H35"]
    assert (h35["row"], h35["col"], h35["formula"]) == (35, 8, cells[0]["formula"])
    assert (h35["cached"], h35["expected"]) == (r"C:\in\2026Q1_2026_BBA_01.xlsm", "2026Q1_2026_BBA_01.xlsm")
    assert (h35["keys"], h35["target_sheet"], h35["tail"]) == (["BBA_01"], "BBA_01", '$C$22&"_"&D35&".xlsm"')
    h36 = by_cell["H36"]
    assert (h36["cached"], h36["expected"], h36["keys"], h36["target_sheet"]) == ("#VALUE!", None, ["ＢＢＡ＿０２"], "BBA_02")
    h37 = by_cell["H37"]
    assert (h37["keys"], h37["target_sheet"], h37["cached"], h37["expected"]) == (["保険料", "2026"], "保険料", None, None)
    assert ("Ctrl_01", 22, 3) not in asked  # only the cells of the row are keys


def test_find_import_table_takes_typed_paths_with_a_folder_only():
    cells, values = _cells_and_values()
    rows = {(r["sheet"], r["cell"]): r for r in find_import_table(cells, lambda s, r, c: values.get((s, r, c)), ["Ctrl_01", "入力"])}
    assert ("Ctrl_01", "J36") not in rows and ("Ctrl_01", "J37") not in rows and ("Ctrl_01", "H38") not in rows
    j35 = rows[("Ctrl_01", "J35")]
    assert (j35["formula"], j35["cached"], j35["expected"], j35["tail"], j35["keys"], j35["target_sheet"]) == (
        None, r"C:\data\rates.csv", "rates.csv", None, [], None)
    assert rows[("入力", "B2")]["expected"] == "金利.xlsx"


def test_a_check_formula_next_to_the_path_is_not_an_import_row():
    cells = [{"sheet": "Ctrl_01", "cell": "H35", "formula": r'=$B$1&"\"&D35&".xlsm"'},
             {"sheet": "Ctrl_01", "cell": "K35", "formula": '=IF(RIGHT(H35,5)=".xlsm","OK","NG")'}]
    values = {("Ctrl_01", 35, 8): r"C:\in\BBA_01.xlsm", ("Ctrl_01", 35, 4): "BBA_01", ("Ctrl_01", 35, 11): "OK"}
    rows = find_import_table(cells, lambda s, r, c: values.get((s, r, c)), ["Ctrl_01", "BBA_01"])
    assert [r["cell"] for r in rows] == ["H35"]


def test_import_table_from_a_scan_without_excel(tmp_path):
    from app.inventory import build_analysis

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ctrl_01"
    ws["A1"] = "SYNTHETIC_TEST_FIXTURE"
    ws["B1"] = r"C:\inputs\2026_1Q"
    ws["C22"] = "2026"
    ws["D35"], ws["H35"] = "BBA_01", r'=$B$1&"\"&$C$22&"_"&D35&".xlsm"'
    ws["D36"], ws["H36"] = "ｂｂａ＿０２", r'=$B$1&"\"&$C$22&"_"&D36&".xlsm"'
    ws["J40"], ws["J41"] = r"C:\data\金利.csv", "no folder.csv"
    wb.create_sheet("BBA_01")
    wb.create_sheet("BBA_02")
    path = tmp_path / "synthetic_outside_inputs.xlsx"
    wb.save(path)
    wb.close()
    rows = import_table_from_analysis(build_analysis(path, tmp_path / "scan", "oi_synthetic"))
    assert [(r["cell"], r["keys"], r["target_sheet"], r["expected"]) for r in rows] == [
        ("H35", ["BBA_01"], "BBA_01", None),  # openpyxl wrote no computed value: the name is not known yet
        ("H36", ["ｂｂａ＿０２"], "BBA_02", None),
        ("J40", [], None, "金利.csv"),
    ]
    assert rows[0]["tail"] == '$C$22&"_"&D35&".xlsm"'


# --- 3. matching the picked files --------------------------------------------------------------------
def test_normalize_name_and_input_extensions():
    assert normalize_name("  ＢＢＡ（１）.XLSM ") == "bba(1).xlsm"
    assert normalize_name(None) == ""
    assert INPUT_FILE_EXTENSIONS == {".xlsx", ".xlsm", ".xlsb", ".xls", ".csv", ".txt", ".prn"}


def test_exact_name_matches_across_full_width_and_case():
    m = match_files([_table_row("H35", ["BBA_01"], "IF_当期末（非金融前提変更後）_１.xlsm")],
                    ["if_当期末(非金融前提変更後)_1.XLSM", "other.xlsx"])
    assert [(r["file"], r["how"]) for r in m["rows"]] == [("if_当期末(非金融前提変更後)_1.XLSM", "exact")]
    assert (m["matched"], m["missing"], m["extra"]) == (1, [], ["other.xlsx"])
    assert describe_match(m) == "1 of 1 input files matched; not used: other.xlsx."


def test_a_key_matches_as_a_whole_word_only():
    m = match_files([_table_row("H35", ["BBA_01"])], ["2026Q1_2026_BBA_011.xlsm", "2026Q1_2026_XBBA_01.xlsm", "2026Q1_2026_BBA_01.xlsm"])
    assert [(r["file"], r["how"], r["candidates"]) for r in m["rows"]] == [("2026Q1_2026_BBA_01.xlsm", "key", ["2026Q1_2026_BBA_01.xlsm"])]
    m = match_files([_table_row("H35", ["BBA_01"])], ["2026Q1_2026_BBA_011.xlsm"])
    assert m["rows"][0]["file"] is None and m["rows"][0]["candidates"] == []
    assert describe_match(m) == "0 of 1 input files matched; missing: BBA_01; not used: 2026Q1_2026_BBA_011.xlsm."
    m = match_files([_table_row("H35", ["BBA_01"])], ["2026Q1_2026_ＢＢＡ＿０１.xlsm"])  # full-width in the file name
    assert (m["rows"][0]["file"], m["rows"][0]["how"]) == ("2026Q1_2026_ＢＢＡ＿０１.xlsm", "key")


def test_keys_in_japanese_and_hebrew_match_between_separators():
    m = match_files([_table_row("H35", ["死亡"]), _table_row("H36", ["חיים"])], ["2026Q1_死亡_BBA.xlsm", "תיק_חיים_2026.xlsx", "死亡保険.xlsm"])
    assert [(r["file"], r["how"]) for r in m["rows"]] == [("2026Q1_死亡_BBA.xlsm", "key"), ("תיק_חיים_2026.xlsx", "key")]
    assert m["extra"] == ["死亡保険.xlsm"]  # 死亡 inside a longer word is not the key


def test_the_expected_extension_is_preferred_among_key_matches():
    m = match_files([_table_row("H35", ["BBA_01"], "old_BBA_01.xlsm")], ["2026_BBA_01.csv", "2026_BBA_01.xlsm"])
    assert (m["rows"][0]["file"], m["rows"][0]["how"]) == ("2026_BBA_01.xlsm", "key")
    assert m["extra"] == ["2026_BBA_01.csv"]
    # without an expected name both stay candidates and the row is missing
    m = match_files([_table_row("H35", ["BBA_01"])], ["2026_BBA_01.csv", "2026_BBA_01.xlsm"])
    assert m["rows"][0]["file"] is None and m["rows"][0]["candidates"] == ["2026_BBA_01.csv", "2026_BBA_01.xlsm"]


def test_two_rows_claiming_one_file():
    # an exact claim keeps the file; the key claim is left open with the file as a candidate
    m = match_files([_table_row("H35", ["BBA_01"], "2026_BBA_01.xlsm"), _table_row("H36", ["BBA_01"])], ["2026_BBA_01.xlsm"])
    assert [(r["cell"], r["file"], r["how"], r["candidates"]) for r in m["rows"]] == [
        ("H35", "2026_BBA_01.xlsm", "exact", []), ("H36", None, None, ["2026_BBA_01.xlsm"])]
    assert (m["matched"], [r["cell"] for r in m["missing"]], m["extra"]) == (1, ["H36"], [])
    # two key claims: both rows stay open
    m = match_files([_table_row("H35", ["BBA"]), _table_row("H36", ["01"])], ["2026_BBA_01.xlsm"])
    assert [(r["file"], r["candidates"]) for r in m["rows"]] == [(None, ["2026_BBA_01.xlsm"])] * 2
    assert m["matched"] == 0 and len(m["missing"]) == 2


def test_a_file_two_rows_want_is_not_reported_as_extra():
    m = match_files([_table_row("H35", ["BBA"]), _table_row("H36", ["01"])], ["2026_BBA_01.xlsm"])
    assert m["extra"] == []


def test_missing_and_extra_files_and_the_sentence():
    m = match_files([_table_row("H35", ["BBA_01"]), _table_row("H36", ["BBA_02"], "2026Q1_2026_BBA_02.xlsm"), _table_row("H37")],
                    ["2026Q1_2026_BBA_01.xlsm", "notes.txt", "2026Q1_2026_BBA_01.xlsm"])  # the same file picked twice counts once
    assert m["matched"] == 1 and [r["cell"] for r in m["missing"]] == ["H36", "H37"] and m["extra"] == ["notes.txt"]
    assert describe_match(m) == "1 of 3 input files matched; missing: 2026Q1_2026_BBA_02.xlsm, Ctrl_01!H37; not used: notes.txt."
    rows = [_table_row(f"H{35 + i}", [s]) for i, s in enumerate(TOKYO_SHEETS)]
    m = match_files(rows, [f"2026Q1_2026_{s}.xlsm" for s in reversed(TOKYO_SHEETS)])
    assert [r["file"] for r in m["rows"]] == [f"2026Q1_2026_{s}.xlsm" for s in TOKYO_SHEETS]
    assert describe_match(m) == "6 of 6 input files matched."


# --- 4. reading the macro's run from Excel's events ----------------------------------------------------
def _tokyo_log(picked: set[str]) -> list[tuple]:
    """calcCF's run as Excel's events reported it: per row, the input area cleared (1,200 months), the
    file opened, its block pasted (1,201 months), the file closed, the status cell stamped. A file not
    picked does not exist: no open, the status cell says so."""
    log: list[tuple] = []
    for i, (sheet, cleared, pasted) in enumerate(TOKYO_RUN):
        path = rf"C:\in\IF_当期末（非金融前提変更後）\2026Q1_2026_{sheet}.xlsm"
        log.append(("change", BOOK, sheet, cleared))
        if sheet in picked:
            log += [("open", path), ("change", BOOK, sheet, pasted), ("close", path)]
        log.append(("change", BOOK, "Ctrl_01", f"G{35 + i}"))
    return log


def _names(sheets) -> set[str]:
    return {normalize_name(f"2026Q1_2026_{s}.xlsm") for s in sheets}


def test_each_files_area_on_a_tokyo_shaped_run():
    seen = analyse_events(_tokyo_log(set(TOKYO_SHEETS)), BOOK, _names(TOKYO_SHEETS))
    assert seen["orphans"] == [] and set(seen["files"]) == _names(TOKYO_SHEETS)
    bba = seen["files"][normalize_name("2026Q1_2026_BBA_04.xlsm")]
    assert bba == {"areas": [{"sheet": "BBA_04", "ref": "AB3:DP1203", "r1": 3, "c1": 28, "r2": 1203, "c2": 120, "rows": 1201, "cols": 93}],
                   "cleared": [{"sheet": "BBA_04", "ref": "AB3:DP1202"}]}
    vfa = seen["files"][normalize_name("2026Q1_2026_VFA_01.xlsm")]
    assert vfa["areas"] == [{"sheet": "VFA_01", "ref": "AE3:ABS1203", "r1": 3, "c1": 31, "r2": 1203, "c2": 747, "rows": 1201, "cols": 717}]
    assert vfa["cleared"] == [{"sheet": "VFA_01", "ref": "AE3:ABS1202"}]


def test_a_file_not_picked_leaves_its_clear_as_an_orphan():
    picked = set(TOKYO_SHEETS) - {"BBA_02"}
    seen = analyse_events(_tokyo_log(picked), BOOK, _names(picked))
    assert seen["orphans"] == [("BBA_02", (3, 28, 1202, 120))]
    assert normalize_name("2026Q1_2026_BBA_02.xlsm") not in seen["files"]
    assert seen["files"][normalize_name("2026Q1_2026_BBA_03.xlsm")]["cleared"] == [{"sheet": "BBA_03", "ref": "AB3:DP1202"}]


def test_one_file_writing_two_sheets_has_two_areas_and_other_events_are_ignored():
    log = [
        ("open", r"C:\IN\2026Q1_2026_bba_01.XLSM"),  # the macro's own spelling of the picked name
        ("change", BOOK, "BBA_01", "$AB$3:$DP$1203"),
        ("change", "2026Q1_2026_bba_01.XLSM", "CF_Report", "A1:B2"),  # a change in the input file itself
        ("change", BOOK, "BBA_02", "AB3:AB10"),
        ("change", BOOK, "BBA_01", "A:A"),  # a whole column: no area
        ("close", r"C:\IN\2026Q1_2026_bba_01.XLSM"),
        ("open", r"C:\IN\helper.xlsx"),  # a file nobody picked
        ("change", BOOK, "BBA_03", "C3:D4"),  # written while no picked file is open: a clear for nobody
        ("close", r"C:\IN\helper.xlsx"),
        ("change", BOOK, "Ctrl_01", "G35"),  # one cell outside a file: a status stamp, not a clear
    ]
    seen = analyse_events(log, BOOK, _names(["BBA_01"]))
    areas = seen["files"][normalize_name("2026Q1_2026_BBA_01.xlsm")]["areas"]
    assert [(a["sheet"], a["ref"]) for a in areas] == [("BBA_01", "AB3:DP1203"), ("BBA_02", "AB3:AB10")]
    assert seen["orphans"] == [("BBA_03", (3, 3, 4, 4))]


def test_a_clear_over_two_written_pieces_makes_one_area():
    log = [("change", BOOK, "S", "B3:Z100"), ("open", r"C:\a.xlsm"),
           ("change", BOOK, "S", "B3:B101"), ("change", BOOK, "S", "D3:D101"), ("close", r"C:\a.xlsm")]
    areas = analyse_events(log, BOOK, {"a.xlsm"})["files"]["a.xlsm"]["areas"]
    assert [(a["sheet"], a["ref"]) for a in areas] == [("S", "B3:Z101")]


def test_columns_written_out_of_order_make_one_area():
    log = [("open", r"C:\a.xlsm"), ("change", BOOK, "S", "E3:E10"), ("change", BOOK, "S", "C3:C10"),
           ("change", BOOK, "S", "D3:D10"), ("close", r"C:\a.xlsm")]
    areas = analyse_events(log, BOOK, {"a.xlsm"})["files"]["a.xlsm"]["areas"]
    assert [(a["sheet"], a["ref"]) for a in areas] == [("S", "C3:E10")]


# --- 5. pointing the readers at the Mind input grid ----------------------------------------------------
def test_tokyo_readers_are_pointed_at_the_input_grid():
    o5 = '=(-SUMIF($AB$1:$DP$1,"保険料(＋)",$AB5:$DP5)+SUMIF($AB$1:$DP$1,"保険料(－)",$AB5:$DP5))'
    assert redirect_formula(o5, "BBA_04", [BBA_04]) == (
        '=(-SUMIF($AB$1:$DP$1,"保険料(＋)",IN_BBA_04!$A5:$CO5)+SUMIF($AB$1:$DP$1,"保険料(－)",IN_BBA_04!$A5:$CO5))', 2, None)
    n4 = "=IF($B$8=0,N6*1/3,SUMIF($AD$1:$DP$1,N$2,$AD4:$DP4)-SUMIF($AD$1:$DP$1,N$2,$AD5:$DP5))"
    assert redirect_formula(n4, "BBA_04", [BBA_04]) == (
        "=IF($B$8=0,N6*1/3,SUMIF($AD$1:$DP$1,N$2,IN_BBA_04!$C4:$CO4)-SUMIF($AD$1:$DP$1,N$2,IN_BBA_04!$C5:$CO5))", 2, None)


def test_references_from_other_sheets_and_japanese_sheet_names():
    assert redirect_formula("=BBA_04!AC5*2", "集計", [BBA_04]) == ("=IN_BBA_04!B5*2", 1, None)
    jp = _area("IN_FILE1", "保険料 明細", "B3:D10")
    assert redirect_formula("=SUM('保険料 明細'!B5:D5)", "集計", [jp]) == ("=SUM(IN_FILE1!A5:C5)", 1, None)
    assert redirect_formula("=B5+1", "保険料 明細", [jp]) == ("=IN_FILE1!A5+1", 1, None)  # unqualified, on its own sheet
    assert redirect_formula("=Ctrl_01!AC5", "BBA_04", [BBA_04]) == ("=Ctrl_01!AC5", 0, None)  # same cell, other sheet
    assert redirect_formula("=BBA_04!AC5+VFA_01!AF9", "集計", [BBA_04, VFA_01]) == ("=IN_BBA_04!B5+IN_VFA_01!B9", 2, None)
    assert redirect_formula("=BBA_04!$AB$3:$DP$1203", "", [BBA_04]) == ("=IN_BBA_04!$A$3:$CO$1203", 1, None)  # a defined name


def test_what_lies_outside_the_area_or_in_a_string_is_untouched():
    assert redirect_formula("=AA5+AC5+AB2+DQ5+AC1204", "BBA_04", [BBA_04]) == ("=AA5+IN_BBA_04!B5+AB2+DQ5+AC1204", 1, None)
    assert redirect_formula('="AC5"&AC5', "BBA_04", [BBA_04]) == ('="AC5"&IN_BBA_04!B5', 1, None)
    assert redirect_formula("=$AB5+DP$5+$AC$7", "BBA_04", [BBA_04]) == ("=IN_BBA_04!$A5+IN_BBA_04!CO$5+IN_BBA_04!$B$7", 3, None)
    assert redirect_formula("=SUM(A:A)", "BBA_04", [BBA_04]) == ("=SUM(A:A)", 0, None)
    assert redirect_formula(r"='C:\x\[b.xlsm]BBA_04'!AC5", "Sheet1", [BBA_04]) == (r"='C:\x\[b.xlsm]BBA_04'!AC5", 0, None)
    assert redirect_formula("AC5", "BBA_04", [BBA_04]) == ("AC5", 0, None)  # not a formula
    assert redirect_formula("", "BBA_04", [BBA_04]) == ("", 0, None)


@pytest.mark.parametrize("formula, why", [
    ("=SUM(AA5:AC5)", "BBA_04!AA5:AC5 covers only part of the import area BBA_04!AB3:DP1203"),
    ("=SUM(AC:AC)", "BBA_04!AC:AC covers only part of the import area BBA_04!AB3:DP1203"),
    ("=SUM(5:5)", "BBA_04!5:5 covers only part of the import area BBA_04!AB3:DP1203"),
    ("=OFFSET(AC5,0,1)", "OFFSET(IN_BBA_04!B5, ...) may step outside the import area"),
])
def test_a_reader_that_cannot_follow_says_why_and_changes_nothing(formula, why):
    assert redirect_formula(formula, "BBA_04", [BBA_04]) == (formula, 0, why)


@pytest.mark.parametrize("formula", ["=SUM(BBA_01:BBA_04!AC5)", "=SUM('BBA_04:合計>>'!AC5)"])
def test_a_3d_reference_over_an_area_cannot_follow(formula):
    new, n, why = redirect_formula(formula, "集計", [BBA_04])
    assert (new, n) == (formula, 0) and why


def test_a_link_to_another_workbook_is_not_a_reader():
    assert redirect_formula("=[1]BBA_04!AC5", "Sheet1", [BBA_04]) == ("=[1]BBA_04!AC5", 0, None)


def test_reads_area_and_partly_covered():
    assert oi._reads_area("=AC5", "BBA_04", BBA_04) and oi._reads_area("=SUM(AB:AB)", "BBA_04", BBA_04)
    assert not oi._reads_area("=Ctrl_01!AC5", "BBA_04", BBA_04) and not oi._reads_area("=SUM(AA:AA)", "BBA_04", BBA_04)
    culprits = oi._partly_covered("=SUM(BBA_04!AA5:AC5)+VFA_01!AF5+OFFSET(VFA_01!AF6,1,1)", "S", [BBA_04, VFA_01])
    assert [a["name"] for a in culprits] == ["IN_BBA_04", "IN_VFA_01"]  # a partial range; OFFSET from inside
    assert oi._partly_covered("=SUM(BBA_04!AC5:AD9)", "S", [BBA_04]) == []


# --- SUMIF / SUMIFS running past the side of an area -----------------------------------------------------
def test_the_vfa_reader_is_split_at_the_side_of_the_area_then_pointed_at_the_grid():
    n4 = "=SUMIF($AD$1:$ABS$1,N$2,$AD4:$ABS4)-SUMIF($AD$1:$ABS$1,N$2,$AD5:$ABS5)"
    split = split_partial_sums(n4, "VFA_01", [VFA_01])
    assert split == ("=(SUMIF($AD$1:$AD$1,N$2,$AD4:$AD4)+SUMIF($AE$1:$ABS$1,N$2,$AE4:$ABS4))"
                     "-(SUMIF($AD$1:$AD$1,N$2,$AD5:$AD5)+SUMIF($AE$1:$ABS$1,N$2,$AE5:$ABS5))")
    assert redirect_formula(n4, "VFA_01", [VFA_01])[2]  # unsplit it cannot follow
    assert redirect_formula(split, "VFA_01", [VFA_01]) == (
        "=(SUMIF($AD$1:$AD$1,N$2,$AD4:$AD4)+SUMIF($AE$1:$ABS$1,N$2,IN_VFA_01!$A4:$AAO4))"
        "-(SUMIF($AD$1:$AD$1,N$2,$AD5:$AD5)+SUMIF($AE$1:$ABS$1,N$2,IN_VFA_01!$A5:$AAO5))", 2, None)


@pytest.mark.parametrize("formula, area, split", [
    ("=SUMIFS($AD5:$ABS5,$AD$1:$ABS$1,N$2)", VFA_01, "=(SUMIFS($AD5:$AD5,$AD$1:$AD$1,N$2)+SUMIFS($AE5:$ABS5,$AE$1:$ABS$1,N$2))"),
    ('=SUMIF($AD5:$ABS5,">0")', VFA_01, '=(SUMIF($AD5:$AD5,">0")+SUMIF($AE5:$ABS5,">0"))'),
    ("=SUMIF($AD$1:$ABS$1,N$2,$AD5:$ABS5)", _area("IN_VFA_01", "VFA_01", "AE3:ABR1203"),  # past both sides
     "=(SUMIF($AD$1:$AD$1,N$2,$AD5:$AD5)+SUMIF($AE$1:$ABR$1,N$2,$AE5:$ABR5)+SUMIF($ABS$1:$ABS$1,N$2,$ABS5:$ABS5))"),
    ("=SUMIF('変額 VFA'!$AD$1:$ABS$1,N$2,'変額 VFA'!$AD4:$ABS4)", _area("IN_VFA", "変額 VFA", "AE3:ABS1203"),
     "=(SUMIF('変額 VFA'!$AD$1:$AD$1,N$2,'変額 VFA'!$AD4:$AD4)+SUMIF('変額 VFA'!$AE$1:$ABS$1,N$2,'変額 VFA'!$AE4:$ABS4))"),
    ('=ROUND(_xlfn.SUMIFS($AD5:$ABS5,$AD$1:$ABS$1,"a,b"),2)', VFA_01,
     '=ROUND((SUMIFS($AD5:$AD5,$AD$1:$AD$1,"a,b")+SUMIFS($AE5:$ABS5,$AE$1:$ABS$1,"a,b")),2)'),
])
def test_sumif_and_sumifs_splits(formula, area, split):
    assert split_partial_sums(formula, "集計" if "変額" in formula else "VFA_01", [area]) == split


@pytest.mark.parametrize("formula", [
    "=SUMIF($AD$1:$ABS$1,N$2,$AD5:$ABT5)",  # ranges of different shapes
    "=SUMIF($AE$1:$ABS$1,N$2,$AE5:$ABS5)",  # inside already
    "=SUMIF($AD$1:$ABS$1,N$2,$AD2:$ABS2)",  # rows outside the area
    "=SUMIF($AD5:$ABS6,N$2,$AD2:$ABS3)",  # rows partly outside
    "=SUM($AD5:$ABS5)",
    "",
    "SUMIF(A1:B1,1)",
])
def test_nothing_to_split(formula):
    assert split_partial_sums(formula, "VFA_01", [VFA_01]) is None


def test_rows_partly_outside_stop_the_split_whatever_the_argument_order():
    assert split_partial_sums("=SUMIF($AD2:$ABS3,N$2,$AD5:$ABS6)", "VFA_01", [VFA_01]) is None


# --- planning every reader at once ------------------------------------------------------------------------
TOKYO_READERS = [
    {"sheet": "BBA_04", "cell": "N4", "formula": "=IF($B$8=0,N6*1/3,SUMIF($AD$1:$DP$1,N$2,$AD4:$DP4)-SUMIF($AD$1:$DP$1,N$2,$AD5:$DP5))"},
    {"sheet": "BBA_04", "cell": "O5", "formula": '=(-SUMIF($AB$1:$DP$1,"保険料(＋)",$AB5:$DP5)+SUMIF($AB$1:$DP$1,"保険料(－)",$AB5:$DP5))'},
    {"sheet": "VFA_01", "cell": "N4", "formula": "=SUMIF($AD$1:$ABS$1,N$2,$AD4:$ABS4)-SUMIF($AD$1:$ABS$1,N$2,$AD5:$ABS5)"},
    {"sheet": "VFA_01", "cell": "O4", "formula": "=-SUMIF($AD$1:$ABS$1,O$2,$AD4:$ABS4)"},
    {"sheet": "集計", "cell": "B2", "formula": "=SUM(BBA_04!N4:N10)"},  # reads no area
]


def test_plan_rewrites_on_tokyo_readers_and_names():
    p = plan_rewrites(TOKYO_READERS, [BBA_04, VFA_01], [("CF_BBA_04", "=BBA_04!$AB$3:$DP$1203"), ("Header", "=BBA_04!$AB$1:$DP$1")])
    assert [a["name"] for a in p["areas"]] == ["IN_BBA_04", "IN_VFA_01"] and p["kept"] == []
    assert p["readers"] == {"IN_BBA_04": 2, "IN_VFA_01": 2} and p["sums_split"] == 2
    assert p["names"] == {"CF_BBA_04": "=IN_BBA_04!$A$3:$CO$1203"}
    assert p["rewrites"] == {
        ("BBA_04", "N4"): "=IF($B$8=0,N6*1/3,SUMIF($AD$1:$DP$1,N$2,IN_BBA_04!$C4:$CO4)-SUMIF($AD$1:$DP$1,N$2,IN_BBA_04!$C5:$CO5))",
        ("BBA_04", "O5"): '=(-SUMIF($AB$1:$DP$1,"保険料(＋)",IN_BBA_04!$A5:$CO5)+SUMIF($AB$1:$DP$1,"保険料(－)",IN_BBA_04!$A5:$CO5))',
        ("VFA_01", "N4"): ("=(SUMIF($AD$1:$AD$1,N$2,$AD4:$AD4)+SUMIF($AE$1:$ABS$1,N$2,IN_VFA_01!$A4:$AAO4))"
                           "-(SUMIF($AD$1:$AD$1,N$2,$AD5:$AD5)+SUMIF($AE$1:$ABS$1,N$2,IN_VFA_01!$A5:$AAO5))"),
        ("VFA_01", "O4"): "=-(SUMIF($AD$1:$AD$1,O$2,$AD4:$AD4)+SUMIF($AE$1:$ABS$1,O$2,IN_VFA_01!$A4:$AAO4))",
    }


def test_a_reader_that_cannot_follow_keeps_only_its_own_area():
    readers = TOKYO_READERS + [{"sheet": "集計", "cell": "C2", "formula": "=SUM(VFA_01!AD5:AF5)+BBA_04!AC9"}]
    p = plan_rewrites(readers, [BBA_04, VFA_01])
    assert [a["name"] for a in p["areas"]] == ["IN_BBA_04"]
    assert p["kept"] == [{"sheet": "VFA_01", "ref": "AE3:ABS1203", "name": "IN_VFA_01",
                          "why": "集計!C2: VFA_01!AD5:AF5 covers only part of the import area VFA_01!AE3:ABS1203"}]
    assert p["rewrites"][("集計", "C2")] == "=SUM(VFA_01!AD5:AF5)+IN_BBA_04!B9"  # the moving area only
    assert ("VFA_01", "N4") not in p["rewrites"] and p["sums_split"] == 0


def test_array_formulas_partial_names_and_offset_keep_their_area():
    p = plan_rewrites([{"sheet": "S", "cell": "C1", "formula": "=SUM(BBA_04!AC5:AC6)", "array_ref": "C1:C2"}],
                      [BBA_04, VFA_01], [("Wide", "=VFA_01!$AD$3:$AF$10")])
    assert p["areas"] == [] and p["kept"] == [
        {"sheet": "BBA_04", "ref": "AB3:DP1203", "name": "IN_BBA_04", "why": "the array formula at S!C1 reads it"},
        {"sheet": "VFA_01", "ref": "AE3:ABS1203", "name": "IN_VFA_01", "why": "the name Wide covers part of it"}]
    p = plan_rewrites([{"sheet": "BBA_04", "cell": "C1", "formula": "=OFFSET(AC5,1,0)+VFA_01!AF5"}], [BBA_04, VFA_01])
    assert [k["name"] for k in p["kept"]] == ["IN_BBA_04"] and "OFFSET" in p["kept"][0]["why"]
    assert p["rewrites"] == {("BBA_04", "C1"): "=OFFSET(AC5,1,0)+IN_VFA_01!B5"}


# --- the CSV Mind's Input Manager takes ----------------------------------------------------------------
@pytest.mark.parametrize("value, text", [
    (1 / 3, "0.3333333333333333"),
    (157.66085262000001, "157.66085262000001"),
    (5.5826458e-05, "0.000055826458"),
    (-1234.5678901234567, "-1234.5678901234567"),
    (1e-12, "0.000000000001"),
    (1e20, "100000000000000000000"),
    (1.5e16, "15000000000000000"),
    (2026.0, "2026"),
    (-5.0, "-5"),
    (0.0, "0"),
    (7, "7"),
    (0.1, "0.1"),
])
def test_csv_number_is_plain_decimal_that_reads_back_exactly(value, text):
    assert csv_number(value) == text
    assert float(text) == value


def test_csv_number_round_trips_any_double_without_an_exponent():
    rng = random.Random(20261109)
    values = [rng.uniform(-1, 1) * 10.0 ** rng.randint(-30, 30) for _ in range(5000)] + [2.0 ** -1074, 1.7976931348623157e308]
    for v in values:
        t = csv_number(v)
        assert "e" not in t.lower() and float(t) == v, (v, t)
    assert csv_number(float("nan")) == "" and csv_number(float("inf")) == "" and csv_number(float("-inf")) == ""


def test_write_input_csv_comma_utf8_without_bom(tmp_path):
    info = write_input_csv(tmp_path / "IN_BBA_01", ["保険料(＋)", "年月"], [[1, 2.5], [None, "x"], [True, False], [5.5826458e-05, "007"]])
    path = tmp_path / "IN_BBA_01.csv"
    assert info == {"path": str(path), "rows": 4, "cols": 2, "separator": "comma", "text_that_looks_like_numbers": 1}
    raw = path.read_bytes()
    assert not raw.startswith(b"\xef\xbb\xbf") and b"\r" not in raw
    assert raw.decode("utf-8") == "保険料(＋),年月\n1,2.5\n,x\nTRUE,FALSE\n0.000055826458,007\n"


def test_write_input_csv_switches_to_tab_txt_for_a_comma_or_a_quote(tmp_path):
    info = write_input_csv(tmp_path / "IN_X", ["a", "b"], [["1,5", "q"]])
    assert (info["path"], info["separator"]) == (str(tmp_path / "IN_X.txt"), "TAB")
    assert (tmp_path / "IN_X.txt").read_bytes() == b"a\tb\n1,5\tq\n"
    assert write_input_csv(tmp_path / "IN_Q.csv", ["a", "b"], [['say "hi"', 1]])["path"] == str(tmp_path / "IN_Q.txt")
    info = write_input_csv(tmp_path / "IN_T", ["a", "b"], [["x\ty", 1]])  # a TAB alone is fine in a comma file
    assert info["separator"] == "comma" and (tmp_path / "IN_T.csv").read_bytes() == b"a,b\nx\ty,1\n"


@pytest.mark.parametrize("data, says", [
    ([["a\nb", 1]], "line break"),
    ([["a\rb", 1]], "line break"),
    ([["a,b", "c\td"]], "TAB and a comma"),
])
def test_write_input_csv_refuses_what_mind_cannot_read(tmp_path, data, says):
    with pytest.raises(ValueError, match=says):
        write_input_csv(tmp_path / "IN_E", ["a", "b"], data)
    assert list(tmp_path.iterdir()) == []


# --- names of the Mind inputs ---------------------------------------------------------------------------
def test_plan_mind_inputs_names_are_unique_ascii_whatever_the_sheet_name():
    def area(sheet: str, ref: str) -> dict:
        a = _area("", sheet, ref)
        return {k: a[k] for k in ("sheet", "ref", "r1", "c1", "r2", "c2", "rows", "cols")}

    report = {"files": [
        {"file": "a.xlsm", "areas": [area("BBA_01", "AB3:DP1203"), area("BBA_01", "A2000:B2001")]},
        {"file": "b.xlsm", "areas": [area("保険料", "B3:C4")]},
        {"file": "c.xlsm", "areas": [area("ＶＦＡ＿０１", "B3:C4")]},
        {"file": "d.xlsm", "areas": [area("BBA_結合前 (死亡)", "B3:C4")]},
        {"file": "e.xlsm", "areas": [area("דוח שנתי", "B3:C4")]},
        {"file": "f.xlsm", "areas": [area("A" * 31, "B3:C4"), area("A" * 31, "E3:F4")]},
        {"file": None, "areas": []},  # a row without a file: nothing to plan
    ]}
    plan = plan_mind_inputs(report, ["BBA_01", "in_bba_01_2", "IN_FILE3", "保険料"])
    assert [(p["name"], p["sheet"], p["ref"], p["file"]) for p in plan] == [
        ("IN_BBA_01", "BBA_01", "AB3:DP1203", "a.xlsm"),
        ("IN_BBA_01_3", "BBA_01", "A2000:B2001", "a.xlsm"),  # _2 is a sheet already (any case)
        ("IN_FILE3_2", "保険料", "B3:C4", "b.xlsm"),  # no ASCII in the name; IN_FILE3 is taken
        ("IN_VFA_01", "ＶＦＡ＿０１", "B3:C4", "c.xlsm"),  # full-width read as ASCII
        ("IN_BBA", "BBA_結合前 (死亡)", "B3:C4", "d.xlsm"),
        ("IN_FILE6", "דוח שנתי", "B3:C4", "e.xlsm"),
        ("IN_" + "A" * 28, "A" * 31, "B3:C4", "f.xlsm"),  # a sheet name holds 31 characters at most
        ("IN_" + "A" * 25 + "_2", "A" * 31, "E3:F4", "f.xlsm"),
    ]
    assert plan[0] == {"name": "IN_BBA_01", "file": "a.xlsm", "sheet": "BBA_01", "ref": "AB3:DP1203",
                       "r1": 3, "c1": 28, "r2": 1203, "c2": 120, "rows": 1201, "cols": 93}
    names = [p["name"] for p in plan]
    assert len({normalize_name(n) for n in names}) == len(names)
    assert all(len(n) <= 31 and n.isascii() and n.replace("_", "").isalnum() for n in names)


# --- the whole chain on the Tokyo fake data (Excel) ---------------------------------------------------------
FAKE = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "MindReady" / "tokyo_inputs_2026-11-09"
FAKE_MAIN = FAKE / "SBI_IFRS17_Anonymized_original.xlsm"
FAKE_INPUTS = FAKE / "fake_inputs" / "2026_1Q" / "IF_当期末（非金融前提変更後）"
needs_tokyo = pytest.mark.skipif(
    not (FAKE_MAIN.is_file() and all((FAKE_INPUTS / f"2026Q1_2026_{s}.xlsm").is_file() for s in TOKYO_SHEETS)),
    reason="the Tokyo fake test data is not on this machine")


@needs_excel
@needs_tokyo
def test_tokyo_fake_inputs_end_to_end(tmp_path):
    """About 8 minutes: two scans, the tool's own calcCF in an Excel of its own, the Mind inputs, the numbers check."""
    from app.inventory import build_analysis
    from app.model_map import read_vba
    from app.numbers_check import Moves, compare_workbooks

    main = tmp_path / FAKE_MAIN.name
    shutil.copy2(FAKE_MAIN, main)
    inputs = tmp_path / FAKE_INPUTS.name  # the client's Japanese folder name
    inputs.mkdir()
    for s in TOKYO_SHEETS:
        shutil.copy2(FAKE_INPUTS / f"2026Q1_2026_{s}.xlsm", inputs / f"2026Q1_2026_{s}.xlsm")

    an = build_analysis(main, tmp_path / "scan0", "oi_e2e_upload")
    vba = scan_vba(read_vba(an["workbooks"][0]["copy_path"]))
    assert vba["import_macro"] == "calcCF" and vba["chain"] == ["calcCF", "InputCFFile"]
    table = import_table_from_analysis(an)
    assert [(r["sheet"], r["cell"], r["keys"], r["target_sheet"]) for r in table] == [
        ("Ctrl_01", f"H{35 + i}", [s], s) for i, s in enumerate(TOKYO_SHEETS)]
    match = match_files(table, sorted(p.name for p in inputs.iterdir()))
    assert describe_match(match) == "6 of 6 input files matched."
    assert all(r["how"] in ("exact", "key") and r["file"] == f"2026Q1_2026_{r['target_sheet']}.xlsm" for r in match["rows"])

    said: list[str] = []
    imported = tmp_path / ("import_" + main.name)
    files = {f"{r['sheet']}!{r['cell']}": inputs / r["file"] for r in match["rows"]}
    rep = oi.run_import(main, match["rows"], files, vba["import_macro"], imported, tmp_path / "work",
                        progress=said.append, messages=vba["messages"])
    assert rep["error"] is None and rep["ok"] and not rep["stopped_early"], rep
    assert rep["dialogs"] == [] and rep["kept"] == [] and rep["saved"] == str(imported) and imported.is_file()
    assert all(e["exact"] and e["opened"] for e in rep["files"]), rep["files"]  # Excel's name for each row = the picked file
    assert {e["cell"]: [(a["sheet"], a["ref"]) for a in e["areas"]] for e in rep["files"]} == {
        f"Ctrl_01!H{35 + i}": [(s, "AE3:ABS1203" if s == "VFA_01" else "AB3:DP1203")] for i, s in enumerate(TOKYO_SHEETS)}
    assert said and all(isinstance(m, str) for m in said)

    an1 = build_analysis(imported, tmp_path / "scan1", "oi_e2e_import")
    plan = plan_mind_inputs(rep, [s["name"] for s in an1["workbooks"][0]["sheets"]])
    assert [(p["name"], p["sheet"]) for p in plan] == [(f"IN_{s}", s) for s in TOKYO_SHEETS]
    final = tmp_path / ("mind_" + main.name)
    res = oi.apply_mind_inputs(imported, plan, an1["workbooks"][0]["formulas"], final, tmp_path / "csv")
    assert res["kept"] == [], res["kept"]
    assert [(i["name"], i["rows"], i["cols"]) for i in res["inputs"]] == [
        (f"IN_{s}", 1201, 717 if s == "VFA_01" else 93) for s in TOKYO_SHEETS]
    assert all(i["readers"] > 0 and Path(i["csv"]).is_file() for i in res["inputs"])
    # VFA_01's item names hold commas and quotes ('Total benefits, expenses, ...'): a TAB .txt
    assert {i["name"]: (i["separator"], Path(i["csv"]).suffix) for i in res["inputs"]} == {
        **{f"IN_{s}": ("comma", ".csv") for s in TOKYO_SHEETS[:5]}, "IN_VFA_01": ("TAB", ".txt")}
    assert res["formulas_changed"] > 0 and res["sums_split"] > 0 and res["saved"] == str(final)
    for i, sep, cols in ((res["inputs"][3], ",", 93), (res["inputs"][5], "\t", 717)):
        lines = Path(i["csv"]).read_text(encoding="utf-8").splitlines()
        assert len(lines) == 1 + 1201 and len(lines[0].split(sep)) == cols

    wb = openpyxl.load_workbook(final, read_only=True)
    try:
        for s in TOKYO_SHEETS:
            assert wb[f"IN_{s}"]["A1"].value == f"#IN_{s} /Input"
    finally:
        wb.close()

    numbers = compare_workbooks(imported, final, Moves())
    assert numbers["ran"] and numbers["clean"] is True, numbers["verdict"]
