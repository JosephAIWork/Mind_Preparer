"""Stand-in input files made with the assistant (app/stand_in_inputs.py): the recipe, the proof, the
loop that tries again with what happened. The Excel steps and the assistant are replaced here; the
real ones run in the Tokyo check (scratchpad stand_ins_live.py, evidence in the deliverables)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from app import stand_in_inputs as si

KEYS = ["BBA_01", "VFA_01"]
ROWS = [{"sheet": "Ctrl_01", "cell": f"H{35 + i}", "row": 35 + i, "col": 8, "expected": f"2026Q1_2026_{k}.xlsm", "keys": [k],
         "target_sheet": k, "file": None} for i, k in enumerate(KEYS)]
NEED = ["Ctrl_01!H35", "Ctrl_01!H36"]
SHEETS = ["Ctrl_01", "BBA_01", "VFA_01"]
VBA = [("M1_Main", '''Sub calcCF()
    Dim i As Long
    For i = 1 To 2
        Call InputCFFile(Sheets("Ctrl_01").Cells(34 + i, 8).Value)
    Next i
End Sub
'''), ("M2_InputCFFile", '''Sub InputCFFile(ByVal filePath As String)
    ' read the file
    Workbooks.Open filePath
    Set myRng = Cells.Find("Description").Cells(5, 1)
    ActiveWorkbook.Close
End Sub
Function Unused()
End Function
''')]


def recipe(**over) -> str:
    files = [
        {"row": "Ctrl_01!H35", "name": "2026Q1_2026_BBA_01.xlsm", "sheet": "CF", "cells": {"A1": "Description"},
         "copies": [{"to": "A5", "from": "BBA_01!AB2:DP2", "transpose": True}, {"to": "B5", "from": "BBA_01!AB3:DP1203", "transpose": True}]},
        {"row": "'Ctrl_01'!$H$36", "sheet": "CF", "cells": {"A1": "Product"},
         "zeros": {"to": "B3", "rows": 717, "cols": 1200}, "copies": [{"to": "A3", "from": "VFA_01!AE2:ABS2", "transpose": True}]},
    ]
    obj = {"files": files, "explain": "Items down, months across; the macro transposes them into row 3 on."}
    obj.update(over)
    return "Here it is:\n```json\n" + json.dumps(obj) + "\n```"


# --- the assistant's view of the model ------------------------------------------------------------
def test_a_row_is_summed_up_in_runs_of_the_same_kind():
    values = [None, "csv file", None, *[f"item {i}" for i in range(93)], 4.0]
    line = si.compress_row(2, 26, values)  # from column Z
    assert line.startswith('AA2="csv file"  AC2:DQ2 = 93 texts ("item 0", "item 1", "item 2", …, "item 92")  DR2=4')
    assert si.compress_row(3, 1, [None, None]) == "(empty)"
    assert si.compress_row(4, 1, [1.5, 2.0, "x"]) == 'A4=1.5  B4=2  C4="x"'


def test_the_macro_text_is_the_import_chain_and_what_it_calls():
    scan = si.core.scan_vba(VBA)
    code = si.chain_code(VBA, scan)
    assert "Sub calcCF()" in code and "Sub InputCFFile(ByVal filePath As String)" in code and code.count("End Sub") == 2
    assert "Unused" not in code and "read the file" not in code  # comments left out, uncalled code left out


def test_the_prompt_has_the_macro_the_table_the_sheets_and_the_last_attempt():
    model = {"sheets": SHEETS, "names": [{"name": "ReadCF", "refers_to": "=Ctrl_01!$C$34"}],
             "table": [{"key": "Ctrl_01!H35", "expected": "2026Q1_2026_BBA_01.xlsm", "keys": ["BBA_01"], "target_sheet": "BBA_01",
                        "formula": '="C:\\in\\"&D35&".xlsm"', "row": 'D35="BBA_01"'}],
             "shapes": [{"name": "BBA_01", "used": "A1:DP1203", "rows": ['row 2: AA2="年月"'], "numbers": ["AB3:DP1203"], "numbers_areas": 1,
                         "texts": ["AA2:DP2"], "texts_areas": 30, "formulas": [], "formulas_areas": 0}]}
    hist = [{"n": 1, "problem": "1,201 of 111,693 cells of BBA_01!AB3:DP1203 differ from the model", "details": ["- BBA_01!AB1203: model 0, file gave (empty)"],
             "recipe": {"files": []}}]
    text = si.build_prompt("Sub calcCF()\nEnd Sub", "calcCF", model, NEED, hist)
    for part in ("## The import macro", "`calcCF`", "ReadCF -> =Ctrl_01!$C$34", "Ctrl_01!H35: expected file \"2026Q1_2026_BBA_01.xlsm\"; fills sheet BBA_01",
                 "### BBA_01 -- used range A1:DP1203", "constant numbers: AB3:DP1203", "(+29 more areas)", "Ctrl_01!H35, Ctrl_01!H36",
                 "## Attempt 1 did not work", "BBA_01!AB1203: model 0, file gave (empty)", "answer with the whole recipe again"):
        assert part in text, part


# --- the recipe ------------------------------------------------------------------------------------
def test_a_recipe_is_read_and_checked():
    r = si.parse_recipe(recipe(), ROWS, NEED, SHEETS)
    bba, vfa = r["files"]
    assert bba["key"] == "Ctrl_01!H35" and bba["name"] == "2026Q1_2026_BBA_01.xlsm" and bba["sheet"] == "CF" and bba["cells"] == [(1, 1, "Description")]
    assert bba["copies"][1] == {"to": (5, 2), "sheet": "BBA_01", "rect": (3, 28, 1203, 120), "transpose": True, "rows": 93, "cols": 1201}
    # a row written another way still finds its row; a file without a name takes the expected one; zeros
    assert vfa["key"] == "Ctrl_01!H36" and vfa["name"] == "2026Q1_2026_VFA_01.xlsm"
    assert vfa["copies"][0] == {"to": (3, 2), "zeros": True, "rows": 717, "cols": 1200}
    assert r["explain"].startswith("Items down, months across")


@pytest.mark.parametrize("bad, says", [
    ({"files": [{"row": "Ctrl_01!H99", "copies": []}]}, "is not one of the rows asked for"),
    ({"files": [{"row": "Ctrl_01!H35", "copies": [{"to": "A1", "from": "BBA_01!A1:B2"}]}]}, "no file for Ctrl_01!H36"),
    ({"files": [{"row": "Ctrl_01!H35", "name": "..\\x.xlsm", "copies": []}]}, "a plain name"),
    ({"files": [{"row": "Ctrl_01!H35", "copies": [{"to": "A1", "from": "Nope!A1:B2"}]}]}, "the model has no sheet 'Nope'"),
    ({"files": [{"row": "Ctrl_01!H35", "copies": [{"to": "A1", "from": "A1:B2"}]}]}, "a range of the model with its sheet"),
    ({"files": [{"row": "Ctrl_01!H35", "cells": {"A1": "x"}}]}, "no block copied from the model and no zeros"),
    ({"files": [{"row": "Ctrl_01!H35", "sheet": "a/b", "copies": [{"to": "A1", "from": "BBA_01!A1"}]}]}, "a character Excel refuses"),
    ({"files": [{"row": "Ctrl_01!H35", "copies": [{"to": "A1:B2", "from": "BBA_01!A1"}]}]}, "is not one cell"),
])
def test_a_recipe_that_cannot_be_followed_says_why(bad, says):
    with pytest.raises(ValueError, match=r".*" + __import__("re").escape(says)):
        si.parse_recipe(json.dumps(bad), ROWS, NEED, SHEETS)
    with pytest.raises(ValueError, match="no JSON object|does not parse"):
        si.parse_recipe("I cannot tell", ROWS, NEED, SHEETS)


# --- the proof ----------------------------------------------------------------------------------------
def test_cells_compare_exactly_and_a_zero_where_the_model_is_empty_is_the_same():
    same = si.compare_matrices("BBA_01", (3, 28, 4, 29), [[1.5, None], ["a", 2.0]], [[1.5, 0.0], ["a", 2.0]])
    assert same["differ"] == 0 and same["zero_for_empty"] == 1 and same["model_empty"] is False and same["cells"] == 4 and same["ref"] == "AB3:AC4"
    empty = si.compare_matrices("VFA_01", (3, 31, 3, 32), [[None, None]], [[0, 0]])
    assert empty["differ"] == 0 and empty["model_empty"] is True
    bad = si.compare_matrices("BBA_01", (3, 28, 3, 30), [[1.0, 2.0, 3.0]], [[1.0, None, "3"]])
    assert bad["differ"] == 2 and bad["samples"][0] == {"cell": "BBA_01!AC3", "model": 2.0, "got": None} and bad["samples"][1]["cell"] == "BBA_01!AD3"


def test_the_verdict_names_what_went_wrong_for_the_assistant():
    report = {"files": [{"cell": "Ctrl_01!H35", "file": "a_BBA.xlsm", "opened": True, "areas": [{"sheet": "BBA_01", "ref": "AB3:DP1203"}]},
                        {"cell": "Ctrl_01!H36", "file": "b_VFA.xlsm", "opened": False, "areas": []}],
              "dialogs": [{"message": "x", "message_repaired": "BBA_01のinput領域には93項目まで入力可能"}], "stopped_early": True, "error": None}
    compared = [{**si.compare_matrices("BBA_01", (3, 28, 3, 28), [[1.0]], [[2.0]]), "key": "Ctrl_01!H35", "file": "a_BBA.xlsm"}]
    problem, details = si.judge(report, NEED, compared)
    assert problem == "b_VFA.xlsm was not opened; 1 of 1 cells of BBA_01!AB3 differ from the model"
    assert details[0] == '- The macro showed a message box: "BBA_01のinput領域には93項目まで入力可能"'
    assert any("the macro did not open it (it stopped before)" in d for d in details)
    assert any("BBA_01!AB3: model 1, file gave 2" in d for d in details)
    ok, _ = si.judge({"files": [{"cell": "Ctrl_01!H35", "file": "a", "opened": True, "areas": [{"sheet": "BBA_01", "ref": "A1"}]}], "dialogs": []},
                     ["Ctrl_01!H35"], [{**si.compare_matrices("BBA_01", (1, 1, 1, 1), [[1]], [[1]]), "key": "Ctrl_01!H35", "file": "a"}])
    assert ok is None


# --- the loop -------------------------------------------------------------------------------------------
class FakeExcel:
    """describe / build / run_import / compare / numbers, as the job calls them."""

    def __init__(self, differ_first=True, numbers=None):
        self.runs, self.prompts, self.differ_first = [], [], differ_first
        self.numbers_answers = list(numbers or [])

    def describe(self, model, rows, vba):
        return {"sheets": SHEETS, "names": [], "table": [], "shapes": []}

    def build(self, model, recipe, folder):
        folder.mkdir(parents=True, exist_ok=True)
        out = {}
        for f in recipe["files"]:
            p = folder / f["name"]
            p.write_bytes(b"built")
            out[f["key"]] = p
        return out

    def run_import(self, workbook, rows, files, macro, out_path, work_dir, messages=None, **_k):
        self.runs.append(sorted(files))
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(b"imported")
        area = {"Ctrl_01!H35": ("BBA_01", "AB3:DP1203"), "Ctrl_01!H36": ("VFA_01", "AE3:ABS1202")}
        return {"saved": str(out_path), "dialogs": [], "error": None, "stopped_early": False,
                "files": [{"cell": k, "file": Path(p).name, "opened": True, "areas": [{"sheet": area[k][0], "ref": area[k][1]}]} for k, p in files.items()]}

    def compare(self, model, imported, areas):
        first = len(self.runs) == 1 and self.differ_first
        out = []
        for a in areas:
            if a["key"] == "Ctrl_01!H35":
                c = si.compare_matrices(a["sheet"], (3, 28, 3, 28), [[5.0]], [[None if first else 5.0]])
            else:  # VFA_01: the model holds nothing there (its file was missing)
                c = si.compare_matrices(a["sheet"], (3, 31, 3, 31), [[None]], [[0.0]])
            out.append({**c, "key": a["key"], "file": a["file"]})
        return out

    def numbers(self, model, imported):
        return self.numbers_answers.pop(0) if self.numbers_answers else {"ran": True, "clean": True}


def _run(fake, answers, **kw):
    asked = list(answers)

    def ask(system, user):
        fake.prompts.append(user)
        assert "Never invent numbers" in system
        return asked.pop(0)

    return si.make_stand_ins(Path("model.xlsm"), ROWS, NEED, "calcCF", VBA, kw.pop("work"), ask, run_import=fake.run_import,
                             describe=fake.describe, build=fake.build, compare=fake.compare, numbers=fake.numbers,
                             attempts=kw.pop("attempts", 3), **kw)


def test_the_assistant_tries_again_with_what_happened_and_the_proof_passes(tmp_path):
    fake = FakeExcel(differ_first=True)
    said = []
    res = _run(fake, ["not json", recipe(), recipe()], work=tmp_path, progress=said.append)
    assert res["ok"] is True and sorted(res["files"]) == NEED and res["left_out"] == [] and res["explain"].startswith("Items down")
    assert [a["problem"] for a in res["attempts"]][:2] == [
        "its answer could not be used: no JSON object in the answer",
        "1 of 1 cells of BBA_01!AB3 differ from the model"]
    assert "## Attempt 1 did not work" in fake.prompts[1] and "## Attempt 2 did not work" in fake.prompts[2]
    assert "BBA_01!AB3: model 5, file gave (empty)" in fake.prompts[2]
    assert "Checking that no computed value of the model changes" in said


def test_zeros_where_the_model_is_empty_are_left_out_when_they_change_a_result(tmp_path):
    changed = {"ran": True, "clean": False, "counts": {"good_to_other": 2},
               "samples": {"good_to_other": [{"original": "Calc!B9", "was": 1.0, "is": 0.0}]}}
    fake = FakeExcel(differ_first=False, numbers=[changed, {"ran": True, "clean": True}])
    res = _run(fake, [recipe()], work=tmp_path)
    assert res["ok"] is True and list(res["files"]) == ["Ctrl_01!H35"] and res["left_out"] == ["Ctrl_01!H36"]
    assert fake.runs == [NEED, ["Ctrl_01!H35"]]  # checked again without the empty area's stand-in
    # and when that does not help either: a failed attempt, said with the first value that changes
    fake2 = FakeExcel(differ_first=False, numbers=[changed, changed] * 3)
    res2 = _run(fake2, [recipe()] * 3, work=tmp_path / "2")
    assert res2["ok"] is False and res2["problem"] == "2 computed values of the model would change, e.g. Calc!B9 would show 0 instead of 1"


def test_no_answer_is_a_plain_failure_with_the_gateways_words(tmp_path):
    res = _run(FakeExcel(), [None, None, None], work=tmp_path)
    assert res["ok"] is False and res["problem"] == "the assistant did not answer" and len(res["attempts"]) == 3 and res["files"] == {}
    assert si.MAX_ATTEMPTS == 4 and si.ASSISTANT_EFFORT == "low"

    def ask(system, user):
        raise RuntimeError("Request failed: Read timed out.")

    res2 = si.make_stand_ins(Path("model.xlsm"), ROWS, NEED, "calcCF", VBA, tmp_path / "2", ask, describe=FakeExcel().describe, attempts=1)
    assert res2["problem"] == "the assistant did not answer (Request failed: Read timed out.)"


def test_the_apps_assistant_says_why_it_returned_nothing(monkeypatch):
    from app import llm

    monkeypatch.setattr(llm, "llm_available", lambda: True)
    seen = {}

    def fake(messages, system, **kw):
        seen.update(kw)
        return {"available": True, "text": None, "message": "The model returned no text (stop_reason max_tokens)."}

    monkeypatch.setattr(llm, "chat_completion", fake)
    ask = si.assistant_ask()
    with pytest.raises(RuntimeError, match="stop_reason max_tokens"):
        ask("system", "user")
    assert seen["effort"] == "low" and seen["max_tokens"] == 12000
    monkeypatch.setattr(llm, "llm_available", lambda: False)
    assert si.assistant_ask() is None


def test_a_block_of_numbers_gets_0_in_its_empty_cells_labels_stay_as_they_are():
    assert si.zero_fill([[None, 1.5], ["", None]]) == [[0, 1.5], [0, 0]]
    assert si.zero_fill([[None, None]]) == [[0, 0]]  # an area that was empty in the model
    labels = [["Premium", None], [None, "Claims"]]
    assert si.zero_fill(labels) is labels


def test_a_walk_that_ran_past_the_data_is_named():
    report = {"files": [{"cell": "Ctrl_01!H40", "file": "v_VFA.xlsm", "opened": True, "areas": [{"sheet": "VFA_01", "ref": "AE3:ABS16385"}]}], "dialogs": []}
    compared = [{"sheet": "VFA_01", "ref": "AE3:ABS16385", "cells": 10, "differ": 1, "samples": [], "key": "Ctrl_01!H40", "file": "v_VFA.xlsm"}]
    problem, details = si.judge(report, ["Ctrl_01!H40"], compared)
    assert any("walk ran past the data" in d and "ABS16385" in d for d in details)
