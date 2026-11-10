"""Outside input files through the API (app/web/server.py + app/outside_inputs.py): the workbook and
the files its macro opens picked together, the main workbook found among them, its import table
matched (any language), POST .../inputs, the bring-in job (its Excel steps replaced here: no Excel
runs), INP-010 and the readiness next step, the numbers baselines that move to the version the files
were brought into, the downloads, and Run in Mind uploading the Mind inputs."""
from __future__ import annotations

import io
import shutil
import threading
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

from app import model_map, outside_inputs, stand_in_inputs
from app.estimate import estimate_mind_run
from app.readiness import compute_readiness
from app.validators.io import inp_010
from app.web import server

MAIN = "IFRS17計算ツール_ver7（テスト）.xlsx"  # any language: a Japanese name with full-width brackets
KEYS = ["BBA_01", "BBA_02", "BBA_03", "BBA_04", "BBA_05", "VFA_01"]
INPUTS = [f"2026Q1_2026_{k}.xlsm" for k in KEYS]
VBA = '''
Sub calcCF()
    Dim i As Long
    For i = 1 To 6
        Call InputCFFile(Sheets("Ctrl_01").Cells(34 + i, 8).Value)
    Next i
End Sub

Sub InputCFFile(ByVal filePath As String)
    Workbooks.Open filePath
    ActiveWorkbook.Close
End Sub
'''


@pytest.fixture(scope="module")
def client():
    return TestClient(server.app)


@pytest.fixture
def macros(monkeypatch):
    """The main workbook's VBA (olevba is replaced): calcCF opens one file per row of Ctrl_01!H35:H40."""
    monkeypatch.setattr(model_map, "read_vba", lambda p: [("M1_Main", VBA)] if Path(p).name == MAIN else [])
    monkeypatch.setattr(server, "has_vba_project", lambda p: Path(p).name == MAIN)


def _main_book(folder: Path, name: str = MAIN) -> Path:
    """A Tokyo-like tool: the import table on Ctrl_01 (key in D, path formula in H), the sheet the
    first file fills (BBA_01) and a formula that reads it."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ctrl_01"
    ws["A1"] = "SYNTHETIC_TEST_FIXTURE"
    ws["D34"], ws["H34"] = "Sheet", "File"
    for i, key in enumerate(KEYS):
        r = 35 + i
        ws[f"D{r}"] = key
        ws[f"H{r}"] = f'="C:\\inputs\\"&"2026Q1_2026_"&D{r}&".xlsm"'
    cf = wb.create_sheet("BBA_01")
    cf["A1"], cf["A2"], cf["B2"] = "#CF", "Month", "Amount"
    for r in range(3, 6):
        cf[f"A{r}"], cf[f"B{r}"] = r - 2, r * 10
    calc = wb.create_sheet("Calc")
    calc["A1"], calc["A2"], calc["A3"] = "#Total", "Total", "=SUM(BBA_01!B3:B5)"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    wb.save(path)
    wb.close()
    return path


def _part(name: str, data: bytes) -> tuple[str, tuple[str, io.BytesIO, str]]:
    return ("files", (name, io.BytesIO(data), "application/octet-stream"))


def _upload_picked(client, main: Path, inputs: list[str], extra: list[str] = (), defer: bool = False) -> dict:
    parts = [_part(n, b"an input file") for n in inputs] + [_part(n, b"x") for n in extra] + [_part(main.name, main.read_bytes())]
    res = client.post("/api/sessions", files=parts, data={"mode": "plan", **({"defer": "1"} if defer else {})})
    assert res.status_code == 200, res.text
    return res.json()


def _finding(report: dict, rule_id: str = "INP-010") -> dict:
    return next(f for f in report["findings"] if f["rule_id"] == rule_id)


def _wait(client, sid: str, timeout: float = 120) -> dict:
    deadline = time.time() + timeout
    while True:
        st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
        if not st["running"]:
            return st
        assert time.time() < deadline, st
        time.sleep(0.1)


# --- the Excel steps, replaced ------------------------------------------------------------------
def _fake_run_import(calls: list, gate: threading.Event | None = None, stopped: bool = False):
    def run_import(workbook, rows, files, macro, out_path, work_dir, timeout_s=900, progress=None, messages=None, restore_rows=None):
        calls.append({"workbook": Path(workbook), "rows": rows, "files": dict(files), "macro": macro, "messages": messages,
                      "restore_rows": sorted(restore_rows or [])})
        if gate is not None:
            assert gate.wait(60)
        if progress:
            progress(f"Running the workbook's import macro {macro}")
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(workbook, out_path)
        entries = []
        for r in rows:
            key = f"{r['sheet']}!{r['cell']}"
            p = files.get(key)
            opened = bool(p) and not (stopped and not key.endswith("H35"))
            areas = [{"sheet": "BBA_01", "ref": "B3:B5", "r1": 3, "c1": 2, "r2": 5, "c2": 2, "rows": 3, "cols": 1}] if p and key.endswith("H35") and not stopped else []
            entries.append({"cell": key, "file": p.name if p else None, "expected": None, "exact": False, "opened": opened, "areas": areas, "cleared": []})
        # a hidden Excel shows a Japanese message box as mojibake; the core repairs it (message_repaired)
        dialogs = [{"title": "Microsoft Excel", "message": "BBA_01\x82\xcc", "message_repaired": "BBA_01のinput領域には93項目まで入力可能です",
                    "buttons": ["OK"], "pressed": "OK"}] if stopped else []
        return {"macro": macro, "seconds": 0.1, "files": entries, "kept": [], "dialogs": dialogs, "stopped_early": stopped,
                "error": "Exception occurred." if stopped else None, "saved": str(out_path), "ok": not stopped}

    return run_import


def _fake_apply_mind_inputs(calls: list):
    def apply_mind_inputs(workbook, plan, formulas, out_path, csv_dir, progress=None):
        calls.append({"workbook": Path(workbook), "plan": plan, "formulas": list(formulas)})
        wb = openpyxl.load_workbook(workbook)
        csv_dir = Path(csv_dir)
        csv_dir.mkdir(parents=True, exist_ok=True)
        inputs = []
        for a in plan:
            ws = wb.create_sheet(a["name"])
            ws["A1"], ws["A2"] = f"#{a['name']} /Input", "Amount"
            for i in range(a["rows"]):
                ws.cell(3 + i, 1, (i + 3) * 10)
            csv = csv_dir / f"{a['name']}.csv"
            csv.write_bytes(b"Amount\n30\n40\n50\n")  # as the core writes it: UTF-8, LF, no BOM
            inputs.append({"name": a["name"], "sheet": a["sheet"], "ref": a["ref"], "file": a.get("file"), "rows": a["rows"], "cols": a["cols"],
                           "csv": str(csv), "separator": "comma", "readers": 1})
        wb["Calc"]["A3"] = f"=SUM({plan[0]['name']}!A3:A5)"
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
        wb.close()
        return {"inputs": inputs, "kept": [], "formulas_changed": 1, "names_changed": 0, "sums_split": 0, "saved": str(out_path)}

    return apply_mind_inputs


@pytest.fixture
def excel_steps(monkeypatch):
    """run_import / apply_mind_inputs / compare_workbooks replaced; Excel 'available'."""
    calls: dict[str, list] = {"import": [], "mind_inputs": [], "compare": []}
    monkeypatch.setattr(server, "com_available", lambda: True)
    monkeypatch.setattr(outside_inputs, "run_import", _fake_run_import(calls["import"]))
    monkeypatch.setattr(outside_inputs, "apply_mind_inputs", _fake_apply_mind_inputs(calls["mind_inputs"]))

    def compare(original, prepared, moves, progress=None):
        calls["compare"].append((Path(original), Path(prepared)))
        return {"ran": True, "clean": True, "verdict": "The preparation changed no computed value (1 formula cells compared with the original).",
                "counts": {"compared": 1, "same": 1, "good_to_error": 0, "good_to_other": 0, "missing_good": 0}, "regressions": []}

    monkeypatch.setattr(server, "compare_workbooks", compare)
    return calls


# --- 1. several files picked at once ---------------------------------------------------------------
def test_files_picked_together_find_the_workbook_keep_the_inputs_and_match_the_import_table(client, macros, tmp_path):
    main = _main_book(tmp_path)
    body = _upload_picked(client, main, INPUTS[:5], extra=["notes.csv"])
    sid = body["sessionId"]
    s = server.SESSIONS[sid]
    assert body["main"] == MAIN and [v["source"] for v in body["versions"]] == ["upload"]
    assert {i["name"] for i in body["inputs"]} == set(INPUTS[:5]) | {"notes.csv"} and all(i["bytes"] > 0 for i in body["inputs"])
    assert s.current_path.parent == s.work_dir / "upload" and s.current_path.name == MAIN
    for i in s.inputs:  # kept beside the session, one folder per file, never as versions
        p = Path(i["path"])
        assert p.is_file() and p.parent.parent == s.work_dir / "inputs" and len(i["sha256"]) == 64
    oi = body["outsideInputs"]
    assert oi["imports"] is True and oi["macro"] == "calcCF" and oi["calls"][0]["procedure"] == "InputCFFile"
    assert [r["cell"] for r in oi["rows"]] == [f"Ctrl_01!H{r}" for r in range(35, 41)]
    first = oi["rows"][0]
    assert first["keys"] == ["BBA_01"] and first["target_sheet"] == "BBA_01" and first["file"] == INPUTS[0] and first["how"] == "key"
    assert oi["matched"] == 5 and oi["missing"] == ["Ctrl_01!H40"] and oi["extra"] == ["notes.csv"] and oi["brought_in"] is False
    assert oi["summary"].startswith("5 of 6 input files matched; missing: VFA_01")
    # INP-010 blocks, and the next step is to bring the files in (before Prep moves anything)
    f = _finding(body["report"])
    assert f["status"] == "ERROR" and f["priority"] == "REQUIRED" and "5 of 6 picked, missing: VFA_01" in f["message"] and "Bring in the outside files" in f["message"]
    rd = body["readiness"]
    assert rd["state"] == "blocked" and rd["next_step"]["label"] == "Bring in the outside files" and rd["next_step"]["detail"] == "6 files the workbook's macro opens"
    assert next(b for b in rd["blocking"] if b["rule_id"] == "INP-010")["fix"] == "outside_inputs"
    # the same on a reopened session and on GET .../outside-inputs
    again = client.get(f"/api/sessions/{sid}").json()
    assert again["outsideInputs"] == oi and again["main"] == MAIN and len(again["inputs"]) == 6
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["outsideInputs"] == oi and st["running"] is False and st["state"] == "idle" and st["result"] is None and st["error"] is None


def test_a_deferred_upload_matches_after_its_scan_and_the_single_file_field_still_works(client, macros, tmp_path):
    main = _main_book(tmp_path)
    pending = _upload_picked(client, main, INPUTS, defer=True)
    assert pending["pending"] is True and pending["main"] == MAIN and len(pending["inputs"]) == 6
    out = client.post(f"/api/sessions/{pending['sessionId']}/analyze", json={})
    assert out.status_code == 200, out.text
    assert out.json()["outsideInputs"]["matched"] == 6 and _finding(out.json()["report"])["status"] == "ERROR"
    # "file" (scripts, tests) + "files": the workbook sent in both is not its own input
    with main.open("rb") as fh:
        res = client.post("/api/sessions", files=[("file", (MAIN, fh, "application/octet-stream")), _part(INPUTS[0], b"x"), _part(MAIN, b"same workbook")], data={"mode": "plan"})
    assert res.status_code == 200, res.text
    body = res.json()
    assert [i["name"] for i in body["inputs"]] == [INPUTS[0]] and body["outsideInputs"]["matched"] == 1


def test_a_workbook_whose_macros_open_nothing_has_no_outside_inputs(client, plain_grid_xlsx):
    with plain_grid_xlsx.open("rb") as fh:
        body = client.post("/api/sessions", files={"file": (plain_grid_xlsx.name, fh, "application/octet-stream")}, data={"mode": "plan"}).json()
    assert body["outsideInputs"] is None and body["inputs"] == [] and body["main"] == plain_grid_xlsx.name
    f = _finding(body["report"])
    assert f["status"] == "PASS" and f["message"].startswith("Nothing to bring in")
    sid = body["sessionId"]
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 422
    assert client.get(f"/api/sessions/{sid}/outside-inputs").json()["outsideInputs"] is None


def test_files_of_another_kind_are_refused_in_plain_words(client, macros, tmp_path):
    main = _main_book(tmp_path)
    res = client.post("/api/sessions", files=[_part("report.pdf", b"%PDF"), _part(MAIN, main.read_bytes())], data={"mode": "plan"})
    assert res.status_code == 400 and res.json()["detail"].startswith("report.pdf: Mind Ready takes the workbook and the files its macros open")
    res = client.post("/api/sessions", files=[_part(INPUTS[0].replace(".xlsm", ".csv"), b"a,b")], data={"mode": "plan"})
    assert res.status_code == 422 and "Pick the workbook itself" in res.json()["detail"]
    assert client.post("/api/sessions", data={"mode": "plan"}).status_code == 422


def test_the_main_workbook_is_the_one_whose_macros_open_outside_files(monkeypatch, tmp_path):
    big = tmp_path / "big_report.xlsx"
    big.write_bytes(b"x" * 5000)
    tool = tmp_path / "tool.xlsm"
    tool.write_bytes(b"x" * 100)
    other = tmp_path / "other_macros.xlsm"
    other.write_bytes(b"x" * 300)
    data = tmp_path / "data.csv"
    data.write_bytes(b"x" * 9000)
    vba = {"tool.xlsm": [("M", VBA)], "other_macros.xlsm": [("M", "Sub Hello()\n  MsgBox \"hi\"\nEnd Sub")]}
    monkeypatch.setattr(model_map, "read_vba", lambda p: vba.get(Path(p).name, []))
    monkeypatch.setattr(server, "has_vba_project", lambda p: Path(p).name in vba)
    main, found = server._pick_main([big, tool, other, data])
    assert main == tool and found["import_macro"] == "calcCF"
    del vba["tool.xlsm"]  # no macro opens a file: the only workbook with macros
    main, found = server._pick_main([big, tool, other, data])
    assert main == other and found["imports"] == []
    vba.clear()  # no macros at all: the largest workbook (a .csv is never the workbook)
    assert server._pick_main([big, tool, other, data]) == (big, None)


# --- 2. adding files later -----------------------------------------------------------------------------
def test_adding_the_missing_file_matches_it_without_a_new_scan(client, macros, tmp_path):
    main = _main_book(tmp_path)
    body = _upload_picked(client, main, INPUTS[:5])
    sid = body["sessionId"]
    scans = len(server.SESSIONS[sid].plan_history)
    res = client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[5], b"the VFA file"), _part(INPUTS[0], b"BBA_01 again, newer")])
    assert res.status_code == 200, res.text
    out = res.json()
    assert out["outsideInputs"]["matched"] == 6 and out["outsideInputs"]["missing"] == [] and out["outsideInputs"]["summary"] == "6 of 6 input files matched."
    names = [i["name"] for i in out["inputs"]]
    assert sorted(names) == sorted(INPUTS) and len(names) == 6  # replaced by name, not added twice
    s = server.SESSIONS[sid]
    assert Path(server._input_path(s, INPUTS[0])).read_bytes() == b"BBA_01 again, newer"
    assert len(s.plan_history) == scans  # no rescan
    f = _finding(s.result["validation_report"])  # INP-010 says the new counts at once
    assert f["status"] == "ERROR" and "6 of 6 picked." in f["message"]
    assert out["readiness"]["next_step"]["label"] == "Bring in the outside files"
    assert client.post(f"/api/sessions/{sid}/inputs", files=[_part("notes.docx", b"x")]).status_code == 400


# --- 3. the bring-in job ---------------------------------------------------------------------------------
def test_bring_in_makes_two_versions_the_mind_inputs_and_their_downloads(client, macros, excel_steps, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert res.status_code == 202 and res.json()["running"] is True and res.json()["base_version_id"] == "ver-001"
    st = _wait(client, sid)
    assert st["state"] == "done" and st["error"] is None, st
    (imp_call,) = excel_steps["import"]
    assert imp_call["macro"] == "calcCF" and imp_call["workbook"] == s.paths["ver-001"]
    assert imp_call["messages"] == []  # the VBA's own MsgBox lines (none in this one) go with the run
    assert sorted(imp_call["files"]) == [f"Ctrl_01!H{r}" for r in range(35, 41)] and all(p.is_file() for p in imp_call["files"].values())
    # the Mind inputs are made from the import version's own scan, every sheet read
    (mi_call,) = excel_steps["mind_inputs"]
    assert mi_call["workbook"] == s.paths["ver-002"] and [p["name"] for p in mi_call["plan"]] == ["IN_BBA_01"]
    assert any(f["sheet"] == "Calc" and f["cell"] == "A3" for f in mi_call["formulas"])
    result = st["result"]
    assert result["summary"] == "6 outside files brought in; 1 Mind input made (IN_BBA_01); no computed value changed."
    assert result["versions"] == ["ver-002", "ver-003"] == st["versions"]
    assert result["import"]["macro"] == "calcCF" and result["import"]["saved"] and all(f["opened"] for f in result["import"]["files"])
    assert result["mindInputs"]["inputs"][0]["csv"] == "IN_BBA_01.csv" and result["download"] == {"zip": "Mind inputs.zip", "files": ["IN_BBA_01.csv"]}
    assert result["numbers"]["clean"] is True and excel_steps["compare"] == [(s.paths["ver-002"], s.paths["ver-003"])]
    versions = {v["id"]: v for v in s.versions}
    assert versions["ver-002"]["source"] == "import" and versions["ver-002"]["label"] == "v2 — Outside files brought in" and versions["ver-002"]["parent"] == "ver-001"
    assert versions["ver-003"]["source"] == "mind_inputs" and versions["ver-003"]["label"] == "v3 — Outside files made Mind inputs" and versions["ver-003"]["parent"] == "ver-002"
    assert versions["ver-002"]["change_log"][0]["op"] == "bring_in_file" and versions["ver-003"]["change_log"][0]["after"] == "#IN_BBA_01 /Input"
    assert s.current_version_id == "ver-003"
    # the CSV folder of this line holds the CSVs alone (Run in Mind uploads the whole folder)
    assert s.mind_inputs_dir and server._mind_inputs_dir_for(s) == s.mind_inputs_dir
    assert sorted(p.name for p in Path(s.mind_inputs_dir).iterdir()) == ["IN_BBA_01.csv"]
    # brought in: INP-010 passes, the verdict no longer asks for it
    after = client.get(f"/api/sessions/{sid}").json()
    assert after["outsideInputs"]["brought_in"] is True and after["outsideInputs"]["brought_in_reason"] is None and _finding(after["report"])["status"] == "PASS"
    assert list(s.brought_in_files) == ["ver-003"]  # the bring-in got to the end, with the files it read
    assert all(b["rule_id"] != "INP-010" for b in after["readiness"]["blocking"])
    assert (after["readiness"]["next_step"] or {}).get("label") != "Bring in the outside files"
    # the numbers baseline is now the version the files were brought into; CELL/INFO stay with the upload
    assert server._original_path(s) == s.paths["ver-002"] and server._original_and_moves(s)[0] == s.paths["ver-002"]
    assert server._upload_and_moves(s)[0] == s.paths["ver-001"]
    # downloads: the zip (the CSVs at its root) and each CSV under the name Mind matches
    z = client.get(f"/api/sessions/{sid}/files/Mind inputs.zip")
    assert z.status_code == 200 and z.headers["content-type"] == "application/zip"
    assert zipfile.ZipFile(io.BytesIO(z.content)).namelist() == ["IN_BBA_01.csv"]
    c = client.get(f"/api/sessions/{sid}/files/IN_BBA_01.csv")
    assert c.status_code == 200 and c.headers["content-type"].startswith("text/csv") and c.content.startswith(b"Amount\n")
    v3 = client.get(f"/api/sessions/{sid}/files/{versions['ver-003']['file_name']}")
    assert v3.status_code == 200 and v3.content[:2] == b"PK"
    assert client.get(f"/api/sessions/{sid}/estimate").status_code == 200  # priced against the import version's baseline
    # and once more: the line is brought in already, so the macro runs on the version before it
    again = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"checkNumbers": False})
    assert again.status_code == 202 and again.json()["base_version_id"] == "ver-001" and "already brought in" in again.json()["note"]
    st2 = _wait(client, sid)
    assert st2["state"] == "done" and st2["result"]["versions"] == ["ver-004", "ver-005"] and st2["result"]["numbers"] is None
    assert st2["result"]["summary"].startswith("6 outside files brought in; 1 Mind input made (IN_BBA_01); the computed values are not checked yet.")
    assert {v["id"]: v for v in s.versions}["ver-004"]["parent"] == "ver-001"


def test_bring_in_is_refused_while_a_job_runs_and_refuses_the_others(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    gate = threading.Event()
    monkeypatch.setattr(outside_inputs, "run_import", _fake_run_import(excel_steps["import"], gate=gate))
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    try:
        st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
        assert st["running"] is True and st["stage"] in ("Starting", "Running the workbook's import macro")
        again = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
        assert again.status_code == 409 and "being brought in" in again.json()["detail"]
        assert client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[0], b"x")]).status_code == 409
        op = {"operations": [{"op": "set_value", "sheet": "Calc", "cell": "Z9", "after": 1}], "reanalyze": False}
        assert client.post(f"/api/sessions/{sid}/apply", json=op).status_code == 409
        assert client.post(f"/api/sessions/{sid}/recalculate").status_code == 409
        assert client.post(f"/api/sessions/{sid}/auto-fix", json={"use_assistant": False}).status_code == 409
        assert "outside files" in client.get(f"/api/sessions/{sid}/mind-loop/options").json()["busy"]
    finally:
        gate.set()
    assert _wait(client, sid)["state"] == "done"
    assert not server.SESSIONS[sid].apply_lock.locked()


def test_a_macro_that_stops_early_changes_nothing(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    monkeypatch.setattr(outside_inputs, "run_import", _fake_run_import(excel_steps["import"], stopped=True))
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "error" and st["result"] is None and st["versions"] == []
    assert st["error"].startswith('The workbook\'s import macro stopped early: "BBA_01のinput領域には93項目まで入力可能です".')
    assert "Not opened: 2026Q1_2026_BBA_02.xlsm" in st["error"] and st["error"].endswith("Nothing was changed: check the input files, then try again.")
    s = server.SESSIONS[sid]
    assert [v["id"] for v in s.versions] == ["ver-001"] and s.current_version_id == "ver-001" and not s.apply_lock.locked()
    assert _finding(s.result["validation_report"])["status"] == "ERROR" and excel_steps["mind_inputs"] == []
    timed_out = {"saved": None, "stopped_early": False, "files": [], "dialogs": [], "error": "the import macro ran longer than 900 s and was stopped"}
    assert server._import_stopped(timed_out) == "The workbook's import macro ran longer than 900 s and was stopped. Nothing was changed: check the input files, then try again."
    # every file opened, the last one written to nowhere, Excel's own words for the VBA error
    files = [{"file": "a.xlsm", "opened": True, "areas": [{"sheet": "A", "ref": "B3:C4"}]}, {"file": "b.xlsm", "opened": True, "areas": []}]
    error = "(-2147352567, 'Exception occurred.', (0, 'Microsoft Visual Basic', 'Subscript out of range', None, 0, -2146828279), None)"
    report = {"saved": "x", "stopped_early": False, "files": files, "dialogs": [], "error": error, "ok": False}
    assert not server._import_ok(report)
    assert server._import_stopped(report) == (
        "The workbook's import macro opened b.xlsm but wrote nothing from it: the macro looked for a sheet, workbook or list item that is not there "
        '(Excel said: "Subscript out of range"). Nothing was changed: check the input files, then try again.'
    )


def test_bring_in_needs_a_matched_file_and_excel(client, macros, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, ["unrelated.xlsx"])["sessionId"]
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert res.status_code == 422 and "add the outside files first" in res.json()["detail"]
    client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[0], b"x")])
    monkeypatch.setattr(server, "com_available", lambda: False)
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert res.status_code == 409 and "needs Excel" in res.json()["detail"]


# --- 4. the baselines and the version the macro runs on -------------------------------------------------
def _session(tmp_path: Path, steps: list[tuple[str, list[dict]]]) -> server.Session:
    """A session with one version per step (source, change log), each a file of its own."""
    s = server.Session(id=f"t{time.time_ns()}", work_dir=tmp_path, mode="plan")
    for i, (source, log) in enumerate(steps, start=1):
        p = tmp_path / f"v{i}" / "model.xlsm"
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"x")
        server._add_version(s, p, source, source, log, None, "0" * 64)
    return s


def test_the_baselines_move_to_the_version_the_files_were_brought_into(client, monkeypatch, tmp_path):
    insert = lambda row: [{"op": "insert_row", "sheet": "Data", "row": row}]  # noqa: E731
    s = _session(tmp_path, [("upload", []), ("prep", insert(2)), ("import", []), ("prep", insert(1))])
    v = {x["id"]: s.paths[x["id"]] for x in s.versions}
    # B: the numbers check and the fixer compare with the import version, through the moves made since
    original, moves = server._original_and_moves(s)
    assert original == v["ver-003"] and moves.forward("Data", 5, 1) == ("Data", 6, 1)
    # A: the recalculation errors are judged against it too
    assert server._original_path(s) == v["ver-003"]
    # C: the CELL / INFO values the user's Excel saved stay with the upload, moved by every insert
    upload, all_moves = server._upload_and_moves(s)
    assert upload == v["ver-001"] and all_moves.forward("Data", 5, 1) == ("Data", 7, 1)
    s.original_environment = {("Data", "A5"): "C:\\models\\tool.xlsm"}
    assert server._original_environment(s) == {"Data!A7": "C:\\models\\tool.xlsm"}
    # on the import version itself there is nothing to compare, and the numbers check says why
    s.current_version_id = "ver-003"
    assert server._original_and_moves(s) is None and server._original_path(s) == v["ver-003"]
    server.SESSIONS[s.id] = s
    out = client.post(f"/api/sessions/{s.id}/numbers-check").json()
    assert out["ran"] is False and out["verdict"] == "This is the version the outside files were brought into: nothing to compare."
    # a line without the import (v2) is compared with the upload as before
    s.current_version_id = "ver-002"
    assert server._original_path(s) == v["ver-001"] and server._original_and_moves(s)[0] == v["ver-001"]
    # each baseline file is recalculated once, the upload's and the import's apart
    runs = []

    def recalculate(path, save=True):
        runs.append(Path(path).parent.name)
        return {"ran": True, "status": "ERROR", "formula_errors": [{"sheet": "Data", "cell": "B2", "error": "#N/A"}], "addin_gap_errors": [], "message": "ok"}

    monkeypatch.setattr(server, "recalculate", recalculate)
    for key in v.values():
        s.ref_formulas[str(key)] = set()
    s.current_version_id = "ver-004"
    first = server._original_baseline(s, server._original_path(s))
    assert first["path"] == str(v["ver-003"]) and server._original_baseline(s, server._original_path(s)) is first
    s.current_version_id = "ver-002"
    assert server._original_baseline(s, server._original_path(s))["path"] == str(v["ver-001"])
    s.current_version_id = "ver-004"
    server._original_baseline(s, server._original_path(s))
    assert len(runs) == 2 and set(s.baselines) == {str(v["ver-001"]), str(v["ver-003"])}


def test_the_macro_runs_where_the_layout_is_still_the_uploads(tmp_path):
    formula_fix = [{"op": "set_formula", "sheet": "Data", "cell": "B2", "after": "=1"}]
    s = _session(tmp_path, [("upload", []), ("prep", formula_fix)])
    assert server._bring_in_base(s) == (s.versions[1], None)  # nothing moved: the current version
    s = _session(tmp_path / "moved", [("upload", []), ("prep", formula_fix), ("grid_namer", [{"op": "insert_row", "sheet": "Data", "row": 1}])])
    base, note = server._bring_in_base(s)
    assert base["id"] == "ver-002" and note.startswith("Brought in on v2: rows, columns or sheet names changed after it")
    assert note.endswith("run Do everything again.")
    s = _session(tmp_path / "again", [("upload", []), ("import", []), ("mind_inputs", [])])
    base, note = server._bring_in_base(s)
    # only the earlier bring-in's own versions follow v1: the new line leaves nothing else out
    assert base["id"] == "ver-001" and note == "Brought in on v1: the outside files were already brought in after it."
    s = _session(tmp_path / "half", [("upload", []), ("import", [])])  # one that did not finish: tried again from v1
    assert server._bring_in_base(s) == (s.versions[0], "Brought in on v1: the outside files were already brought in after it.")
    s = _session(tmp_path / "later", [("upload", []), ("import", []), ("mind_inputs", []), ("prep", formula_fix)])
    base, note = server._bring_in_base(s)
    assert base["id"] == "ver-001" and "already brought in" in note and note.endswith("The changes made after v1 are not in this new line: run Do everything again.")


# --- 5. INP-010 and the next step -------------------------------------------------------------------------
def test_inp_010_says_what_is_left_to_do():
    rows = [{"cell": f"Ctrl_01!H{r}", "expected": None, "keys": [k], "file": f"{k}.xlsm" if k != "VFA_01" else None} for r, k in zip(range(35, 41), KEYS)]
    oi = {"imports": True, "macro": "calcCF", "calls": [], "rows": rows, "matched": 5, "missing": ["Ctrl_01!H40"], "extra": [], "brought_in": False}
    assert inp_010(None, {}, {})["status"] == "PASS"
    f = inp_010(None, {}, {"outside_inputs": oi})
    assert f["status"] == "ERROR" and f["message"] == "The macro calcCF fills the workbook from 6 outside files, and Mind runs no macro: 5 of 6 picked, missing: VFA_01. Use Bring in the outside files."
    assert f["location"] == {"sheet": "Ctrl_01", "cell": "H35"} and len(f["observed"]["sites"]) == 6
    assert inp_010(None, {}, {"outside_inputs": {**oi, "brought_in": True}})["status"] == "PASS"
    assert inp_010(None, {}, {"outside_inputs": {**oi, "rows": []}})["status"] == "WARNING"  # no cell names the files
    assert inp_010(None, {}, {"outside_inputs": {**oi, "macro": None}})["status"] == "WARNING"
    # brought in before, and that no longer counts: said, and asked again
    stale = inp_010(None, {}, {"outside_inputs": {**oi, "brought_in_reason": "the input files changed since they were brought in"}})
    assert stale["status"] == "ERROR" and stale["message"] == (
        "The macro calcCF fills the workbook from 6 outside files, and Mind runs no macro: the input files changed since they were brought in "
        "(5 of 6 picked, missing: VFA_01). Use Bring in the outside files to bring them in again.")
    assert stale["observed"]["brought_in_reason"] == "the input files changed since they were brought in" and f["observed"]["brought_in_reason"] is None
    unfinished = inp_010(None, {}, {"outside_inputs": {**oi, "brought_in_reason": "the last bring-in did not finish"}})
    assert unfinished["status"] == "ERROR" and ": the last bring-in did not finish (5 of 6 picked" in unfinished["message"]
    assert inp_010(None, {}, {"outside_inputs": {**oi, "brought_in": True, "brought_in_reason": "ignored"}})["status"] == "PASS"


def test_the_next_step_brings_the_outside_files_in_before_prep():
    def f(rid, status, sites=None):
        return {"rule_id": rid, "status": status, "priority": "REQUIRED", "message": rid, "observed": {"sites": sites} if sites else None}

    plan = [{"id": "create_grid_titles", "rule_ids": ["STR-004"], "count": 3, "level": "blocking", "skipped": []}]
    report = {"findings": [f("STR-004", "ERROR"), f("INP-010", "ERROR", [1, 2, 3, 4, 5, 6]), f("READY-001", "NOT_SUPPORTED")]}
    r = compute_readiness(report, plan, None, "v")
    assert r["next_step"] == {"screen": "outside-inputs", "label": "Bring in the outside files", "detail": "6 files the workbook's macro opens",
                              "rule_id": "INP-010", "rule_name": "macros bring in data from other files"}
    assert {b["rule_id"]: b["fix"] for b in r["blocking"]} == {"STR-004": "prep", "INP-010": "outside_inputs"}


# --- 6. Run in Mind uploads the Mind inputs of the line ------------------------------------------------------
@dataclass
class _LoopConfig:
    source: Path
    work_dir: Path
    max_iterations: int = 4
    use_assistant: bool = True
    run_model: bool = True
    check_numbers: bool = True
    enable: list = field(default_factory=list)
    disable: list = field(default_factory=list)
    delete_projects: bool = True
    skip_mind: bool = False
    inputs_dir: Path | None = None


@dataclass
class _OldLoopConfig:
    source: Path
    work_dir: Path
    max_iterations: int = 4
    use_assistant: bool = True
    run_model: bool = True
    check_numbers: bool = True
    enable: list = field(default_factory=list)
    disable: list = field(default_factory=list)
    delete_projects: bool = True
    skip_mind: bool = False


def test_run_in_mind_uploads_the_mind_inputs_of_the_current_line(client, macros, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    csv_dir = tmp_path / "csv"
    csv_dir.mkdir()
    (csv_dir / "IN_BBA_01.csv").write_text("Amount\n1\n", encoding="utf-8")
    s.mind_inputs_dirs[s.current_version_id] = str(csv_dir)
    seen = []

    def run_loop(cfg, progress):
        seen.append(cfg)
        return {"verdict": "converged", "iterations": []}

    monkeypatch.setattr(server, "run_loop", run_loop)
    monkeypatch.setattr(server, "LoopConfig", _OldLoopConfig)  # a loop that cannot upload them: refused, nothing left running
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={"skipMind": True, "actions": [], "useAssistant": False, "checkNumbers": False})
    assert res.status_code == 409 and "Mind inputs" in res.json()["detail"] and sid not in server.MIND_LOOPS
    monkeypatch.setattr(server, "LoopConfig", _LoopConfig)
    options = client.get(f"/api/sessions/{sid}/mind-loop/options").json()
    assert options["inputs"] == ["IN_BBA_01.csv"] and "inputs" in [x["id"] for x in options["estimate"]["stages"]]
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={"skipMind": True, "actions": [], "useAssistant": False, "checkNumbers": False})
    assert res.status_code == 200, res.text
    for _ in range(200):
        if server.MIND_LOOPS[sid]["state"] != "running":
            break
        time.sleep(0.05)
    assert server.MIND_LOOPS[sid]["state"] == "done" and server.MIND_LOOPS[sid]["inputs"] == ["IN_BBA_01.csv"]
    assert seen[0].inputs_dir == csv_dir
    # another line (a version made from the upload without the bring-in) has no Mind inputs
    s.mind_inputs_dirs.clear()
    assert client.get(f"/api/sessions/{sid}/mind-loop/options").json()["inputs"] == []


def test_the_mind_run_estimate_lists_the_input_upload_only_with_inputs():
    plain = [x["id"] for x in estimate_mind_run(13.0, 60.0, 0, 200)["stages"]]
    assert "inputs" not in plain
    ids = [x["id"] for x in estimate_mind_run(13.0, 60.0, 0, 200, inputs=6)["stages"]]
    assert ids.index("template") < ids.index("inputs") < ids.index("run")


def test_rzs_006_leaves_the_mind_inputs_made_from_outside_files_alone(tmp_path):
    """A reader over part of an app-made Mind input (`IN_X!$C4:$E4`, ending on the grid's last
    column) is flagged by RZS-006 for an ordinary /Input grid; the grids the bring-in made keep the
    import area's size (the model's macro refuses bigger files, the app writes same-size CSVs)."""
    import openpyxl

    from app.inventory import build_analysis
    from app.validators.loop import rzs_006

    wb = openpyxl.Workbook()
    calc = wb.active
    calc.title = "Calc"
    calc["A1"] = "#Totals"
    calc["A2"] = "Item"
    calc["B2"] = "Value"
    calc["A3"] = "Sum"
    calc["B3"] = "=SUM(IN_X!$C4:$E4)"
    inp = wb.create_sheet("IN_X")
    inp["A1"] = "#IN_X /Input"
    for c, h in enumerate(["a", "b", "c", "d", "e"], start=1):
        inp.cell(2, c, h)
        for r in range(3, 6):
            inp.cell(r, c, r * 10 + c)
    path = tmp_path / "m.xlsx"
    wb.save(path)
    analysis = build_analysis(path, tmp_path / "w", "t")
    assert rzs_006(None, analysis, {})["status"] == "ERROR"
    res = rzs_006(None, analysis, {"outside_inputs": {"imports": True, "mind_inputs": ["IN_X"]}})
    assert res["status"] == "PASS", res["message"]
    assert "keep the import area's size" in res["message"]


# --- 7. a bring-in counts once it got to the end, with the files picked now ---------------------------------------
def _clean_compare(original, prepared, moves, progress=None):
    return {"ran": True, "clean": True, "verdict": "The preparation changed no computed value (1 formula cells compared with the original).",
            "counts": {"compared": 1, "same": 1, "good_to_error": 0, "good_to_other": 0, "missing_good": 0}, "regressions": []}


def test_mind_inputs_that_would_change_a_value_are_not_kept_and_the_step_runs_again(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    seen = []

    def changed(original, prepared, moves, progress=None):
        seen.append((Path(original), Path(prepared)))
        return {"ran": True, "clean": False, "verdict": "The preparation changed 3 computed value(s): 1 became errors, 2 became other values (9 formula cells compared with the original).",
                "counts": {"compared": 9, "same": 6, "good_to_error": 1, "good_to_other": 2, "missing_good": 0},
                "samples": {"good_to_error": [], "good_to_other": [{"original": "Calc!A3", "now": "Calc!A3", "was": "120", "is": "0"}], "missing": []},
                "regressions": [("Calc", 3, 1, 120)]}

    monkeypatch.setattr(server, "compare_workbooks", changed)
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "error" and st["result"] is None, st
    assert st["error"] == ("Making the Mind inputs would change 3 computed values; the outside data stays inside the model (v2); nothing else was changed. "
                           "For example Calc!A3 would show 0 instead of 120.")
    # compared before anything was kept: only the import version, current; the Mind inputs never a version
    (made_file,) = [p for _, p in seen]
    assert seen[0][0] == s.paths["ver-002"] and made_file.parent.name == "mind_inputs" and made_file not in s.paths.values()
    assert [v["id"] for v in s.versions] == ["ver-001", "ver-002"] and s.current_version_id == "ver-002" and st["versions"] == ["ver-002"]
    # their CSVs are never registered: no download, nothing for Run in Mind
    assert s.mind_inputs_dirs == {} and s.mind_inputs_dir is None and server._mind_inputs_dir_for(s) is None
    assert "Mind inputs.zip" not in s.files and "IN_BBA_01.csv" not in s.files
    assert client.get(f"/api/sessions/{sid}/files/Mind inputs.zip").status_code == 404
    assert client.get(f"/api/sessions/{sid}/mind-loop/options").json()["inputs"] == []
    # the line is not brought in: INP-010 and the verdict say so, and Do everything runs the step again
    oi = st["outsideInputs"]
    assert oi["brought_in"] is False and oi["brought_in_reason"] == "the last bring-in did not finish" and oi["mind_inputs"] == []
    assert s.brought_in_files == {} and not s.apply_lock.locked()
    f = _finding(s.result["validation_report"])
    assert f["status"] == "ERROR" and "Mind runs no macro: the last bring-in did not finish (6 of 6 picked)" in f["message"]
    page = client.get(f"/api/sessions/{sid}").json()
    assert page["outsideInputs"]["brought_in"] is False and page["readiness"]["next_step"]["label"] == "Bring in the outside files"
    # Try again: a new line from the version before (nothing else is left out of it); kept once no value changes
    monkeypatch.setattr(server, "compare_workbooks", _clean_compare)
    again = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert again.status_code == 202 and again.json()["base_version_id"] == "ver-001"
    assert again.json()["note"] == "Brought in on v1: the outside files were already brought in after it."
    st2 = _wait(client, sid)
    assert st2["state"] == "done" and st2["error"] is None and st2["result"]["versions"] == ["ver-003", "ver-004"]
    assert st2["outsideInputs"]["brought_in"] is True and st2["outsideInputs"]["brought_in_reason"] is None
    assert server._mind_inputs_dir_for(s) and _finding(s.result["validation_report"])["status"] == "PASS"


def test_a_comparison_that_could_not_run_keeps_the_mind_inputs_and_says_so(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    not_run = {"ran": False, "clean": None, "verdict": "Excel is not available here: the numbers could not be compared.", "counts": {}, "regressions": []}
    monkeypatch.setattr(server, "compare_workbooks", lambda *a, **k: not_run)
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "done", st
    r = st["result"]
    assert r["versions"] == ["ver-002", "ver-003"] and s.current_version_id == "ver-003"
    # a field the front end tests: Do everything stops there, the user goes on on purpose
    assert r["numbers"]["ran"] is False and r["numbers"]["clean"] is None and r["numbers"]["version_id"] == "ver-003"
    assert r["summary"] == ("6 outside files brought in; 1 Mind input made (IN_BBA_01); "
                            "the computed values were NOT checked (Excel is not available here: the numbers could not be compared).")
    assert r["download"]["zip"] == "Mind inputs.zip" and server._mind_inputs_dir_for(s)
    assert st["outsideInputs"]["brought_in"] is True and _finding(s.result["validation_report"])["status"] == "PASS"


def test_a_comparison_that_crashed_keeps_nothing_but_the_import(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]

    def crash(*a, **k):
        raise RuntimeError("Excel went away")

    monkeypatch.setattr(server, "compare_workbooks", crash)
    monkeypatch.setattr(server, "_report_problem", lambda *a, **k: None)  # no incident file from a test
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "error" and st["error"] == "v2 holds the outside data, but the rest did not finish: RuntimeError: Excel went away"
    assert [v["id"] for v in s.versions] == ["ver-001", "ver-002"] and s.current_version_id == "ver-002"
    assert server._mind_inputs_dir_for(s) is None and "Mind inputs.zip" not in s.files
    assert st["outsideInputs"]["brought_in"] is False and st["outsideInputs"]["brought_in_reason"] == "the last bring-in did not finish"


def test_a_bring_in_that_stopped_half_way_does_not_count(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    real_scan = server._run_scan

    def scan(sess, path=None):
        if path is not None and Path(path).parent.name == "mind_inputs":  # the scan of the Mind inputs version is refused
            raise server.HTTPException(status_code=409, detail="a scan is already running for this session")
        return real_scan(sess, path)

    monkeypatch.setattr(server, "_run_scan", scan)
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "error" and st["error"] == "v3 holds the outside data, but the rest did not finish: a scan is already running for this session"
    assert st["versions"] == ["ver-002", "ver-003"] and s.brought_in_files == {}
    assert st["outsideInputs"]["brought_in"] is False and st["outsideInputs"]["brought_in_reason"] == "the last bring-in did not finish"
    f = _finding(s.result["validation_report"])  # patched once the job stopped (its last scan ran while it was going)
    assert f["status"] == "ERROR" and "the last bring-in did not finish" in f["message"]
    # Do everything runs the step again, from the version before it
    monkeypatch.setattr(server, "_run_scan", real_scan)
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert res.status_code == 202 and res.json()["base_version_id"] == "ver-001"
    st = _wait(client, sid)
    assert st["state"] == "done" and st["outsideInputs"]["brought_in"] is True and st["result"]["versions"] == ["ver-004", "ver-005"]


def test_input_files_changed_since_the_bring_in_are_brought_in_again(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS[:5])["sessionId"]  # VFA_01 missing
    s = server.SESSIONS[sid]
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    assert _wait(client, sid)["state"] == "done"
    assert client.get(f"/api/sessions/{sid}/outside-inputs").json()["outsideInputs"]["brought_in"] is True
    # a file no row wants, or the same file again with the same bytes: still brought in
    out = client.post(f"/api/sessions/{sid}/inputs", files=[_part("notes.csv", b"a,b")]).json()
    assert out["outsideInputs"]["brought_in"] is True and out["outsideInputs"]["brought_in_reason"] is None
    out = client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[0], b"an input file")]).json()
    assert out["outsideInputs"]["brought_in"] is True
    # a corrected file under the same name: the line no longer counts as brought in
    out = client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[0], b"BBA_01, corrected")]).json()
    oi = out["outsideInputs"]
    assert oi["brought_in"] is False and oi["brought_in_reason"] == "the input files changed since they were brought in"
    f = _finding(s.result["validation_report"])
    assert f["status"] == "ERROR" and f["message"] == (
        "The macro calcCF fills the workbook from 6 outside files, and Mind runs no macro: the input files changed since they were brought in "
        "(5 of 6 picked, missing: VFA_01). Use Bring in the outside files to bring them in again.")
    assert out["readiness"]["next_step"]["label"] == "Bring in the outside files"
    # the page reopened says the same: the card and Do everything follow it
    page = client.get(f"/api/sessions/{sid}").json()
    assert page["outsideInputs"]["brought_in"] is False and page["outsideInputs"]["brought_in_reason"] == oi["brought_in_reason"]
    assert _finding(page["report"])["status"] == "ERROR"
    # Run in Mind refuses: it would upload the version and the Mind inputs made from the old files
    started = []
    monkeypatch.setattr(server, "LoopConfig", _LoopConfig)
    monkeypatch.setattr(server, "run_loop", lambda cfg, progress: started.append(cfg) or {"verdict": "converged", "iterations": []})
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={"skipMind": True, "actions": [], "useAssistant": False, "checkNumbers": False})
    assert res.status_code == 409 and res.json()["detail"].startswith("The input files changed since they were brought in: Run in Mind would use the data of the old files.")
    assert sid not in server.MIND_LOOPS and started == []
    # the missing file too, then Bring in again: a new line that reads the files picked now
    client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[5], b"the VFA file")])
    assert client.get(f"/api/sessions/{sid}/outside-inputs").json()["outsideInputs"]["brought_in"] is False
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert res.status_code == 202 and res.json()["base_version_id"] == "ver-001"
    st = _wait(client, sid)
    assert st["state"] == "done" and st["outsideInputs"]["brought_in"] is True and st["outsideInputs"]["brought_in_reason"] is None
    call = excel_steps["import"][-1]
    assert call["files"]["Ctrl_01!H35"].read_bytes() == b"BBA_01, corrected" and call["files"]["Ctrl_01!H40"].read_bytes() == b"the VFA file"
    assert _finding(s.result["validation_report"])["status"] == "PASS"
    assert client.post(f"/api/sessions/{sid}/mind-loop", json={"skipMind": True, "actions": [], "useAssistant": False, "checkNumbers": False}).status_code == 200
    for _ in range(200):
        if server.MIND_LOOPS[sid]["state"] != "running":
            break
        time.sleep(0.05)
    assert server.MIND_LOOPS[sid]["state"] == "done" and len(started) == 1


def test_the_fingerprint_follows_the_content_of_the_matched_files(tmp_path):
    s = server.Session(id=f"t{time.time_ns()}", work_dir=tmp_path, mode="plan")
    rows = [{"sheet": "Ctrl_01", "cell": "H35", "file": "a.xlsm"}, {"sheet": "Ctrl_01", "cell": "H36", "file": None}]
    s.inputs = [{"name": "a.xlsm", "path": "x", "bytes": 1, "sha256": "1" * 64}]
    first = server._inputs_fingerprint(s, rows)
    assert server._inputs_fingerprint(s, list(reversed(rows))) == first  # the order of the rows does not matter
    s.inputs = [{"name": "a.xlsm", "path": "y", "bytes": 1, "sha256": "2" * 64}]  # same name, other content
    assert server._inputs_fingerprint(s, rows) != first
    s.inputs = [{"name": "a.xlsm", "path": "x", "bytes": 1, "sha256": "1" * 64}]
    assert server._inputs_fingerprint(s, rows) == first
    assert server._inputs_fingerprint(s, [rows[0], {**rows[1], "file": "a.xlsm"}]) != first  # a row matched since


# --- 7. no input files: stand-ins made with the assistant ------------------------------------------------
def _fake_stand_ins(calls: list, ok: bool = True, left_out: tuple = ()):
    """make_stand_ins replaced: one file per row asked for, named as the model expects (or nothing)."""
    def make(model, rows, need, macro, modules, work_dir, ask, real_files=None, messages=None, progress=None, **_k):
        calls.append({"need": list(need), "real": sorted(real_files or {}), "macro": macro, "ask": ask("system", "user")})
        if progress:
            progress("Asking the assistant how the input files look (attempt 1 of 3)")
        if not ok:
            return {"ok": False, "files": {}, "left_out": [], "explain": None, "problem": "the assistant did not answer",
                    "attempts": [{"n": n, "problem": "the assistant did not answer", "details": []} for n in (1, 2, 3)]}
        work_dir = Path(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)
        by_key = {f"{r['sheet']}!{r['cell']}": r for r in rows}
        files = {}
        for k in need:
            if k in left_out:
                continue
            p = work_dir / f"2026Q1_2026_{by_key[k]['keys'][0]}.xlsm"
            p.write_bytes(b"a stand-in")
            files[k] = p
        return {"ok": True, "files": files, "left_out": list(left_out), "explain": "Items down, months across.", "problem": None,
                "numbers": {"ran": True, "clean": True}, "attempts": [{"n": 1, "problem": None, "details": []}]}

    return make


def _assistant_on(monkeypatch, calls: list, **kw) -> None:
    monkeypatch.setattr(server, "llm_available", lambda: True)
    monkeypatch.setattr(stand_in_inputs, "assistant_ask", lambda: (lambda system, user: "the assistant's answer"))
    monkeypatch.setattr(stand_in_inputs, "make_stand_ins", _fake_stand_ins(calls, **kw))


def test_with_only_the_workbook_the_assistant_makes_stand_ins_that_are_brought_in(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, [])["sessionId"]  # the workbook alone
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "llm_available", lambda: False)
    assert client.get(f"/api/sessions/{sid}").json()["outsideInputs"]["assistant"] is False
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"standIns": True})
    assert res.status_code == 409 and "The assistant is not set up" in res.json()["detail"]
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 422  # nothing picked, nothing to make
    calls: list = []
    _assistant_on(monkeypatch, calls)
    assert client.get(f"/api/sessions/{sid}").json()["outsideInputs"]["assistant"] is True
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"standIns": True})
    assert res.status_code == 202, res.text
    st = _wait(client, sid)
    assert st["state"] == "done", st
    (call,) = calls
    assert call["need"] == [f"Ctrl_01!H{r}" for r in range(35, 41)] and call["real"] == [] and call["macro"] == "calcCF"
    assert call["ask"] == "the assistant's answer"  # the app's assistant is what the job asks
    oi = st["outsideInputs"]
    assert oi["matched"] == 6 and oi["brought_in"] is True and sorted(oi["stand_ins"]) == sorted(INPUTS) and all(r["stand_in"] for r in oi["rows"])
    assert all(i["stand_in"] for i in st["inputs"])
    assert len(excel_steps["import"][0]["files"]) == 6  # the macro ran on the stand-ins
    # and their rows of the import table keep the model's own status (a stand-in is not a real import)
    assert excel_steps["import"][0]["restore_rows"] == [f"Ctrl_01!H{r}" for r in range(35, 41)]
    result = st["result"]
    assert result["summary"].endswith("best results come with the actual input files (add them, then bring them in again).")
    assert result["standIns"]["ok"] is True and result["standIns"]["explain"] == "Items down, months across."
    f = _finding(s.result["validation_report"])
    assert f["status"] == "WARNING" and "6 of 6 are stand-ins the assistant made from the numbers the model already holds" in f["message"]
    assert "Best results come with the actual input files" in f["message"] and f["observed"]["stand_ins"] == oi["stand_ins"]
    assert all(b["rule_id"] != "INP-010" for b in s.readiness()["blocking"])
    # the actual file arrives: it replaces its stand-in, and the line has to be brought in again
    out = client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[0], b"the real BBA_01")]).json()
    assert INPUTS[0] not in out["outsideInputs"]["stand_ins"] and len(out["outsideInputs"]["stand_ins"]) == 5
    assert out["outsideInputs"]["brought_in"] is False and out["outsideInputs"]["brought_in_reason"] == "the input files changed since they were brought in"
    assert _finding(s.result["validation_report"])["status"] == "ERROR"


def test_stand_ins_only_for_the_missing_files_a_failure_changes_nothing_and_empty_areas_are_said(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS[:4])["sessionId"]
    s = server.SESSIONS[sid]
    calls: list = []
    _assistant_on(monkeypatch, calls, ok=False)
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"standIns": True}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "error" and st["error"] == (
        "The assistant could not make input files the workbook's macro accepts (3 attempts): the assistant did not answer. "
        "Nothing was changed: add the actual input files.")
    assert calls[0]["need"] == ["Ctrl_01!H39", "Ctrl_01!H40"] and len(calls[0]["real"]) == 4
    assert [v["id"] for v in s.versions] == ["ver-001"] and not any(i.get("stand_in") for i in s.inputs) and excel_steps["import"] == []
    # the model holds nothing for VFA_01: no stand-in for it, its block stays; BBA_05's stand-in comes in with the 4 real files
    _assistant_on(monkeypatch, calls, left_out=("Ctrl_01!H40",))
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"standIns": True}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "done", st
    assert st["outsideInputs"]["stand_ins"] == [INPUTS[4]] and st["outsideInputs"]["missing"] == ["Ctrl_01!H40"]
    assert excel_steps["import"][-1]["restore_rows"] == ["Ctrl_01!H39"]  # the real files' rows show their import
    notes = " ".join(st["result"]["notes"])
    assert "No stand-in for VFA_01: the model holds no data for it and zeros would change a result, so its block stays as the model holds it." in notes
    # standIns with every file there: nothing to make, the plain bring-in
    calls.clear()
    client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[5], b"the real VFA_01")])
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"standIns": True}).status_code == 202
    assert _wait(client, sid)["state"] == "done" and calls == []


# --- 8. Do everything stops and asks for the input files (1.14.0) -------------------------------------
# Do everything (the browser's chain) starts the bring-in with waitForInputs: with files missing the job
# waits instead of making stand-ins quietly; the files are picked meanwhile (POST .../inputs), then
# .../continue goes on from the import itself -- or makes the stand-ins on purpose; .../stop ends it.
def _run_waiting(client, sid: str, **payload) -> dict:
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"waitForInputs": True, **payload})
    assert res.status_code == 202, res.text
    return res.json()


def test_do_everything_waits_for_the_missing_files_then_goes_on_from_the_import_only(client, macros, excel_steps, monkeypatch, tmp_path):
    from app import session_store

    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS[:4])["sessionId"]  # BBA_05 and VFA_01 missing
    s = server.SESSIONS[sid]
    stand_in_calls: list = []
    _assistant_on(monkeypatch, stand_in_calls)  # an AI key is there, and still nothing is made by itself
    body = _run_waiting(client, sid)
    assert body["state"] == "waiting_for_inputs" and body["running"] is False
    assert body["message"] == "Waiting for 2 input files: BBA_05, VFA_01"
    assert [(n["name"], n["present"], n["fills"]) for n in body["needed"]] == [
        ("BBA_01", True, "BBA_01"), ("BBA_02", True, None), ("BBA_03", True, None), ("BBA_04", True, None), ("BBA_05", False, None), ("VFA_01", False, None)]
    assert body["needed"][0]["file"] == INPUTS[0] and body["needed"][0]["cell"] == "Ctrl_01!H35"
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["waiting"] is True and st["running"] is False and st["state"] == "waiting_for_inputs" and st["stage"] == "Waiting for the input files"
    assert st["stages"] == ["waiting"] and st["needed"] == body["needed"] and st["error"] is None and st["result"] is None
    # nothing ran, nothing changed, no stand-in asked for
    assert [v["id"] for v in s.versions] == ["ver-001"] and excel_steps["import"] == [] and stand_in_calls == [] and not s.apply_lock.locked()
    assert st["outsideInputs"]["brought_in"] is False and st["outsideInputs"]["missing"] == ["Ctrl_01!H39", "Ctrl_01!H40"]
    # the wait is saved with the session: reopened after a restart it is over, in plain words
    assert session_store.read_state(sid)["outside_run"]["state"] == "waiting_for_inputs"
    back, _ = session_store.load(sid, server.Session)
    assert back.outside_run["state"] == "stopped" and back.outside_run["error"] == session_store.CLOSED_WHILE_WAITING
    # Go on needs every file
    res = client.post(f"/api/sessions/{sid}/outside-inputs/continue", json={})
    assert res.status_code == 409 and res.json()["detail"] == "2 input files still missing: BBA_05, VFA_01. Pick them, or let the assistant make stand-ins."
    # the files are picked while it waits (the usual inputs route); the list follows
    out = client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[4], b"BBA_05"), _part(INPUTS[5], b"VFA_01")]).json()
    assert out["outsideInputs"]["missing"] == []
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["waiting"] is True and all(n["present"] for n in st["needed"]) and st["message"] == "Every input file is there: press Go on."
    # Go on: the job starts at the import -- the stages after the wait, and only those
    res = client.post(f"/api/sessions/{sid}/outside-inputs/continue", json={})
    assert res.status_code == 202, res.text
    assert res.json()["state"] == "running"
    st = _wait(client, sid)
    assert st["state"] == "done", st
    assert st["stages"] == ["start", "import", "scan", "mind_inputs", "numbers", "rescan", "done"]
    assert st["waiting"] is False and st["needed"] is None
    assert len(excel_steps["import"]) == 1 and len(excel_steps["import"][0]["files"]) == 6 and stand_in_calls == []
    assert st["outsideInputs"]["brought_in"] is True and st["outsideInputs"]["stand_ins"] == []
    assert [v["source"] for v in s.versions] == ["upload", "import", "mind_inputs"]
    assert "stand-in" not in st["result"]["summary"]
    # not waiting any more: Go on and Stop say so
    assert client.post(f"/api/sessions/{sid}/outside-inputs/continue", json={}).status_code == 409
    assert client.post(f"/api/sessions/{sid}/outside-inputs/stop").json() == {"running": False, "state": "done"}


def test_while_waiting_make_them_with_ai_takes_the_stand_in_path_with_its_warning(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, [])["sessionId"]  # the workbook alone
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "llm_available", lambda: False)
    body = _run_waiting(client, sid)
    assert body["state"] == "waiting_for_inputs" and body["message"] == "Waiting for 6 input files: BBA_01, BBA_02, BBA_03, BBA_04, BBA_05, VFA_01"
    # without an AI key the assistant cannot make them: the wait goes on
    res = client.post(f"/api/sessions/{sid}/outside-inputs/continue", json={"makeStandIns": True})
    assert res.status_code == 409 and "The assistant is not set up" in res.json()["detail"]
    assert client.get(f"/api/sessions/{sid}/outside-inputs").json()["waiting"] is True
    calls: list = []
    _assistant_on(monkeypatch, calls)
    assert client.post(f"/api/sessions/{sid}/outside-inputs/continue", json={"makeStandIns": True}).status_code == 202
    st = _wait(client, sid)
    assert st["state"] == "done", st
    assert st["stages"] == ["start", "stand_ins", "import", "scan", "mind_inputs", "numbers", "rescan", "done"]
    (call,) = calls
    assert call["need"] == [f"Ctrl_01!H{r}" for r in range(35, 41)] and call["real"] == []
    assert st["result"]["summary"].endswith("best results come with the actual input files (add them, then bring them in again).")
    assert st["result"]["standIns"]["ok"] is True and sorted(st["outsideInputs"]["stand_ins"]) == sorted(INPUTS)
    assert [v["source"] for v in s.versions] == ["upload", "import", "mind_inputs"]
    # with every file there, waitForInputs waits for nothing: the plain run starts at once
    sid2 = _upload_picked(client, _main_book(tmp_path / "two"), INPUTS)["sessionId"]
    body = _run_waiting(client, sid2)
    assert body["state"] == "running" and body["running"] is True
    st = _wait(client, sid2)
    assert st["state"] == "done" and st["stages"][:2] == ["start", "import"] and calls == [call]


def test_stop_while_waiting_ends_the_job_and_changes_nothing(client, macros, excel_steps, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS[:5])["sessionId"]
    s = server.SESSIONS[sid]
    assert client.post(f"/api/sessions/{sid}/outside-inputs/stop").json() == {"running": False, "state": "idle"}  # nothing to stop
    _run_waiting(client, sid)
    res = client.post(f"/api/sessions/{sid}/outside-inputs/stop")
    assert res.status_code == 200 and res.json() == {"running": False, "state": "stopped", "error": server._WAIT_STOPPED}
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["state"] == "stopped" and st["running"] is False and st["waiting"] is False and st["needed"] is None
    assert st["error"] == "Stopped while waiting for the input files. Nothing was changed: pick the files, then press Do everything again."
    assert st["stages"] == ["waiting", "stopped"] and st["result"] is None
    assert [v["id"] for v in s.versions] == ["ver-001"] and excel_steps["import"] == [] and not s.apply_lock.locked()
    assert client.post(f"/api/sessions/{sid}/outside-inputs/continue", json={}).status_code == 409
    # the workbook is free: Do everything again waits again, the matched files alone still come in
    assert _run_waiting(client, sid)["state"] == "waiting_for_inputs"
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202  # the card's "Bring in the 5 matched files"
    st = _wait(client, sid)
    assert st["state"] == "done" and st["waiting"] is False and len(excel_steps["import"][0]["files"]) == 5
    # a bring-in that runs is not stopped half-way
    gate = threading.Event()
    monkeypatch.setattr(outside_inputs, "run_import", _fake_run_import(excel_steps["import"], gate=gate))
    client.post(f"/api/sessions/{sid}/inputs", files=[_part(INPUTS[5], b"VFA_01")])
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    try:
        res = client.post(f"/api/sessions/{sid}/outside-inputs/stop")
        assert res.status_code == 409 and "cannot be stopped half-way" in res.json()["detail"]
    finally:
        gate.set()
    assert _wait(client, sid)["state"] == "done"


# --- 6. next quarter's Mind input files (app/next_quarter_inputs.py) -------------------------------
def _wait_next(client, sid: str, timeout: float = 60) -> dict:
    deadline = time.time() + timeout
    while True:
        st = client.get(f"/api/sessions/{sid}/next-quarter-inputs").json()
        if not st["running"]:
            return st
        assert time.time() < deadline, st
        time.sleep(0.05)


@pytest.fixture
def next_quarter_export(monkeypatch):
    """export_grids replaced (no Excel): each grid's new block, under the grid's header."""
    calls: list = []

    def export_grids(workbook, grids, out_dir, progress=None):
        calls.append({"workbook": Path(workbook), "grids": [dict(g) for g in grids]})
        made = []
        for g in grids:
            info = outside_inputs.write_input_csv(Path(out_dir) / g["name"], g["header"], [[33], [44], [55]])
            made.append({"grid": g["name"], "path": info["path"], "rows": info["rows"], "cols": info["cols"], "separator": info["separator"]})
        return made

    from app import next_quarter_inputs

    monkeypatch.setattr(next_quarter_inputs, "export_grids", export_grids)
    return calls


def test_next_quarter_files_come_from_the_raw_files_through_the_model_s_own_macro(client, macros, excel_steps, next_quarter_export, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    # before the bring-in: no grid to fill, said plainly
    before = client.get(f"/api/sessions/{sid}/next-quarter-inputs").json()
    assert before["available"] is False and before["grids"] == [] and before["state"] == "idle"
    assert before["why_not"] == "Bring in the outside files first: that makes the Mind input grids these files fill."
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("2026Q2_2026_BBA_01.xlsm", b"raw")])
    assert res.status_code == 422 and res.json()["detail"] == before["why_not"]
    client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert _wait(client, sid)["state"] == "done"
    versions = [v["id"] for v in s.versions]
    st = client.get(f"/api/sessions/{sid}/next-quarter-inputs").json()
    assert st["available"] is True and st["why_not"] is None
    assert st["grids"] == [{"name": "IN_BBA_01", "sheet": "BBA_01", "ref": "B3:B5", "rows": 3, "cols": 1, "file": "IN_BBA_01.csv"}]
    # the new quarter's raw file: the model's own macro runs on the version the files were first brought into
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("2026Q2_2026_BBA_01.xlsm", b"next quarter's raw file")])
    assert res.status_code == 202, res.text
    assert res.json() == {"running": True, "state": "running", "raw": ["2026Q2_2026_BBA_01.xlsm"], "grid_files": [], "notes": [], "unused": []}
    done = _wait_next(client, sid)
    assert done["state"] == "done" and done["error"] is None, done
    nq_import = excel_steps["import"][-1]
    assert nq_import["workbook"] == s.paths["ver-001"] and nq_import["macro"] == "calcCF"
    assert {k: p.read_bytes() for k, p in nq_import["files"].items()} == {"Ctrl_01!H35": b"next quarter's raw file"}
    (exp,) = next_quarter_export
    assert exp["grids"][0]["name"] == "IN_BBA_01" and exp["grids"][0]["header"] == ["Amount"]  # the header Mind's grid has
    r = done["result"]
    assert r["files"] == [{"file": "IN_BBA_01.csv", "grid": "IN_BBA_01", "from": "2026Q2_2026_BBA_01.xlsm", "source": "macro", "rows": 3, "cols": 1,
                           "separator": "comma", "changed": 3}]
    assert r["summary"] == "1 Mind input file made (IN_BBA_01.csv): put it in Mind's Inputs window, then Run."
    assert r["zip"] == "Next quarter Mind inputs.zip" and r["not_in_zip"] == [] and r["import"]["files"][0]["opened"] is True
    assert "Raw input files cannot go into Mind directly." in r["readme"]
    # no new version, no change to the session's line or its Mind inputs
    assert [v["id"] for v in s.versions] == versions and s.current_version_id == versions[-1]
    assert sorted(p.name for p in Path(s.mind_inputs_dir).iterdir()) == ["IN_BBA_01.csv"]
    z = client.get(f"/api/sessions/{sid}/files/Next quarter Mind inputs.zip")
    assert z.status_code == 200
    zf = zipfile.ZipFile(io.BytesIO(z.content))
    assert zf.namelist() == ["README.md", "Mind inputs/IN_BBA_01.csv"]
    assert zf.read("Mind inputs/IN_BBA_01.csv") == b"Amount\n33\n44\n55\n"


def test_next_quarter_takes_grid_files_as_they_are_and_recognises_raw_ones(client, macros, excel_steps, next_quarter_export, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert _wait(client, sid)["state"] == "done"
    imports = len(excel_steps["import"])
    # a file already shaped like the grid (its header): into the zip as it is, no macro run
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("IN_BBA_01.csv", b"Amount\n31\n41\n51\n")])
    assert res.status_code == 202 and res.json()["grid_files"] == [{"file": "IN_BBA_01.csv", "grid": "IN_BBA_01"}] and res.json()["raw"] == []
    done = _wait_next(client, sid)
    assert done["state"] == "done" and len(excel_steps["import"]) == imports and next_quarter_export == []
    assert [(f["file"], f["source"], f["changed"]) for f in done["result"]["files"]] == [("IN_BBA_01.csv", "as_given", 3)]
    # the raw file renamed after the grid (what Mind refused): recognised, read with the macro, said so
    wide = ",".join(f"m{i}" for i in range(1340)).encode("utf-8")
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("IN_BBA_01.csv", wide + b"\n1,2\n")])
    assert res.status_code == 202, res.text
    assert res.json()["raw"] == ["IN_BBA_01.csv"] and res.json()["notes"] == [
        "IN_BBA_01.csv has 1,340 columns; the Mind input grid IN_BBA_01 has 1. It looks like a raw input file, so the app reads it with the model's import macro."]
    done = _wait_next(client, sid)
    assert done["state"] == "done" and len(excel_steps["import"]) == imports + 1
    assert done["result"]["files"][0]["source"] == "macro" and done["result"]["notes"][0].startswith("IN_BBA_01.csv has 1,340 columns")
    # nothing usable: refused at once, in plain words; other kinds of file refused as at the upload
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("notes.csv", b"a,b\n1,2\n")])
    assert res.status_code == 422 and res.json()["detail"] == (
        "None of the picked files fills a Mind input grid (notes.csv: no row of the model's import table names it, "
        "and it is not shaped like a Mind input grid).")
    assert client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("report.pdf", b"%PDF")]).status_code == 400
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("a.csv", b"Amount\n1\n"), _part("A.CSV", b"Amount\n2\n")])
    assert res.status_code == 422 and res.json()["detail"] == "Two files are named A.CSV: pick one."


def test_next_quarter_waits_for_a_running_job_and_says_a_macro_s_refusal(client, macros, excel_steps, next_quarter_export, monkeypatch, tmp_path):
    main = _main_book(tmp_path)
    sid = _upload_picked(client, main, INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert _wait(client, sid)["state"] == "done"
    s.next_quarter = {"state": "running", "started": time.time()}
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("2026Q2_2026_BBA_01.xlsm", b"raw")])
    assert res.status_code == 409 and res.json()["detail"] == "Wait a moment: next quarter's Mind input files are still being made."
    # and the other way round: while its import macro runs, no second bring-in, Run in Mind, fixer or recalculation
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={})
    assert res.status_code == 409 and res.json()["detail"] == "Wait a moment: next quarter's Mind input files are still being made."
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={})
    assert res.status_code == 409 and res.json()["detail"].startswith("Wait a moment: next quarter's Mind input files are still being made.")
    assert client.get(f"/api/sessions/{sid}/mind-loop/options").json()["busy"] == "next quarter's Mind input files are still being made"
    for route in ("recalculate", "auto-fix"):
        res = client.post(f"/api/sessions/{sid}/{route}", json={})
        assert res.status_code == 409 and res.json()["detail"] == "next quarter's Mind input files are being made: wait for it to finish", route
    s.next_quarter = {}
    s.outside_run["state"] = "running"
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("2026Q2_2026_BBA_01.xlsm", b"raw")])
    assert res.status_code == 409 and res.json()["detail"] == "Wait a moment: the outside files are still being brought in."
    s.outside_run["state"] = "done"
    monkeypatch.setattr(outside_inputs, "run_import", _fake_run_import([], stopped=True))
    assert client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("2026Q2_2026_BBA_01.xlsm", b"raw")]).status_code == 202
    done = _wait_next(client, sid)
    assert done["state"] == "error" and done["result"]["files"] == [] and done["result"]["zip"] is None
    assert done["error"] == ("The model's import macro did not take the new files: \"BBA_01のinput領域には93項目まで入力可能です\". "
                             "No file was made: check the input files, then try again.")
    monkeypatch.setattr(server, "com_available", lambda: False)
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[_part("2026Q2_2026_BBA_01.xlsm", b"raw")])
    assert res.status_code == 409 and res.json()["detail"].startswith("Reading raw input files needs Excel on this machine")
