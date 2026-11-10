"""Saved workbooks (app/session_store.py, 1.14.0): a session's files and state live in the data folder
and come back after a restart of the app -- every version, the check result, the plan, the recalculation,
the downloads -- from the main screen's "Recent workbooks" list; Remove deletes the folder; a job that
was running when the app stopped shows as stopped."""
from __future__ import annotations

import json
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app import session_store
from app.settings import SETTINGS
from app.web import server

CLOSED = session_store.CLOSED_WHILE_RUNNING


@pytest.fixture
def data_folder(tmp_path, monkeypatch):
    """The sessions go under tmp_path (Settings is frozen: the field is set on the one SETTINGS object)."""
    folder = tmp_path / "data"
    real = SETTINGS.data_folder
    object.__setattr__(SETTINGS, "data_folder", folder)
    try:
        yield folder
    finally:
        object.__setattr__(SETTINGS, "data_folder", real)


@pytest.fixture
def client(data_folder):
    server.SESSIONS.clear()
    server.AUTOFIX.clear()
    yield TestClient(server.app)
    server.SESSIONS.clear()
    server.AUTOFIX.clear()


def _upload(client, path: Path, defer: bool = False) -> dict:
    with path.open("rb") as f:
        res = client.post("/api/sessions", files={"file": (path.name, f, "application/octet-stream")}, data={"mode": "plan", **({"defer": "1"} if defer else {})})
    assert res.status_code == 200, res.text
    return res.json()


def _restart() -> None:
    """What a restart of the app leaves: nothing in memory, everything on disk."""
    server.SESSIONS.clear()
    server.AUTOFIX.clear()


def test_a_checked_workbook_comes_back_after_a_restart(client, data_folder, flagged_model_broken_xlsx):
    body = _upload(client, flagged_model_broken_xlsx)
    sid = body["sessionId"]
    s = server.SESSIONS[sid]
    # the session's folder is in the data folder, not in %TEMP%, with its state file next to the files
    assert s.work_dir == data_folder / "sessions" / sid and (s.work_dir / "state.json").is_file()
    assert "mind_ready_" not in str(s.work_dir)
    plan = {a["id"]: a for a in body["plan"]}
    applied = client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["flag_spelling"]["operations"], "reanalyze": True})
    assert applied.status_code == 200, applied.text
    assert applied.json()["version"]["id"] == "ver-002"
    # the list, before the restart: one workbook, its name, two versions, a check result
    listed = client.get("/api/saved-workbooks").json()["workbooks"]
    assert [w["id"] for w in listed] == [sid]
    row = listed[0]
    assert row["name"] == "flagged_model_broken.xlsx" and row["versions"] == 2 and row["has_result"] is True and row["pending"] is False
    assert row["bytes"] > 0 and row["uploaded_at"] and row["last_opened"] and row["current_version"] == "ver-002"
    before = client.get(f"/api/sessions/{sid}").json()
    download_name = before["version"]["file_name"]

    _restart()
    assert client.get(f"/api/sessions/{sid}").status_code == 404  # nothing reopens by itself
    assert [w["id"] for w in client.get("/api/saved-workbooks").json()["workbooks"]] == [sid]  # the list reads the disk
    opened = client.post(f"/api/saved-workbooks/{sid}/open")
    assert opened.status_code == 200, opened.text
    after = opened.json()
    # exactly where it was: the versions, the current one, the report, the plan, the verdict, the gauge
    assert after["sessionId"] == sid and after["versions"] == before["versions"] and after["version"] == before["version"]
    assert after["report"]["findings"] == before["report"]["findings"] and after["report"]["summary"] == before["report"]["summary"]
    assert [a["id"] for a in after["plan"]] == [a["id"] for a in before["plan"]]
    assert after["readiness"] == before["readiness"] and after["prep_progress"] == before["prep_progress"]
    assert after["delta"] == before["delta"] and after["pivot_mode"] == before["pivot_mode"] and after["size"] == before["size"]
    assert after["scan"]["state"] == "ready"
    # and the session answers again, with its downloads
    assert client.get(f"/api/sessions/{sid}").status_code == 200
    assert [v["id"] for v in client.get(f"/api/sessions/{sid}/versions").json()] == ["ver-001", "ver-002"]
    dl = client.get(f"/api/sessions/{sid}/files/{download_name}")
    assert dl.status_code == 200 and dl.content[:2] == b"PK"
    # the readiness route and a further apply work on the reopened session (its lineage and paths are back)
    assert client.get(f"/api/sessions/{sid}/readiness").status_code == 200
    again = client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["loop_name_case"]["operations"], "reanalyze": False})
    assert again.status_code == 200 and again.json()["version"]["id"] == "ver-003"
    # opening it again moved it to the top of the list (last opened)
    assert client.get("/api/saved-workbooks").json()["workbooks"][0]["id"] == sid


def test_remove_deletes_the_folder_and_the_session(client, data_folder, plain_grid_xlsx):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    folder = data_folder / "sessions" / sid
    assert folder.is_dir()
    res = client.delete(f"/api/saved-workbooks/{sid}")
    assert res.status_code == 200 and res.json() == {"removed": True, "id": sid}
    assert not folder.exists()
    assert client.get("/api/saved-workbooks").json()["workbooks"] == []
    assert client.get(f"/api/sessions/{sid}").status_code == 404
    assert client.post(f"/api/saved-workbooks/{sid}/open").status_code == 404
    assert client.delete(f"/api/saved-workbooks/{sid}").status_code == 404
    # an id that is not one is refused the same way, never read as a path
    assert client.post("/api/saved-workbooks/..x/open").status_code == 404
    assert client.delete("/api/saved-workbooks/not%20an%20id").status_code == 404
    assert session_store.valid_id("../x") is False and session_store.valid_id("abc_123-x") is True


def test_remove_waits_for_a_running_step(client, plain_grid_xlsx):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    server.SESSIONS[sid].outside_run = {"state": "running", "stage": "import", "started": time.time()}
    res = client.delete(f"/api/saved-workbooks/{sid}")
    assert res.status_code == 409 and "still being brought in" in res.json()["detail"]
    assert (server.SESSIONS[sid].work_dir / "state.json").is_file()


def test_a_session_saved_while_a_job_ran_shows_it_stopped(client, plain_grid_xlsx):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    # the app is closed with three things running: the fixer, a bring-in and a scan
    server.AUTOFIX[sid] = {"state": "running", "stage": "fix", "title": "Fixing", "message": "pass 2", "started": time.time(), "updated": time.time(), "finished": None, "error": None, "stop": False, "version_id": "ver-001"}
    s.outside_run = {"state": "running", "stage": "import", "title": "Running the import macro", "message": "file 2 of 6", "started": time.time()}
    s.scan = {**s.scan, "state": "running", "stage": "rules"}
    server._save(s)
    _restart()
    opened = client.post(f"/api/saved-workbooks/{sid}/open").json()
    assert opened["scan"]["state"] == "error" and opened["scan"]["error"] == CLOSED
    outside = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert outside["running"] is False and outside["state"] == "error" and outside["error"] == CLOSED
    fixer = client.get(f"/api/sessions/{sid}/auto-fix").json()
    assert fixer["state"] == "error" and fixer["error"] == CLOSED and "stop" not in fixer
    # the stopped states are what the file holds now too (a second restart says the same)
    _restart()
    assert client.post(f"/api/saved-workbooks/{sid}/open").json()["scan"]["error"] == CLOSED


def test_a_workbook_uploaded_but_not_checked_reopens_as_pending(client, flagged_model_broken_xlsx):
    body = _upload(client, flagged_model_broken_xlsx, defer=True)
    sid = body["sessionId"]
    assert body["pending"] is True
    row = client.get("/api/saved-workbooks").json()["workbooks"][0]
    assert row["has_result"] is False and row["pending"] is True and row["versions"] == 1
    _restart()
    opened = client.post(f"/api/saved-workbooks/{sid}/open").json()
    assert opened["pending"] is True and opened["sessionId"] == sid and opened["size"]["sheets"] and opened["version"]["id"] == "ver-001"
    # the scan can then run as after the upload
    scanned = client.post(f"/api/sessions/{sid}/analyze", json={"ignore_sheets": []})
    assert scanned.status_code == 200 and scanned.json()["report"]["findings"]
    assert client.get("/api/saved-workbooks").json()["workbooks"][0]["has_result"] is True


def test_the_list_is_last_opened_first_and_skips_a_folder_without_a_state_file(client, data_folder, plain_grid_xlsx):
    first = _upload(client, plain_grid_xlsx)["sessionId"]
    second = _upload(client, plain_grid_xlsx)["sessionId"]
    (data_folder / "sessions" / "stray_folder").mkdir()
    (data_folder / "sessions" / "broken").mkdir()
    (data_folder / "sessions" / "broken" / "state.json").write_text("{not json", encoding="utf-8")
    assert [w["id"] for w in client.get("/api/saved-workbooks").json()["workbooks"]] == [second, first]
    server.SESSIONS[first].last_opened = time.time() + 10
    server._save(server.SESSIONS[first])
    assert [w["id"] for w in client.get("/api/saved-workbooks").json()["workbooks"]] == [first, second]


def test_a_save_that_fails_never_breaks_the_step(client, plain_grid_xlsx, monkeypatch, capsys):
    def broken(s, autofix=None):
        raise RuntimeError("disk full")

    monkeypatch.setattr(session_store, "to_state", broken)
    body = _upload(client, plain_grid_xlsx)
    assert body["report"]["findings"]  # the upload and its scan answered as before
    assert "Could not save the session" in capsys.readouterr().out
    assert client.get("/api/saved-workbooks").json()["workbooks"] == []  # nothing half-written


def test_the_state_file_keeps_every_odd_shape(data_folder):
    """Sets of (sheet, cell), tuple keys, Paths inside and outside the folder, the baselines' sets:
    out through the file and back as they were."""
    folder = data_folder / "sessions" / "abc123"
    folder.mkdir(parents=True)
    s = server.Session(id="abc123", work_dir=folder, mode="plan")
    s.versions = [{"id": "ver-001", "file_name": "m.xlsx"}, {"id": "ver-002", "file_name": "m_v2.xlsx"}]
    s.current_version_id = "ver-002"
    s.paths = {"ver-001": folder / "upload" / "m.xlsx", "ver-002": folder / "apply" / "m_v2.xlsx"}
    s.files = {"m.xlsx": folder / "upload" / "m.xlsx", "elsewhere.csv": Path("C:/elsewhere/elsewhere.csv")}
    s.recalc = {"status": "OK", "formula_errors": [], "ran": True}
    s.recalc_version_id = "ver-002"
    s.recalc_path = folder / "recalc" / "m_v2.xlsx"
    s.original_ref_formulas = {("Calc", "B2"), ("Calc", "A1")}
    s.ref_formulas = {str(folder / "upload" / "m.xlsx"): {("Calc", "B2")}}
    s.original_environment = {("Ctrl", "A1"): "C:\\x\\[m.xlsx]Ctrl", ("Ctrl", "A2"): 3}
    base = {"ran": True, "errors": [{"sheet": "Calc", "cell": "B2", "error": "#REF!"}], "ref_formulas": {("Calc", "B2")}, "message": None, "path": "p"}
    s.baselines = {"p": base}
    s.original_baseline = base
    s.instances_zip = folder / "instances" / "ver-002" / "m_instances.zip"
    s.mind_inputs_dir = str(folder / "outside_inputs" / "1" / "mind_inputs")
    s.mind_inputs_dirs = {"ver-002": str(folder / "outside_inputs" / "1" / "mind_inputs")}
    s.inputs = [{"name": "in.xlsm", "path": str(folder / "inputs" / "01" / "in.xlsm"), "bytes": 3, "sha256": "ab"}]
    s.grid_names = {"Calc!A1:B2": "Cashflows"}
    s.outside_run = {"state": "done", "result": {"summary": "ok"}}
    s.name = "m.xlsx"
    s.last_opened = 123.0
    autofix = {"state": "done", "stage": "done", "stop": False, "result": {"summary": "2 cells fixed"}}
    path = session_store.save(s, autofix)
    assert path == folder / "state.json"
    state = json.loads(path.read_text(encoding="utf-8"))
    assert state["paths"] == {"ver-001": "upload/m.xlsx", "ver-002": "apply/m_v2.xlsx"}  # relative: the data folder may move
    assert state["files"]["elsewhere.csv"] == str(Path("C:/elsewhere/elsewhere.csv"))
    assert state["original_ref_formulas"] == [["Calc", "A1"], ["Calc", "B2"]] and state["original_baseline"] == "p"
    assert "stop" not in state["autofix"]

    back, autofix_back = session_store.load("abc123", server.Session)
    assert back.id == "abc123" and back.work_dir == folder and back.mode == "plan"
    assert back.versions == s.versions and back.current_version_id == "ver-002" and back.current_path == s.paths["ver-002"]
    assert back.paths == s.paths and back.files == s.files and back.recalc_path == s.recalc_path and back.instances_zip == s.instances_zip
    assert back.original_ref_formulas == s.original_ref_formulas and back.ref_formulas == s.ref_formulas
    assert back.original_environment == s.original_environment
    assert back.baselines["p"]["ref_formulas"] == {("Calc", "B2")} and back.original_baseline is back.baselines["p"]
    assert back.mind_inputs_dir == s.mind_inputs_dir and back.mind_inputs_dirs == s.mind_inputs_dirs and back.inputs == s.inputs
    assert back.grid_names == s.grid_names and back.recalc == s.recalc and back.recalc_version_id == "ver-002"
    assert back.outside_run == s.outside_run and back.name == "m.xlsx" and back.last_opened == 123.0
    assert autofix_back == {"state": "done", "stage": "done", "result": {"summary": "2 cells fixed"}}
    # the locks are fresh ones, never saved
    assert not back.lock.locked() and not back.apply_lock.locked()
    assert session_store.load("nobody", server.Session) is None
