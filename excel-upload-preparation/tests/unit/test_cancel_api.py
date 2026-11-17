"""17 Nov 2026: the Cancel button on the server (app/web/server.py on top of app/cancel.py), through
the test client. No Excel and no Mind run here: the Apply, the scan's wait, the recalculation, the
fixer, the bring-in's macro, the model calls and the Mind browser are fakes that wait on events, and
an "Excel" is a number the fake kill sets an event for.

The promise these tests pin: a cancelled job leaves the workbook exactly as it was before the job,
never commits a version later, and never leaves the session saying "running". Every cancel test
ends with the same three checks (`_clean`): nothing is said to change the workbook any more, GET
.../status shows no "running", and no incident was written."""
from __future__ import annotations

import contextlib
import json
import shutil
import threading
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app import cancel, chat_context, excel_com, excel_worker, mind_client, outside_inputs, prep, session_store, stand_in_inputs
from app.settings import SETTINGS
from app.web import server
from test_outside_inputs_api import INPUTS, _fake_apply_mind_inputs, _fake_run_import, _main_book, _upload_picked  # the Tokyo-like tool and its fake Excel steps
from test_outside_inputs_api import macros  # noqa: F401  (a fixture: the main workbook's VBA, olevba replaced)
from test_grid_namer import _wb as _book  # a small workbook from a dict of cells

WAIT = 30  # seconds a test waits for a thread or an event before it fails
STOPPED = "Stopped. Nothing was changed."
OPS = [{"op": "set_value", "sheet": "Data", "cell": "Z9", "after": 1}]  # one change on the plain_grid fixture


# --- the test's own tools ---------------------------------------------------------------------------
class Gate:
    """A piece of fake work: it says when it started and waits to be let go. Meanwhile it calls `tick`
    -- the work's own progress call -- as real work reaches its next stop every few milliseconds."""

    def __init__(self) -> None:
        self.started = threading.Event()
        self.go = threading.Event()

    def hold(self, tick=None) -> None:
        self.started.set()
        deadline = time.time() + WAIT
        while not self.go.wait(0.02):
            if tick is not None:
                tick()
            assert time.time() < deadline, "the fake work was never let go"


def _until_cancelled(gate: Gate | None = None) -> None:
    """Work that only a Cancel ends (a wait inside Mind, a model call): it returns once the flag of
    this thread's job is set, as if nothing had happened -- the code after it must read the flag."""
    if gate is not None:
        gate.started.set()
    deadline = time.time() + WAIT
    while not cancel.flagged():
        assert time.time() < deadline, "the Cancel never came"
        time.sleep(0.02)


class FakeExcels:
    """The Excels the fakes "start": a pid each, noted for the current job as app/excel_com.py does
    (`cancel.note_excel`), and an event the fake `kill_excel` sets."""

    def __init__(self) -> None:
        self.killed: list[int] = []
        self.gone: dict[int, threading.Event] = {}

    def start(self, pid: int) -> threading.Event:
        self.gone[pid] = threading.Event()
        cancel.note_excel(pid)
        return self.gone[pid]

    def kill(self, pid: int) -> bool:
        self.killed.append(pid)
        if pid in self.gone:
            self.gone[pid].set()
        return pid in self.gone


class FakeWorker:
    """Stands in for the session's Excel worker (app/excel_worker.py): the calls the server makes on
    it, and the worker's rule for a Cancel -- only the item of a cancelled job is stopped, or, with
    `free`, an item that belongs to no job (the recalculation made ahead)."""

    def __init__(self, running: object = "nothing") -> None:
        self._stopped = False
        self.running = running      # the owner of the item running now: a job, None (no job's), or "nothing"
        self.waiting: list = []     # the jobs whose items wait for their turn
        self.cancels: list[tuple[list[str], bool]] = []
        self.killed = False         # its Excel was ended
        self.gate = Gate()
        self.let_go = threading.Event()

    def take_recalc(self, version):  # waits behind the item that runs, as ExcelWorker._submit does
        self.waiting.append(cancel.current())
        self.gate.started.set()
        assert self.let_go.wait(WAIT)
        return None

    def apply(self, source, operations, progress=None, defer_verify=False, recalc_ahead=False):
        self.running = cancel.current()
        self.gate.hold(lambda: progress("apply_write", "Writing", None, done=0, total=len(operations)))
        return _applied(source, Path(source).parent / "worker", operations)

    def release(self) -> None:
        pass

    def cancel_jobs(self, jobs, free):
        self.cancels.append(([j.kind for j in jobs], bool(free)))
        if any(w in jobs for w in self.waiting):
            self.let_go.set()
        if self.running in jobs:
            self.killed = True
        if free and self.running is None:
            self.killed = True
            return 1
        return 0


def _applied(source, folder, operations) -> dict:
    """What apply_operations answers after it wrote: a copy of the version, as the new file."""
    folder = Path(folder)
    n = 1
    while (folder / f"v{n}").exists():
        n += 1
    out = folder / f"v{n}" / Path(source).name
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, out)
    return {"status": "APPLIED", "method": "excel_com", "output_path": str(out), "applied": list(operations), "failed": [], "warnings": [],
            "verified_opens_in_excel": True, "change_log_entry": {"output_sha256": "0" * 64}}


def _fake_apply(calls: list, gate: Gate | None = None, deaf: bool = False):
    """app.prep.apply_operations without Excel. With a gate it waits there, between two written
    changes; `deaf`: it never reports progress while it waits (work that ignores the flag)."""
    def apply_operations(source_path, work_dir, operations, prefer_excel=True, progress=None, defer_verify=False, run=None, verify=None, excel=None):
        calls.append(Path(source_path))
        progress("apply_copy", "Copying the version", None, done=0, total=len(operations))
        if gate is not None:
            gate.hold(None if deaf else (lambda: progress("apply_write", "Writing", None, done=0, total=len(operations))))
        res = _applied(source_path, work_dir, operations)
        if not deaf:
            progress("apply_save", "Saving", None, done=len(operations), total=len(operations))
        return res

    return apply_operations


def _slow_scan(monkeypatch, gate: Gate, when=lambda s, path: True) -> None:
    """The scan waits at the gate (reporting its stage as a real scan does at every sheet and rule)."""
    real = server._analyze

    def analyze(s, path, progress=None):
        if when(s, Path(path)):
            gate.hold(lambda: progress("load", "Reading the workbook", None))
        return real(s, path, progress)

    monkeypatch.setattr(server, "_analyze", analyze)


def _bg(call):
    """One request in a thread of its own, with a client of its own: (thread, box)."""
    box: dict = {}

    def run() -> None:
        try:
            box["res"] = call(TestClient(server.app))
        except BaseException as exc:  # noqa: BLE001 -- shown by _answer
            box["error"] = exc

    t = threading.Thread(target=run, daemon=True)
    t.start()
    return t, box


def _answer(pending, seconds: float = WAIT):
    t, box = pending
    t.join(seconds)
    assert not t.is_alive(), "the request did not answer"
    assert "error" not in box, repr(box.get("error"))
    return box["res"]


def _upload(client, path: Path, defer: bool = False) -> dict:
    with path.open("rb") as f:
        res = client.post("/api/sessions", files={"file": (path.name, f, "application/octet-stream")}, data={"mode": "plan", **({"defer": "1"} if defer else {})})
    assert res.status_code == 200, res.text
    return res.json()


def _cancel(client, sid: str, kind: str | None = None) -> dict:
    res = client.post(f"/api/sessions/{sid}/cancel", json={"kind": kind} if kind else {})
    assert res.status_code == 200, res.text
    return res.json()


def _jobs(client, sid: str) -> list[dict]:
    return client.get(f"/api/sessions/{sid}/jobs").json()["jobs"]


def _incidents() -> int:
    path = Path(SETTINGS.data_folder) / "feedback" / "incidents.jsonl"
    return len(path.read_text(encoding="utf-8").splitlines()) if path.is_file() else 0


def _look(s) -> dict:
    """The workbook as a job must leave it after a Cancel: the versions, the pointer, the very same
    check result, plan and recalculation, the files, the two gauges, and the report word for word."""
    return {
        "versions": [dict(v) for v in s.versions], "current": s.current_version_id, "result": id(s.result), "plan": id(s.plan), "recalc": id(s.recalc),
        "recalc_version": s.recalc_version_id, "paths": dict(s.paths), "files": dict(s.files), "inputs": [dict(i) for i in s.inputs],
        "histories": (len(s.plan_history), len(s.apply_history)), "pending": s.pending, "zip": s.instances_zip, "grid_names": dict(s.grid_names),
        "report": json.dumps(s.result["validation_report"], sort_keys=True, default=str) if s.result else None,
    }


def _clean(client, sid: str, busy: str | None = None) -> None:
    """The three checks every cancel test ends with."""
    s = server.SESSIONS[sid]
    assert server._busy_with(s) == busy  # 1. nothing is said to change the workbook any more
    status = client.get(f"/api/sessions/{sid}/status").json()  # 2. no "running" left behind
    assert status["state"] != "running" and (status["apply"] or {}).get("state") != "running", status
    assert _jobs(client, sid) == [] and cancel.jobs(sid) == [] and not s.lock.locked() and not s.apply_lock.locked() and sid not in server.RECALCULATING
    assert _incidents() == 0  # 3. a Cancel is no problem to report


def _recalculated(s) -> None:
    """A recalculation of the current version, as the Recalculate step leaves it."""
    server._recalc_out(s, {"status": "OK", "message": "recalculated", "ran": True, "formula_errors": [], "addin_gap_errors": []}, s.current_path)


def _is_stopped(res, kind: str) -> None:
    assert res.status_code == 409, res.text
    assert res.json() == {"detail": STOPPED, "cancelled": True, "kind": kind}


# --- fixtures ---------------------------------------------------------------------------------------
@pytest.fixture
def excels(monkeypatch) -> FakeExcels:
    """`excel_com.kill_excel` replaced: no process is ever ended by these tests."""
    fake = FakeExcels()
    monkeypatch.setattr(excel_com, "kill_excel", fake.kill)
    return fake


@pytest.fixture
def client(tmp_path, monkeypatch, excels):
    """A client on a data folder of its own (sessions, incidents, mind_projects.json), with no Excel
    worker (worker_for answers None: the Apply goes to `server.apply_operations`, a fake here)."""
    real = SETTINGS.data_folder
    object.__setattr__(SETTINGS, "data_folder", tmp_path / "data")
    for registry in (server.SESSIONS, server.AUTOFIX, server.MIND_LOOPS, server.RECALCULATING, excel_worker.WORKERS):
        registry.clear()
    monkeypatch.setattr(excel_worker, "worker_for", lambda sid, apply_dir: None)
    monkeypatch.setattr(server, "CANCEL_WAIT_S", WAIT)  # a busy machine must not turn a slow restore into "stopping"
    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    monkeypatch.setitem(server._MIND_DELETE, "name", None)
    try:
        yield TestClient(server.app)
    finally:
        for registry in (server.SESSIONS, server.AUTOFIX, server.MIND_LOOPS, server.RECALCULATING, excel_worker.WORKERS):
            registry.clear()
        session_store.release_all()
        object.__setattr__(SETTINGS, "data_folder", real)


# --- the snapshot and the restore (the next builder's job uses them too) ------------------------------
def test_restore_puts_back_everything_a_job_may_touch_and_can_be_used_twice(client, plain_grid_xlsx, monkeypatch, tmp_path):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _recalculated(s)  # READY-001 now speaks of this recalculation: one of the two findings the server changes in place
    client.get(f"/api/sessions/{sid}/model-map")
    client.get(f"/api/sessions/{sid}/instance-finder")
    client.get(f"/api/sessions/{sid}/mind-instances-plan")
    before = _look(s)
    kept_for_v1 = (s.model_map, s.instance_finder, s.instance_plan_default)
    snap = server._snapshot(s)

    def spoil() -> None:
        """What a job that changes the workbook does, all of it at once."""
        new = tmp_path / f"spoiled_{len(s.versions)}.xlsx"
        shutil.copyfile(s.current_path, new)
        server._add_version(s, new, "prep", "a version that is going", [], True, "1" * 64)
        report = s.result["validation_report"]
        ready = next(f for f in report["findings"] if f["rule_id"] == "READY-001")
        ready["status"], ready["message"] = "ERROR", "changed in place"
        report["status"], report["summary"]["status_counts"] = "ERROR", {"ERROR": 99}
        s.recalc, s.recalc_version_id, s.recalc_path = {"status": "ERROR", "message": "did not run", "ran": False}, s.current_version_id, new
        s.plan, s.last_delta, s.pending = [], {"version_id": s.current_version_id}, True
        s.plan_history.append({"version_id": s.current_version_id})
        s.apply_history.append({"version_id": s.current_version_id})
        s.grid_names["Data!A2:B4"] = "A name"
        s.inputs.append({"name": "stand_in.xlsx", "path": str(new), "bytes": 1, "sha256": "x", "stand_in": True})
        s.input_seq += 1
        s.ignore_sheets.append("Data")
        s.files["Mind inputs.zip"], s.instances_zip = new, new
        s.mind_inputs_dirs[s.current_version_id], s.mind_inputs_dir = str(tmp_path), str(tmp_path)
        s.brought_in_files[s.current_version_id] = "a fingerprint"
        # what was kept for the new version: its id is used again by the next one
        s.model_map_version_id = s.instance_finder_version_id = s.instance_plan_default_version_id = s.current_version_id
        s.instance_finder_assistant = {"asked": True}

    for _ in range(2):  # the same snapshot twice: a job may go back to it after every try
        spoil()
        assert s.current_version_id == "ver-002" and _look(s) != before
        server._restore(s, snap)
        assert _look(s) == before
        assert (s.input_seq, s.ignore_sheets, s.mind_inputs_dirs, s.mind_inputs_dir, s.brought_in_files, s.recalc_path) == (0, [], {}, None, {}, before["paths"]["ver-001"])
        assert s.model_map is None and s.instance_finder is None and s.instance_finder_assistant is None and s.instance_plan_default is None
        saved = session_store.read_state(sid)  # the restore saves
        assert [v["id"] for v in saved["versions"]] == ["ver-001"] and saved["current_version_id"] == "ver-001" and saved["pending"] is False
    # what was kept for a version that stays is not thrown away
    s.model_map, s.model_map_version_id = kept_for_v1[0], "ver-001"
    spoil()
    s.model_map_version_id = "ver-001"
    # the session's Excel worker, at work for no job (the recalculation made ahead of the version that
    # is going): let go when a version is dropped, left alone when none is
    worker = FakeWorker(running=None)
    monkeypatch.setitem(excel_worker.WORKERS, sid, worker)
    server._restore(s, snap)
    assert s.model_map is kept_for_v1[0] and _look(s) == before
    assert worker.cancels == [([], True)] and worker.killed
    still = FakeWorker(running=None)
    excel_worker.WORKERS[sid] = still
    server._restore(s, snap)  # nothing to drop this time
    assert still.cancels == [] and not still.killed and _look(s) == before


def test_an_apply_asked_not_to_save_writes_nothing_to_the_state_file(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    snap = server._snapshot(s)
    out = server._apply_and_version(s, [{"action_id": "assistant", "rule_id": "UI", **OPS[0]}], "assistant", "Assistant", reanalyze=False, save=False)
    assert out["version"]["id"] == "ver-002" and s.current_version_id == "ver-002" and not s.apply_lock.locked()
    assert [v["id"] for v in session_store.read_state(sid)["versions"]] == ["ver-001"]  # judged first: nothing of it is on disk yet
    server._restore(s, snap)  # ... and taken back: the state file never knew it
    assert s.current_version_id == "ver-001" and [v["id"] for v in session_store.read_state(sid)["versions"]] == ["ver-001"]
    out = server._apply_and_version(s, [{"action_id": "assistant", "rule_id": "UI", **OPS[0]}], "assistant", "Assistant", reanalyze=False)
    assert out["version"]["id"] == "ver-002" and [v["id"] for v in session_store.read_state(sid)["versions"]] == ["ver-001", "ver-002"]  # as ever


# --- the Apply --------------------------------------------------------------------------------------
def test_a_cancelled_apply_answers_409_and_leaves_the_workbook_exactly_as_it_was(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _recalculated(s)
    before = _look(s)
    gate, calls = Gate(), []
    monkeypatch.setattr(server, "apply_operations", _fake_apply(calls, gate))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": True}))
    assert gate.started.wait(WAIT)
    job = _jobs(client, sid)
    assert [(j["kind"], j["title"], j["stopping"]) for j in job] == [("apply", "Writing the changes", False)] and job[0]["elapsed_s"] >= 0
    assert server._busy_with(s) == "changes are still being written"
    # another job that changes the workbook is refused while this one is registered
    again = client.post(f"/api/sessions/{sid}/reanalyze", json={})
    assert again.status_code == 409 and again.json()["detail"] == "Wait a moment: changes are still being written."
    out = _cancel(client, sid, "apply")
    assert out["cancelled"] == ["apply"] and out["late"] == [] and out["stopping"] == [] and out["ahead"] is False and out["mind"] is None
    assert out["message"] == STOPPED
    # the answer carries the workbook after the restore, as GET /api/sessions/{id} gives it
    assert out["session"] == client.get(f"/api/sessions/{sid}").json() and out["session"]["version"]["id"] == "ver-001"
    _is_stopped(_answer(pending), "apply")
    assert _look(s) == before and len(calls) == 1
    apply = client.get(f"/api/sessions/{sid}/status").json()["apply"]
    assert apply["state"] == "cancelled" and apply["title"] == "Stopped" and apply["message"] == "Stopped." and apply["error"] is None
    # nothing comes later: the next Apply starts from the same version and makes the same number
    gate.go.set()
    res = client.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": False})
    assert res.status_code == 200 and res.json()["version"]["id"] == "ver-002" and res.json()["version"]["parent"] == "ver-001"
    _clean(client, sid)


def test_a_cancel_during_the_scan_after_the_new_version_takes_the_version_away(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _recalculated(s)
    client.get(f"/api/sessions/{sid}/model-map")  # kept for ver-001: it must survive, the dropped version's must not
    before = _look(s)
    gate = Gate()
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    _slow_scan(monkeypatch, gate)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": True}))
    assert gate.started.wait(WAIT)
    # the point of no return of before tonight is behind: the version exists and is current
    assert [v["id"] for v in s.versions] == ["ver-001", "ver-002"] and s.current_version_id == "ver-002"
    s.model_map_version_id = "ver-002"  # as if the map had been built for the version that is going
    out = _cancel(client, sid)  # no kind: everything that runs for this workbook
    assert out["cancelled"] == ["apply"] and out["message"] == STOPPED and out["session"]["versions"] == before["versions"]
    _is_stopped(_answer(pending), "apply")
    assert _look(s) == before
    assert s.model_map is None and s.model_map_version_id is None  # the next version is ver-002 again: nothing of the old one may look current
    status = client.get(f"/api/sessions/{sid}/status").json()
    assert status["state"] == "cancelled" and status["title"] == "Stopped" and status["message"] == "Stopped." and status["error"] is None
    assert status["apply"]["state"] == "cancelled"
    # the state file holds the restored workbook: a restart never brings the dropped version back
    saved = session_store.read_state(sid)
    assert [v["id"] for v in saved["versions"]] == ["ver-001"] and saved["current_version_id"] == "ver-001"
    _clean(client, sid)


def test_the_cold_retry_does_not_run_after_a_cancel(client, plain_grid_xlsx, monkeypatch):
    """The worker's Excel ended by a Cancel used to mean "do the whole Apply again in a fresh Excel"."""
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    gate, cold = Gate(), []

    class Worker(FakeWorker):
        def apply(self, source, operations, progress=None, defer_verify=False, recalc_ahead=False):
            self.running = cancel.current()
            _until_cancelled(gate)
            raise excel_worker.WorkerUnavailable("Excel stopped answering")

    worker = Worker()
    monkeypatch.setattr(excel_worker, "worker_for", lambda sid, apply_dir: worker)
    monkeypatch.setitem(excel_worker.WORKERS, sid, worker)
    monkeypatch.setattr(server, "apply_operations", _fake_apply(cold))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS}))
    assert gate.started.wait(WAIT)
    assert _cancel(client, sid, "apply")["cancelled"] == ["apply"]
    _is_stopped(_answer(pending), "apply")
    assert cold == [] and worker.killed and worker.cancels == [(["apply"], False)]  # its own item: the worker's Excel is ended
    assert _look(s) == before and client.get(f"/api/sessions/{sid}/status").json()["apply"]["state"] == "cancelled"
    excel_worker.WORKERS.pop(sid, None)
    _clean(client, sid)


def test_the_openpyxl_fallback_stops_at_its_first_step_and_a_file_written_anyway_never_becomes_a_version(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    # 1. the real apply_operations: its Excel "dies" on the Cancel, and it falls back to openpyxl
    gate, wrote = Gate(), []

    def excel_dies(copy_path, operations, progress=None, excel=None, wb=None):
        _until_cancelled(gate)
        raise RuntimeError("the Excel was closed")

    real_fallback = prep._apply_with_openpyxl

    def fallback(copy_path, operations, progress=None):
        wrote.append("called")
        out = real_fallback(copy_path, operations, progress)
        wrote.append("written")
        return out

    monkeypatch.setattr(prep, "com_available", lambda: True)
    monkeypatch.setattr(prep, "_apply_with_excel", excel_dies)
    monkeypatch.setattr(prep, "_apply_with_openpyxl", fallback)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS}))
    assert gate.started.wait(WAIT)
    assert _cancel(client, sid, "apply")["cancelled"] == ["apply"]
    _is_stopped(_answer(pending), "apply")
    assert wrote == ["called"] and _look(s) == before  # it left at its first progress call: nothing written
    # 2. the belt: work that never looks at the flag and hands a finished file back
    gate2 = Gate()

    def deaf_apply(source_path, work_dir, operations, prefer_excel=True, progress=None, defer_verify=False, run=None, verify=None, excel=None):
        _until_cancelled(gate2)
        return _applied(source_path, work_dir, operations)

    monkeypatch.setattr(server, "apply_operations", deaf_apply)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS}))
    assert gate2.started.wait(WAIT)
    assert _cancel(client, sid, "apply")["cancelled"] == ["apply"]
    _is_stopped(_answer(pending), "apply")
    assert _look(s) == before and client.get(f"/api/sessions/{sid}/status").json()["apply"]["state"] == "cancelled"
    _clean(client, sid)


def test_work_that_ignores_the_flag_past_the_wait_answers_stopping_with_no_session(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    gate = Gate()
    monkeypatch.setattr(server, "CANCEL_WAIT_S", 0.3)
    monkeypatch.setattr(server, "apply_operations", _fake_apply([], gate, deaf=True))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS}))
    assert gate.started.wait(WAIT)
    out = _cancel(client, sid)
    assert out == {"cancelled": [], "late": [], "stopping": ["apply"], "ahead": False, "message": "Stopping…", "session": None, "mind": None}
    assert [(j["kind"], j["stopping"]) for j in _jobs(client, sid)] == [("apply", True)]
    assert server._busy_with(s) == "changes are still being written"  # until its restore is done
    again = _cancel(client, sid, "apply")  # a second click repeats harmless steps
    assert again["stopping"] == ["apply"] and again["session"] is None
    gate.go.set()  # the work comes back at last: it meets the flag at its next step
    _is_stopped(_answer(pending), "apply")
    assert _look(s) == before
    assert _cancel(client, sid)["message"] == "Nothing was running."
    _clean(client, sid)


def test_cancel_with_nothing_running_twice_after_the_end_and_too_late(client, plain_grid_xlsx, monkeypatch):
    assert client.post("/api/sessions/nope/cancel", json={}).status_code == 404
    assert client.get("/api/sessions/nope/jobs").status_code == 404
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    nothing = {"cancelled": [], "late": [], "stopping": [], "ahead": False, "message": "Nothing was running.", "mind": None}
    for kind in (None, None, "apply", "mind", "no-such-kind"):
        out = _cancel(client, sid, kind)
        assert {k: v for k, v in out.items() if k != "session"} == nothing and out["session"]["sessionId"] == sid
    # after the job has ended: its version stays
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    assert client.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": False}).status_code == 200
    out = _cancel(client, sid, "apply")
    assert out["message"] == "Nothing was running." and out["session"]["version"]["id"] == "ver-002"
    # too late: the job passed its gate (the reports are being registered) -- its result stands
    gate = Gate()
    monkeypatch.setattr(server, "build_standalone_report", lambda report, path, **kw: Path(path))
    monkeypatch.setattr(server, "build_report_workbook", lambda copy, report, path, **kw: type("Built", (), {"path": Path(path), "method": "openpyxl", "verified_opens_in_excel": None, "warnings": []})())
    real_register = server._register_file

    def slow_register(s_, path):
        gate.hold()
        return real_register(s_, path)

    monkeypatch.setattr(server, "_register_file", slow_register)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/reports"))
    assert gate.started.wait(WAIT)
    out = _cancel(client, sid)
    assert out["cancelled"] == [] and out["late"] == ["reports"] and out["stopping"] == [] and out["message"] == "It had already finished."
    gate.go.set()
    res = _answer(pending)
    assert res.status_code == 200 and res.json()["standalone_name"] in s.files and res.json()["workbook_name"] in s.files
    assert [v["id"] for v in s.versions] == ["ver-001", "ver-002"]
    _clean(client, sid)


# --- the check (scan) -------------------------------------------------------------------------------
def test_a_cancelled_reanalyze_puts_the_pointer_back(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    assert client.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": True}).status_code == 200
    before = _look(s)
    assert before["current"] == "ver-002"
    gate = Gate()
    _slow_scan(monkeypatch, gate)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/reanalyze", json={"versionId": "ver-001"}))  # Versions > Make current
    assert gate.started.wait(WAIT)
    assert s.current_version_id == "ver-001" and [j["title"] for j in _jobs(client, sid)] == ["Checking the workbook"]
    out = _cancel(client, sid, "scan")
    assert out["cancelled"] == ["scan"] and out["session"]["version"]["id"] == "ver-002"
    _is_stopped(_answer(pending), "scan")
    assert _look(s) == before
    assert client.get(f"/api/sessions/{sid}/status").json()["state"] == "cancelled"
    _clean(client, sid)


def test_a_cancelled_first_check_leaves_the_workbook_pending_in_recent_workbooks(client, plain_grid_xlsx, monkeypatch):
    body = _upload(client, plain_grid_xlsx, defer=True)
    sid = body["sessionId"]
    s = server.SESSIONS[sid]
    assert body["pending"] is True
    before = _look(s)
    gate = Gate()
    _slow_scan(monkeypatch, gate)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/analyze", json={"ignore_sheets": []}))
    assert gate.started.wait(WAIT)
    out = _cancel(client, sid, "scan")
    assert out["cancelled"] == ["scan"] and out["message"] == STOPPED
    assert out["session"] is None  # before the first check there is no analysed workbook to send
    _is_stopped(_answer(pending), "scan")
    assert _look(s) == before and s.pending and s.result is None
    status = client.get(f"/api/sessions/{sid}/status").json()
    assert status["state"] == "cancelled" and status["message"] == "Stopped." and status["error"] is None
    row = next(w for w in client.get("/api/saved-workbooks").json()["workbooks"] if w["id"] == sid)
    assert row["pending"] is True and row["has_result"] is False and row["name"] == plain_grid_xlsx.name
    # the three checks, here with what the workbook said before the check: uploaded, "not analysed yet"
    _clean(client, sid, busy="the workbook has not been analysed yet")
    # ... and it can be checked again, from Recent workbooks: then nothing at all is left of the Cancel
    gate.go.set()
    assert client.post(f"/api/sessions/{sid}/analyze", json={}).status_code == 200 and not s.pending
    _clean(client, sid)


# --- Recalculate ------------------------------------------------------------------------------------
def _fake_recalculate(excels: FakeExcels, calls: list, pid: int = 4101):
    """app.recalc.recalculate: an Excel of its own, and when that Excel is ended the answer the real
    one gives -- "did not run", never an exception."""
    def recalculate(path, save=True):
        calls.append(Path(path))
        assert excels.start(pid).wait(WAIT)
        return {"status": "ERROR", "message": "Excel COM recalculation failed: the Excel was closed", "ran": False, "formula_errors": [], "addin_gap_errors": []}

    return recalculate


def test_a_cancelled_recalculate_keeps_the_last_recalculation(client, plain_grid_xlsx, monkeypatch, excels):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _recalculated(s)
    before, last = _look(s), s.recalc
    calls: list = []
    monkeypatch.setattr(server, "recalculate", _fake_recalculate(excels, calls))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/recalculate"))
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert sid in server.RECALCULATING and server._busy_with(s) == "a recalculation is still running"
    assert [j["title"] for j in _jobs(client, sid)] == ["Recalculating"]
    out = _cancel(client, sid, "recalc")
    assert out["cancelled"] == ["recalc"] and out["message"] == STOPPED and out["session"]["recalc"]["message"] == "recalculated"
    _is_stopped(_answer(pending), "recalc")
    assert excels.killed == [4101]  # only the Excel this job started
    assert s.recalc is last and _look(s) == before  # "did not run" was not stored over it
    _clean(client, sid)


def test_a_recalculate_waiting_behind_the_recalculation_made_ahead_is_let_go_at_once(client, plain_grid_xlsx, monkeypatch, excels):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    worker, own = FakeWorker(running=None), []  # the item running in the worker belongs to no job: the recalculation made ahead
    monkeypatch.setitem(excel_worker.WORKERS, sid, worker)
    monkeypatch.setattr(server, "recalculate", _fake_recalculate(excels, own))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/recalculate"))
    assert worker.gate.started.wait(WAIT)
    t0 = time.time()
    out = _cancel(client, sid, "recalc")
    assert out["cancelled"] == ["recalc"] and out["ahead"] is True and out["message"] == STOPPED
    _is_stopped(_answer(pending), "recalc")
    assert time.time() - t0 < 5
    assert worker.cancels == [(["recalc"], True)] and worker.killed  # the recalculation made ahead was stopped too
    assert own == [] and excels.killed == []  # ... and the route did not fall through to an Excel of its own
    assert server.RECALCULATING == set() and _look(s) == before
    excel_worker.WORKERS.pop(sid, None)
    _clean(client, sid)


def test_only_a_recalculation_made_ahead_counts_as_stopped(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    worker = FakeWorker(running=None)
    monkeypatch.setitem(excel_worker.WORKERS, sid, worker)
    out = _cancel(client, sid)  # Start's Cancel between two steps: no job, only the worker at work
    assert out["cancelled"] == [] and out["late"] == [] and out["stopping"] == [] and out["ahead"] is True and out["message"] == STOPPED
    assert worker.cancels == [([], True)] and out["session"]["sessionId"] == sid
    worker2 = FakeWorker(running=None)
    excel_worker.WORKERS[sid] = worker2
    assert _cancel(client, sid, "apply")["ahead"] is False and worker2.cancels == [([], False)] and not worker2.killed  # another kind leaves it alone
    excel_worker.WORKERS.pop(sid, None)
    _clean(client, sid)


def test_a_cancelled_recalculate_never_drops_the_version_an_apply_made_meanwhile(client, plain_grid_xlsx, monkeypatch, excels):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "recalculate", _fake_recalculate(excels, []))
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/recalculate"))
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    res = client.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": True})  # commits while the Recalculate waits
    assert res.status_code == 200 and res.json()["version"]["id"] == "ver-002"
    after_apply = _look(s)
    assert _cancel(client, sid, "recalc")["cancelled"] == ["recalc"]
    _is_stopped(_answer(pending), "recalc")
    assert _look(s) == after_apply and s.current_version_id == "ver-002"  # a job that writes only at its end never restores
    _clean(client, sid)


# --- the numbers check, the reports, the names, the finder, the assistant ----------------------------
def test_cancelling_the_numbers_check_leaves_the_apply_running_in_the_worker_alone(client, plain_grid_xlsx, monkeypatch, excels):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    assert client.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": False}).status_code == 200  # ver-002: something to compare

    def compare(original, prepared, moves, progress=None):
        assert excels.start(4202).wait(WAIT)  # its own Excel, ended by the Cancel
        return {"ran": False, "clean": False, "verdict": "The numbers could not be compared: the Excel was closed.", "counts": {}, "regressions": []}

    monkeypatch.setattr(server, "compare_workbooks", compare)
    worker = FakeWorker()
    monkeypatch.setattr(excel_worker, "worker_for", lambda sid, apply_dir: worker)
    monkeypatch.setitem(excel_worker.WORKERS, sid, worker)
    applying = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": False}))
    assert worker.gate.started.wait(WAIT)
    checking = _bg(lambda c: c.post(f"/api/sessions/{sid}/numbers-check"))
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert sorted(j["kind"] for j in _jobs(client, sid)) == ["apply", "numbers"]
    out = _cancel(client, sid, "numbers")
    assert out["cancelled"] == ["numbers"] and out["late"] == [] and out["message"] == STOPPED
    _is_stopped(_answer(checking), "numbers")  # "could not be compared" is no verdict: the request says it was stopped
    assert excels.killed == [4202]
    assert worker.cancels == [(["numbers"], False)] and not worker.killed  # the worker's Excel, at work for the Apply, is not ended
    assert [j["kind"] for j in _jobs(client, sid)] == ["apply"]
    worker.gate.go.set()
    res = _answer(applying)
    assert res.status_code == 200 and res.json()["version"]["id"] == "ver-003" and s.current_version_id == "ver-003"  # the Apply commits
    excel_worker.WORKERS.pop(sid, None)
    _clean(client, sid)


def test_cancelled_reports_register_no_file(client, plain_grid_xlsx, monkeypatch, excels):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    monkeypatch.setattr(server, "build_standalone_report", lambda report, path, **kw: Path(path))

    def report_workbook(copy, report, path, **kw):
        assert excels.start(4303).wait(WAIT)
        raise RuntimeError("the Excel was closed")

    monkeypatch.setattr(server, "build_report_workbook", report_workbook)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/reports"))
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert [j["title"] for j in _jobs(client, sid)] == ["Making the reports"]
    assert _cancel(client, sid, "reports")["cancelled"] == ["reports"]
    _is_stopped(_answer(pending), "reports")  # the error of the ended Excel is not shown as a failure
    assert _look(s) == before and excels.killed == [4303]
    _clean(client, sid)


def test_cancelled_suggest_names_leaves_the_names_and_the_plan(client, section_heading_grids_xlsx, monkeypatch):
    sid = _upload(client, section_heading_grids_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    gate = Gate()

    def suggest(contexts, **kw):
        gate.started.set()
        cancel.call(lambda: gate.go.wait(WAIT))  # as app/llm.py makes its request: given up when Cancel is pressed
        return {"available": True, "names": {"CF!B8:C8": "Too late"}, "message": None}

    monkeypatch.setattr(server, "suggest_names", suggest)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/grid-names", json={"apply": True}))
    assert gate.started.wait(WAIT)
    assert [j["title"] for j in _jobs(client, sid)] == ["Suggesting names"]
    assert _cancel(client, sid, "names")["cancelled"] == ["names"]
    _is_stopped(_answer(pending), "names")
    assert _look(s) == before and s.grid_names == {}
    gate.go.set()
    _clean(client, sid)


def test_a_cancelled_finder_keeps_no_explanation(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    assert client.get(f"/api/sessions/{sid}/instance-finder").status_code == 200  # the finder alone: kept for this version
    found = s.instance_finder
    gate = Gate()

    def advise(result):
        gate.started.set()
        cancel.call(lambda: gate.go.wait(WAIT))  # the assistant's call, given up when Cancel is pressed
        return {"asked": True, "available": True, "message": None, "summary": "too late"}

    monkeypatch.setattr(server, "make_finder_advisor", lambda: advise)
    pending = _bg(lambda c: c.get(f"/api/sessions/{sid}/instance-finder", params={"assistant": "1"}))
    assert gate.started.wait(WAIT)
    assert [j["title"] for j in _jobs(client, sid)] == ["Looking for instances"]
    assert _cancel(client, sid, "finder")["cancelled"] == ["finder"]
    _is_stopped(_answer(pending), "finder")
    assert s.instance_finder is found and s.instance_finder_assistant is None
    gate.go.set()
    _clean(client, sid)


def test_a_cancelled_answer_ends_the_stream_with_cancelled(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    gate = Gate()

    def fake_chat(messages, system_prompt, **kw):
        if kw.get("on_delta"):
            kw["on_delta"]("Thinking")
        gate.started.set()
        cancel.call(lambda: gate.go.wait(WAIT))  # the model call in flight
        return {"available": True, "text": "too late", "message": None}

    monkeypatch.setattr(chat_context, "chat_completion", fake_chat)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/chat/stream", json={"history": [], "question": "what is wrong here?"}))
    assert gate.started.wait(WAIT)
    assert [(j["kind"], j["title"]) for j in _jobs(client, sid)] == [("chat", "The assistant is answering")]
    out = _cancel(client, sid, "chat")
    assert out["cancelled"] == ["chat"] and out["message"] == STOPPED
    res = _answer(pending)
    lines = [json.loads(x) for x in res.text.splitlines() if x.strip()]
    assert res.status_code == 200 and lines[-1] == {"type": "cancelled"}
    assert not any(x["type"] in ("done", "error") for x in lines)
    gate.go.set()
    _clean(client, sid)


def test_a_cancelled_grid_namer_and_instances_build_take_everything_back(client, monkeypatch, excels, tmp_path):
    book = _book(tmp_path / "namer.xlsx", {"Model": {"B3": "Age", "C3": "Rate", "B4": 30, "C4": 0.01, "E7": "=C4*2"}})
    sid = _upload(client, book)["sessionId"]
    s = server.SESSIONS[sid]
    old_zip = tmp_path / "older.zip"
    old_zip.write_bytes(b"zip")
    s.instances_zip = old_zip  # the zip of an earlier build
    before = _look(s)
    monkeypatch.setattr(server, "apply_operations", _fake_apply([]))
    # the Grid Namer's Submit: stopped during the scan of its new version
    gate = Gate()
    real_analyze = server._analyze
    _slow_scan(monkeypatch, gate)
    area = {"sheet": "Model", "ref": "B3:C4", "name": "Mortality", "flags": ["input"]}
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/grid-namer", json={"areas": [area], "nameRest": True}))
    assert gate.started.wait(WAIT)
    assert [v["source"] for v in s.versions] == ["upload", "grid_namer"]
    assert _cancel(client, sid, "apply")["cancelled"] == ["apply"]
    _is_stopped(_answer(pending), "apply")
    assert _look(s) == before
    monkeypatch.setattr(server, "_analyze", real_analyze)
    # Build the Mind instances: stopped while its main workbook is written (an Excel of this request's own)
    op = {"op": "set_value", "action_id": "mind_instances", "rule_id": "INST-001", "sheet": "Model", "cell": "Z9", "before": None, "after": 1}
    monkeypatch.setattr(server, "build_instance_plan", lambda analysis, mp, **kw: {
        "found": True, "status": "ready", "notes": [], "operations": [op], "main": {"file": "m_Main.xlsx", "model_file": "m.xlsx", "keys": ["a"], "grids": []}})

    def main_workbook(plan, model_path, main_path):
        assert excels.start(4404).wait(WAIT)
        raise RuntimeError("the Excel was closed")

    monkeypatch.setattr(server, "write_main_workbook", main_workbook)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/mind-instances-plan/apply", json={"reanalyze": False}))
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert [v["source"] for v in s.versions] == ["upload", "instances"] and s.instances_zip is None
    assert _cancel(client, sid)["cancelled"] == ["apply"]
    _is_stopped(_answer(pending), "apply")
    assert excels.killed == [4404]
    after = _look(s)
    # the plan kept for the dropped version went with it (the model map of ver-001 was built meanwhile: it stays)
    assert after == before and s.instances_zip == old_zip and s.instance_plan_default_version_id in (None, "ver-001")
    _clean(client, sid)


# --- Fix all issues ---------------------------------------------------------------------------------
def _wait_fixer(client, sid: str) -> dict:
    deadline = time.time() + WAIT
    while True:
        st = client.get(f"/api/sessions/{sid}/auto-fix").json()
        if st["state"] != "running":
            return st
        assert time.time() < deadline, st
        time.sleep(0.05)


def test_the_fixer_ends_cancelled_with_no_version(client, plain_grid_xlsx, monkeypatch, excels):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _recalculated(s)
    before = _look(s)
    monkeypatch.setattr(server, "com_available", lambda: True)

    def run_autofix(source, work_dir, cfg):
        cfg.progress("read", "Reading the workbook", None)
        assert excels.start(4505).wait(WAIT)  # the fixer's own Excel, ended by the Cancel
        return {"status": "error", "message": "Excel stopped answering", "log": []}  # what the real one answers then

    monkeypatch.setattr(server, "run_autofix", run_autofix)
    assert client.post(f"/api/sessions/{sid}/auto-fix", json={"use_assistant": False}).json()["state"] == "running"
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert [(j["kind"], j["title"]) for j in _jobs(client, sid)] == [("autofix", "Fixing the formula errors")]
    assert client.post(f"/api/sessions/{sid}/apply", json={"operations": OPS}).status_code == 409
    out = _cancel(client, sid, "autofix")
    assert out["cancelled"] == ["autofix"] and out["message"] == STOPPED and out["session"]["version"]["id"] == "ver-001"
    st = client.get(f"/api/sessions/{sid}/auto-fix?full=1").json()
    assert st["state"] == "cancelled" and st["message"] == STOPPED and st["error"] is None and st["version"] is None
    assert st.get("incident") is None and st.get("result") is None
    assert excels.killed == [4505] and _look(s) == before
    _clean(client, sid)


def test_a_fixer_cancelled_during_its_last_check_drops_the_version_it_had_added(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _recalculated(s)
    before, last = _look(s), s.recalc
    monkeypatch.setattr(server, "com_available", lambda: True)

    def run_autofix(source, work_dir, cfg):
        res = _applied(source, work_dir, [{"op": "set_formula", "action_id": "autofix", "rule_id": "READY-001", "sheet": "Data", "cell": "B4", "after": "=A4"}])
        return {**res, "status": "clean", "summary": "1 cell fixed", "cells_rewritten": 1, "errors_after": 0, "log": [],
                "recalc": {"status": "OK", "message": "recalculated by the fixer", "ran": True, "formula_errors": [], "addin_gap_errors": []}}

    monkeypatch.setattr(server, "run_autofix", run_autofix)
    gate = Gate()
    _slow_scan(monkeypatch, gate)
    assert client.post(f"/api/sessions/{sid}/auto-fix", json={"use_assistant": False}).json()["state"] == "running"
    assert gate.started.wait(WAIT)
    assert [v["source"] for v in s.versions] == ["upload", "autofix"] and s.recalc is not last  # the version and its recalculation are in
    assert _cancel(client, sid)["cancelled"] == ["autofix"]
    st = _wait_fixer(client, sid)
    assert st["state"] == "cancelled" and st["version"] is None and st["message"] == STOPPED
    assert _look(s) == before and s.recalc is last  # no version, and the recalculation is back
    assert client.get(f"/api/sessions/{sid}/status").json()["state"] == "cancelled"
    # "Stop, keep what is fixed" is another thing and still answers
    assert client.post(f"/api/sessions/{sid}/auto-fix/stop").json() == {"state": "cancelled"}
    _clean(client, sid)


# --- the outside files ------------------------------------------------------------------------------
@pytest.fixture
def excel_steps(monkeypatch):
    """The bring-in's Excel steps replaced (as in test_outside_inputs_api.py); Excel "available"."""
    calls: dict[str, list] = {"import": [], "mind_inputs": []}
    monkeypatch.setattr(server, "com_available", lambda: True)
    monkeypatch.setattr(outside_inputs, "run_import", _fake_run_import(calls["import"]))
    monkeypatch.setattr(outside_inputs, "apply_mind_inputs", _fake_apply_mind_inputs(calls["mind_inputs"]))
    monkeypatch.setattr(server, "compare_workbooks", lambda original, prepared, moves, progress=None: {
        "ran": True, "clean": True, "verdict": "no computed value changed", "counts": {"compared": 1, "same": 1}, "regressions": []})
    return calls


def _wait_outside(client, sid: str) -> dict:
    deadline = time.time() + WAIT
    while True:
        st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
        if not st["running"]:
            return st
        assert time.time() < deadline, st
        time.sleep(0.05)


def test_a_running_bring_in_ends_cancelled_with_both_versions_gone_and_the_lock_free(client, macros, excel_steps, monkeypatch, tmp_path):  # noqa: F811
    sid = _upload_picked(client, _main_book(tmp_path), INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    gate = Gate()
    _slow_scan(monkeypatch, gate, when=lambda s_, path: path.parent.name == "mind_inputs")  # its last scan: both versions exist
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    assert gate.started.wait(WAIT)
    assert [v["source"] for v in s.versions] == ["upload", "import", "mind_inputs"] and s.apply_lock.locked() and s.mind_inputs_dirs
    assert [(j["kind"], j["title"]) for j in _jobs(client, sid)] == [("outside", "Bringing in the input files")]
    # the old Stop still ends only the wait: a running bring-in answers 409 there
    assert client.post(f"/api/sessions/{sid}/outside-inputs/stop").status_code == 409
    out = _cancel(client, sid, "outside")
    assert out["cancelled"] == ["outside"] and out["message"] == STOPPED and out["session"]["versions"] == before["versions"]
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["state"] == "cancelled" and st["running"] is False and st["error"] is None and st["versions"] == [] and st["message"] == STOPPED
    assert st["result"] is None and st.get("incident") is None and st["outsideInputs"]["brought_in"] is False and st["outsideInputs"]["brought_in_reason"] is None
    assert _look(s) == before and s.mind_inputs_dirs == {} and s.mind_inputs_dir is None and s.brought_in_files == {}
    assert not s.apply_lock.locked()
    # nothing of it is left: the same step runs again from the same version, to the same numbers
    gate.go.set()
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    assert _wait_outside(client, sid)["state"] == "done" and [v["id"] for v in s.versions] == ["ver-001", "ver-002", "ver-003"]
    _clean(client, sid)


def test_a_bring_in_cancelled_in_the_macro_throws_away_the_stand_in_files_it_picked(client, macros, excel_steps, monkeypatch, tmp_path, excels):  # noqa: F811
    sid = _upload_picked(client, _main_book(tmp_path), INPUTS[:5])["sessionId"]  # one file is missing
    s = server.SESSIONS[sid]
    before = _look(s)
    monkeypatch.setattr(server, "llm_available", lambda: True)
    monkeypatch.setattr(stand_in_inputs, "assistant_ask", lambda: (lambda system, user: "an answer"))

    def make_stand_ins(base_path, rows, need, macro, vba, work, ask, real_files=None, messages=None, progress=None):
        made = Path(work) / INPUTS[5]
        made.parent.mkdir(parents=True, exist_ok=True)
        made.write_bytes(b"a stand-in")
        return {"ok": True, "explain": "made from the model's own numbers", "files": {need[0]: made}, "left_out": [], "attempts": [{"n": 1, "problem": None, "details": None}]}

    def run_import(workbook, rows, files, macro, out_path, work_dir, timeout_s=900, progress=None, messages=None, restore_rows=None):
        assert excels.start(4606).wait(WAIT)  # the macro's Excel, ended by the Cancel
        return {"macro": macro, "seconds": 1.0, "files": [], "kept": [], "dialogs": [], "stopped_early": True, "error": "the Excel was closed", "saved": None, "ok": False}

    monkeypatch.setattr(stand_in_inputs, "make_stand_ins", make_stand_ins)
    monkeypatch.setattr(outside_inputs, "run_import", run_import)
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"standIns": True}).status_code == 202
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert any(i.get("stand_in") for i in s.inputs)  # the stand-in is among the picked files now
    assert _cancel(client, sid)["cancelled"] == ["outside"]
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["state"] == "cancelled" and st["error"] is None and st["standIns"] is None and st["outsideInputs"]["stand_ins"] == []
    assert excels.killed == [4606] and _look(s) == before and not any(i.get("stand_in") for i in s.inputs)
    _clean(client, sid)


def test_a_wait_for_the_input_files_ends_stopped(client, macros, excel_steps, tmp_path):  # noqa: F811
    sid = _upload_picked(client, _main_book(tmp_path), INPUTS[:5])["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    res = client.post(f"/api/sessions/{sid}/outside-inputs/run", json={"waitForInputs": True})
    assert res.status_code == 202 and res.json()["state"] == "waiting_for_inputs"
    assert _jobs(client, sid) == []  # the wait is no job: it holds nothing
    out = _cancel(client, sid)
    assert out["cancelled"] == ["outside"] and out["stopping"] == [] and out["message"] == STOPPED
    st = client.get(f"/api/sessions/{sid}/outside-inputs").json()
    assert st["state"] == "stopped" and st["error"] == server._WAIT_STOPPED and st["stages"] == ["waiting", "stopped"]
    assert _look(s) == before and excel_steps["import"] == []
    assert _cancel(client, sid, "outside")["message"] == "Nothing was running."
    _clean(client, sid)


def test_cancelled_next_quarter_files_register_no_zip(client, macros, excel_steps, monkeypatch, tmp_path, excels):  # noqa: F811
    from app import next_quarter_inputs

    sid = _upload_picked(client, _main_book(tmp_path), INPUTS)["sessionId"]
    s = server.SESSIONS[sid]
    assert client.post(f"/api/sessions/{sid}/outside-inputs/run", json={}).status_code == 202
    assert _wait_outside(client, sid)["state"] == "done"
    before = _look(s)

    def make_next_quarter(base, table, paths, grids, macro, folder, messages=None, progress=None, model_name=None, made_at=None):
        progress("Running the model's import macro")
        assert excels.start(4707).wait(WAIT)
        zip_path = Path(folder) / "Next quarter Mind inputs.zip"
        zip_path.parent.mkdir(parents=True, exist_ok=True)
        zip_path.write_bytes(b"zip")
        return {"ok": True, "zip": str(zip_path), "summary": "made", "files": []}  # an answer that ignores the ended Excel

    monkeypatch.setattr(next_quarter_inputs, "make_next_quarter", make_next_quarter)
    monkeypatch.setattr(next_quarter_inputs, "classify_files", lambda paths, grids, table: {"grid_files": [], "raw": [p.name for p in paths], "notes": [], "unused": []})
    res = client.post(f"/api/sessions/{sid}/next-quarter-inputs", files=[("files", (INPUTS[0], b"new quarter", "application/octet-stream"))])
    assert res.status_code == 202, res.text
    for _ in range(500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert [(j["kind"], j["title"]) for j in _jobs(client, sid)] == [("next_quarter", "Making next quarter's input files")]
    assert _cancel(client, sid, "next_quarter")["cancelled"] == ["next_quarter"]
    st = client.get(f"/api/sessions/{sid}/next-quarter-inputs").json()
    assert st["state"] == "cancelled" and st["running"] is False and st["error"] is None and st["result"] is None and st["message"] == STOPPED
    assert excels.killed == [4707] and _look(s) == before and "Next quarter Mind inputs.zip" not in s.files
    _clean(client, sid)


# --- Run in Mind ------------------------------------------------------------------------------------
RUN = {"skipMind": False, "checkNumbers": False, "useAssistant": False, "actions": []}


class FakeMind:
    """app.mind_client for the loop (`mind_loop.mc` is that module: its functions are replaced one by
    one). Each step is a function of the test; the default answers lead to a model that was run."""

    def __init__(self, monkeypatch, **steps) -> None:
        self.calls: list[str] = []
        self.sessions = {"opened": 0, "closed": 0}
        answers = {
            "open_manager": True, "ensure_folder": True, "create_blank_project": "https://mind.example/init/1", "open_project": "empty",
            "upload_and_convert": {"success": True, "steps": {}, "errors": [], "text": ""}, "add_template": True, "template_names": [],
            "run_model": {"completed": True, "audit_consistent": True, "text": ""}, "export_summary": {"opened": True}, "screenshot": None,
            "delete_project": "deleted",
        }
        for name, default in answers.items():
            monkeypatch.setattr(mind_client, name, self._step(name, steps.get(name, default)))
        monkeypatch.setattr(mind_client, "mind_session", self._session)

    def _step(self, name: str, answer):
        def step(*args, **kwargs):
            self.calls.append(name)
            return answer(*args, **kwargs) if callable(answer) else answer

        return step

    @contextlib.contextmanager
    def _session(self, headed: bool = False):
        self.sessions["opened"] += 1
        try:
            yield None, type("Page", (), {"set_viewport_size": lambda self, size: None})()
        finally:
            self.sessions["closed"] += 1  # Edge closes on the loop's own thread, whatever ended the step


def _mind(client, sid: str) -> dict:
    return client.get(f"/api/sessions/{sid}/mind-loop").json()


def _wait_mind(client, sid: str, states=("done", "error", "cancelled")) -> dict:
    deadline = time.time() + WAIT
    while True:
        st = _mind(client, sid)
        if st["state"] in states:
            return st
        assert time.time() < deadline, st
        time.sleep(0.05)


def _projects_file() -> dict:
    path = Path(SETTINGS.data_folder) / "mind_projects.json"
    return json.loads(path.read_text(encoding="utf-8")) if path.is_file() else {}


def test_cancel_before_the_project_is_made_leaves_nothing_in_mind(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    gate, hold = Gate(), threading.Event()

    def open_manager(page):  # a wait inside Mind; then, still stopping, it holds a little so the state can be read
        _until_cancelled(gate)
        assert hold.wait(WAIT)
        return True

    mind = FakeMind(monkeypatch, open_manager=open_manager)
    assert client.post(f"/api/sessions/{sid}/mind-loop", json=RUN).status_code == 200
    assert gate.started.wait(WAIT)
    assert [(j["kind"], j["title"]) for j in _jobs(client, sid)] == [("mind", "Uploading to Mind")]
    stopping = _bg(lambda c: c.post(f"/api/sessions/{sid}/cancel", json={"kind": "mind"}))
    st = _wait_mind(client, sid, states=("cancelling",))
    assert st["message"] == "Stopping…" and st["project"] is None
    # while it is stopping Mind and the workbook are still taken: Remove and the bring-in are refused
    refused = client.delete(f"/api/saved-workbooks/{sid}")
    assert refused.status_code == 409 and "Run in Mind is still running" in refused.json()["detail"]
    monkeypatch.setattr(server, "com_available", lambda: True)
    with pytest.raises(server.HTTPException) as bring_in:
        server._free_for_bring_in(s, sid)
    assert bring_in.value.status_code == 409 and bring_in.value.detail == "Wait a moment: Run in Mind is using this workbook."
    assert client.post(f"/api/sessions/{sid}/mind-loop", json=RUN).status_code == 409  # one loop for the whole app
    assert client.post("/api/mind/sign-in", json={}).status_code == 409 and client.get("/api/mind/sign-in", params={"check": True}).status_code == 409
    hold.set()
    out = _answer(stopping).json()
    assert out["cancelled"] == ["mind"] and out["message"] == "Stopped. Nothing was uploaded to Mind." and out["mind"] is None
    st = _mind(client, sid)
    assert st["state"] == "cancelled" and st["message"] == "Stopped. Nothing was uploaded to Mind." and st["project"] is None
    assert st["kept"] == [] and st["earlier"] == [] and st["error"] is None and st.get("incident") is None
    assert "create_blank_project" not in mind.calls and "delete_project" not in mind.calls and _projects_file() == {}
    assert mind.sessions == {"opened": 1, "closed": 1} and _look(s) == before  # Edge closed; nothing in the session changed
    report = json.loads((Path(st["work_dir"]) / "loop_report.json").read_text(encoding="utf-8"))
    assert report["verdict"] == "cancelled" and report["reason"] == "stopped on request" and report["leftover_projects"] == []
    _clean(client, sid)


def test_cancel_while_mind_makes_the_project_says_it_may_be_there_and_asks_twice(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    gate = Gate()

    def create(page, name):
        _until_cancelled(gate)
        return None  # the click was made; the tile was not seen yet when the run was stopped

    mind = FakeMind(monkeypatch, create_blank_project=create, delete_project="not found")
    assert client.post(f"/api/sessions/{sid}/mind-loop", json=RUN).status_code == 200
    assert gate.started.wait(WAIT)
    live = _mind(client, sid)["project"]
    assert live["status"] == "creating" and live["name"].startswith(mind_client.PROJECT_PREFIX) and live["url"] == mind_client.MAIN_URL
    out = _cancel(client, sid)
    name = live["name"]
    assert out["cancelled"] == ["mind"] and out["message"] == f"Stopped. A project named {name} may be in Mind." and out["mind"]["status"] == "creating"
    st = _mind(client, sid)
    assert st["state"] == "cancelled" and st["message"] == out["message"] and [p["name"] for p in st["kept"]] == [name]
    assert _projects_file()[name] == {**_projects_file()[name], "session": sid, "status": "creating", "ran": False, "cancelled": True, "deleted": None, "tries": 0}
    assert "delete_project" not in mind.calls
    # Mind may still be making it: the first "not there" keeps the button, the second settles it
    first = client.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name}).json()
    assert first == {"name": name, "result": "not found", "message": "It is not in Mind yet. Try again in a minute.", "url": mind_client.MAIN_URL}
    assert [p["name"] for p in _mind(client, sid)["kept"]] == [name] and _projects_file()[name]["tries"] == 1
    second = client.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name}).json()
    assert second["result"] == "not found" and second["message"] == "It was not in Mind: nothing to delete."
    st = _mind(client, sid)
    assert st["kept"] == [] and _projects_file()[name]["deleted"] == "not found" and st["message"] == f"Stopped. The project {name} is not in Mind."
    _clean(client, sid)


def test_cancel_after_the_project_is_made_keeps_it_names_it_and_offers_to_delete_it(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    gate = Gate()

    def open_project(page, init):
        _until_cancelled(gate)
        return "empty"

    mind = FakeMind(monkeypatch, open_project=open_project)
    assert client.post(f"/api/sessions/{sid}/mind-loop", json=RUN).status_code == 200
    assert gate.started.wait(WAIT)
    out = _cancel(client, sid, "mind")
    name = out["mind"]["name"]
    assert out["cancelled"] == ["mind"] and out["message"] == f"Stopped. The project {name} is still in Mind."
    assert out["mind"] == {"name": name, "status": "created", "ran": False, "deleted": None, "url": mind_client.MAIN_URL}
    st = _mind(client, sid)
    assert st["state"] == "cancelled" and st["project"] == out["mind"] and st["kept"] == [out["mind"]] and st["earlier"] == []
    assert st["message"] == out["message"] and st["timing"]["remaining_s"] == 0
    # nothing was deleted without a click, and nothing after the stop was sent to Mind
    assert "delete_project" not in mind.calls and "upload_and_convert" not in mind.calls and mind.sessions == {"opened": 1, "closed": 1}
    report = json.loads((Path(st["work_dir"]) / "loop_report.json").read_text(encoding="utf-8"))
    assert report["verdict"] == "cancelled" and report["leftover_projects"] == [name]
    # a second run after it -- this one runs to its end -- still lists the kept project ...
    time.sleep(1.1)  # a run's folder and its projects are named by the second it started in
    FakeMind(monkeypatch)
    assert client.post(f"/api/sessions/{sid}/mind-loop", json={**RUN, "maxIterations": 1}).status_code == 200
    st2 = _wait_mind(client, sid)
    assert st2["state"] == "done" and st2["message"] is None and st2["project"]["status"] == "run" and st2["project"]["ran"] is True
    assert [p["name"] for p in st2["kept"]] == [name] and st2["project"]["name"] != name
    # ... a Cancel right after the loop returned leaves "done" ...
    late = _cancel(client, sid, "mind")
    assert late["cancelled"] == [] and late["message"] == "Nothing was running." and _mind(client, sid)["state"] == "done"
    # ... and so does the idle answer, after a restart emptied what the server remembers of the runs
    server.MIND_LOOPS.clear()
    idle = _mind(client, sid)
    assert idle["state"] == "idle" and idle["project"] is None and [p["name"] for p in idle["kept"]] == [name] and idle["message"] is None
    # the person's own click deletes it
    done = client.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name})
    assert done.status_code == 200 and done.json() == {"name": name, "result": "deleted", "message": "Deleted from Mind.", "url": mind_client.MAIN_URL}
    assert _mind(client, sid)["kept"] == [] and _projects_file()[name]["deleted"] == "deleted"
    assert server._busy_with(s) is None
    _clean(client, sid)


def test_a_project_that_was_not_made_is_not_listed_as_left_over(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    mind = FakeMind(monkeypatch, create_blank_project=None)  # Mind made nothing
    assert client.post(f"/api/sessions/{sid}/mind-loop", json={**RUN, "maxIterations": 1}).status_code == 200
    st = _wait_mind(client, sid)
    assert st["state"] == "done" and st["report"]["verdict"] == "stuck" and "project was not created" in st["report"]["reason"]
    assert st["report"]["leftover_projects"] == [] and st["project"] is None and st["kept"] == [] and st["earlier"] == []
    assert [e["status"] for e in st["events"] if e["event"] == "project"] == ["creating", "not created"]
    assert [e["status"] for e in _projects_file().values()] == ["not created"] and "delete_project" not in mind.calls
    late = _cancel(client, sid, "mind")
    assert late["cancelled"] == [] and _mind(client, sid)["state"] == "done"
    _clean(client, sid)


def test_delete_from_mind_refuses_a_foreign_name_and_says_each_result_in_plain_words(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    prefix, folder = mind_client.PROJECT_PREFIX, "MindReady_Sandbox"

    def entry(session=sid, status="created", **more):
        return {"session": session, "folder": folder, "status": status, "ran": False, "made": "2026-11-17T01:00:00Z", "cancelled": True, "deleted": None, "tries": 0, **more}

    names = {k: f"{prefix}_20261117_0100{i:02d}_i1" for i, k in enumerate(("deleted", "gone", "creating", "stuck", "refused", "closed", "theirs"))}
    with server._MIND_LOOP_LOCK:
        server._mind_projects_write({
            names["deleted"]: entry(), names["gone"]: entry(), names["creating"]: entry(status="creating"), names["stuck"]: entry(status="run", ran=True),
            names["refused"]: entry(), names["closed"]: entry(), names["theirs"]: entry(session="another-workbook"), "somebody_elses_project": entry(),
        })  # the last one: a name without the sandbox prefix, which the app never writes down -- refused all the same
    answers = {names["deleted"]: "deleted", names["gone"]: "not found", names["creating"]: "not found", names["stuck"]: "still present", names["refused"]: "refused"}
    asked: list[tuple[str, str]] = []
    opens = {"manager": True}

    def delete_project(page, name, folder_=None, attempts=2):
        asked.append((name, folder_))
        return answers[name]

    mind = FakeMind(monkeypatch, delete_project=delete_project, open_manager=lambda page: opens["manager"])

    def delete(name: str):
        return client.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name})

    def kept() -> set[str]:
        return {p["name"] for p in _mind(client, sid)["kept"]}

    # only a name the app wrote down for this workbook, and with the sandbox prefix
    for foreign in ("", "never_heard_of", names["theirs"], "somebody_elses_project"):
        assert delete(foreign).status_code == 404
    assert client.post("/api/sessions/nope/mind-loop/delete-project", json={"name": names["deleted"]}).status_code == 404
    assert asked == [] and mind.sessions["opened"] == 0
    assert kept() == {names[k] for k in ("deleted", "gone", "creating", "stuck", "refused", "closed")}
    by_hand = f"Mind did not delete it. Delete it by hand in Mind: open the Project manager, folder {folder}."
    table = [  # name, result, message, what the file keeps, does the button stay
        ("deleted", "deleted", "Deleted from Mind.", "deleted", False),
        ("gone", "not found", "It was not in Mind: nothing to delete.", "not found", False),
        ("creating", "not found", "It is not in Mind yet. Try again in a minute.", None, True),
        ("stuck", "still present", by_hand, None, True),
        ("refused", "refused", by_hand, None, True),
    ]
    for key, result, message, deleted, stays in table:
        res = delete(names[key])
        assert res.status_code == 200 and res.json() == {"name": names[key], "result": result, "message": message, "url": mind_client.MAIN_URL}, key
        assert _projects_file()[names[key]]["deleted"] == deleted and (names[key] in kept()) is stays, key
    assert asked == [(names[k], folder) for k in ("deleted", "gone", "creating", "stuck", "refused")]
    # Mind did not open: nothing is called "not found", the button stays
    opens["manager"] = False
    res = delete(names["closed"])
    assert res.json() == {"name": names["closed"], "result": "unreachable", "message": "Mind did not open. Sign in to Mind, then try again.", "url": mind_client.MAIN_URL}
    assert names["closed"] in kept() and len(asked) == 5

    def no_edge(headed=False):
        raise RuntimeError("Chromium distribution 'msedge' is not found")

    monkeypatch.setattr(mind_client, "mind_session", no_edge)
    assert delete(names["closed"]).json()["result"] == "unreachable" and server._MIND_DELETE["name"] is None
    assert mind.sessions["opened"] == mind.sessions["closed"] and _incidents() == 0


def test_while_a_project_is_being_deleted_the_other_users_of_mind_wait(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    name = f"{mind_client.PROJECT_PREFIX}_20261117_010000_i1"
    with server._MIND_LOOP_LOCK:
        server._mind_projects_write({name: {"session": sid, "folder": "MindReady_Sandbox", "status": "created", "ran": False, "made": "x", "cancelled": True, "deleted": None, "tries": 0}})
    gate = Gate()

    def delete_project(page, name_, folder=None, attempts=2):
        gate.hold()
        return "deleted"

    FakeMind(monkeypatch, delete_project=delete_project)
    deleting = _bg(lambda c: c.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name}))
    assert gate.started.wait(WAIT)
    busy = "Mind is busy. Try again in a moment."
    for res in (
        client.post(f"/api/sessions/{sid}/mind-loop", json=RUN),
        client.post("/api/mind/sign-in", json={}),
        client.get("/api/mind/sign-in", params={"check": True}),
        client.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name}),
    ):
        assert res.status_code == 409 and res.json()["detail"] == busy
    assert sid not in server.MIND_LOOPS and server.MIND_SIGN_IN["state"] == "idle"
    gate.go.set()
    assert _answer(deleting).json()["result"] == "deleted" and server._MIND_DELETE["name"] is None
    # and the other way round: a delete waits while the sign-in window is open
    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "waiting")
    res = client.post(f"/api/sessions/{sid}/mind-loop/delete-project", json={"name": name})
    assert res.status_code == 409 and res.json()["detail"] == busy
    # Cancel closes a waiting sign-in window (Upload to Mind was stopped before anything reached Mind)
    monkeypatch.setitem(server.MIND_SIGN_IN, "stop", False)
    out = _cancel(client, sid, "mind")
    assert server.MIND_SIGN_IN["stop"] is True and out["cancelled"] == [] and out["message"] == "Stopped. Nothing was uploaded to Mind."
    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    _clean(client, sid)


def test_remove_and_the_bring_in_are_refused_while_a_run_is_cancelling(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    monkeypatch.setattr(server, "com_available", lambda: True)
    monkeypatch.setitem(server.MIND_LOOPS, sid, {"state": "cancelling", "started": time.time(), "events": [], "work_dir": str(s.work_dir), "projects": []})
    assert server._loop_busy(sid) and server._loop_running() == sid and server._mind_in_use() == "loop"
    assert client.delete(f"/api/saved-workbooks/{sid}").status_code == 409 and sid in server.SESSIONS
    with pytest.raises(server.HTTPException) as refused:
        server._free_for_bring_in(s, sid)
    assert refused.value.status_code == 409
    server.MIND_LOOPS[sid]["state"] = "cancelled"
    assert not server._loop_busy(sid) and server._mind_in_use() is None
    server._free_for_bring_in(s, sid)  # free again


# --- 17 Nov 2026, after the review ------------------------------------------------------------------
BUSY_APPLY = "Wait a moment: changes are still being written."


def _slow_model_map(monkeypatch) -> Gate:
    """The model map of a new version is built slowly (38 s on Palermo): it waits at the gate."""
    gate = Gate()

    def build_model_map(analysis):
        gate.started.set()
        assert gate.go.wait(WAIT), "the model map was never let go"
        return {"inputs": [], "loops": [], "outputs": []}

    monkeypatch.setattr(server, "build_model_map", build_model_map)
    return gate


def test_a_cancel_while_the_loop_plan_is_built_stops_the_loop_step(client, plain_grid_xlsx, monkeypatch):
    # R1-1: the job is registered before the plan is built, and the build of the model map hears the Cancel
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    calls: list = []
    monkeypatch.setattr(server, "apply_operations", _fake_apply(calls))
    monkeypatch.setattr(server, "build_loop_plan", lambda *a, **k: {"operations": [dict(o, action_id="loop", rule_id="LOOP") for o in OPS], "notes": []})
    gate = _slow_model_map(monkeypatch)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/mind-loop-plan/apply", json={}))
    try:
        assert gate.started.wait(WAIT)
        assert [j["kind"] for j in _jobs(client, sid)] == ["apply"]  # the job is there while the plan is built
        out = _cancel(client, sid, "apply")
        assert out["cancelled"] == ["apply"] and out["message"] == STOPPED  # not "Nothing was running."
        _is_stopped(_answer(pending), "apply")  # at once: the model map is still being built
        assert calls == [] and _look(s) == before
    finally:
        gate.go.set()
    _clean(client, sid)


def test_a_cancel_while_the_instances_plan_is_built_stops_the_instances_step(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    gate = _slow_model_map(monkeypatch)
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/mind-instances-plan/apply", json={}))
    try:
        assert gate.started.wait(WAIT)
        assert [j["kind"] for j in _jobs(client, sid)] == ["apply"]
        assert _cancel(client, sid, "apply")["cancelled"] == ["apply"]
        _is_stopped(_answer(pending), "apply")
        assert _look(s) == before
    finally:
        gate.go.set()
    _clean(client, sid)


def test_a_loop_or_instances_plan_with_nothing_to_write_still_answers_422_and_leaves_no_job(client, plain_grid_xlsx, monkeypatch):
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    monkeypatch.setattr(server, "_loop_plan_of", lambda s, payload: {"operations": [], "notes": ["no loop in the macros"]})
    res = client.post(f"/api/sessions/{sid}/mind-loop-plan/apply", json={})
    assert res.status_code == 422 and res.json()["detail"] == "nothing to write: no loop in the macros"
    monkeypatch.setattr(server, "_instance_plan_of", lambda s, *a: {"operations": [], "notes": []})
    res = client.post(f"/api/sessions/{sid}/mind-instances-plan/apply", json={})
    assert res.status_code == 422 and res.json()["detail"] == "nothing to write: no instances could be built"
    _clean(client, sid)


def _hold_an_apply(client, sid, monkeypatch):
    gate = Gate()
    monkeypatch.setattr(server, "apply_operations", _fake_apply([], gate))
    pending = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": OPS, "reanalyze": True}))
    assert gate.started.wait(WAIT)
    return gate, pending


def test_what_a_cancelled_apply_would_take_back_is_refused_while_it_runs(client, plain_grid_xlsx, monkeypatch):
    # R1-2: the restore of a cancelled Apply puts back the recalculation, the input files, the names, the
    # files and the plan: written meanwhile, each was lost after the person was told it was done
    import io

    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    recalcs: list = []
    monkeypatch.setattr(server, "recalculate", lambda path, save=True: recalcs.append(path) or {"status": "OK", "message": "recalculated clean", "ran": True, "formula_errors": [], "addin_gap_errors": []})
    monkeypatch.setattr(server, "suggest_names", lambda contexts, **kw: {"available": True, "names": {"CF!B8:C8": "Cash flows"}, "message": None})
    before = _look(s)
    gate, applying = _hold_an_apply(client, sid, monkeypatch)
    try:
        refused = {
            "recalculate": client.post(f"/api/sessions/{sid}/recalculate"),
            "inputs": client.post(f"/api/sessions/{sid}/inputs", files=[("files", ("rates_2026Q3.csv", io.BytesIO(b"a,b\n1,2\n"), "text/csv"))]),
            "grid-names": client.post(f"/api/sessions/{sid}/grid-names", json={"apply": True}),
            "reports": client.post(f"/api/sessions/{sid}/reports"),
            "pivot-mode": client.post(f"/api/sessions/{sid}/pivot-mode", json={"mode": "values"}),
        }
        assert {k: (r.status_code, r.json()["detail"]) for k, r in refused.items()} == {k: (409, BUSY_APPLY) for k in refused}
        assert recalcs == [] and s.inputs == [] and s.grid_names == {}
        out = _cancel(client, sid, "apply")
    finally:
        gate.go.set()
    _is_stopped(_answer(applying), "apply")
    assert out["cancelled"] == ["apply"] and _look(s) == before
    _clean(client, sid)


def test_names_whose_answer_comes_after_an_apply_began_are_not_written(client, section_heading_grids_xlsx, monkeypatch):
    # R1-2, the other order: Suggest names waits for the model, an Apply begins meanwhile. Its Cancel
    # would take the names back, so they are not written (409); the Apply is never refused for them.
    sid = _upload(client, section_heading_grids_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    plan = id(s.plan)
    asked, answer = Gate(), threading.Event()

    def suggest(contexts, **kw):
        asked.started.set()
        assert answer.wait(WAIT)
        return {"available": True, "names": {"CF!B8:C8": "Cash flows"}, "message": None}

    monkeypatch.setattr(server, "suggest_names", suggest)
    naming = _bg(lambda c: c.post(f"/api/sessions/{sid}/grid-names", json={"apply": True}))
    assert asked.started.wait(WAIT)
    gate, applying = _hold_an_apply(client, sid, monkeypatch)
    try:
        answer.set()
        res = _answer(naming)
        assert (res.status_code, res.json()["detail"]) == (409, BUSY_APPLY)
        assert s.grid_names == {} and id(s.plan) == plan
    finally:
        gate.go.set()
    assert _answer(applying).status_code == 200
    _clean(client, sid)


def test_the_cancel_beside_start_sends_its_kinds_and_leaves_a_run_in_mind_and_an_answer_alone(client, plain_grid_xlsx, monkeypatch):
    # R1-5: the Cancel beside Start (and Do everything) sends the kinds of its own steps
    sid = _upload(client, plain_grid_xlsx)["sessionId"]
    mind = cancel.begin(sid, "mind")
    chat = cancel.begin(sid, "chat")
    gate, applying = _hold_an_apply(client, sid, monkeypatch)
    try:
        res = client.post(f"/api/sessions/{sid}/cancel", json={"kinds": ["outside", "autofix", "prep", "apply", "recalc", "scan"]})
        assert res.status_code == 200, res.text
        out = res.json()
        assert out["cancelled"] == ["apply"] and out["message"] == STOPPED
        assert not mind.flag.is_set() and not chat.flag.is_set()  # they go on
        assert sorted(j["kind"] for j in _jobs(client, sid)) == ["chat", "mind"]
    finally:
        gate.go.set()
        cancel.end(mind)
        cancel.end(chat)
    _is_stopped(_answer(applying), "apply")
    assert client.post(f"/api/sessions/{sid}/cancel", json={"kinds": "apply"}).status_code == 200  # one kind as text: a list of one
    assert client.post(f"/api/sessions/{sid}/cancel", json={"kinds": {"apply": 1}}).status_code == 422
    _clean(client, sid)
