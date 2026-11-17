"""1.11.9: one Excel per session (app/excel_worker.py). The worker must give exactly what the old path
gave -- same versions, same labels, same cells -- while the check that a version opens is the open of
the next step's copy, and fall back to the old path when its Excel dies."""
import json
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

from app import excel_worker
from app.excel_com import com_available, kill_excel
from app.web import server

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")


def _upload(client, path: Path) -> dict:
    with path.open("rb") as f:
        res = client.post("/api/sessions", files={"file": (path.name, f, "application/octet-stream")}, data={"mode": "plan"})
    assert res.status_code == 200, res.text
    return res.json()


def _cells(path: Path) -> dict:
    out = {}
    for data_only in (False, True):
        wb = openpyxl.load_workbook(path, data_only=data_only)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        out[(data_only, ws.title, c.coordinate)] = repr(c.value)
        wb.close()
    return out


def _chain(client, path: Path) -> dict:
    """Upload, the plan's ops in two applies (the second on the first's version), Recalculate."""
    body = _upload(client, path)
    sid = body["sessionId"]
    plan = {a["id"]: a for a in body["plan"]}
    first = plan["flag_spelling"]["operations"] + plan["special_headers"]["operations"]
    r1 = client.post(f"/api/sessions/{sid}/apply", json={"operations": first, "reanalyze": True})
    assert r1.status_code == 200, r1.text
    r2 = client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["loop_name_case"]["operations"], "reanalyze": True})
    assert r2.status_code == 200, r2.text
    rc = client.post(f"/api/sessions/{sid}/recalculate").json()
    s = server.SESSIONS[sid]
    versions = client.get(f"/api/sessions/{sid}/versions").json()
    return {
        "sid": sid,
        "labels": [(v["id"], v["source"], v["label"].split(" — ", 1)[-1], v.get("verified_opens_in_excel")) for v in versions],
        "applied": [len(r1.json()["result"]["applied"]), len(r2.json()["result"]["applied"])],
        "recalc": (rc.get("status"), len(rc.get("formula_errors") or [])),
        "cells": _cells(s.current_path),
        "paths": [str(Path(p).relative_to(s.work_dir)) for p in (s.paths[v["id"]] for v in versions[1:])],
    }


@needs_excel
def test_the_worker_gives_the_same_versions_and_cells(flagged_model_broken_xlsx, monkeypatch, tmp_path):
    runs = {}
    for flag in ("0", "1"):
        monkeypatch.setenv("MINDREADY_EXCEL_WORKER", flag)
        monkeypatch.setenv("MINDREADY_PERF_LOG", str(tmp_path / f"perf_{flag}.jsonl"))
        runs[flag] = _chain(TestClient(server.app), flagged_model_broken_xlsx)
        excel_worker.stop_all_workers()
    off, on = runs["0"], runs["1"]
    assert on["labels"] == off["labels"] and all(v[3] for v in on["labels"][1:])  # every version still verified
    assert on["applied"] == off["applied"] and on["recalc"] == off["recalc"]
    assert on["paths"] == off["paths"]  # the copies sit where they did
    assert on["cells"] == off["cells"]
    # with the worker, the checks are opens on the worker's thread, and the second apply found its copy open
    perf = [json.loads(line) for line in (tmp_path / "perf_1.jsonl").read_text(encoding="utf-8").splitlines()]
    checks = [r for r in perf if r["stage"] == "verify_opens"]
    assert checks and all(r["thread"].startswith("excel-") for r in checks)
    opens_on = sum(1 for r in perf if r["stage"] == "excel_open")
    opens_off = sum(1 for line in (tmp_path / "perf_0.jsonl").read_text(encoding="utf-8").splitlines() if json.loads(line)["stage"] == "excel_open")
    assert opens_on < opens_off


@needs_excel
def test_a_dead_excel_falls_back_to_the_old_path(flagged_model_broken_xlsx, monkeypatch, tmp_path):
    monkeypatch.setenv("MINDREADY_EXCEL_WORKER", "1")
    monkeypatch.setenv("MINDREADY_PERF_LOG", str(tmp_path / "perf.jsonl"))
    client = TestClient(server.app)
    body = _upload(client, flagged_model_broken_xlsx)
    sid = body["sessionId"]
    plan = {a["id"]: a for a in body["plan"]}
    r1 = client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["flag_spelling"]["operations"], "reanalyze": True})
    assert r1.status_code == 200 and r1.json()["version"]["verified_opens_in_excel"] is True
    w = excel_worker.WORKERS[sid]
    assert w._held is not None  # the next copy is open, waiting
    kill_excel(w._pid)
    r2 = client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["special_headers"]["operations"], "reanalyze": True})
    assert r2.status_code == 200, r2.text
    assert r2.json()["result"]["status"] == "APPLIED" and r2.json()["version"]["id"] == "ver-003"
    perf = [json.loads(line) for line in (tmp_path / "perf.jsonl").read_text(encoding="utf-8").splitlines()]
    assert any(r["stage"] == "excel_worker_fallback" for r in perf)


@needs_excel
def test_recalculate_after_a_loop_version_is_made_ahead_and_matches(flagged_model_broken_xlsx, monkeypatch, tmp_path):
    runs = {}
    for flag in ("0", "1"):
        monkeypatch.setenv("MINDREADY_EXCEL_WORKER", flag)
        monkeypatch.setenv("MINDREADY_PERF_LOG", str(tmp_path / f"perf_{flag}.jsonl"))
        client = TestClient(server.app)
        body = _upload(client, flagged_model_broken_xlsx)
        sid = body["sessionId"]
        plan = {a["id"]: a for a in body["plan"]}
        s = server.SESSIONS[sid]
        out = server._apply_and_version(s, plan["flag_spelling"]["operations"], "loop", "Mind loop", True)  # a loop version
        rc = client.post(f"/api/sessions/{sid}/recalculate").json()
        runs[flag] = {
            "label": out["version"]["label"].split(" — ", 1)[-1],
            "recalc": (rc.get("status"), rc.get("formula_errors"), rc.get("mmforexcel_loaded")),
            "recalc_path": str(Path(s.recalc_path).relative_to(s.work_dir)),
            "cells": _cells(Path(s.recalc_path)),
        }
        excel_worker.stop_all_workers()
    assert runs["1"] == runs["0"]
    perf = [json.loads(line) for line in (tmp_path / "perf_1.jsonl").read_text(encoding="utf-8").splitlines()]
    assert any(r["stage"] == "recalc_ahead_used" for r in perf)


def test_a_job_that_waits_too_long_is_cancelled_and_never_runs(tmp_path):
    import threading
    import time as _time

    w = excel_worker.ExcelWorker("t-queue", tmp_path / "apply")
    try:
        ran = []
        threading.Thread(target=lambda: w._submit(lambda: _time.sleep(2.0)), daemon=True).start()
        _time.sleep(0.2)
        with pytest.raises(excel_worker.WorkerUnavailable):
            w._submit(lambda: ran.append(1), queue_timeout=0.3)
        _time.sleep(2.5)
        assert ran == [] and w._submit(lambda: 7) == 7  # cancelled, never ran; the worker goes on
    finally:
        w.stop()


def test_a_job_that_hangs_retires_the_worker(tmp_path, monkeypatch):
    import time as _time

    monkeypatch.setenv("MINDREADY_EXCEL_WORKER", "1")
    w = excel_worker.worker_for("t-hang", tmp_path / "apply")
    if w is None:
        pytest.skip("no COM here")
    with pytest.raises(excel_worker.WorkerUnavailable):
        w._submit(lambda: _time.sleep(1.5), timeout=0.3)
    assert w._stopped and excel_worker.WORKERS.get("t-hang") is not w
    fresh = excel_worker.worker_for("t-hang", tmp_path / "apply")
    assert fresh is not w and fresh._submit(lambda: 3) == 3


@needs_excel
def test_a_failed_warm_apply_falls_back_at_the_same_path(flagged_model_broken_xlsx, monkeypatch, tmp_path):
    from app import prep

    monkeypatch.setenv("MINDREADY_EXCEL_WORKER", "1")
    monkeypatch.setenv("MINDREADY_PERF_LOG", str(tmp_path / "perf.jsonl"))
    client = TestClient(server.app)
    body = _upload(client, flagged_model_broken_xlsx)
    sid = body["sessionId"]
    plan = {a["id"]: a for a in body["plan"]}
    assert client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["flag_spelling"]["operations"], "reanalyze": True}).status_code == 200
    real = prep._apply_with_excel

    def refuse_in_the_open_copy(copy_path, operations, progress=None, excel=None, wb=None):
        if wb is not None:
            raise prep.NotSavedError("refused in the open copy")  # not "Excel gone": a plain failure
        return real(copy_path, operations, progress, excel=excel, wb=wb) if excel is not None else real(copy_path, operations, progress)

    monkeypatch.setattr(prep, "_apply_with_excel", refuse_in_the_open_copy)
    r2 = client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["special_headers"]["operations"], "reanalyze": True})
    assert r2.status_code == 200, r2.text
    assert r2.json()["result"]["status"] == "APPLIED"
    s = server.SESSIONS[sid]
    name = Path(s.paths["ver-002"]).name
    assert Path(s.paths["ver-003"]).relative_to(s.work_dir).as_posix() == f"apply/v2/{name}"
    perf = [json.loads(line) for line in (tmp_path / "perf.jsonl").read_text(encoding="utf-8").splitlines()]
    assert any(r["stage"] == "excel_worker_fallback" for r in perf)


@needs_excel
def test_an_unused_recalculation_ahead_leaves_no_copy_behind(flagged_model_broken_xlsx, monkeypatch, tmp_path):
    monkeypatch.setenv("MINDREADY_EXCEL_WORKER", "1")
    client = TestClient(server.app)
    body = _upload(client, flagged_model_broken_xlsx)
    sid = body["sessionId"]
    plan = {a["id"]: a for a in body["plan"]}
    s = server.SESSIONS[sid]
    server._apply_and_version(s, plan["flag_spelling"]["operations"], "loop", "Mind loop", True)  # recalculated ahead
    assert client.post(f"/api/sessions/{sid}/apply", json={"operations": plan["special_headers"]["operations"], "reanalyze": True}).status_code == 200
    excel_worker.stop_all_workers()
    leftovers = [p for p in (s.work_dir / "recalc").rglob("*") if p.is_file()] if (s.work_dir / "recalc").exists() else []
    assert leftovers == []


# --- 17 Nov 2026 (Cancel, DESIGN R4b): the worker lets go of a cancelled job's items, and only of those.
# Plain sleep jobs: no Excel is started, and every kill the worker asks for is a recorded fake.
class _Caller:
    """One thread that asks the worker for `fn` (as a route or a job's thread does) and waits."""

    def __init__(self, w, fn, job=None):
        import threading

        self.value = self.error = None
        self.thread = threading.Thread(target=self._run, args=(w, fn, job), daemon=True)
        self.thread.start()

    def _run(self, w, fn, job) -> None:
        from app import cancel

        try:
            with cancel.bind(job):
                self.value = w._submit(fn)
        except BaseException as exc:  # noqa: BLE001 -- a Cancelled is what some of these tests expect
            self.error = exc

    def back_within(self, seconds: float) -> bool:
        self.thread.join(seconds)
        return not self.thread.is_alive()


@pytest.fixture
def lone_worker(tmp_path, monkeypatch):
    """(worker, kills): a worker with no Excel behind it that excel_worker.cancel finds, and the pids it asked to end."""
    from app import cancel, excel_com

    kills: list = []
    monkeypatch.setattr(excel_worker, "kill_excel", lambda pid: bool(kills.append(pid)))
    monkeypatch.setattr(excel_com, "kill_excel", lambda pid: bool(kills.append(pid)))
    w = excel_worker.ExcelWorker("t-cancel", tmp_path / "apply")
    monkeypatch.setitem(excel_worker.WORKERS, "t-cancel", w)
    try:
        yield w, kills
    finally:
        w.stop()
        for job in cancel.jobs("t-cancel"):
            cancel.end(job)


def _slow(running, seconds: float, result=None, checks: bool = False):
    """A job that takes `seconds`. `checks`: it reads its job's flag as the progress closures of a real
    apply do (a real one also fails when its Excel is ended; here there is no Excel)."""
    import time as _time

    from app import cancel

    def job():
        running.set()
        end = _time.time() + seconds
        while _time.time() < end:
            _time.sleep(0.05)
            if checks:
                cancel.check()
        return result

    return job


def test_an_item_runs_for_the_job_of_the_thread_that_queued_it(lone_worker):
    import concurrent.futures
    import threading
    import time as _time

    from app import cancel

    w, _kills = lone_worker
    job = cancel.begin("t-cancel", "apply", changes=True)
    with cancel.bind(job):
        assert w._submit(cancel.current) is job  # the worker thread is bound to the item's owner
    assert w._submit(cancel.current) is None  # ... and to nobody for the next item
    seen: list = []
    w._q.put((lambda: seen.append(cancel.current()), concurrent.futures.Future(), threading.Event(), None))  # as preopen_async
    w._submit(lambda: None)
    for _ in range(200):  # the worker thread notes "nothing runs" just after it hands the answer over
        if not w._busy:
            break
        _time.sleep(0.01)
    assert seen == [None] and (w._busy, w._owner) == (False, None)
    # a check inside an item reads its owner's flag, and the caller gets the Cancelled
    gate, running = threading.Event(), threading.Event()
    blocker = _Caller(w, lambda: (running.set(), gate.wait(10)))  # runs first, so that the next item waits
    assert running.wait(5)
    caller = _Caller(w, cancel.check, job)
    _time.sleep(0.3)
    job.flag.set()  # the flag alone (a full Cancel would let the waiting item go before it ran)
    gate.set()
    assert blocker.back_within(3.0) and caller.back_within(3.0) and isinstance(caller.error, cancel.Cancelled)
    assert w._submit(lambda: 5) == 5  # a Cancelled raised by an item does not break the worker
    ran: list = []
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        w._submit(lambda: ran.append(1))  # a cancelled job queues nothing more
    assert ran == [] and w._q.empty()


def test_cancelling_the_running_job_returns_the_caller_that_waits_behind_it(lone_worker):
    import threading
    import time as _time

    from app import cancel

    w, kills = lone_worker
    apply_job, numbers_job = cancel.begin("t-cancel", "apply", changes=True), cancel.begin("t-cancel", "numbers")
    running, ran = threading.Event(), []
    first = _Caller(w, _slow(running, 5.0, "finished", checks=True), apply_job)
    assert running.wait(5)
    w._pid = 4242  # the Excel the running item works in
    second = _Caller(w, lambda: ran.append(1), numbers_job)
    _time.sleep(0.3)
    assert not second.back_within(0)  # a second caller waits for its turn
    out = cancel.request("t-cancel", "apply")
    assert out["flagged"] == [apply_job] and out["ahead"] is False
    assert second.back_within(1.0)  # ... and returns within a second (before: after QUEUE_TIMEOUT_S, half an hour)
    assert isinstance(second.error, excel_worker.WorkerUnavailable)  # its own job goes on: it falls back to an Excel of its own
    assert kills == [4242] and w._stopped and excel_worker.WORKERS.get("t-cancel") is not w  # the two steps of the time limit
    assert first.back_within(2.0) and isinstance(first.error, cancel.Cancelled)
    assert ran == []  # what waited never ran


def test_cancelling_only_the_waiting_job_leaves_the_running_job_and_its_excel_alone(lone_worker):
    import threading
    import time as _time

    from app import cancel

    w, kills = lone_worker
    apply_job, numbers_job = cancel.begin("t-cancel", "apply", changes=True), cancel.begin("t-cancel", "numbers")
    running, ran = threading.Event(), []
    first = _Caller(w, _slow(running, 1.2, 7, checks=True), apply_job)
    assert running.wait(5)
    w._pid = 4242
    second = _Caller(w, lambda: ran.append(1), numbers_job)  # e.g. the release before the numbers check
    _time.sleep(0.3)
    out = cancel.request("t-cancel", "numbers")
    assert out["flagged"] == [numbers_job] and out["ahead"] is False
    assert second.back_within(1.0) and isinstance(second.error, cancel.Cancelled)  # cancelled: no fallback for it
    assert kills == [] and not w._stopped and excel_worker.WORKERS.get("t-cancel") is w  # the Apply's Excel is not ended
    assert first.back_within(3.0) and first.error is None and first.value == 7  # ... and its result arrives
    assert ran == [] and w._submit(lambda: 3) == 3  # the worker goes on


def test_a_worker_cancelled_while_its_excel_starts_ends_that_excel(lone_worker, monkeypatch):
    import threading

    from app import cancel

    w, kills = lone_worker
    starting, go = threading.Event(), threading.Event()

    def slow_start():
        starting.set()
        go.wait(10)
        return object()

    monkeypatch.setattr(excel_worker, "start_excel", slow_start)
    monkeypatch.setattr(excel_worker, "excel_pid", lambda excel: 777)
    job = cancel.begin("t-cancel", "apply", changes=True)
    caller = _Caller(w, w._ensure_excel, job)
    assert starting.wait(5)
    cancel.request("t-cancel", "apply")  # retires the worker; its Excel has no pid yet: nothing to end
    assert w._stopped and 777 not in kills
    go.set()
    assert caller.back_within(3.0) and isinstance(caller.error, cancel.Cancelled)
    assert kills.count(777) == 1 and w._excel is None and w._pid is None  # ended as soon as it had started


def test_a_cancel_of_everything_stops_the_recalculation_made_ahead_and_another_kind_does_not(lone_worker):
    import concurrent.futures
    import threading

    from app import cancel

    w, kills = lone_worker
    running = threading.Event()
    w._q.put((_slow(running, 1.5), concurrent.futures.Future(), threading.Event(), None))  # as preopen queues it: no owner
    assert running.wait(5)
    w._pid = 4242
    assert excel_worker.cancel("t-cancel", []) == 0  # free=False
    assert cancel.request("t-cancel", "apply") == {"flagged": [], "late": [], "ahead": False}
    assert kills == [] and not w._stopped  # it belongs to no job: an Apply's Cancel does not touch it
    assert cancel.request("t-cancel") == {"flagged": [], "late": [], "ahead": True}  # free=True
    assert kills == [4242] and w._stopped and excel_worker.WORKERS.get("t-cancel") is not w
    assert excel_worker.cancel("no-such-workbook", [], free=True) == 0


def test_cancelling_a_recalculate_that_waits_stops_what_was_made_ahead_and_not_the_running_apply(lone_worker):
    import concurrent.futures
    import threading
    import time as _time

    from app import cancel

    w, kills = lone_worker
    apply_job, recalc_job = cancel.begin("t-cancel", "apply", changes=True), cancel.begin("t-cancel", "recalc")
    running, ran = threading.Event(), []
    first = _Caller(w, _slow(running, 1.2, 7), apply_job)
    assert running.wait(5)
    w._pid = 4242
    ahead: concurrent.futures.Future = concurrent.futures.Future()
    w._q.put((lambda: ran.append("ahead"), ahead, threading.Event(), None))
    taker = _Caller(w, lambda: ran.append("take"), recalc_job)  # take_recalc, waiting for its turn
    _time.sleep(0.3)
    out = cancel.request("t-cancel", "recalc")
    assert out["flagged"] == [recalc_job] and out["ahead"] is True
    assert taker.back_within(1.0) and isinstance(taker.error, cancel.Cancelled)
    assert ahead.cancelled()
    assert kills == [] and not w._stopped  # the Apply that runs belongs to a job nobody cancelled
    assert first.back_within(3.0) and first.value == 7 and ran == []
    # an item whose job has ended (the check of a version, still on its way) is nobody's: it goes with "everything"
    done_job = cancel.begin("t-cancel", "scan")
    late_running = threading.Event()
    late = _Caller(w, _slow(late_running, 1.5), done_job)
    assert late_running.wait(5)
    cancel.end(done_job)
    assert excel_worker.cancel("t-cancel", [], free=False) == 0 and not w._stopped
    assert excel_worker.cancel("t-cancel", [], free=True) == 1 and w._stopped and kills == [4242]
    assert late.back_within(3.0) and isinstance(late.error, excel_worker.WorkerUnavailable)  # its Excel was ended under it


def test_an_answer_made_after_its_excel_was_ended_is_not_an_answer(lone_worker):
    import threading

    from app import cancel

    # the open check swallows the error of an Excel that was ended and answers "this version does not
    # open in Excel": after a Cancel that must reach nobody as a fact
    w, kills = lone_worker
    job = cancel.begin("t-cancel", "apply", changes=True)
    running = threading.Event()
    caller = _Caller(w, _slow(running, 0.8, {"opens": False}), job)  # it never looks at the flag
    assert running.wait(5)
    w._pid = 4242
    assert cancel.request("t-cancel", "apply")["flagged"] == [job] and kills == [4242]
    assert caller.back_within(3.0) and caller.value is None and isinstance(caller.error, cancel.Cancelled)


def test_a_retired_or_stopped_worker_lets_its_waiting_callers_go_at_once(tmp_path):
    import threading
    import time as _time

    for how in ("retire", "stop"):
        w = excel_worker.ExcelWorker(f"t-{how}", tmp_path / "apply")
        try:
            running, ran = threading.Event(), []
            first = _Caller(w, _slow(running, 1.2, "finished"))
            assert running.wait(5)
            second = _Caller(w, lambda: ran.append(1))
            _time.sleep(0.3)
            if how == "retire":
                w._retire()
            else:
                threading.Thread(target=w.stop, daemon=True).start()  # stop() itself waits for the running item
            assert second.back_within(0.8), how  # before 17 Nov 2026: blocked for QUEUE_TIMEOUT_S
            assert isinstance(second.error, excel_worker.WorkerUnavailable) and ran == []
            assert first.back_within(3.0)
            if how == "stop":
                assert first.value == "finished"  # a plain stop lets the running item finish and answer
            else:  # retired: its Excel is ended next, so what it answers after that is no answer
                assert first.value is None and isinstance(first.error, excel_worker.WorkerUnavailable)
        finally:
            w.stop()
