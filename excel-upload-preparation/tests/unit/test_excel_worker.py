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
