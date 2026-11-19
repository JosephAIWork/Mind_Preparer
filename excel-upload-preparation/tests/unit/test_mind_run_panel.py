"""1.12.1 (the Mind team's notes): Run in Mind says which version goes up (the latest), lets the user
tick the changes made before uploading, waits for anything still changing the workbook, and reports
its stage, time spent and a conservative time left."""
import time

import pytest
from fastapi.testclient import TestClient

from app.estimate import MIND_SAFETY, estimate_mind_run, mind_run_remaining
from app.web import server


@pytest.fixture(scope="module")
def client():
    return TestClient(server.app)


def _upload(client, path):
    with path.open("rb") as f:
        res = client.post("/api/sessions", files={"file": (path.name, f, "application/octet-stream")}, data={"mode": "plan"})
    assert res.status_code == 200, res.text
    return res.json()["sessionId"]


def test_the_estimate_is_on_the_long_side_of_the_measured_run():
    # the one measured run: a 13 MB model, 6 min 10 s end to end with Mind (scan about 60 s)
    r = estimate_mind_run(13.0, 60.0, 0, 200)
    assert r["first_s"] >= 370 * 1.2
    ids = [s["id"] for s in r["stages"]]
    assert ids[:2] == ["analyse", "names"] and "upload" in ids and "run" in ids
    assert "cleanup" not in ids  # a project that was run is kept
    assert r["retry_s"] == r["first_s"] - next(s["seconds"] for s in r["stages"] if s["id"] == "names")
    assert all(s["seconds"] > 0 for s in r["stages"])
    no_run = estimate_mind_run(13.0, 60.0, 0, 200, use_assistant=False, run_model=False)
    assert no_run["first_s"] < r["first_s"] and "cleanup" in [s["id"] for s in no_run["stages"]]
    assert MIND_SAFETY >= 1.5


def test_time_left_never_drops_to_zero_while_a_stage_runs_past_its_estimate():
    stages = [{"id": "a", "seconds": 100}, {"id": "b", "seconds": 50}]
    assert mind_run_remaining(stages, None, 0) == 150
    assert mind_run_remaining(stages, "a", 40) == 60 + 50
    assert mind_run_remaining(stages, "a", 500) == 10 + 50  # a tenth of the stage is always left
    assert mind_run_remaining(stages, "b", 0) == 50
    assert mind_run_remaining(stages, "cleanup", 0) == 0


def test_options_name_the_latest_version_and_the_changes_to_tick(client, section_heading_grids_xlsx):
    sid = _upload(client, section_heading_grids_xlsx)
    o = client.get(f"/api/sessions/{sid}/mind-loop/options").json()
    s = server.SESSIONS[sid]
    assert o["version"]["id"] == s.current_version_id == o["newest_version_id"]
    assert o["busy"] is None
    assert {c["id"] for c in o["choices"]} == {a["id"] for a in s.plan if a["count"] > 0}
    assert o["estimate"]["first_s"] > 0 and o["estimate"]["caveat"]


def test_run_in_mind_waits_while_something_still_changes_the_workbook(client, section_heading_grids_xlsx):
    sid = _upload(client, section_heading_grids_xlsx)
    server.AUTOFIX[sid] = {"state": "running"}
    try:
        assert "Fix all issues" in client.get(f"/api/sessions/{sid}/mind-loop/options").json()["busy"]
        res = client.post(f"/api/sessions/{sid}/mind-loop", json={"skipMind": True})
        assert res.status_code == 409 and "latest version" in res.json()["detail"]
    finally:
        server.AUTOFIX.pop(sid, None)


def test_nothing_ticked_uploads_the_version_as_it_is_and_the_status_carries_the_timing(client, section_heading_grids_xlsx):
    sid = _upload(client, section_heading_grids_xlsx)
    s = server.SESSIONS[sid]
    plan_ids = {a["id"] for a in s.plan}
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={"actions": [], "skipMind": True, "useAssistant": False, "checkNumbers": False, "maxIterations": 1})
    assert res.status_code == 200, res.text
    assert res.json()["version"]["id"] == s.current_version_id
    st = client.get(f"/api/sessions/{sid}/mind-loop").json()
    assert st["timing"]["stages"] and st["timing"]["elapsed_s"] >= 0
    for _ in range(300):
        st = client.get(f"/api/sessions/{sid}/mind-loop").json()
        if st["state"] != "running":
            break
        time.sleep(0.5)
    assert st["state"] == "done", st.get("error")
    it = st["report"]["iterations"][0]
    assert it["actions"] == []  # nothing ticked: nothing changed before the upload
    assert set(it["disabled"]) == plan_ids
    assert st["actions"] == [] and st["version"]["id"] == s.current_version_id
    t = st["timing"]
    assert t["remaining_s"] == 0 and not t["over_estimate"]
    assert all(x["state"] == "done" for x in t["stages"])
    assert any(e.get("stage") == "analyse" for e in st["events"]) and any(e.get("stage") == "prepare" for e in st["events"])


def test_ticked_changes_are_the_only_ones_made(client, section_heading_grids_xlsx):
    sid = _upload(client, section_heading_grids_xlsx)
    s = server.SESSIONS[sid]
    doable = [a["id"] for a in s.plan if a["count"] > 0]
    if not doable:
        pytest.skip("the fixture plans no change")
    pick = doable[:1]
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={"actions": pick, "skipMind": True, "useAssistant": False, "checkNumbers": False, "maxIterations": 1})
    assert res.status_code == 200, res.text
    for _ in range(600):
        st = client.get(f"/api/sessions/{sid}/mind-loop").json()
        if st["state"] != "running":
            break
        time.sleep(0.5)
    assert st["state"] == "done", st.get("error")
    it = st["report"]["iterations"][0]
    assert [a["id"] for a in it["actions"]] == pick


def _wait_sign_in(client, t=50):
    for _ in range(t):
        st = client.get("/api/mind/sign-in").json()
        if st["state"] not in ("waiting", "checking"):
            return st
        time.sleep(0.1)
    raise AssertionError("the sign-in did not finish")


def test_sign_in_opens_the_window_in_the_background_and_reports_back(client, monkeypatch):
    from app import mind_client as mc

    calls = []

    def fake_sign_in(timeout_s, progress, should_stop):
        calls.append(timeout_s)
        progress("An Edge window is open on Mind")
        time.sleep(0.3)
        return {"ok": True, "reason": None, "url": mc.MAIN_URL, "last_login": "2026-11-05T12:00:00Z"}

    monkeypatch.setattr(mc, "sign_in", fake_sign_in)
    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    st = client.post("/api/mind/sign-in", json={}).json()
    assert st["state"] == "waiting"
    assert client.post("/api/mind/sign-in", json={}).status_code == 409  # one window at a time
    st = _wait_sign_in(client)
    assert st["state"] == "done" and st["checked"]["signed_in"] is True and calls == [600]


def test_a_failed_or_impossible_sign_in_says_why(client, monkeypatch):
    from app import mind_client as mc

    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    monkeypatch.setattr(mc, "sign_in", lambda **_k: {"ok": False, "reason": "the Edge window was closed before the sign-in finished"})
    client.post("/api/mind/sign-in", json={})
    st = _wait_sign_in(client)
    assert st["state"] == "failed" and "window was closed" in st["message"]

    def no_edge(**_k):
        raise RuntimeError("Chromium distribution 'msedge' is not found")

    monkeypatch.setattr(mc, "sign_in", no_edge)
    client.post("/api/mind/sign-in", json={})
    st = _wait_sign_in(client)
    # 19 Nov 2026: a plain sentence on the screen; the raw error is kept for Send feedback only
    assert st["state"] == "failed" and st["message"] == mc.EDGE_NOT_FOUND
    assert st["detail"] == "RuntimeError: Chromium distribution 'msedge' is not found" and st["feedback_ref"]


class TargetClosedError(Exception):
    """Stands in for Playwright's error of the same name."""


_EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
_PROFILE_IN_USE = ("BrowserType.launch_persistent_context: Target page, context or browser has been closed\nBrowser logs:\n\n"
                   f"<launching> {_EDGE_EXE} --disable-field-trial-config --no-first-run")


@pytest.mark.parametrize("error, sentence", [
    (RuntimeError(f"Chromium distribution 'msedge' is not found at {_EDGE_EXE}"), "EDGE_NOT_FOUND"),
    (TargetClosedError(_PROFILE_IN_USE), "EDGE_ALREADY_OPEN"),  # the real text of a second start, 19 Nov 2026
    (OSError("[WinError 5] Access is denied"), "EDGE_DID_NOT_OPEN"),
    (ModuleNotFoundError("No module named 'playwright'"), "EDGE_DID_NOT_OPEN"),
])
def test_each_way_the_window_fails_to_open_gets_its_plain_sentence(client, monkeypatch, error, sentence):
    from app import feedback
    from app import mind_client as mc

    def fails(**_k):
        raise error

    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    monkeypatch.setattr(mc, "sign_in", fails)
    client.post("/api/mind/sign-in", json={})
    st = _wait_sign_in(client)
    assert st["state"] == "failed" and st["message"] == getattr(mc, sentence)
    # nothing of the error itself on the screen: no type name, no path, no Playwright wording
    for word in (type(error).__name__, "Error", "Traceback", "launch_persistent_context", "msedge.exe", "\\", "playwright"):
        assert word not in st["message"], word
    assert st["detail"] == f"{type(error).__name__}: {error}"[:2000]
    incident = feedback.find(st["feedback_ref"])
    assert incident and incident["step"] == "Sign in to Mind"
    # the next start begins clean
    monkeypatch.setattr(mc, "sign_in", lambda **_k: {"ok": True, "reason": None, "url": mc.MAIN_URL, "last_login": "2026-11-19T20:00:00Z"})
    client.post("/api/mind/sign-in", json={})
    st = _wait_sign_in(client)
    assert st["state"] == "done" and st["detail"] is None and st["feedback_ref"] is None


def test_a_window_the_person_closed_early_is_no_error_to_report(client, monkeypatch):
    from app import mind_client as mc

    def closed(**_k):
        raise TargetClosedError("Page.goto: Target page, context or browser has been closed")

    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    monkeypatch.setattr(mc, "sign_in", closed)
    client.post("/api/mind/sign-in", json={})
    st = _wait_sign_in(client)
    assert st["state"] == "failed" and st["message"] == "Not signed in: the Edge window was closed before the sign-in finished."
    assert st["feedback_ref"] is None and st["detail"].startswith("TargetClosedError: Page.goto")


def test_the_screen_is_told_where_the_sign_in_is_kept(client):
    from app import mind_client as mc

    st = client.get("/api/mind/sign-in").json()
    assert st["kept_in"] == str(mc.PROFILE_DIR) and st["kept_in"].endswith("edge-profile")


@pytest.mark.parametrize("front_fails", [False, True])
def test_the_window_is_asked_to_the_front_and_the_waiting_sentence_says_where_to_look(monkeypatch, front_fails):
    from contextlib import contextmanager

    from app import mind_client as mc

    seen = []

    class Page:
        url = mc.MAIN_URL

        def goto(self, url, wait_until=None):
            seen.append(("goto", url))

        def bring_to_front(self):
            seen.append(("front",))
            if front_fails:
                raise RuntimeError("Windows said no")

        def wait_for_timeout(self, _ms):
            pass

    @contextmanager
    def fake_session(headed=False):
        seen.append(("headed", headed))
        yield None, Page()

    said = []
    monkeypatch.setattr(mc, "mind_session", fake_session)
    monkeypatch.setattr(mc, "_ensure_dirs", lambda: None)
    out = mc.sign_in(timeout_s=5, progress=said.append, should_stop=lambda: True)
    assert out == {"ok": False, "reason": "stopped", "url": None}  # a failed bring_to_front never fails the sign-in
    assert seen == [("headed", True), ("goto", mc.MAIN_URL), ("front",)]
    assert said == [mc.SIGN_IN_WAITING] and "taskbar" in mc.SIGN_IN_WAITING and "Chrome" in mc.SIGN_IN_WAITING


def test_check_loads_mind_headless_and_says_whether_the_sign_in_still_works(client, monkeypatch):
    from app import mind_client as mc

    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "idle")
    monkeypatch.setattr(mc, "is_authenticated", lambda: {"authenticated": False, "url": "https://login.example/authorize"})
    st = client.get("/api/mind/sign-in", params={"check": True}).json()
    assert st["checked"]["signed_in"] is False and st["state"] == "idle"
    monkeypatch.setattr(mc, "is_authenticated", lambda: {"authenticated": True})
    assert client.get("/api/mind/sign-in", params={"check": True}).json()["checked"]["signed_in"] is True


def test_run_in_mind_waits_while_the_sign_in_window_is_open(client, section_heading_grids_xlsx, monkeypatch):
    sid = _upload(client, section_heading_grids_xlsx)
    monkeypatch.setitem(server.MIND_SIGN_IN, "state", "waiting")
    res = client.post(f"/api/sessions/{sid}/mind-loop", json={"maxIterations": 1})
    assert res.status_code == 409 and "signing in" in res.json()["detail"]
    assert client.get("/api/mind/sign-in", params={"check": True}).status_code == 409
