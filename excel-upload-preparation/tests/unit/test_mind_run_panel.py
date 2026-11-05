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
