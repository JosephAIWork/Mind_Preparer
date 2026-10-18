"""1.11.6: the "Do everything" time estimate (app/estimate.py) and its endpoint."""
from fastapi.testclient import TestClient

from app.estimate import CAVEAT, estimate_chain
from app.web import server


def test_phoenix_numbers_give_the_range_it_took():
    # Phoenix TVOG: 179 MB, scan 344 s, 6,147 error cells -- Fix all issues took ~37 min
    r = estimate_chain(179, 344, 305_785, 6147, 844, True)
    fix = next(s for s in r["steps"] if s["id"] == "fixAll")
    assert fix["low_s"] <= 37 * 60 <= fix["high_s"]
    assert [s["id"] for s in r["steps"]] == ["fixAll", "prepare", "loop", "recalc"]
    assert r["total"]["low_s"] <= r["total"]["mid_s"] <= r["total"]["high_s"]
    assert r["caveat"] == CAVEAT and "depends" in CAVEAT


def test_a_clean_workbook_and_no_loop_cost_less():
    clean = estimate_chain(27, 300, 57_000, 0, 0, False)
    broken = estimate_chain(27, 300, 57_000, 500, 200, True)
    assert clean["total"]["mid_s"] < broken["total"]["mid_s"]
    assert next(s for s in clean["steps"] if s["id"] == "loop")["mid_s"] == 0


def test_an_unscanned_file_falls_back_to_its_formula_count():
    r = estimate_chain(10, None, 100_000, None)
    assert r["basis"]["scan_measured"] is False and r["basis"]["scan_s"] == 350


def test_the_endpoint_answers_after_a_scan(flagged_model_broken_xlsx):
    client = TestClient(server.app)
    with flagged_model_broken_xlsx.open("rb") as f:
        sid = client.post("/api/sessions", files={"file": (flagged_model_broken_xlsx.name, f, "application/octet-stream")}, data={"mode": "plan"}).json()["sessionId"]
    body = client.get(f"/api/sessions/{sid}/estimate?loop=false").json()
    assert body["basis"]["scan_measured"] and body["basis"]["stored_errors"] is not None
    assert next(s for s in body["steps"] if s["id"] == "loop")["mid_s"] == 0
    assert client.get("/api/sessions/nope/estimate").status_code == 404
