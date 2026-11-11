"""Problem reports behind "Send feedback" (app/feedback.py): codes, formula shapes, nothing from the workbook."""
from __future__ import annotations

import json
import time

import pytest
from fastapi import HTTPException
from fastapi.testclient import TestClient

from app import feedback, llm
from app.perf import stage
from app.settings import SETTINGS
from app.web import server


@pytest.fixture(autouse=True)
def _files(tmp_path, monkeypatch):
    # the incidents go under tmp_path through tests/unit/conftest.py (SETTINGS.data_folder); the timings too,
    # but one dated now is needed: feedback._timings() keeps the last day only, so a fixed date would age out
    timings = tmp_path / "timings.jsonl"
    timings.write_text(json.dumps({"at": time.strftime("%Y-%m-%d %H:%M:%S"), "stage": "scan", "s": 12.5, "thread": "excel-abc", "file": "Secret Client Model.xlsm"}) + "\n", encoding="utf-8")
    monkeypatch.setenv("MINDREADY_PERF_LOG", str(timings))


def test_the_incidents_of_a_test_never_reach_the_users_real_file(tmp_path):
    # tests/unit/conftest.py: the data folder of every test is its own tmp_path, so a recorded incident
    # lands there and the user's %LOCALAPPDATA%\MindReady\feedback\incidents.jsonl is never touched
    assert SETTINGS.data_folder == tmp_path / "MindReady"
    inc = feedback.record(None, step="Build the Mind instances", what="no plan today")
    assert feedback._path() == tmp_path / "MindReady" / "feedback" / "incidents.jsonl"
    assert feedback._path().is_file() and feedback.find(inc["ref"]) == inc


def test_codes_are_stable():
    # the developers' decoder computes the same codes: a change here breaks every message already sent
    assert feedback.code("stage", "scan") == "RT515"
    assert feedback.code("fn", "app.prep._write_one") == "HA467"


@pytest.mark.parametrize("formula, expected", [
    ('=MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,255)', '=MID(CELL("filename",<cell>),FIND("]",CELL("filename",<cell>))+1,<n>)'),
    ('=IF(Base!R5="N",1,0)', '=IF(<sheet>!<cell>="<text>",1,0)'),
    ("=SUM('My Sheet'!A1:B20)*0.035+colLHModel", "=SUM(<sheet>!<range>)*<n>+<name>"),
    ("=INDEX(Data!$A:$A,MATCH(H9,colLHModel,0))", "=INDEX(<sheet>!<columns>,MATCH(<cell>,<name>,0))"),
    ("=[1]Sheet1!A1+Table1[Col]", "=<external>!<cell>+<table ref>"),
    ("=_xlfn.XLOOKUP(A2,B:B,C:C)+ClientUdf(A1)", "=_xlfn.XLOOKUP(<cell>,<columns>,<columns>)+<function>(<cell>)"),
    ("=SUM(2:2)", "=SUM(<rows>)"),
    ("=A1#", "=<cell>"),
    ("12.5", "<value>"),
])
def test_a_formula_keeps_its_kind_not_its_content(formula, expected):
    assert feedback.shape(formula) == expected


def test_a_message_loses_paths_files_cells_and_the_workbooks_names():
    said = feedback.mask(
        r"Unable to set the Formula2 property at 'Base'!J414 in C:\Users\someone\AppData\Local\Temp\mind_ready_1\Client 500_v3.xlsm "
        "=IF(A1>0,1,0) value 0.035 on LifeHealth", names=["LifeHealth", "Base"],
    )
    assert "Unable to set the Formula2 property" in said
    for gone in ("Base", "J414", "someone", "Client", "xlsm", "IF(A1", "0.035", "LifeHealth"):
        assert gone not in said


def _crash_route(path, exc_factory):
    """A POST route that fails inside a timed stage (GET would meet the front end's catch-all)."""
    if any(getattr(r, "path", None) == path for r in server.app.routes):
        return

    @server.app.post(path)
    def _boom():  # pragma: no cover - raises
        with stage("unit_test_stage"):
            formula = '=CELL("filename",Inputs!A1)'  # noqa: F841 -- read back from the frame by the report
            raise exc_factory()


def test_a_crash_answers_with_an_incident_and_the_message_holds_codes_only():
    _crash_route("/api/test-crash", lambda: ValueError(r"bad cell 'Inputs'!A1 in C:\work\Client Model.xlsm"))
    client = TestClient(server.app, raise_server_exceptions=False)
    res = client.post("/api/test-crash")
    assert res.status_code == 500
    ref = res.json()["incident"]
    # the user's own screen says what stopped, in plain words (their words may appear: it is not sent)
    assert res.json()["problem"].startswith("The app stopped on an error: bad cell")
    assert res.json()["problem"].endswith("Your original file is not changed.")
    msg = client.post("/api/feedback", json={"ref": ref, "note": "I clicked Do everything", "useAssistant": False}).json()
    text = msg["text"]
    assert msg["ref"] == ref and msg["file_name"] == f"MindReady-feedback-{ref}.txt"
    assert msg["email"] and ref in msg["subject"]
    assert f"interrupted at stage {feedback.code('stage', 'unit_test_stage')}" in text
    assert f"Error {feedback.code('error', 'builtins.ValueError')}" in text
    assert 'Formula shape: =CELL("filename",<sheet>!<cell>)' in text
    assert "I clicked Do everything" in text
    assert feedback.code("stage", "scan") in text  # the stage timings, coded
    for gone in ("Inputs", "Client", "xlsm", r"C:\work", "Secret", "test_feedback"):
        assert gone not in text


def test_an_http_error_wrapping_a_crash_reports_the_crash():
    def wrapped():
        try:
            raise KeyError("Sheet9")
        except KeyError as exc:
            raise HTTPException(status_code=500, detail="analysis failed") from exc

    if not any(getattr(r, "path", None) == "/api/test-wrapped" for r in server.app.routes):
        @server.app.post("/api/test-wrapped")
        def _wrapped():  # pragma: no cover - raises
            wrapped()

    client = TestClient(server.app, raise_server_exceptions=False)
    res = client.post("/api/test-wrapped")
    assert res.status_code == 500 and res.json()["detail"] == "analysis failed"
    assert res.json()["problem"] == "The app stopped on an error: 'Sheet9'. Your original file is not changed."
    incident = feedback.find(res.json()["incident"])
    assert incident["error"] == feedback.code("error", "builtins.KeyError")
    assert "Sheet9" not in json.dumps(incident)


def test_a_4xx_stays_as_it_was():
    client = TestClient(server.app)
    res = client.get("/api/sessions/nope")
    assert res.status_code == 404 and "incident" not in res.json()


def test_changes_excel_did_not_write_are_summed_up_by_rule_with_their_shapes():
    ops = [
        {"rule_id": "FRM-012", "op": "set_formula", "sheet": "Base", "cell": "C2", "before": '=CELL("filename",Base!A1)', "after": '="Client folder"',
         "error": "Excel refused this formula at 'Base'!C2", "error_code": "AB123", "excel_code": "0x800A03EC", "trace": ["CD456:2213"], "one_cell_array": True},
        {"rule_id": "FRM-012", "op": "set_formula", "sheet": "Base", "cell": "C3", "before": "=C2", "after": "=1", "error": "refused"},
    ]
    inc = feedback.record(None, step="Fix the formulas", what="not written", stage="apply_writes", failed_ops=ops, names=["Base"])
    text = feedback.compose(inc)["text"]
    assert "FRM-012 not written at stage" in text
    assert "2 change(s) not written: FRM-012 set_formula x2" in text
    assert 'before: =CELL("filename",<sheet>!<cell>)' in text
    assert "one-cell array" in text and "Excel code 0x800A03EC" in text and "trace CD456:2213" in text
    for gone in ("Base", "C2", "Client folder"):
        assert gone not in text


def test_the_plain_sentence_for_changes_excel_did_not_write():
    text = server._plain_problem(None, "Fix the formulas", [{"rule_id": "FRM-012"}, {"rule_id": "STR-008"}])
    assert text.startswith("Fix the formulas: Excel did not accept 2 of the changes, for FRM-012 (a formula asks Excel about its own file), STR-008 (a formula reads a lone text cell).")


def test_feedback_from_the_header_has_no_incident():
    client = TestClient(server.app)
    msg = client.post("/api/feedback", json={"note": "Works well", "useAssistant": False}).json()
    assert "Feedback (no step failed)" in msg["text"] and "Works well" in msg["text"] and "Machine:" in msg["text"]


@pytest.fixture
def fake_assistant(monkeypatch):
    """A chat_completion that answers with a fixed text and keeps what it was asked."""
    asked: dict[str, str] = {}

    def install(answer: str):
        def chat_completion(messages, system_prompt, **kw):
            asked["system"] = system_prompt
            asked["user"] = messages[-1]["content"]
            return {"available": True, "text": answer}

        monkeypatch.setattr(llm, "llm_available", lambda: True)
        monkeypatch.setattr(llm, "chat_completion", chat_completion)
        return asked

    return install


NOTE = "I clicked Do everything and it stopped at the second step."


def test_the_assistant_is_told_to_say_it_in_its_own_words_and_never_copy_the_note(fake_assistant):
    asked = fake_assistant("I was running Fix the formulas when the app stopped.")
    inc = feedback.record(ValueError("boom"), step="Fix the formulas", names=[])
    assert feedback.assistant_lines(inc, NOTE, "a formula asks Excel about its own file") == "I was running Fix the formulas when the app stopped."
    assert "own words" in asked["system"] and "never copy it" in asked["system"]
    assert "do not repeat or quote it" in asked["user"] and NOTE in asked["user"]  # the note is background, labelled as such
    assert "The step the user ran: Fix the formulas." in asked["user"]


def test_a_sentence_of_the_opening_that_repeats_the_note_is_dropped(fake_assistant):
    fake_assistant(f"I was running Fix the formulas. {NOTE} The app reported an error.")
    inc = feedback.record(ValueError("boom"), step="Fix the formulas", names=[])
    ai = feedback.assistant_lines(inc, NOTE)
    text = feedback.compose(inc, NOTE, ai)["text"]
    assert "I was running Fix the formulas. The app reported an error.\n" in text
    assert text.count(NOTE) == 1  # once, under "The user's note"
    assert "The user's note\n  " + NOTE in text


def test_the_opening_goes_when_it_is_only_the_note(fake_assistant):
    fake_assistant(f'"{NOTE}"')
    msg = feedback.compose(None, NOTE, feedback.assistant_lines(None, NOTE))
    assert msg["text"].count(NOTE) == 1 and msg["text"].index("Feedback (no step failed)") < msg["text"].index(NOTE)


@pytest.mark.parametrize("ai, note, expected", [
    ("I was fixing. The scan stopped early. Nothing else.", "the scan stopped early", "I was fixing. Nothing else."),  # case, no full stop
    ("I was fixing.\nIt stopped. Thanks.", "It stopped.\nThanks.", "I was fixing. Thanks."),  # a short line of the note ("Thanks.") does not match
    ("I was fixing the formulas.", "", "I was fixing the formulas."),
    (None, "anything", None),
])
def test_without_note(ai, note, expected):
    assert feedback.without_note(ai, note) == expected


def test_a_stage_error_carries_the_innermost_stage():
    with pytest.raises(RuntimeError) as info:
        with stage("outer_stage"):
            with stage("inner_stage"):
                raise RuntimeError("x")
    assert info.value.mind_ready_stage == "inner_stage"


def test_the_message_says_how_many_tokens_the_assistant_used_on_the_workbook():
    """1.14.1: one line in the log part, from the open workbook's count; none without a workbook."""
    used = {"calls": 7, "input_tokens": 12000, "output_tokens": 340, "cache_read_tokens": 0, "cache_write_tokens": 0}
    msg = feedback.compose(None, "Works well", None, None, ai_usage=used)
    assert "\n  AI used: 12,340 tokens in 7 calls\n" in msg["text"] and msg["aiUsage"] == used
    fresh = {"calls": 0, "input_tokens": 0, "output_tokens": 0, "cache_read_tokens": 0, "cache_write_tokens": 0}
    assert "\n  AI not used yet on this workbook\n" in feedback.compose(None, "Works well", ai_usage=fresh)["text"]
    none = feedback.compose(None, "Works well")
    assert "AI used" not in none["text"] and "AI not used" not in none["text"] and none["aiUsage"] is None
