"""The app shows how many tokens its assistant used on this workbook (1.14.1): every completed call
of app/llm.py reports its tokens to the usage sink of the current context; the server sets one per
session around every step that may ask the assistant (app/web/server.py `_ai_scope` / `_ai_thread`),
the count is saved with the workbook (app/session_store.py) and shown as "AI used: 12,340 tokens in 7
calls" (GET /api/sessions/{id}, GET .../ai-usage, the feedback message).

The assistant here is a fake gateway: `requests.post` answers like the Anthropic Messages passthrough
(plain JSON, or a server-sent event stream when the call streams), with a `usage` object -- so the real
`chat_completion` runs, and the counting is proven where it lives."""
from __future__ import annotations

import json
import threading
from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app import llm, session_store
from app.web import server

USAGE = {"input_tokens": 100, "output_tokens": 20, "cache_read_input_tokens": 5, "cache_creation_input_tokens": 2}
ONE_CALL = {"calls": 1, "input_tokens": 100, "output_tokens": 20, "cache_read_tokens": 5, "cache_write_tokens": 2}


class _Response:
    """What requests.post returns: plain JSON, or an SSE stream when the payload asked for one."""

    def __init__(self, text: str, stream: bool):
        self.status_code = 200
        self.text = ""
        self._answer = text
        self.headers = {"content-type": "text/event-stream" if stream else "application/json"}

    def raise_for_status(self) -> None:
        pass

    def close(self) -> None:
        pass

    def json(self) -> dict[str, Any]:
        return {"content": [{"type": "text", "text": self._answer}], "stop_reason": "end_turn", "usage": USAGE}

    def iter_lines(self, decode_unicode: bool = False):
        events = [
            {"type": "message_start", "message": {"usage": {"input_tokens": USAGE["input_tokens"], "cache_read_input_tokens": 5, "cache_creation_input_tokens": 2}}},
            {"type": "content_block_delta", "delta": {"type": "text_delta", "text": self._answer[:4]}},
            {"type": "content_block_delta", "delta": {"type": "text_delta", "text": self._answer[4:]}},
            {"type": "message_delta", "delta": {"stop_reason": "end_turn"}, "usage": {"output_tokens": USAGE["output_tokens"]}},
            {"type": "message_stop"},
        ]
        for ev in events:
            yield b"event: " + ev["type"].encode()
            yield b"data: " + json.dumps(ev).encode("utf-8")
            yield b""


@pytest.fixture
def gateway(monkeypatch):
    """A saved key and a gateway that answers every call with `answer` (default: a plain sentence), counting the calls."""
    import requests

    calls: list[dict[str, Any]] = []
    answer = {"text": "The workbook looks fine."}

    def post(url, headers=None, json=None, timeout=None, stream=False, **kw):
        calls.append(json)
        return _Response(answer["text"], bool(json.get("stream")))

    monkeypatch.setattr(llm, "find_api_key", lambda: ("key", "https://gateway.example"))
    monkeypatch.setattr(requests, "post", post)
    monkeypatch.setenv("MIND_READY_STREAM", "1")

    def say(text: str) -> None:
        answer["text"] = text

    return type("Gateway", (), {"calls": calls, "say": staticmethod(say)})


@pytest.fixture
def client():
    server.SESSIONS.clear()
    server.AUTOFIX.clear()
    yield TestClient(server.app)
    server.SESSIONS.clear()
    server.AUTOFIX.clear()
    session_store.release_all()


def _upload(client, path: Path) -> dict:
    with path.open("rb") as f:
        res = client.post("/api/sessions", files={"file": (path.name, f, "application/octet-stream")}, data={"mode": "plan"})
    assert res.status_code == 200, res.text
    return res.json()


# --- the count itself (app/llm.py) -------------------------------------------------------------
def test_the_gateways_usage_becomes_the_four_counts_and_odd_values_count_nothing():
    assert llm.usage_of(USAGE) == {"input_tokens": 100, "output_tokens": 20, "cache_read_tokens": 5, "cache_write_tokens": 2}
    assert llm.usage_of({"prompt_tokens": 7, "completion_tokens": 3}) == {"input_tokens": 7, "output_tokens": 3, "cache_read_tokens": 0, "cache_write_tokens": 0}
    assert llm.usage_of({"input_tokens": "x", "output_tokens": None, "cache_read_input_tokens": -4}) == {"input_tokens": 0, "output_tokens": 0, "cache_read_tokens": 0, "cache_write_tokens": 0}
    assert llm.usage_of(None) == {"input_tokens": 0, "output_tokens": 0, "cache_read_tokens": 0, "cache_write_tokens": 0}


def test_the_line_says_tokens_and_calls_or_that_the_assistant_was_not_used():
    assert llm.usage_line(None) == "AI not used yet on this workbook"
    assert llm.usage_line(llm.empty_usage()) == "AI not used yet on this workbook"
    assert llm.usage_line({"calls": 1, "input_tokens": 90, "output_tokens": 10}) == "AI used: 100 tokens in 1 call"
    assert llm.usage_line({"calls": 7, "input_tokens": 12000, "output_tokens": 340, "cache_read_tokens": 5000}) == "AI used: 12,340 tokens in 7 calls"
    total = llm.add_usage(llm.add_usage(llm.empty_usage(), llm.usage_of(USAGE)), llm.usage_of(USAGE))
    assert total == {"calls": 2, "input_tokens": 200, "output_tokens": 40, "cache_read_tokens": 10, "cache_write_tokens": 4}


def test_a_completed_call_reports_to_the_sink_of_its_context_streamed_or_not(gateway):
    got: list[dict[str, int]] = []
    with llm.usage_scope(got.append):
        plain = llm.chat_completion([{"role": "user", "content": "hi"}], "be brief")
        streamed = llm.chat_completion([{"role": "user", "content": "hi"}], "be brief", on_delta=lambda _t: None)
    outside = llm.chat_completion([{"role": "user", "content": "hi"}], "be brief")  # no sink here: nothing reported
    assert plain["text"] == "The workbook looks fine." and streamed["text"] == "The workbook looks fine." and outside["text"]
    assert len(gateway.calls) == 3 and gateway.calls[1].get("stream") is True
    assert got == [llm.usage_of(USAGE), llm.usage_of(USAGE)]


def test_a_sink_that_fails_never_breaks_the_call_and_a_failed_call_counts_nothing(gateway, monkeypatch, capsys):
    def broken(_usage: dict[str, int]) -> None:
        raise RuntimeError("disk full")

    with llm.usage_scope(broken):
        res = llm.chat_completion([{"role": "user", "content": "hi"}], "be brief")
    assert res["text"] == "The workbook looks fine." and "disk full" in capsys.readouterr().out

    import requests

    def down(*a, **k):
        raise ConnectionError("gateway down")

    monkeypatch.setattr(requests, "post", down)
    got: list[dict[str, int]] = []
    with llm.usage_scope(got.append):
        res = llm.chat_completion([{"role": "user", "content": "hi"}], "be brief")
    assert res["text"] is None and "gateway down" in res["message"] and got == []


def test_a_thread_runs_under_the_scope_it_was_given(gateway):
    got: list[dict[str, int]] = []
    t = threading.Thread(target=llm.under_usage_scope(got.append, lambda: llm.chat_completion([{"role": "user", "content": "hi"}], "x")))
    t.start()
    t.join(10)
    assert got == [llm.usage_of(USAGE)]
    assert llm._USAGE_SINK.get() is None  # the scope stayed in the thread


# --- per workbook, through the app (app/web/server.py) ------------------------------------------
def test_a_chat_a_suggestion_and_the_grid_namer_all_add_to_the_workbooks_count(client, flagged_model_broken_xlsx, gateway):
    body = _upload(client, flagged_model_broken_xlsx)
    sid = body["sessionId"]
    assert body.get("aiUsage", llm.empty_usage()) == llm.empty_usage()
    assert client.get(f"/api/sessions/{sid}/ai-usage").json() == {"aiUsage": llm.empty_usage(), "line": "AI not used yet on this workbook"}

    # the chat (one call: a plain answer, no lookup, no proposal)
    res = client.post(f"/api/sessions/{sid}/chat", json={"history": [], "question": "What blocks this workbook?"})
    assert res.status_code == 200 and res.json()["text"] == "The workbook looks fine."
    assert server.SESSIONS[sid].ai_usage == ONE_CALL

    # the streamed chat runs in a thread of its own: counted all the same
    res = client.post(f"/api/sessions/{sid}/chat/stream", json={"history": [], "question": "And now?"})
    events = [json.loads(line) for line in res.text.splitlines() if line.strip()]
    assert events[-1]["type"] == "done" and events[-1]["text"] == "The workbook looks fine."
    assert server.SESSIONS[sid].ai_usage["calls"] == 2

    # a suggestion for one finding
    finding = next(f for f in body["report"]["findings"] if f["rule_id"].startswith("FRM") or f["location"].get("cell"))
    res = client.post(f"/api/sessions/{sid}/suggest", json={"ruleId": finding["rule_id"], "sheet": finding["location"].get("sheet"), "cell": finding["location"].get("cell")})
    assert res.status_code == 200, res.text
    assert server.SESSIONS[sid].ai_usage["calls"] == 3

    # the grid namer: one batched call for the whole workbook (the job path, run with the real suggest_names)
    gateway.say('{"Model!A4": "Assumptions"}')
    res = client.post(f"/api/sessions/{sid}/grid-names", json={"apply": False})
    assert res.status_code == 200, res.text
    assert res.json()["available"] is True
    usage = server.SESSIONS[sid].ai_usage
    assert usage["calls"] == 4 and usage["input_tokens"] == 400 and usage["output_tokens"] == 80 and usage["cache_read_tokens"] == 20 and usage["cache_write_tokens"] == 8
    assert len(gateway.calls) == 4

    # the answers the screens read
    assert client.get(f"/api/sessions/{sid}").json()["aiUsage"] == usage
    assert client.get(f"/api/sessions/{sid}/ai-usage").json() == {"aiUsage": usage, "line": "AI used: 480 tokens in 4 calls"}
    assert client.get("/api/sessions/nope/ai-usage").status_code == 404


def test_the_count_survives_a_save_and_a_restart(client, flagged_model_broken_xlsx, gateway):
    sid = _upload(client, flagged_model_broken_xlsx)["sessionId"]
    assert client.post(f"/api/sessions/{sid}/chat", json={"history": [], "question": "hi"}).status_code == 200
    assert client.post(f"/api/sessions/{sid}/chat", json={"history": [], "question": "again"}).status_code == 200
    before = dict(server.SESSIONS[sid].ai_usage)
    assert before["calls"] == 2
    # saved at once, with the state file (no other step needed)
    state = json.loads((session_store.folder_of(sid) / session_store.STATE_FILE).read_text(encoding="utf-8"))
    assert state["ai_usage"] == before
    # the app restarts: nothing in memory; Reopen brings the count back, and the next call adds to it
    server.SESSIONS.clear()
    server.AUTOFIX.clear()
    session_store.release_all()
    opened = client.post(f"/api/saved-workbooks/{sid}/open")
    assert opened.status_code == 200, opened.text
    assert opened.json()["aiUsage"] == before
    assert client.post(f"/api/sessions/{sid}/chat", json={"history": [], "question": "once more"}).status_code == 200
    assert server.SESSIONS[sid].ai_usage["calls"] == 3 and server.SESSIONS[sid].ai_usage["input_tokens"] == 300


def test_a_state_file_from_before_the_count_opens_with_an_empty_count(client, flagged_model_broken_xlsx):
    sid = _upload(client, flagged_model_broken_xlsx)["sessionId"]
    path = session_store.folder_of(sid) / session_store.STATE_FILE
    state = json.loads(path.read_text(encoding="utf-8"))
    state.pop("ai_usage", None)
    path.write_text(json.dumps(state), encoding="utf-8")
    server.SESSIONS.clear()
    session_store.release_all()
    assert client.post(f"/api/saved-workbooks/{sid}/open").json()["aiUsage"] == llm.empty_usage()
    # an odd value in the file: the count starts again rather than failing the reopen
    state["ai_usage"] = "lots"
    path.write_text(json.dumps(state), encoding="utf-8")
    server.SESSIONS.clear()
    session_store.release_all()
    assert client.post(f"/api/saved-workbooks/{sid}/open").json()["aiUsage"] == llm.empty_usage()


def test_the_feedback_message_says_what_the_assistant_used_on_the_open_workbook(client, flagged_model_broken_xlsx, gateway):
    sid = _upload(client, flagged_model_broken_xlsx)["sessionId"]
    # nothing used yet: the message says so, in the log part
    msg = client.post("/api/feedback", json={"note": "Works well", "useAssistant": False, "sessionId": sid}).json()
    assert "  AI not used yet on this workbook\n" in msg["text"] and msg["aiUsage"] == llm.empty_usage()
    # the assistant writes the opening lines: that call counts for the workbook too, and the line is in the same message
    gateway.say("I was sending feedback from the header; no step failed.")
    msg = client.post("/api/feedback", json={"note": "Works well", "useAssistant": True, "sessionId": sid}).json()
    assert msg["assistant"] is True and msg["aiUsage"] == ONE_CALL
    assert "  AI used: 120 tokens in 1 call\n" in msg["text"]
    assert msg["text"].index("Workbook:") < msg["text"].index("AI used:")
    # no open workbook: no line, nothing to count
    msg = client.post("/api/feedback", json={"note": "Works well", "useAssistant": False}).json()
    assert "AI used" not in msg["text"] and "AI not used" not in msg["text"] and msg["aiUsage"] is None
