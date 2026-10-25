"""The gateway key pasted in the web app: saved encrypted for the Windows user, never sent back
whole, refused when the gateway refuses it (app/ai_key.py, app/llm.py, /api/assistant-key)."""
from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app import ai_key, llm
from app.web import server

win32crypt = pytest.importorskip("win32crypt")

KEY = "0123456789abcdef0123456789abWXYZ"


@pytest.fixture(autouse=True)
def key_file(tmp_path, monkeypatch):
    """Every test saves into its own folder, never the real %LOCALAPPDATA%."""
    path = tmp_path / "MindReady" / "ai_key.bin"
    monkeypatch.setattr(ai_key, "KEY_FILE", path)
    return path


@pytest.fixture
def client():
    return TestClient(server.app)


class FakeResp:
    def __init__(self, status_code):
        self.status_code = status_code


def test_save_and_load_round_trip_without_the_key_in_the_file(key_file):
    assert ai_key.load() is None
    ai_key.save(KEY)
    assert ai_key.load() == KEY
    assert KEY.encode() not in key_file.read_bytes()
    assert llm.find_api_key() == (KEY, llm.DEFAULT_AZURE_ENDPOINT)


def test_a_damaged_file_reads_as_no_key(key_file):
    key_file.parent.mkdir(parents=True)
    key_file.write_bytes(b"not a DPAPI blob")
    assert ai_key.load() is None and llm.find_api_key() is None


def test_check_format():
    assert ai_key.check_format("") is not None
    assert ai_key.check_format("two words here") is not None
    assert ai_key.check_format("short") is not None
    assert ai_key.check_format(KEY) is None


def test_check_api_key_refuses_only_on_401_or_403(monkeypatch):
    import requests

    for status, accepted in ((401, False), (403, False), (200, True), (400, True), (429, True)):
        monkeypatch.setattr(requests, "post", lambda *a, _s=status, **k: FakeResp(_s))
        assert llm.check_api_key(KEY)[0] is accepted, status

    def offline(*a, **k):
        raise requests.ConnectionError("offline")

    monkeypatch.setattr(requests, "post", offline)
    accepted, note = llm.check_api_key(KEY)
    assert accepted and "could not be reached" in note


def test_endpoint_saves_a_key_and_shows_only_its_last_4(client, monkeypatch):
    monkeypatch.setattr(server, "check_api_key", lambda key: (True, None))
    assert client.get("/api/assistant-key").json() == {"saved": False, "last4": None}
    res = client.post("/api/assistant-key", json={"key": f"  {KEY}\n"})
    assert res.status_code == 200 and res.json()["last4"] == "WXYZ"
    assert KEY not in res.text
    after = client.get("/api/assistant-key")
    assert after.json() == {"saved": True, "last4": "WXYZ"} and KEY not in after.text
    assert client.get("/api/health").json()["assistant"] is True


def test_endpoint_does_not_save_a_refused_or_malformed_key(client, monkeypatch):
    monkeypatch.setattr(server, "check_api_key", lambda key: (False, "The gateway refused this key."))
    res = client.post("/api/assistant-key", json={"key": KEY})
    assert res.status_code == 422 and "refused" in res.json()["detail"]
    assert client.post("/api/assistant-key", json={"key": "has a space"}).status_code == 422
    assert ai_key.load() is None


def test_endpoint_needs_a_json_body(client):
    """A plain-text POST (the kind another website can send without asking) is refused."""
    res = client.post("/api/assistant-key", content=f'{{"key": "{KEY}"}}', headers={"Content-Type": "text/plain"})
    assert res.status_code == 422 and ai_key.load() is None
