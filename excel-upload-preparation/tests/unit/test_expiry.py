"""A shared copy's end date (app/expiry.py): before / on / after the date, a clock moved back, the
developer version without one, the server's 410 gate, the program's start message and the package
build's EXPIRY file. The clock and the remembered date are injected; no real network."""
from __future__ import annotations

import datetime as dt
import importlib.util
import json
import urllib.error
from email.message import Message
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app import expiry
from app.expiry import Checker, EndDate, decide, read_end_date
from app.web import server

LAST_DAY = dt.date(2026, 12, 15)
JOSEPH = EndDate(LAST_DAY, "Joseph Gestetner")
STOPPED = "This copy of Mind Ready stopped working on 15 December 2026. Ask Joseph Gestetner for a new copy."
BUILD_PY = Path(__file__).resolve().parents[2] / "tools" / "package" / "build.py"


@pytest.fixture(autouse=True)
def _program_checker_back():
    """Each test that replaces this program's checker gets the real one back (no end date in the source)."""
    yield
    expiry.use(None)


def checker(tmp_path: Path, today: dt.date, end: EndDate | None = JOSEPH) -> Checker:
    return Checker(end, today=lambda: today, state_file=tmp_path / "latest_date_seen.json")


def remembered(tmp_path: Path) -> str | None:
    path = tmp_path / "latest_date_seen.json"
    return json.loads(path.read_text(encoding="utf-8"))["latest"] if path.exists() else None


# --- the date --------------------------------------------------------------------------------------------
def test_before_the_date_it_works_and_counts_the_days():
    v = decide(JOSEPH, dt.date(2026, 12, 12), None)
    assert (v.expired, v.days_left, v.message) == (False, 3, None)
    assert v.health() == {"expires": "2026-12-15", "days_left": 3, "expired": False, "expired_message": None}


def test_it_works_all_day_on_the_last_day():
    v = decide(JOSEPH, LAST_DAY, None)
    assert (v.expired, v.days_left) == (False, 0)


def test_from_the_next_day_it_stops_with_the_plain_message():
    v = decide(JOSEPH, dt.date(2026, 12, 16), None)
    assert (v.expired, v.reason, v.message) == (True, "date", STOPPED)
    assert v.refusal() == {"detail": STOPPED, "expired": True, "expires": "2026-12-15"}
    # no contact in the build: whoever shared it
    assert decide(EndDate(LAST_DAY), dt.date(2027, 1, 3), None).message == (
        "This copy of Mind Ready stopped working on 15 December 2026. Ask the person who shared it for a new copy."
    )


def test_the_developer_version_has_no_end_date_and_writes_nothing(tmp_path):
    # the source has no EXPIRY file next to VERSION, and no MIND_READY_EXPIRES: the developer servers never stop
    assert not expiry.EXPIRY_FILE.exists()
    assert read_end_date(environ={}) is None
    c = checker(tmp_path, dt.date(2099, 1, 1), end=None)
    assert c.verdict() == expiry.Verdict()
    assert c.verdict().health() == {"expires": None, "days_left": None, "expired": False, "expired_message": None}
    assert remembered(tmp_path) is None
    assert expiry.start_network_check(c, fetch=lambda: pytest.fail("no network for the developer version")) is None


# --- the clock moved back -------------------------------------------------------------------------------
def test_the_latest_date_seen_is_written_at_most_once_a_day(tmp_path, monkeypatch):
    writes: list[str] = []
    real_replace = expiry.os.replace
    monkeypatch.setattr(expiry.os, "replace", lambda a, b: (writes.append(str(b)), real_replace(a, b)))
    day = [dt.date(2026, 11, 9)]
    c = Checker(JOSEPH, today=lambda: day[0], state_file=tmp_path / "latest_date_seen.json")
    for _ in range(3):
        assert not c.verdict().expired
    assert remembered(tmp_path) == "2026-11-09" and len(writes) == 1
    day[0] = dt.date(2026, 11, 10)
    c.verdict()
    c.verdict()
    assert remembered(tmp_path) == "2026-11-10" and len(writes) == 2


def test_a_clock_moved_back_more_than_2_days_stops_the_copy(tmp_path):
    checker(tmp_path, dt.date(2026, 12, 10)).verdict()  # the copy ran on 10 December
    back = checker(tmp_path, dt.date(2026, 12, 7)).verdict()  # then the clock says 7 December
    assert (back.expired, back.reason) == (True, "clock")
    assert back.message == (
        "This copy of Mind Ready stopped working: this computer says it is 7 December 2026, but it was already "
        "10 December 2026. Set the right date, or ask Joseph Gestetner for a new copy."
    )
    assert remembered(tmp_path) == "2026-12-10"  # never moved back
    # 2 days behind is believed (time zones, a clock put right); the days are counted from the later date
    ok = checker(tmp_path, dt.date(2026, 12, 8)).verdict()
    assert (ok.expired, ok.days_left) == (False, 5)


def test_a_date_seen_after_the_last_day_stops_it_even_with_the_clock_back_a_day(tmp_path):
    checker(tmp_path, dt.date(2026, 12, 16)).verdict()
    v = checker(tmp_path, LAST_DAY).verdict()
    assert (v.expired, v.reason, v.message) == (True, "date", STOPPED)


def test_an_unreadable_remembered_date_believes_the_clock(tmp_path):
    (tmp_path / "latest_date_seen.json").write_text("not json", encoding="utf-8")
    assert not checker(tmp_path, dt.date(2026, 12, 1)).verdict().expired
    assert remembered(tmp_path) == "2026-12-01"


# --- where the date comes from --------------------------------------------------------------------------
def test_the_env_var_first_then_the_expiry_file_the_earlier_date_wins(tmp_path):
    file = tmp_path / "EXPIRY"
    assert read_end_date(environ={"MIND_READY_EXPIRES": "2026-12-15"}, path=file) == EndDate(LAST_DAY)
    file.write_text(expiry.end_date_text(LAST_DAY, "Joseph Gestetner"), encoding="utf-8")
    assert read_end_date(environ={}, path=file) == JOSEPH
    # a developer's test date brings the copy's date forward; it can never push it back
    assert read_end_date(environ={"MIND_READY_EXPIRES": "2026-11-12"}, path=file) == EndDate(dt.date(2026, 11, 12), "Joseph Gestetner")
    assert read_end_date(environ={"MIND_READY_EXPIRES": "2027-06-30"}, path=file) == JOSEPH


@pytest.mark.parametrize("env, text", [
    ({"MIND_READY_EXPIRES": "15/12/2026"}, None),
    ({}, "not json"),
    ({}, '{"contact": "Joseph Gestetner"}'),
    ({}, '{"expires": "2026-02-30"}'),
])
def test_a_date_that_cannot_be_read_stops_the_copy(tmp_path, env, text):
    file = tmp_path / "EXPIRY"
    if text is not None:
        file.write_text(text, encoding="utf-8")
    end = read_end_date(environ=env, path=file)
    assert end is not None and end.last_day is None
    v = checker(tmp_path, dt.date(2026, 11, 9), end).verdict()
    assert (v.expired, v.reason) == (True, "unreadable")
    assert v.message == (
        "MIND_READY_EXPIRES should be a date such as 2026-12-15, not '15/12/2026'." if env
        else "This copy of Mind Ready cannot read its end date. Ask the person who shared it for a new copy."
    )


# --- the decision the server reuses for a minute --------------------------------------------------------
def test_the_server_decides_at_most_once_a_minute(tmp_path):
    day = [LAST_DAY]
    expiry.use(Checker(JOSEPH, today=lambda: day[0], state_file=tmp_path / "seen.json"))
    assert not expiry.current(now=lambda: 1000.0).expired
    day[0] = dt.date(2026, 12, 16)  # midnight passes
    assert not expiry.current(now=lambda: 1059.0).expired  # the minute's decision
    assert expiry.current(now=lambda: 1060.0).expired
    assert expiry.at_start().expired


# --- the internet's date --------------------------------------------------------------------------------
def test_the_date_header_is_read_as_a_local_date():
    noon = dt.datetime(2026, 11, 16, 12, 0, tzinfo=dt.timezone.utc).astimezone().date()
    assert expiry.date_from_header("Mon, 16 Nov 2026 12:00:00 GMT") == noon
    for bad in (None, "", "tomorrow", "Fri, 01 Jan 1999 12:00:00 GMT"):
        assert expiry.date_from_header(bad) is None


class _Answer:
    def __init__(self, date: str | None):
        self.headers = Message()
        if date:
            self.headers["Date"] = date

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def test_one_quick_head_request_a_refusal_still_says_the_date():
    asked: list[tuple[str, str, float]] = []

    def opener(request, timeout):
        asked.append((request.full_url, request.get_method(), timeout))
        if "gateway" in request.full_url:
            raise urllib.error.URLError("no network")
        headers = Message()
        headers["Date"] = "Wed, 02 Dec 2026 12:00:00 GMT"
        raise urllib.error.HTTPError(request.full_url, 404, "Not Found", headers, None)

    day = expiry.internet_date(["https://gateway.example", "https://www.microsoft.com"], opener=opener)
    assert day == dt.datetime(2026, 12, 2, 12, 0, tzinfo=dt.timezone.utc).astimezone().date()
    assert asked == [("https://gateway.example", "HEAD", 2.0), ("https://www.microsoft.com", "HEAD", 2.0)]
    assert expiry.internet_date(["https://a.example"], opener=lambda r, timeout: _Answer(None)) is None  # no Date header

    def down(request, timeout):
        raise OSError("the network is down")

    assert expiry.internet_date(["https://a.example", "https://b.example"], opener=down) is None


def test_the_internet_date_only_moves_the_latest_date_forward(tmp_path):
    c = checker(tmp_path, dt.date(2026, 12, 1))
    c.verdict()
    expiry.start_network_check(c, fetch=lambda: dt.date(2026, 11, 20)).join(5)
    assert remembered(tmp_path) == "2026-12-01"  # an earlier date never moves it back
    expiry.start_network_check(c, fetch=lambda: dt.date(2026, 12, 16)).join(5)
    assert remembered(tmp_path) == "2026-12-16"
    assert c.verdict().expired  # the internet says the last day is over, whatever the clock says

    def down():
        raise OSError("no network")

    other = checker(tmp_path / "other", dt.date(2026, 12, 1))
    expiry.start_network_check(other, fetch=down).join(5)  # no error, no change
    assert not other.verdict().expired


def test_the_sites_asked_are_the_ai_gateway_then_microsoft():
    from app.settings import SETTINGS

    assert expiry.date_sites() == [SETTINGS.ai_endpoint, "https://www.microsoft.com"]


# --- the server: 410 on every API call after the date, /api/health still answers ------------------------
def _server_on(tmp_path: Path, today: dt.date, end: EndDate | None = JOSEPH) -> TestClient:
    expiry.use(checker(tmp_path, today, end))
    return TestClient(server.app)


def test_after_the_date_every_api_call_answers_410_with_the_message(tmp_path):
    client = _server_on(tmp_path, dt.date(2026, 12, 16))
    for res in (client.get("/api/sessions/s1/versions"), client.get("/api/assistant-key"), client.post("/api/sessions", data={"mode": "plan"})):
        assert res.status_code == 410
        assert res.json() == {"detail": STOPPED, "expired": True, "expires": "2026-12-15"}
    health = client.get("/api/health")
    assert health.status_code == 200
    body = health.json()
    assert (body["expires"], body["expired"], body["expired_message"]) == ("2026-12-15", True, STOPPED)
    assert "version" in body and "rules" in body  # a second start still recognises Mind Ready
    assert client.get("/").status_code != 410  # the screens still load, to show the notice


def test_before_the_date_the_api_answers_as_usual_and_health_says_the_date(tmp_path):
    client = _server_on(tmp_path, dt.date(2026, 12, 12))
    assert client.get("/api/sessions/does-not-exist/versions").status_code == 404
    body = client.get("/api/health").json()
    assert (body["expires"], body["days_left"], body["expired"], body["expired_message"]) == ("2026-12-15", 3, False, None)


def test_the_developer_version_says_no_date(tmp_path):
    body = _server_on(tmp_path, dt.date(2099, 1, 1), end=None).get("/api/health").json()
    assert (body["expires"], body["days_left"], body["expired"], body["expired_message"]) == (None, None, False, None)


# --- the program: the message, and nothing starts ------------------------------------------------------
def test_the_program_stops_on_the_message_with_the_window_open(monkeypatch, capsys):
    import mind_ready_app

    prompts: list[str] = []
    monkeypatch.setattr(expiry, "at_start", lambda: decide(JOSEPH, dt.date(2026, 12, 16), None))
    monkeypatch.setattr("builtins.input", lambda prompt="": prompts.append(prompt) or "")
    monkeypatch.setattr(server, "main", lambda: pytest.fail("the server must not start"))
    mind_ready_app.main()
    assert STOPPED in capsys.readouterr().out
    assert prompts == ["\nPress Enter to close this window."]


def test_the_program_starts_and_says_until_when(monkeypatch, capsys):
    import mind_ready_app

    started: list[bool] = []
    monkeypatch.setattr(expiry, "at_start", lambda: decide(JOSEPH, dt.date(2026, 12, 12), None))
    monkeypatch.setattr(server, "_port_is_taken", lambda port: True)  # no browser from a test
    monkeypatch.setattr(server, "main", lambda: started.append(True))
    mind_ready_app.main()
    assert started == [True]
    assert "This copy works until 15 December 2026." in capsys.readouterr().out


# --- the package build writes the EXPIRY file (Nuitka itself never runs here) -----------------------------
@pytest.fixture(scope="module")
def build():
    spec = importlib.util.spec_from_file_location("mind_ready_build", BUILD_PY)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_the_build_writes_the_expiry_file_the_app_reads(build, tmp_path):
    path = build.write_expiry_file(tmp_path / "package-data", LAST_DAY, "Joseph Gestetner")
    assert path.name == "EXPIRY"
    assert json.loads(path.read_text(encoding="utf-8")) == {"expires": "2026-12-15", "contact": "Joseph Gestetner"}
    assert read_end_date(environ={}, path=path) == JOSEPH
    assert read_end_date(environ={}, path=build.write_expiry_file(tmp_path / "b", LAST_DAY)) == EndDate(LAST_DAY)


def test_the_build_packages_it_next_to_version(build, tmp_path):
    file = tmp_path / "package-data" / "EXPIRY"
    cmd = build.nuitka_command(tmp_path, "1.12.4", True, file)
    assert f"--include-data-files={file}=EXPIRY" in cmd
    assert any(c.endswith("=VERSION") for c in cmd)
    assert not any(c.endswith("=EXPIRY") for c in build.nuitka_command(tmp_path, "1.12.4", True, None))


def test_the_build_expires_on_15_december_unless_told_otherwise(build):
    today = dt.date(2026, 11, 9)
    args = build.parse_args(["out", "--onefile"])
    assert (args.out, args.onefile, args.contact) == ("out", True, "")
    assert build.end_date(args, today) == LAST_DAY
    assert build.end_date(build.parse_args(["out", "--expires", "2027-01-31", "--contact", "Joseph Gestetner"]), today) == dt.date(2027, 1, 31)
    assert build.end_date(build.parse_args(["out", "--no-expiry"]), today) is None  # none only when said
    with pytest.raises(SystemExit, match="should be a date"):
        build.end_date(build.parse_args(["out", "--expires", "15.12.2026"]), today)
    with pytest.raises(SystemExit, match="is past"):
        build.end_date(build.parse_args(["out", "--expires", "2026-11-08"]), today)
    with pytest.raises(SystemExit):
        build.parse_args(["out", "--expires", "2026-12-15", "--no-expiry"])  # one or the other
