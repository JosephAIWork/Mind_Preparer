"""An end date for a shared copy of Mind Ready (docs/SHARING.md).

A copy built to share (`tools/package/build.py --expires 2026-12-15`) carries a small file EXPIRY next
to VERSION: the last day it works, who to ask for a new copy, and when it was built. It works all day
on that date (the computer's local time) and stops from the next day: the program shows a plain
message and does not start (mind_ready_app.py), and a server still running answers every API call with
410 and the same message (app/web/server.py), which the screens show instead of the app.

The source has no EXPIRY file, so the developer servers never stop. MIND_READY_EXPIRES=2026-12-15
gives them a date to try it with; next to an EXPIRY file the earlier date wins (the variable can make a
copy stop sooner, never later).

A clock moved back: each copy remembers the latest date this computer has shown it (a small file in
%LOCALAPPDATA%\\MindReady, written at most once a day, and only by a copy that has an end date), and a
clock more than 2 days behind it counts as expired. At start one quick HTTPS request, in the background,
reads the internet's date (the Date header of the answer): it moves that date forward, and puts right
one that has not come yet. Nothing waits for it and nothing fails without a network.

A clock set ahead while a copy ran leaves it a date that has not come yet: put back right, the clock
would look moved back. So a copy that would stop only because of the date it remembers first asks the
internet's date, at start (a few seconds at most), and is put right by it. Without a network, a new
copy still works: every build remembers its own date, never another copy's. A deterrent, not copy
protection: someone who edits the program's own files can get round it.
"""
from __future__ import annotations

import contextlib
import datetime as dt
import json
import os
import re
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
EXPIRY_FILE = PACKAGE_ROOT / "EXPIRY"  # next to VERSION; only the package build writes one
ENV_VAR = "MIND_READY_EXPIRES"
# The default data folder, not the one settings.yaml may name: a settings.yaml sits next to the shared
# .exe, and editing it must not forget the dates already seen. One file for every copy on this
# computer: {copy: its latest date seen}.
STATE_FILE = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "MindReady" / "latest_date_seen.json"
# The name the developer servers (a date from MIND_READY_EXPIRES alone, no EXPIRY file) keep their
# latest date seen under: theirs, never a shared copy's.
DEVELOPER_COPY = "developer"
SLACK_DAYS = 2  # a clock up to 2 days behind the latest date seen is believed (time zones, a clock put right)
CACHE_SECONDS = 60.0  # every API call asks: the server decides at most once a minute
FALLBACK_DATE_SITE = "https://www.microsoft.com"
NETWORK_TIMEOUT = 2.0
_DAY_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# English month names whatever the computer's language (strftime's %B follows the locale)
_MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")


@dataclass(frozen=True)
class EndDate:
    """The last day a copy works, and who its message says to ask for a new copy. `problem` instead of
    a day when the date could not be read: the copy then counts as expired (a broken date must not
    mean no date). `copy` is the name its latest date seen is kept under: every build is its own copy
    (its EXPIRY says when it was built), so a date one copy saw never stops another."""

    last_day: dt.date | None
    contact: str | None = None
    problem: str | None = None
    copy: str = DEVELOPER_COPY


@dataclass(frozen=True)
class Verdict:
    """What the program and the server do now. `days_left` counts from the latest date seen: 0 on the
    last day, negative after it."""

    expires: dt.date | None = None
    expired: bool = False
    days_left: int | None = None
    message: str | None = None  # the plain message once expired
    reason: str | None = None  # why it is expired: "date", "clock" (moved back) or "unreadable"

    def health(self) -> dict[str, Any]:
        """The fields of GET /api/health (None everywhere for the developer version)."""
        return {
            "expires": self.expires.isoformat() if self.expires else None,
            "days_left": self.days_left,
            "expired": self.expired,
            "expired_message": self.message,
        }

    def refusal(self) -> dict[str, Any]:
        """The body of the 410 every API call gets once expired (`detail`, as any refused call has)."""
        return {"detail": self.message, "expired": True, "expires": self.expires.isoformat() if self.expires else None}


def long_date(day: dt.date) -> str:
    """15 December 2026."""
    return f"{day.day} {_MONTHS[day.month - 1]} {day.year}"


def parse_day(text: str) -> dt.date:
    """A date written 2026-12-15 (nothing else: fromisoformat would also take 20261215 or a week date)."""
    text = str(text).strip()
    if not _DAY_RE.match(text):
        raise ValueError(f"not a date such as 2026-12-15: {text!r}")
    return dt.date.fromisoformat(text)


def end_date_text(last_day: dt.date, contact: str | None = None, built: str | None = None) -> str:
    """The content of the EXPIRY file (tools/package/build.py writes it; read_end_date reads it).
    `built`, when the build ran, makes each build its own copy for the dates it remembers."""
    data = {"expires": last_day.isoformat(), "contact": (contact or "").strip() or None, "built": (built or "").strip() or None}
    return json.dumps(data, indent=2) + "\n"


def _who(contact: str | None) -> str:
    return contact or "the person who shared it"


def read_end_date(environ: Mapping[str, str] | None = None, path: Path | None = None) -> EndDate | None:
    """The copy's end date: MIND_READY_EXPIRES first, then the EXPIRY file next to VERSION; with both,
    the earlier date (the contact and the copy's name come from the file: the variable cannot make a
    copy forget the dates it saw). None: no end date (the developer version). An EXPIRY file always
    holds a date: one without, or one that cannot be read, is a problem, not "none"."""
    env = os.environ if environ is None else environ
    path = EXPIRY_FILE if path is None else path
    days: list[dt.date] = []
    contact: str | None = None
    copy = DEVELOPER_COPY
    text = (env.get(ENV_VAR) or "").strip()
    if text:
        try:
            days.append(parse_day(text))
        except ValueError:
            return EndDate(None, problem=f"{ENV_VAR} should be a date such as 2026-12-15, not {text!r}.")
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("not an object")
            day = parse_day(data.get("expires") or "")
            days.append(day)
            contact = str(data.get("contact") or "").strip() or None
            built = str(data.get("built") or "").strip()
            copy = f"{day.isoformat()}, built {built}" if built else day.isoformat()
        except (OSError, ValueError):  # JSONDecodeError is a ValueError
            return EndDate(None, problem="This copy of Mind Ready cannot read its end date. Ask the person who shared it for a new copy.")
    return EndDate(min(days), contact, copy=copy) if days else None


def decide(end: EndDate | None, today: dt.date, latest_seen: dt.date | None) -> Verdict:
    """The verdict for `today` (the computer's clock) and the latest date this copy has seen."""
    if end is None:
        return Verdict()
    if end.last_day is None:
        return Verdict(expired=True, message=end.problem, reason="unreadable")
    who = _who(end.contact)
    seen = max(today, latest_seen) if latest_seen else today
    days_left = (end.last_day - seen).days
    over = f"This copy of Mind Ready stopped working on {long_date(end.last_day)}. Ask {who} for a new copy."
    if today > end.last_day:  # the clock itself says the last day is over
        return Verdict(end.last_day, True, days_left, over, "date")
    if latest_seen and (latest_seen - today).days > SLACK_DAYS:
        # Said with both dates, not as "stopped on <last day>": the later one may be the wrong one (a
        # clock set ahead while the copy ran), and then the message shows it.
        message = (
            f"This copy of Mind Ready stopped working: this computer says it is {long_date(today)}, "
            f"but it was already {long_date(latest_seen)}. Set the right date, or ask {who} for a new copy."
        )
        return Verdict(end.last_day, True, days_left, message, "clock")
    if seen > end.last_day:  # a day after the last one was seen; the clock is now at most 2 days behind it
        return Verdict(end.last_day, True, days_left, over, "date")
    return Verdict(end.last_day, False, days_left)


class Checker:
    """The verdict for one end date, with the clock and the remembered dates injectable (tests)."""

    def __init__(self, end: EndDate | None, today: Callable[[], dt.date] = dt.date.today, state_file: Path = STATE_FILE):
        self.end = end
        self.today = today
        self.state_file = state_file
        self._lock = threading.Lock()

    @property
    def copy(self) -> str:
        """The name this copy's latest date seen is kept under (EndDate.copy)."""
        return self.end.copy if self.end is not None else DEVELOPER_COPY

    def _every_copy(self) -> dict[str, Any]:
        """The latest date seen of every copy on this computer, {copy: "2026-11-09"}; {} when there is
        none yet or the file cannot be read."""
        try:
            data = json.loads(self.state_file.read_text(encoding="utf-8"))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def latest_seen(self) -> dt.date | None:
        """This copy's latest date remembered; None when there is none yet or it cannot be read (then the
        clock is believed: the next date seen is written over it)."""
        try:
            return parse_day(self._every_copy()[self.copy])
        except Exception:
            return None

    def _remember(self, day: dt.date) -> bool:
        """Write `day` as this copy's latest date seen, the other copies' kept (with self._lock held). A
        file that cannot be written is left as it is, never an error."""
        data = self._every_copy()
        data[self.copy] = day.isoformat()
        try:
            self.state_file.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.state_file.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
            os.replace(tmp, self.state_file)
        except OSError:
            return False
        return True

    def saw(self, day: dt.date) -> bool:
        """The clock's date: remembered when it is later than this copy's latest date seen (so the file
        is written at most once a day); True when it was."""
        with self._lock:
            latest = self.latest_seen()
            if latest is not None and day <= latest:
                return False
            return self._remember(day)

    def heard(self, day: dt.date) -> bool:
        """The internet's date: it moves the latest date seen forward as the clock does, and also puts
        right one more than SLACK_DAYS after it (a date that has not come yet: the clock was set ahead
        while the copy ran). Up to SLACK_DAYS after it is kept (time zones). True when it changed it."""
        with self._lock:
            latest = self.latest_seen()
            if latest is not None and day <= latest and (latest - day).days <= SLACK_DAYS:
                return False
            return self._remember(day)

    def verdict(self, ask: Callable[[], dt.date | None] | None = None) -> Verdict:
        """Decide now. A copy without an end date reads and writes nothing. With `ask` (the internet's
        date; the program's start gives it), a copy that would stop only because of the date it
        remembers asks it first: a remembered date that has not come yet is put right, and the copy
        decides again. No network: the remembered date stands."""
        if self.end is None or self.end.last_day is None:
            return decide(self.end, self.today(), None)
        today = self.today()
        latest = self.latest_seen()
        if latest is None or today > latest:
            self.saw(today)
        verdict = decide(self.end, today, latest)
        if ask is not None and verdict.expired and today <= self.end.last_day:  # the clock alone: still works
            try:
                day = ask()
            except Exception:
                day = None
            if day is not None and self.heard(day):
                verdict = decide(self.end, today, self.latest_seen())
        return verdict


# --- this program's checker, and the decision the server reuses for a minute ---------------------------
_default: Checker | None = None
_cached: tuple[float, Verdict] | None = None
_lock = threading.Lock()


def default() -> Checker:
    """This program's checker: its end date is read once (MIND_READY_EXPIRES, then EXPIRY)."""
    global _default
    with _lock:
        if _default is None:
            _default = Checker(read_end_date())
        return _default


def use(checker: Checker | None) -> None:
    """Replace this program's checker (tests); None reads the end date again at the next question."""
    global _default, _cached
    with _lock:
        _default, _cached = checker, None


def forget() -> None:
    """Decide again at the next question (the internet's date moved the latest date seen)."""
    global _cached
    with _lock:
        _cached = None


def current(now: Callable[[], float] = time.monotonic) -> Verdict:
    """The verdict, decided at most once a minute (the server asks on every API call)."""
    global _cached
    checker = default()
    moment = now()
    with _lock:
        if _cached is not None and 0 <= moment - _cached[0] < CACHE_SECONDS:
            return _cached[1]
    verdict = checker.verdict()
    with _lock:
        _cached = (moment, verdict)
    return verdict


def at_start(ask: Callable[[], dt.date | None] | None = None) -> Verdict:
    """The verdict as the program starts: decided now, then reused by the server for a minute. A copy
    that would stop only because of the date it remembers first asks the internet's date (`ask`, by
    default the gateway then Microsoft, 2 seconds each): only then does the start wait for it."""
    global _cached
    verdict = default().verdict(ask=ask or (lambda: internet_date(date_sites())))
    with _lock:
        _cached = (time.monotonic(), verdict)
    return verdict


# --- the internet's date ----------------------------------------------------------------------------------
def date_from_header(value: str | None) -> dt.date | None:
    """The local date of an HTTP Date header ("Mon, 09 Nov 2026 10:00:00 GMT"); None when it is not one."""
    if not value:
        return None
    try:
        moment = parsedate_to_datetime(value)
    except (TypeError, ValueError, IndexError, OverflowError):
        return None
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=dt.timezone.utc)
    day = moment.astimezone().date()
    return day if 2020 <= day.year <= 2100 else None


def internet_date(urls: Iterable[str], timeout: float = NETWORK_TIMEOUT, opener: Callable[..., Any] = urllib.request.urlopen) -> dt.date | None:
    """The date the first site that answers gives (a HEAD request, `timeout` seconds each); None
    without a network. A refusal (401, 404) still says the date."""
    for url in urls:
        try:
            request = urllib.request.Request(url, method="HEAD", headers={"User-Agent": "MindReady"})
            with opener(request, timeout=timeout) as response:
                header = response.headers.get("Date")
        except urllib.error.HTTPError as exc:
            header = exc.headers.get("Date") if exc.headers is not None else None
            with contextlib.suppress(Exception):
                exc.close()
        except Exception:
            continue
        day = date_from_header(header)
        if day is not None:
            return day
    return None


def date_sites() -> list[str]:
    """Where to ask the date: the AI gateway of settings.yaml (the app talks to it anyway), else Microsoft."""
    try:
        from .settings import SETTINGS

        gateway = (SETTINGS.ai_endpoint or "").strip()
    except Exception:  # a broken settings.yaml is reported where the app reads it, not here
        gateway = ""
    return [url for url in (gateway, FALLBACK_DATE_SITE) if url]


def start_network_check(checker: Checker | None = None, fetch: Callable[[], dt.date | None] | None = None) -> threading.Thread | None:
    """In the background: the internet's date moves the latest date seen forward, and puts right one
    that has not come yet (Checker.heard). Nothing for a copy without an end date (the developer
    version); nothing waits for it; no network, no change."""
    checker = checker or default()
    if checker.end is None or checker.end.last_day is None:
        return None
    get = fetch or (lambda: internet_date(date_sites()))

    def run() -> None:
        try:
            day = get()
            if day is not None and checker.heard(day):
                forget()  # the next API call decides again with it
        except Exception:
            pass

    thread = threading.Thread(target=run, name="mind-ready-date", daemon=True)
    thread.start()
    return thread
