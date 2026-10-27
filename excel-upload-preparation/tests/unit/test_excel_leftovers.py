"""Hidden Excels left behind by an app stopped abruptly are ended at the next start (app/excel_com.py)."""
from __future__ import annotations

import json
import subprocess
import sys
import time
from pathlib import Path

import pytest

from app import excel_com
from app.excel_com import _process, clean_leftover_excels, com_available, excel_pid, kill_own_excels, quit_excel, start_excel

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")
APP_DIR = Path(__file__).resolve().parents[2]


@pytest.fixture
def started(tmp_path, monkeypatch):
    path = tmp_path / "excel_started.json"
    monkeypatch.setattr(excel_com, "_started_path", lambda: path)
    return path


def _gone(pid: int, wait_s: float = 15) -> bool:
    end = time.time() + wait_s
    while time.time() < end:
        if _process(pid) is None:
            return True
        time.sleep(0.25)
    return False


@needs_excel
def test_an_excel_left_by_an_app_that_was_killed_is_ended_at_the_next_start(started):
    # an "app" that starts Excel and is then stopped abruptly (as closing its window does: no Quit)
    script = (
        "import os, sys, pythoncom; from pathlib import Path; sys.path.insert(0, sys.argv[1]);"
        "from app import excel_com; excel_com._started_path = lambda: Path(sys.argv[2]);"
        "pythoncom.CoInitialize(); x = excel_com.start_excel(mm_addin=False);"
        "print(excel_com.excel_pid(x), flush=True); os._exit(0)"
    )
    out = subprocess.run([sys.executable, "-c", script, str(APP_DIR), str(started)], capture_output=True, text=True, timeout=120)
    pid = int(out.stdout.strip())
    try:
        assert _process(pid) is not None, "the killed app's Excel should still be running"
        assert str(pid) in json.loads(started.read_text(encoding="utf-8"))
        assert clean_leftover_excels() == 1
        assert _gone(pid)
        assert json.loads(started.read_text(encoding="utf-8")) == {}
    finally:
        excel_com._terminate(pid)


@needs_excel
def test_an_excel_whose_app_still_runs_is_left_alone_and_quitting_forgets_it(started):
    import pythoncom

    pythoncom.CoInitialize()
    excel = start_excel(mm_addin=False)
    pid = excel_pid(excel)
    try:
        assert clean_leftover_excels() == 0  # this process -- its app -- still runs
        assert _process(pid) is not None
        quit_excel(excel)
        excel = None
        assert json.loads(started.read_text(encoding="utf-8")) == {}
    finally:
        if excel is not None:
            quit_excel(excel)
        excel_com._terminate(pid)


@needs_excel
def test_closing_the_window_ends_the_apps_own_excels(started):
    import pythoncom

    pythoncom.CoInitialize()
    excel = start_excel(mm_addin=False)
    pid = excel_pid(excel)
    try:
        assert kill_own_excels() == 1
        assert _gone(pid)
        assert json.loads(started.read_text(encoding="utf-8")) == {}
    finally:
        excel = None
        excel_com._terminate(pid)


def test_a_listed_id_now_used_by_another_program_is_never_ended(started):
    other = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
    try:
        created = _process(other.pid)[1]
        # listed as an Excel of an app that is gone, with this process's own creation time: a reused id
        started.write_text(json.dumps({str(other.pid): {"excel_created": created, "app": 999_999, "app_created": 1}}), encoding="utf-8")
        assert clean_leftover_excels() == 0  # not an excel.exe: never ended
        assert other.poll() is None
        assert json.loads(started.read_text(encoding="utf-8")) == {}  # the stale line is dropped
    finally:
        other.kill()
