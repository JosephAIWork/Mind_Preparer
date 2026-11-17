"""Hidden Excels left behind by an app stopped abruptly are ended at the next start (app/excel_com.py).

17 Nov 2026 (Cancel): kill_excel is the checked form -- it ends only an Excel THIS process started
and has not quit. The person's own Excel, an Excel of another Mind Ready window, an id Windows gave
to another program: never. The tests of that are the second half of this file."""
from __future__ import annotations

import json
import os
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


def _end_my_excel(pid: int, created: int | None) -> None:
    """A test's own Excel, should the test have failed before it was ended. 17 Nov 2026: only while
    the process with that id is still the very excel.exe the test started (same creation time) --
    these `finally` blocks ended whatever had the id."""
    found = _process(pid)
    if created is not None and found is not None and found[1] == created and found[0].lower().endswith("excel.exe"):
        excel_com._terminate(pid)


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
        clean_leftover_excels()  # its app is gone: ended while the list still names it, and only then


@needs_excel
def test_an_excel_whose_app_still_runs_is_left_alone_and_quitting_forgets_it(started):
    import pythoncom

    pythoncom.CoInitialize()
    excel = start_excel(mm_addin=False)
    pid = excel_pid(excel)
    created = (_process(pid) or (None, None))[1]
    try:
        assert clean_leftover_excels() == 0  # this process -- its app -- still runs
        assert _process(pid) is not None
        quit_excel(excel)
        excel = None
        assert json.loads(started.read_text(encoding="utf-8")) == {}
    finally:
        if excel is not None:
            quit_excel(excel)
        _end_my_excel(pid, created)


@needs_excel
def test_closing_the_window_ends_the_apps_own_excels(started):
    import pythoncom

    pythoncom.CoInitialize()
    excel = start_excel(mm_addin=False)
    pid = excel_pid(excel)
    created = (_process(pid) or (None, None))[1]
    try:
        assert kill_own_excels() == 1
        assert _gone(pid)
        assert json.loads(started.read_text(encoding="utf-8")) == {}
    finally:
        excel = None
        _end_my_excel(pid, created)


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


# --- 17 Nov 2026 (Cancel): kill_excel ends only an Excel this process started -------------------
EXCEL_EXE = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
ME = 77  # this process's creation time in the fakes below


@pytest.fixture
def ended(monkeypatch):
    """Every pid kill_excel really ends (the terminate is a fake: no process is touched)."""
    seen: list = []
    monkeypatch.setattr(excel_com, "_terminate", seen.append)
    return seen


def _fake_processes(monkeypatch, table: dict) -> None:
    """pid -> (program, creation time); this process itself is always there."""
    table = {os.getpid(): ("python.exe", ME), **table}
    monkeypatch.setattr(excel_com, "_process", lambda pid: table.get(int(pid)))


def _list(started: Path, entries: dict) -> None:
    started.write_text(json.dumps(entries), encoding="utf-8")


def _listed(started: Path) -> dict:
    return json.loads(started.read_text(encoding="utf-8")) if started.exists() else {}


def test_kill_excel_refuses_a_pid_that_is_not_listed(started, ended, monkeypatch):
    # the person's own Excel looks exactly like this: a running excel.exe that no list names
    _fake_processes(monkeypatch, {4242: (EXCEL_EXE, 1111), 4243: (EXCEL_EXE, 2222)})
    assert excel_com.kill_excel(4242) is False  # there is no list at all
    mine = {"4243": {"excel_created": 2222, "app": os.getpid(), "app_created": ME}}
    _list(started, mine)
    assert excel_com.kill_excel(4242) is False  # a list, and it is not in it
    assert excel_com.kill_excel(None) is False and excel_com.kill_excel(0) is False
    assert ended == []  # _terminate was never called
    assert _listed(started) == mine  # and the list is as it was


def test_kill_excel_refuses_an_excel_listed_for_another_process(started, ended, monkeypatch):
    # another Mind Ready window's Excel: listed, a real excel.exe, the same creation time -- not ours
    _fake_processes(monkeypatch, {4242: (EXCEL_EXE, 1111), os.getpid() + 1: ("python.exe", 5)})
    theirs = {"4242": {"excel_created": 1111, "app": os.getpid() + 1, "app_created": 5}}
    _list(started, theirs)
    assert excel_com.kill_excel(4242) is False
    assert ended == [] and _listed(started) == theirs  # still listed: that window quits it itself


def test_kill_excel_refuses_a_line_left_by_an_earlier_app_that_had_this_process_id(started, ended, monkeypatch):
    _fake_processes(monkeypatch, {4242: (EXCEL_EXE, 1111)})
    old = {"4242": {"excel_created": 1111, "app": os.getpid(), "app_created": ME - 1}}  # same id, another process
    _list(started, old)
    assert excel_com.kill_excel(4242) is False
    assert ended == [] and _listed(started) == old  # the next start cleans it (clean_leftover_excels)


def test_kill_excel_refuses_a_process_that_is_not_that_excel_any_more(started, ended, monkeypatch):
    def entry() -> dict:
        return {"4242": {"excel_created": 1111, "app": os.getpid(), "app_created": ME}}

    # the creation time differs: our Excel is gone and Windows gave its id to a new process -- an
    # excel.exe even (the person opened a workbook)
    _fake_processes(monkeypatch, {4242: (EXCEL_EXE, 9999)})
    _list(started, entry())
    assert excel_com.kill_excel(4242) is False and ended == []
    assert _listed(started) == {}  # the stale line is dropped
    # the same creation time, but not an excel.exe
    _fake_processes(monkeypatch, {4242: (r"C:\Windows\notepad.exe", 1111)})
    _list(started, entry())
    assert excel_com.kill_excel(4242) is False and ended == [] and _listed(started) == {}
    # no such process any more
    _fake_processes(monkeypatch, {})
    _list(started, entry())
    assert excel_com.kill_excel(4242) is False and ended == [] and _listed(started) == {}


def test_kill_excel_ends_an_excel_this_process_started_and_only_once(started, ended, monkeypatch):
    from app import cancel

    _fake_processes(monkeypatch, {4242: (EXCEL_EXE, 1111), 4243: (EXCEL_EXE, 2222)})
    excel_com._remember(4242)  # as start_excel and the import macro's _new_excel write it down
    excel_com._remember(4243)
    assert _listed(started)["4242"] == {"excel_created": 1111, "app": os.getpid(), "app_created": ME}
    job = cancel.begin("s-kill", "numbers")
    try:
        with cancel.bind(job):
            cancel.note_excel(4242)
        assert excel_com.kill_excel(4242) is True
        assert ended == [4242] and list(_listed(started)) == ["4243"]  # ended, forgotten; the other one untouched
        assert job.pids == set()  # and no job ends it again
        assert excel_com.kill_excel(4242) is False  # a second kill finds no line: the id may be anyone's by now
        assert ended == [4242]
    finally:
        cancel.end(job)


def test_kill_excel_leaves_a_real_program_that_holds_a_listed_id(started):
    # no fakes: a real process (not an Excel) under an id the list names for this very process
    other = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
    try:
        created = _process(other.pid)[1]
        me = _process(os.getpid())[1]
        _list(started, {str(other.pid): {"excel_created": created, "app": os.getpid(), "app_created": me}})
        assert excel_com.kill_excel(other.pid) is False
        time.sleep(0.5)
        assert other.poll() is None  # still running
        assert _listed(started) == {}
    finally:
        other.kill()


def test_only_the_three_checked_paths_ever_terminate_a_process():
    """No code of the app ends a process except through kill_excel, clean_leftover_excels and
    kill_own_excels (each checks the list and the process first)."""
    import ast

    callers: dict = {}
    for path in sorted((APP_DIR / "app").rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for fn in ast.walk(tree):
            if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            for node in ast.walk(fn):
                func = getattr(node, "func", None)
                name = getattr(func, "id", None) or getattr(func, "attr", None)
                if isinstance(node, ast.Call) and name in ("_terminate", "TerminateProcess"):
                    callers.setdefault(name, set()).add(f"{path.relative_to(APP_DIR).as_posix()}:{fn.name}")
    assert callers == {
        "_terminate": {"app/excel_com.py:kill_excel", "app/excel_com.py:clean_leftover_excels", "app/excel_com.py:kill_own_excels"},
        "TerminateProcess": {"app/excel_com.py:_terminate"},
    }
    # ... and nothing ends a program by its name, or reaches for an Excel that is already open
    for path in sorted((APP_DIR / "app").rglob("*.py")):
        code = path.read_text(encoding="utf-8").lower()
        for word in ("taskkill", "stop-process", "getactiveobject", "os.kill(", 'dispatch("excel', "dispatch('excel", "ensuredispatch"):
            assert word not in code, f"{path.name}: {word}"


@needs_excel
def test_kill_excel_ends_the_apps_own_real_excel(started):
    import pythoncom

    pythoncom.CoInitialize()
    excel = start_excel(mm_addin=False)
    pid = excel_pid(excel)
    created = (_process(pid) or (None, None))[1]
    try:
        assert str(pid) in _listed(started)
        excel = None
        assert excel_com.kill_excel(pid) is True
        assert _gone(pid) and _listed(started) == {}
        assert excel_com.kill_excel(pid) is False
    finally:
        excel = None
        _end_my_excel(pid, created)


@needs_excel
def test_kill_excel_never_ends_a_real_excel_this_process_did_not_start(started, tmp_path, monkeypatch):
    # an Excel of "another app" (a second process, with a list of its own), running while we try
    their_list = tmp_path / "their_excel_started.json"
    script = (
        "import sys, pythoncom; from pathlib import Path; sys.path.insert(0, sys.argv[1]);"
        "from app import excel_com; excel_com._started_path = lambda: Path(sys.argv[2]);"
        "pythoncom.CoInitialize(); x = excel_com.start_excel(mm_addin=False);"
        "print(excel_com.excel_pid(x), flush=True); sys.stdin.readline(); excel_com.quit_excel(x)"
    )
    app = subprocess.Popen([sys.executable, "-c", script, str(APP_DIR), str(their_list)], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    pid = None
    try:
        pid = int(app.stdout.readline().strip())
        assert _process(pid) is not None and _process(pid)[0].lower().endswith("excel.exe")
        # 1. our list does not name it (as it never names the person's own Excel)
        assert excel_com.kill_excel(pid) is False
        # 2. our list names it, for the process that started it
        _list(started, json.loads(their_list.read_text(encoding="utf-8")))
        assert excel_com.kill_excel(pid) is False
        assert kill_own_excels() == 0 and clean_leftover_excels() == 0  # not ours, and its app still runs
        time.sleep(1.0)
        assert _process(pid) is not None, "an Excel this process did not start was ended"
        assert str(pid) in _listed(started)
    finally:
        try:
            app.stdin.write("\n")
            app.stdin.flush()
            app.wait(60)  # it quits its own Excel
        except Exception:
            app.kill()
            app.wait(30)
        if pid is not None and not _gone(pid, 10):
            monkeypatch.setattr(excel_com, "_started_path", lambda: their_list)
            clean_leftover_excels()  # its app is gone: the leftover path, checked as always


@needs_excel
def test_the_import_macros_real_excel_is_listed_while_it_runs_and_forgotten_when_closed(started):
    import pythoncom

    from app import outside_inputs

    pythoncom.CoInitialize()
    xl, pid = outside_inputs._new_excel(macros=False)
    created = (_process(pid) or (None, None))[1]
    try:
        assert _listed(started)[str(pid)]["app"] == os.getpid()  # 17 Nov 2026: it was in no list before
        outside_inputs._close_excel(xl, None, pid)
        xl = None
        assert _gone(pid) and _listed(started) == {}
    finally:
        xl = None
        _end_my_excel(pid, created)


@needs_excel
def test_a_cancel_ends_the_real_excel_its_job_started_and_the_job_starts_no_other(started):
    import pythoncom

    from app import cancel

    pythoncom.CoInitialize()
    job = cancel.begin("s-real-excel", "numbers")
    excel = pid = created = None
    try:
        with cancel.bind(job):
            excel = start_excel(mm_addin=False)
            pid = excel_pid(excel)
            created = (_process(pid) or (None, None))[1]
            assert job.pids == {pid} and str(pid) in _listed(started)
            excel = None
            assert cancel.request("s-real-excel", "numbers")["flagged"] == [job]
            assert _gone(pid) and _listed(started) == {} and job.pids == set()
            with pytest.raises(cancel.Cancelled):
                start_excel(mm_addin=False)  # cancelled: no Excel is started any more
    finally:
        cancel.end(job, "cancelled")
        excel = None
        if pid is not None:
            _end_my_excel(pid, created)
