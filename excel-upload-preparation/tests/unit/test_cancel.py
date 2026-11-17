"""Cancel, the core (17 Nov 2026): the registry of running jobs (app/cancel.py), and the places that
read the flag without anyone asking (DESIGN.md R3): start_excel, the import macro's Excel, the model
call, the wait of the Mind client, the release of the session's Excel worker.

No Excel and no network here: the kill, DispatchEx and requests.post are fakes. The checked kill
itself is tested in test_excel_leftovers.py, the worker in test_excel_worker.py, the routes in
test_cancel_api.py."""
from __future__ import annotations

import ast
import inspect
import threading
import time

import pytest

from app import cancel, excel_com, excel_worker


@pytest.fixture(autouse=True)
def _empty_registry():
    cancel._JOBS.clear()
    yield
    left = {sid: list(mine) for sid, mine in cancel._JOBS.items()}
    cancel._JOBS.clear()
    assert left == {}, f"a test left jobs registered: {left}"


@pytest.fixture
def kills(monkeypatch):
    """Every pid a Cancel asks excel_com to end (never a real process: the kill is a fake)."""
    seen: list = []
    monkeypatch.setattr(excel_com, "kill_excel", lambda pid: bool(seen.append(pid)))
    return seen


@pytest.fixture
def worker_calls(monkeypatch):
    """What a Cancel tells the session's Excel worker: (sid, jobs, free)."""
    seen: list = []

    def fake(sid, jobs, free=False):
        seen.append((sid, list(jobs), free))
        return 0

    monkeypatch.setattr(excel_worker, "cancel", fake)
    return seen


# --- the registry -----------------------------------------------------------------------------------
def test_the_module_imports_nothing_from_the_app_at_the_top():
    top = [n for n in ast.parse(inspect.getsource(cancel)).body if isinstance(n, (ast.Import, ast.ImportFrom))]
    assert top, "the imports moved: look again"
    for node in top:
        if isinstance(node, ast.ImportFrom):
            assert node.level == 0 and not (node.module or "").startswith("app"), ast.dump(node)
        else:
            assert not any(a.name.startswith("app") for a in node.names), ast.dump(node)


def test_nothing_running(kills):
    # the real excel_worker.cancel: no worker for this workbook, nothing to stop
    assert cancel.request("s1") == {"flagged": [], "late": [], "ahead": False}
    assert cancel.request("s1", "apply") == {"flagged": [], "late": [], "ahead": False}
    assert kills == [] and cancel.jobs("s1") == [] and cancel.changing("s1") is None
    assert cancel.current() is None and cancel.flagged() is False
    cancel.check()  # no job: nothing
    cancel.note_excel(123)  # no job: nothing
    cancel.end(None)  # nothing to end


def test_a_cancel_flags_the_job_ends_its_excels_and_tells_the_worker(kills, worker_calls):
    told: list = []
    job = cancel.begin("s1", "apply", changes=True, on_flag=lambda: told.append(cancel.jobs("s1")))
    assert (job.sid, job.kind, job.changes, job.committed, job.outcome) == ("s1", "apply", True, False, None)
    assert job.started <= time.time() and not job.flag.is_set() and not job.ended.is_set()
    assert cancel.jobs("s1") == [job] and cancel.changing("s1") is job
    with cancel.bind(job):
        cancel.note_excel(111)
        cancel.note_excel(222)
        cancel.note_excel(None)
        assert job.pids == {111, 222}
        cancel.check()  # not cancelled yet
        assert not cancel.flagged()
        out = cancel.request("s1")
        assert out == {"flagged": [job], "late": [], "ahead": False}
        assert cancel.flagged()
        with pytest.raises(cancel.Cancelled):
            cancel.check()
        with pytest.raises(cancel.Cancelled):
            job.check()
    assert sorted(kills) == [111, 222]
    assert told == [[job]]  # called once, under the lock (which it may take again), the job still registered
    assert worker_calls == [("s1", [job], True)]  # no kind: the recalculation made ahead goes too
    assert cancel.jobs("s1") == [job]  # registered until its end (its restore comes first)
    cancel.end(job, "cancelled")
    assert job.outcome == "cancelled" and job.ended.is_set() and cancel.jobs("s1") == [] and cancel.changing("s1") is None


def test_cancelled_is_not_an_exception():
    # the app's many `except Exception` must not swallow it
    assert issubclass(cancel.Cancelled, BaseException) and not issubclass(cancel.Cancelled, Exception)
    assert issubclass(cancel.Busy, Exception)
    try:
        try:
            raise cancel.Cancelled()
        except Exception:  # noqa: BLE001
            pytest.fail("an `except Exception` caught a cancel")
    except cancel.Cancelled as exc:
        assert str(exc) == "Stopped."


def test_a_double_request_does_the_same_harmless_things(kills, worker_calls):
    told: list = []
    job = cancel.begin("s1", "autofix", changes=True, on_flag=lambda: told.append(1))
    with cancel.bind(job):
        cancel.note_excel(111)
    first = cancel.request("s1", "autofix")
    second = cancel.request("s1", "autofix")
    assert first["flagged"] == [job] and second["flagged"] == [job] and first["late"] == second["late"] == []
    assert told == [1]  # the status line is written once
    assert kills == [111]  # the second call finds no Excel left to end
    assert worker_calls == [("s1", [job], False), ("s1", [job], False)]
    cancel.end(job, "cancelled")
    cancel.end(job, "done")  # a second end changes nothing
    assert job.outcome == "cancelled"
    assert cancel.request("s1") == {"flagged": [], "late": [], "ahead": False}  # after the job has ended
    assert told == [1]


def test_a_commit_before_the_cancel_wins(kills, worker_calls):
    told: list = []
    job = cancel.begin("s1", "recalc", on_flag=lambda: told.append(1))
    with cancel.bind(job):
        cancel.note_excel(111)
    job.commit()
    job.commit()  # a second call does nothing
    out = cancel.request("s1")
    assert out["flagged"] == [] and out["late"] == [job]  # too late: its result stands
    assert job.committed and not job.flag.is_set() and told == [] and kills == []
    assert worker_calls == [("s1", [], True)]  # none of its items is touched
    with cancel.bind(job):
        cancel.check()
        assert not cancel.flagged()
    job.commit()  # still nothing
    cancel.end(job)
    assert job.outcome == "done"


def test_a_cancel_before_the_commit_wins(kills, worker_calls):
    job = cancel.begin("s1", "recalc")
    assert cancel.request("s1", "recalc")["flagged"] == [job]
    with pytest.raises(cancel.Cancelled):
        job.commit()
    assert not job.committed
    assert cancel.request("s1")["late"] == []  # flagged, not late: still on its way out
    cancel.end(job, "cancelled")


def test_of_a_commit_and_a_cancel_at_the_same_moment_exactly_one_wins(kills, worker_calls):
    for i in range(150):
        job = cancel.begin(f"race-{i}", "apply", changes=True)
        gate = threading.Barrier(2)
        seen: dict = {}

        def commit() -> None:
            gate.wait()
            try:
                job.commit()
                seen["commit"] = "committed"
            except cancel.Cancelled:
                seen["commit"] = "cancelled"

        def press() -> None:
            gate.wait()
            seen["request"] = cancel.request(job.sid)

        threads = [threading.Thread(target=commit), threading.Thread(target=press)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(5)
        if seen["commit"] == "committed":
            assert job.committed and not job.flag.is_set() and seen["request"]["late"] == [job] and seen["request"]["flagged"] == []
        else:
            assert not job.committed and job.flag.is_set() and seen["request"]["flagged"] == [job] and seen["request"]["late"] == []
        cancel.end(job)


def test_a_kind_cancels_only_that_kind_of_that_workbook(kills, worker_calls):
    apply_job = cancel.begin("s1", "apply", changes=True)
    recalc_job = cancel.begin("s1", "recalc")
    elsewhere = cancel.begin("s2", "recalc")
    assert cancel.jobs("s1") == [apply_job, recalc_job]  # oldest first
    out = cancel.request("s1", "recalc")
    assert out["flagged"] == [recalc_job] and out["late"] == []
    assert not apply_job.flag.is_set() and not elsewhere.flag.is_set()
    assert cancel.request("s1", "apply")["flagged"] == [apply_job]
    assert cancel.request("s1", "numbers") == {"flagged": [], "late": [], "ahead": False}
    # the recalculation made ahead belongs to no job: it goes with a Recalculate (and with everything), not with an Apply
    assert [(c[0], c[2]) for c in worker_calls] == [("s1", True), ("s1", False), ("s1", False)]
    assert not elsewhere.flag.is_set()
    for job in (apply_job, recalc_job, elsewhere):
        cancel.end(job)


def test_a_list_of_kinds_cancels_only_those_kinds(kills, worker_calls):
    # 17 Nov 2026 (R1-5): Start's Cancel names the kinds of its own steps; a Run in Mind and an
    # assistant's answer started elsewhere go on
    steps = ["outside", "autofix", "prep", "apply", "recalc", "scan"]
    apply_job = cancel.begin("s1", "apply", changes=True)
    mind = cancel.begin("s1", "mind")
    chat = cancel.begin("s1", "chat")
    out = cancel.request("s1", steps)
    assert out["flagged"] == [apply_job] and not mind.flag.is_set() and not chat.flag.is_set()
    assert worker_calls[-1][2] is True  # "recalc" among the kinds: the recalculation made ahead goes too
    assert cancel.request("s1", ("outside", "prep"))["flagged"] == [] and worker_calls[-1][2] is False
    assert cancel.request("s1", [])["flagged"] == []  # an empty list stops nothing
    assert not mind.flag.is_set() and not chat.flag.is_set()
    for job in (apply_job, mind, chat):
        cancel.end(job)


def test_a_commit_alone_waits_for_no_job_that_changes_the_workbook_and_holds_them_off_until_its_end(kills, worker_calls):
    # 17 Nov 2026 (R1-2): a Recalculate, Suggest names or the reports write what a cancelled Apply's restore puts back
    writer = cancel.begin("s1", "names")
    apply_job = cancel.begin("s1", "apply", changes=True)  # began while the writer waited for the model
    with pytest.raises(cancel.Busy) as refused:
        writer.commit(alone=True)
    assert refused.value.kind == "apply" and not writer.committed
    cancel.end(apply_job)
    writer.commit(alone=True)
    assert writer.committed
    with pytest.raises(cancel.Busy) as held:  # it writes now: no job that changes the workbook begins
        cancel.begin("s1", "prep", changes=True)
    assert held.value.kind == "names"
    reader = cancel.begin("s1", "numbers")  # a job that changes nothing is never held off
    cancel.end(writer)
    again = cancel.begin("s1", "prep", changes=True)
    plain = cancel.begin("s1", "recalc")
    plain.commit()  # a plain gate holds nothing off
    cancel.end(again)
    for job in (reader, plain):
        cancel.end(job)


def test_ahead_says_the_worker_stopped_an_item_of_no_job(kills, monkeypatch):
    monkeypatch.setattr(excel_worker, "cancel", lambda sid, jobs, free=False: 1 if free else 0)
    assert cancel.request("s1") == {"flagged": [], "late": [], "ahead": True}  # only a recalculation made ahead
    assert cancel.request("s1", "recalc")["ahead"] is True
    assert cancel.request("s1", "apply")["ahead"] is False


def test_one_job_that_changes_the_workbook_at_a_time(kills, worker_calls):
    first = cancel.begin("s1", "apply", changes=True)
    with pytest.raises(cancel.Busy) as refused:
        cancel.begin("s1", "scan", changes=True)
    assert refused.value.kind == "apply"  # the kind of the job that is in the way: the route's words come from it
    reader = cancel.begin("s1", "recalc")  # it changes nothing along the way: never refused
    elsewhere = cancel.begin("s2", "scan", changes=True)  # another workbook
    assert cancel.changing("s1") is first and cancel.changing("s2") is elsewhere
    assert cancel.jobs("s1") == [first, reader]  # the refused one was never registered
    cancel.request("s1", "apply")
    with pytest.raises(cancel.Busy):  # cancelled, but still putting the workbook back: not free yet
        cancel.begin("s1", "prep", changes=True)
    cancel.end(first, "cancelled")
    again = cancel.begin("s1", "prep", changes=True)
    assert cancel.changing("s1") is again
    for job in (reader, elsewhere, again):
        cancel.end(job)


def test_a_thread_without_bind_sees_no_job(kills, worker_calls):
    job = cancel.begin("s1", "scan", changes=True)
    seen: dict = {}

    def elsewhere() -> None:
        seen["current"] = cancel.current()
        seen["flagged"] = cancel.flagged()
        cancel.check()  # no job: nothing, though the workbook's job is cancelled
        cancel.note_excel(555)  # no job: kept by nobody, ended by nobody
        seen["through"] = True

    with cancel.bind(job) as bound:
        assert bound is job and cancel.current() is job
        cancel.request("s1")
        t = threading.Thread(target=elsewhere)  # a new thread starts with an empty context
        t.start()
        t.join(5)
        with cancel.bind(None):  # binds nothing: the thread still works for its job
            assert cancel.current() is job
        assert cancel.current() is job
    assert cancel.current() is None
    assert seen == {"current": None, "flagged": False, "through": True}
    assert kills == [] and job.pids == set()
    cancel.end(job, "cancelled")


def test_bind_puts_the_job_before_it_back():
    outer, inner = cancel.begin("s1", "mind"), cancel.begin("s1", "apply", changes=True)
    with cancel.bind(outer):
        with cancel.bind(inner):
            assert cancel.current() is inner
        assert cancel.current() is outer
        with pytest.raises(RuntimeError), cancel.bind(inner):
            raise RuntimeError("the work failed")
        assert cancel.current() is outer  # also after an error
    assert cancel.current() is None
    cancel.end(outer)
    cancel.end(inner)


def test_an_excel_noted_after_the_cancel_is_ended_at_once(kills, worker_calls):
    job = cancel.begin("s1", "numbers")
    with cancel.bind(job):
        cancel.request("s1")
        with pytest.raises(cancel.Cancelled):  # the Cancel came while that Excel was starting
            cancel.note_excel(4321)
    assert kills == [4321] and job.pids == set()
    cancel.end(job, "cancelled")


def test_an_excel_that_was_quit_is_not_ended_by_a_later_cancel(kills, worker_calls):
    loop, recalc = cancel.begin("s1", "mind"), cancel.begin("s1", "recalc")
    with cancel.bind(loop):
        cancel.note_excel(10)
        cancel.note_excel(11)
    with cancel.bind(recalc):
        cancel.forget_excel(10)  # quit on a thread that works for another job
    cancel.forget_excel(999)  # unknown: nothing
    cancel.forget_excel(None)
    assert loop.pids == {11}
    cancel.request("s1")
    assert kills == [11]
    cancel.end(loop, "cancelled")
    cancel.end(recalc, "cancelled")


def test_an_on_flag_that_fails_never_holds_a_cancel(kills, worker_calls):
    def broken() -> None:
        raise RuntimeError("the status line could not be written")

    job = cancel.begin("s1", "mind", on_flag=broken)
    assert cancel.request("s1", "mind")["flagged"] == [job] and job.flag.is_set()
    cancel.end(job, "cancelled")


def test_call_hands_back_the_value_and_the_error(kills, worker_calls):
    me = threading.current_thread().name
    assert cancel.call(lambda: threading.current_thread().name) == me  # no job: a plain call, on this thread
    job = cancel.begin("s1", "names")
    with cancel.bind(job):
        assert cancel.call(lambda: threading.current_thread().name) == "cancel-call-names"
        assert cancel.call(cancel.current) is job  # the helper sees what the caller sees
        assert cancel.call(lambda: None) is None
        with pytest.raises(ZeroDivisionError):  # an error of the call reaches the caller as before
            cancel.call(lambda: 1 / 0)
    cancel.end(job)


def test_call_leaves_on_a_cancel_and_the_helper_ends_alone(kills, worker_calls):
    let_go, finished = threading.Event(), threading.Event()

    def slow() -> str:
        let_go.wait(20)
        finished.set()
        return "an answer nobody reads"

    job = cancel.begin("s1", "names")
    threading.Timer(0.3, lambda: cancel.request("s1", "names")).start()
    began = time.time()
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        cancel.call(slow)
    assert time.time() - began < 1.5 and not finished.is_set()  # it left; the request is still on its way
    let_go.set()
    assert finished.wait(5)  # ... and ends alone
    ran: list = []
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        cancel.call(lambda: ran.append(1))  # already cancelled: the next call is not even made
    assert ran == []
    cancel.end(job, "cancelled")


# --- R3: where the flag is read without anyone asking -------------------------------------------------
class _FakeExcel:
    """What start_excel and _new_excel touch of an Excel, with nothing behind it."""
    Hwnd = 1

    def __init__(self) -> None:
        self.quit = 0

    def Quit(self) -> None:  # noqa: N802 -- Excel's own name
        self.quit += 1


def _fake_dispatch(monkeypatch, on_start=None):
    client = pytest.importorskip("win32com.client")
    made: list = []

    def dispatch(name):
        if on_start is not None:
            on_start()
        made.append(_FakeExcel())
        return made[-1]

    monkeypatch.setattr(client, "DispatchEx", dispatch)
    return made


def test_start_excel_starts_nothing_for_a_cancelled_job(monkeypatch, kills, worker_calls):
    made = _fake_dispatch(monkeypatch)
    job = cancel.begin("s1", "numbers")
    cancel.request("s1")
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        excel_com.start_excel(mm_addin=False)
    assert made == [] and kills == []
    cancel.end(job, "cancelled")


def test_start_excel_gives_its_excel_to_the_job_and_quitting_takes_it_back(monkeypatch, kills, worker_calls):
    made = _fake_dispatch(monkeypatch)
    written: list = []
    monkeypatch.setattr(excel_com, "excel_pid", lambda excel: 4242)
    monkeypatch.setattr(excel_com, "_remember", lambda pid: written.append(("remember", pid)))
    monkeypatch.setattr(excel_com, "_forget", lambda pid: written.append(("forget", pid)))
    job = cancel.begin("s1", "numbers")
    with cancel.bind(job):
        excel = excel_com.start_excel(mm_addin=False)
    assert excel is made[0] and job.pids == {4242} and written == [("remember", 4242)]
    excel_com.quit_excel(excel)  # on a thread that works for no job
    assert made[0].quit == 1 and job.pids == set() and written == [("remember", 4242), ("forget", 4242)]
    cancel.request("s1")
    assert kills == []  # quit: a Cancel has nothing to end
    cancel.end(job, "cancelled")
    assert excel_com.start_excel(mm_addin=False) is made[1]  # no job: as before


def test_a_cancel_while_excel_starts_ends_that_excel(monkeypatch, kills, worker_calls):
    made = _fake_dispatch(monkeypatch, on_start=lambda: cancel.request("s1"))  # the Cancel lands during DispatchEx
    written: list = []
    monkeypatch.setattr(excel_com, "excel_pid", lambda excel: 4242)
    monkeypatch.setattr(excel_com, "_remember", lambda pid: written.append(pid))
    job = cancel.begin("s1", "numbers")
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        excel_com.start_excel(mm_addin=False)
    assert len(made) == 1 and written == [4242]  # written down first: the checked kill ends only what is listed
    assert kills == [4242] and job.pids == set()
    assert made[0].quit == 1  # 17 Nov 2026 (R1-4): quit through its own object too, should that kill have found nothing
    cancel.end(job, "cancelled")


def test_the_import_macros_excel_is_written_down_given_to_the_job_and_forgotten(monkeypatch, kills, worker_calls):
    from app import outside_inputs

    win32process = pytest.importorskip("win32process")
    made = _fake_dispatch(monkeypatch)
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda hwnd: (0, 5151))
    written: list = []
    monkeypatch.setattr(excel_com, "_remember", lambda pid: written.append(("remember", pid)))
    monkeypatch.setattr(excel_com, "_forget", lambda pid: written.append(("forget", pid)))
    job = cancel.begin("s1", "outside", changes=True)
    with cancel.bind(job):
        xl, pid = outside_inputs._new_excel(macros=True)
    assert xl is made[0] and pid == 5151 and xl.AutomationSecurity == 1
    assert written == [("remember", 5151)] and job.pids == {5151}  # before tonight it was in no list
    monkeypatch.setattr(outside_inputs, "_pid_alive", lambda pid: False)  # it quit
    outside_inputs._close_excel(xl, None, pid)
    assert made[0].quit == 1 and written[-1] == ("forget", 5151) and job.pids == set() and kills == []
    cancel.end(job)


def test_the_import_macros_excel_that_does_not_quit_is_ended_through_the_checked_kill(monkeypatch, kills, worker_calls):
    from app import outside_inputs

    monkeypatch.setattr(outside_inputs, "_pid_alive", lambda pid: True)  # it never goes
    monkeypatch.setattr(outside_inputs.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_com, "_forget", lambda pid: None)
    outside_inputs._close_excel(_FakeExcel(), None, 5151)
    assert kills == [5151]  # excel_com.kill_excel, which ends only an Excel this app started
    outside_inputs._close_excel(None, None, None)  # nothing was started: nothing to end
    assert kills == [5151]


def test_the_import_macros_excel_is_not_started_for_a_cancelled_job_and_is_ended_when_the_cancel_comes_meanwhile(monkeypatch, kills, worker_calls):
    from app import outside_inputs

    win32process = pytest.importorskip("win32process")
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda hwnd: (0, 5151))
    monkeypatch.setattr(excel_com, "_remember", lambda pid: None)
    late = cancel.begin("s1", "outside", changes=True)
    made = _fake_dispatch(monkeypatch, on_start=lambda: cancel.request("s1"))
    with cancel.bind(late), pytest.raises(cancel.Cancelled):
        outside_inputs._new_excel(macros=False)
    assert len(made) == 1 and kills == [5151]  # it had started: ended by its pid
    assert made[0].quit == 1  # 17 Nov 2026 (R1-4): and quit through its own object (nobody else holds it)
    with cancel.bind(late), pytest.raises(cancel.Cancelled):
        outside_inputs._new_excel(macros=False)
    assert len(made) == 1  # cancelled already: no second Excel
    cancel.end(late, "cancelled")


class _Reply:
    status_code = 200
    text = ""
    headers = {"content-type": "application/json"}

    def __init__(self) -> None:
        self.closed = False

    def raise_for_status(self) -> None:
        pass

    def json(self) -> dict:
        return {"content": [{"type": "text", "text": "the answer"}], "stop_reason": "end_turn", "usage": {}}

    def close(self) -> None:
        self.closed = True


@pytest.fixture
def model(monkeypatch):
    from app import llm

    pytest.importorskip("requests")
    monkeypatch.setattr(llm, "find_api_key", lambda: ("a-key", "https://gateway.invalid"))
    monkeypatch.setenv("MIND_READY_STREAM", "0")
    return llm


def test_a_model_call_is_abandoned_on_a_cancel(model, monkeypatch, kills, worker_calls):
    let_go, sent, reply = threading.Event(), threading.Event(), _Reply()

    def post(url, **kwargs):
        sent.set()
        let_go.wait(20)
        return reply

    monkeypatch.setattr("requests.post", post)
    job = cancel.begin("s1", "names")
    threading.Timer(0.3, lambda: sent.wait(5) and cancel.request("s1", "names")).start()
    began = time.time()
    with cancel.bind(job), pytest.raises(cancel.Cancelled):  # not {"message": "Request failed: ..."}
        model.chat_completion([{"role": "user", "content": "hello"}], "be brief")
    assert time.time() - began < 2
    let_go.set()
    for _ in range(50):
        if reply.closed:
            break
        time.sleep(0.1)
    assert reply.closed  # the answer that came after the caller left is closed, unread
    calls: list = []
    monkeypatch.setattr("requests.post", lambda url, **kwargs: calls.append(url) or _Reply())
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        model.chat_completion([{"role": "user", "content": "hello"}], "be brief")
    assert calls == []  # cancelled before the request: it is not sent
    cancel.end(job, "cancelled")


def test_a_model_call_answers_as_before_under_a_job_and_without_one(model, monkeypatch, kills, worker_calls):
    monkeypatch.setattr("requests.post", lambda url, **kwargs: _Reply())
    plain = model.chat_completion([{"role": "user", "content": "hello"}], "be brief")
    job = cancel.begin("s1", "chat")
    pieces: list = []
    with cancel.bind(job):
        bound = model.chat_completion([{"role": "user", "content": "hello"}], "be brief", on_delta=pieces.append)
    assert plain == bound == {"available": True, "text": "the answer", "message": None} and pieces == ["the answer"]

    def refused(url, **kwargs):
        raise ConnectionError("no route to the gateway")

    monkeypatch.setattr("requests.post", refused)
    with cancel.bind(job):  # an error of the call is still an answer, not an exception
        failed = model.chat_completion([{"role": "user", "content": "hello"}], "be brief")
    assert failed["text"] is None and "no route to the gateway" in failed["message"]
    cancel.end(job)


def test_a_cancelled_job_takes_no_more_of_the_answer(model, monkeypatch, kills, worker_calls):
    job = cancel.begin("s1", "chat")

    def post(url, **kwargs):
        cancel.request("s1", "chat")  # the Cancel lands just as the answer arrives
        return _Reply()

    monkeypatch.setattr("requests.post", post)
    pieces: list = []
    with cancel.bind(job), pytest.raises(cancel.Cancelled):
        model.chat_completion([{"role": "user", "content": "hello"}], "be brief", on_delta=pieces.append)
    assert pieces == []  # the text never reached the screen's stream
    cancel.end(job, "cancelled")


def test_the_mind_clients_wait_leaves_on_a_cancel(kills, worker_calls):
    pytest.importorskip("playwright.sync_api")
    from app import mind_client

    class Page:
        waited: list = []

        def wait_for_timeout(self, ms: int) -> None:
            self.waited.append(ms)

    page = Page()
    mind_client._wait(page, 5)  # no job: as before
    job = cancel.begin("s1", "mind")
    with cancel.bind(job):
        mind_client._wait(page, 6)
        cancel.request("s1", "mind")
        with pytest.raises(cancel.Cancelled):
            mind_client._wait(page, 7)  # cancelled: it does not even wait
    assert page.waited == [5, 6]

    class Slow(Page):
        def wait_for_timeout(self, ms: int) -> None:
            cancel.request("s2", "mind")  # the Cancel lands during the wait

    late = cancel.begin("s2", "mind")
    with cancel.bind(late), pytest.raises(cancel.Cancelled):
        mind_client._wait(Slow(), 5)  # ... and the wait does not return into the next click
    cancel.end(job, "cancelled")
    cancel.end(late, "cancelled")


def test_releasing_the_sessions_excel_leaves_on_a_cancel(kills, worker_calls):
    excel_worker.release("s1")  # no worker, no job: nothing
    job = cancel.begin("s1", "numbers")
    with cancel.bind(job):
        excel_worker.release("s1")
        cancel.request("s1")
        with pytest.raises(cancel.Cancelled):  # so the Excel the caller would start next never starts
            excel_worker.release("s1")
    cancel.end(job, "cancelled")
