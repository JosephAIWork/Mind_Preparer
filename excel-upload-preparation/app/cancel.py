"""Cancel:  one registry  of the jobs that run for a workbook, one flag, one exception.

17 Nov 2026. A route begins a job (`begin`) before its work starts, binds the thread that does the
work (`bind`),  and ends it (`end`) after the work, and after the restore  when the job was cancelled.
`request(sid, kind)`  sets the flag of every job of  the session that has not passed its gate
(`Job.commit`), ends the Excels  those jobs started -- only Excels this  app process started:
`excel_com.kill_excel` checks each one -- and lets go of their items in the session's Excel worker.
The work reads the flag at its natural stops (`check`) and leaves with `Cancelled`.

A thread works for a job only when it was bound to it: Python starts a thread with an empty context,
so every thread of a job binds by hand, and a  thread nobody bound sees no job (every check is then a
no-op: the command line scripts, the tests of other modules).

This module imports nothing from the app at the top, so every module may import it.
"""
from __future__ import annotations

import contextlib
import contextvars
import threading
import time
from typing import Any, Callable, Iterator

CALL_SLICE_S = 0.25  # how  often a  waiting `call` looks at the flag


class Cancelled(BaseException):
    """The  person pressed  Cancel. A BaseException on purpose: the many `except Exception` of the app
    (the server's routes, prep, the Excel worker, the model calls, the Mind client) would turn a
    cancel into an error, a  fallback or a retry."""

    def __init__(self, message: str = "Stopped.") -> None:
        super().__init__(message)


class Busy(Exception):
    """Another  job that  changes this workbook is registered. `.kind` is that job's kind."""

    def __init__(self, kind: str) -> None:
        super().__init__(f"another job that changes this workbook is running: {kind}")
        self.kind = kind


# One lock for the registry, every flag and every gate: `Job.commit` and `request` take it, so of a
# Cancel and a job's end  exactly one wins. Re-entrant, so `on_flag` may read the registry.
_LOCK = threading.RLock()
_JOBS: dict[str, list["Job"]] = {}
_CURRENT: contextvars.ContextVar["Job | None"] = contextvars.ContextVar("mindready_job", default=None)


class Job:
    """One  piece of  work for one workbook, from `begin` to `end`."""

    def __init__(self, sid: str, kind: str, changes: bool = False, on_flag: Callable[[], Any] | None = None) -> None:
        self.sid = sid
        self.kind = kind
        self.changes = bool(changes)  # it changes the workbook all along: one such job per workbook
        self.started = time.time()
        self.flag = threading.Event()   # Cancel was pressed for this job
        self.ended = threading.Event()  # `end` was called: nothing of it runs any more
        self.committed = False          # it  passed its gate: its result stands, a Cancel is too late
        self.outcome: str | None = None  # "done"  | "cancelled" | "error", written by  `end`
        self.pids: set[int] = set()     # the Excels this job started and has  not quit
        self.on_flag = on_flag          # called once, under the registry's lock, when the flag is  set

    def check(self) -> None:
        """Leave with `Cancelled` when Cancel was pressed  for this job."""
        if self.flag.is_set():
            raise Cancelled()

    def commit(self) -> None:
        """The gate, the point of no return. Flag set: `Cancelled`. Else the job is committed, and a
        Cancel that comes later finds it in `late` and does not touch it. A second call does nothing."""
        with _LOCK:
            if self.committed:
                return
            if self.flag.is_set():
                raise Cancelled()
            self.committed = True

    def __repr__(self) -> str:
        state = "ended" if self.ended.is_set() else "committed" if self.committed else "flagged" if self.flag.is_set() else "running"
        return f"<Job {self.kind} {self.sid} {state}>"


def begin(sid: str, kind: str, changes: bool = False, on_flag: Callable[[], Any] | None = None) -> Job:
    """Register  a job  for the workbook `sid`. With `changes`: `Busy` while another job that changes
    this workbook is registered (it stays registered until its `end`, which comes after its restore)."""
    job = Job(sid, kind, changes, on_flag)
    with _LOCK:
        if job.changes:
            for other in _JOBS.get(sid, ()):
                if other.changes:
                    raise Busy(other.kind)
        _JOBS.setdefault(sid, []).append(job)
    return job


def end(job: Job | None, outcome: str = "done") -> None:
    """The  job is  over (after its restore, when it was cancelled). Idempotent: the first call's
    outcome stays. After it returns, `on_flag` is never called for this job again."""
    if job is None:
        return
    with _LOCK:
        if job.ended.is_set():
            return
        job.outcome = outcome
        mine = _JOBS.get(job.sid)
        if mine is not None:
            if job in mine:
                mine.remove(job)
            if not mine:
                _JOBS.pop(job.sid, None)
        job.ended.set()


@contextlib.contextmanager
def bind(job: Job | None) -> Iterator[Job | None]:
    """This  thread works  for `job` inside the block. `bind(None)` binds nothing."""
    if job is None:
        yield None
        return
    token = _CURRENT.set(job)
    try:
        yield job
    finally:
        _CURRENT.reset(token)


def current() -> Job | None:
    """The  job this  thread works for, or None."""
    return _CURRENT.get()


def check() -> None:
    """`Cancelled`  when Cancel  was pressed for the job this thread works for. No job: nothing."""
    job = _CURRENT.get()
    if job is not None and job.flag.is_set():
        raise Cancelled()


def flagged() -> bool:
    """True  when Cancel  was pressed for the job this thread works for. No job: False."""
    job = _CURRENT.get()
    return job is not None and job.flag.is_set()


def note_excel(pid: int | None) -> None:
    """The  current job  started the Excel `pid` (already written down by `excel_com._remember`): a
    Cancel ends it. When the Cancel came while that Excel was starting, nobody else knows the pid:
    it is ended here, and the  work leaves with `Cancelled`."""
    job = _CURRENT.get()
    if job is None or not pid:
        return
    with _LOCK:
        job.pids.add(int(pid))
        late = job.flag.is_set()
    if late:
        from . import excel_com  # by its module, at the call: the checked  kill (and what a test puts  there)

        excel_com.kill_excel(int(pid))
        with _LOCK:
            job.pids.discard(int(pid))
        raise Cancelled()


def forget_excel(pid: int | None) -> None:
    """The  Excel `pid`  was quit, or handed to the session's Excel worker: no job ends it any more.
    Taken off every registered job, not only the current one: the thread that quits an Excel is  not
    always bound to the job that started it."""
    if not pid:
        return
    job = _CURRENT.get()
    with _LOCK:
        if job is not None:
            job.pids.discard(int(pid))
        for mine in _JOBS.values():
            for other in mine:
                other.pids.discard(int(pid))


def call(fn: Callable[[], Any]) -> Any:
    """`fn()`  for work  that cannot be interrupted (a web request): it runs in a helper thread while
    this one looks at the flag every quarter second. On a cancel this raises `Cancelled` and the
    helper is  left to end alone (its answer is dropped). The value, or  the error, of `fn` comes
    back as  from a plain call. No job: a  plain `fn()`."""
    job = _CURRENT.get()
    if job is None:
        return fn()
    job.check()
    box: dict[str, Any] = {}
    done = threading.Event()
    context = contextvars.copy_context()  # the helper sees what this thread sees (the job too)

    def run() -> None:
        try:
            box["value"] = context.run(fn)
        except BaseException as exc:  # noqa: BLE001 -- handed to the caller below
            box["error"] = exc
        finally:
            done.set()

    threading.Thread(target=run, name=f"cancel-call-{job.kind}", daemon=True).start()
    while not done.wait(CALL_SLICE_S):
        job.check()
    if "error" in box:
        raise box["error"]
    return box.get("value")


def jobs(sid: str) -> list[Job]:
    """The  registered jobs  of the workbook, oldest first."""
    with _LOCK:
        return list(_JOBS.get(sid, ()))


def changing(sid: str) -> Job | None:
    """The  registered job  that changes this workbook (there is at most one), or None."""
    with _LOCK:
        for job in _JOBS.get(sid, ()):
            if job.changes:
                return job
    return None


def request(sid: str, kind: str | None = None) -> dict[str, Any]:
    """Cancel  was pressed  for the workbook `sid` (for its jobs of `kind` when given).

    Under the lock, every registered job: committed or ended goes to `late` (its result stands);
    else its flag is set, `on_flag` is  called the first time, and it is collected. Then, outside the
    lock:  the Excels those jobs started are  ended (`excel_com.kill_excel` ends only an Excel this
    process  started) and their items in the session's Excel worker are let  go; with no kind, or
    kind "recalc", the  recalculation made ahead too (it belongs to no job).

    -> {"flagged": [jobs], "late": [jobs], "ahead": bool}. Nothing running: empty lists. A second
    call does the same harmless things."""
    hit: list[Job] = []
    late: list[Job] = []
    pids: list[int] = []
    with _LOCK:
        for job in list(_JOBS.get(sid, ())):
            if kind is not None and job.kind != kind:
                continue
            if job.committed or job.ended.is_set():
                late.append(job)
                continue
            first = not job.flag.is_set()
            job.flag.set()
            if first and job.on_flag is not None:
                try:
                    job.on_flag()
                except Exception:
                    pass  # a status line that could not be written never holds a Cancel
            hit.append(job)
            pids.extend(sorted(job.pids))
            job.pids.clear()  # taken: an Excel this job starts from now on is ended by note_excel
    from . import excel_com, excel_worker  # here, not at the top: both import the app

    for pid in pids:
        excel_com.kill_excel(pid)
    stopped = excel_worker.cancel(sid, hit, free=kind in (None, "recalc"))
    return {"flagged": hit, "late": late, "ahead": bool(stopped)}
