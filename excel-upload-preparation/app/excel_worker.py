"""One  Excel per  session, kept by a thread of its own.

Starting an Excel, opening the version (a minute on a 179 MB workbook), writing, saving and quitting
for every step of Do everything, then opening  the saved file again read-only to check it, costs
minutes on a  big file. Instead:

- When a version  is saved, its check is made by  opening a fresh copy of it -- the very copy the  next
  apply will write into, at the very  path make_immutable_copy would give it -- in a freshly started
  Excel. That runs while the server re-analyses the version (minutes on a big file), so it costs no
  wall-clock time, and the workbook stays open ("held").
- The next apply on that version writes into the held workbook: no Excel start, no open, no copy.
- Anything else -- another version, an explicit_colors operation (it  edits the copy before Excel
  opens it), a workbook with data connections, little free memory, the idle timeout, a Recalculate
  or any other Excel user -- closes the held workbook and deletes its copy first, so later paths and
  results are exactly what they were.

Every COM object of a session lives on its thread `excel-<session id>`: other threads hand it jobs
through a queue and wait for plain results (COM proxies must stay on the thread that made them).
When Excel dies or hangs (the wait  has a timeout), the worker  kills it, drops what it held and
raises  WorkerUnavailable; the server  then falls  back to one Excel per step.
MINDREADY_EXCEL_WORKER=0 turns it off.

Cancel (17 Nov 2026, app/cancel.py). Every queued item carries its owner: the job of the thread that
asked (none for the check and the  recalculation made ahead, which outlive their job). The worker
thread runs an  item bound to that job, so  every check inside it reads that job's  flag. `cancel(sid,
jobs, free)` lets go of the items of the cancelled  jobs only: what waits never runs and  its caller
returns at once; when the item running now is theirs the worker is retired and its Excel ended (only
an Excel this app started: excel_com.kill_excel). A caller whose own job was cancelled leaves with
Cancelled, never with WorkerUnavailable, so it cannot fall back to an Excel of its own.
"""
from __future__ import annotations

import atexit
import concurrent.futures
import ctypes
import gc
import os
import queue
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from . import cancel as registry  # 17  Nov 2026:  under another name -- `cancel` here is this module's own function (below)
from .excel_com import _file_stamp, close_quietly, excel_alive, excel_gone, excel_pid, kill_excel, open_check, quit_excel, start_excel
from .inventory import make_immutable_copy, next_copy_path
from .perf import record as perf_record
from .perf import stage as perf_stage

JOB_TIMEOUT_S = float(os.environ.get("MINDREADY_EXCEL_JOB_TIMEOUT_S", "1800"))  # once  a job  has started
QUEUE_TIMEOUT_S = float(os.environ.get("MINDREADY_EXCEL_QUEUE_TIMEOUT_S", "1800"))  # waiting  for its  turn
IDLE_S = float(os.environ.get("MINDREADY_EXCEL_IDLE_S", "900"))
HOLD_MIN_FREE_MB = int(os.environ.get("MINDREADY_EXCEL_HOLD_MIN_FREE_MB", "3000"))


def speculative_recalc() -> bool:
    return os.environ.get("MINDREADY_SPECULATIVE_RECALC", "1").strip().lower() not in ("0", "off", "no", "false")


def enabled() -> bool:
    return os.environ.get("MINDREADY_EXCEL_WORKER", "1").strip().lower() not in ("0", "off", "no", "false")


class WorkerUnavailable(RuntimeError):
    """The  worker's Excel  died, hung or failed: do it the way it was done before (a fresh Excel)."""


def free_mb() -> int:
    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong), ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong), ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong), ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]

    try:
        m = MEMORYSTATUSEX()
        m.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
        return int(m.ullAvailPhys // 2**20)
    except Exception:
        return 1 << 30


@dataclass
class _Held:
    source: Path            # the  version file  the copy was made from
    stamp: tuple[int, int]  # its (mtime_ns, size) then
    copy_path: Path
    source_sha256: str
    wb: Any


@dataclass
class ExcelRun:
    """What  apply_operations needs  to write into a workbook the worker already holds open."""
    excel: Any
    wb: Any
    copy_path: Path
    source_sha256: str
    verify: Callable[[Path], dict[str, Any]] | None


class ExcelWorker:
    def __init__(self, sid: str, apply_dir: Path, idle_s: float = IDLE_S) -> None:
        self.sid = sid
        self.apply_dir = Path(apply_dir)
        self.idle_s = idle_s
        self._q: queue.Queue = queue.Queue()
        self._stopped = False
        self._retired = False  # 17  Nov 2026:  stopped by _retire (its Excel is ended next), not by a plain stop()
        # 17 Nov 2026 (Cancel): which job the item running now belongs to. Written by the worker
        # thread and read by  `cancel_jobs`, both under `_turn`, so a Cancel never mistakes the next
        # item  (another job's) for the one it  means to stop.
        self._turn = threading.Lock()
        self._busy = False
        self._owner: Any = None
        # owned by the worker thread  only
        self._excel: Any = None
        self._pid: int | None = None
        self._held: _Held | None = None
        self._recalc: tuple[Path, tuple[int, int], Path, dict[str, Any]] | None = None  # (version, stamp, copy, result)
        self.thread = threading.Thread(target=self._loop, name=f"excel-{sid}", daemon=True)
        self.thread.start()

    # --- the thread ---------------------------------------------------------------------------
    def _loop(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        last = time.time()
        try:
            while True:
                try:
                    item = self._q.get(timeout=1)
                except queue.Empty:
                    pythoncom.PumpWaitingMessages()
                    if self._held is not None and time.time() - last > self.idle_s:
                        self._drop()
                    continue
                if item is None:
                    break
                fn, fut, started, owner = item
                with self._turn:
                    running = fut.set_running_or_notify_cancel()
                    if running:
                        self._busy, self._owner = True, owner
                if not running:
                    started.set()  # whoever cancelled it has let  its caller go; once more does no  harm
                    fn = fut = item = started = owner = None
                    continue  # cancelled while it waited: it never runs
                started.set()
                try:
                    # 17 Nov 2026: this thread works for the job that queued the item, so the checks
                    # inside it (the progress closures, start_excel) read that job's flag
                    with registry.bind(owner):
                        value = fn()
                    if self._retired:
                        # its Excel was ended under it (a Cancel, the time limit): what it answered is
                        # no answer -- "this version does not open in Excel" would be one
                        value = None
                        raise WorkerUnavailable("the Excel worker was retired while this job ran")
                    fut.set_result(value)
                    value = None
                except BaseException as exc:  # noqa: BLE001 -- handed to the waiting caller (a Cancelled too)
                    if excel_gone(exc) or (self._excel is not None and not excel_alive(self._excel)):
                        self._abandon()
                        fut.set_exception(WorkerUnavailable(f"Excel stopped answering: {str(exc)[:200]}"))
                    elif isinstance(exc, registry.Cancelled):
                        fut.set_exception(registry.Cancelled())  # a fresh one: the frames  of this one hold COM proxies of this thread
                    else:
                        fut.set_exception(exc)
                finally:
                    with self._turn:
                        self._busy, self._owner = False, None
                    fn = fut = item = started = owner = None
                    gc.collect()
                    last = time.time()
                if self._stopped and self._q.empty():
                    break
        finally:
            self._drop()
            pythoncom.CoUninitialize()

    def _ensure_excel(self) -> Any:
        if self._excel is None:
            excel = start_excel()  # a Cancel of this item's own job while it started: ended there, Cancelled
            pid = excel_pid(excel)
            # 17 Nov 2026: this Excel serves the items of any job, so no job keeps its pid: it is
            # ended through the worker, and only while the item running in it is the cancelled job's
            registry.forget_excel(pid)
            self._excel, self._pid = excel, pid  # before the look at _stopped: whoever retires reads _pid after
            excel = None
            if self._stopped:
                # retired while its Excel started (a Cancel, the time limit): the kill found no pid then
                kill_excel(pid)
                self._excel, self._pid = None, None
                raise WorkerUnavailable("the Excel worker was stopped while its Excel started")
        return self._excel

    def _discard_recalc(self) -> None:
        """A recalculation made ahead  that nobody took: delete its  copy, so the next Recalculate's
        copy takes the  path it always  did."""
        found, self._recalc = self._recalc, None
        if found is not None:
            try:
                found[2].unlink()
            except OSError:
                pass

    def _drop(self) -> None:
        """Close what  is held, delete its copy (so the next copy takes the same path), quit Excel."""
        self._discard_recalc()
        held, self._held = self._held, None
        if held is not None:
            close_quietly(held.wb)
            held.wb = None
            try:
                held.copy_path.unlink()
            except OSError:
                pass
        held = None
        excel, self._excel, self._pid = self._excel, None, None
        quit_excel(excel)
        excel = None
        gc.collect()

    def _abandon(self) -> None:
        """Excel died or hangs: kill it first, then forget everything."""
        kill_excel(self._pid)
        held, self._held = self._held, None
        if held is not None:
            held.wb = None
            try:
                held.copy_path.unlink()
            except OSError:
                pass
        self._excel, self._pid = None, None
        gc.collect()
        perf_record("excel_worker_abandon", 0.0, session=self.sid)

    def _submit(self, fn: Callable[[], Any], timeout: float = JOB_TIMEOUT_S, queue_timeout: float = QUEUE_TIMEOUT_S, kill_on_timeout: bool = True) -> Any:
        """Run `fn` on the worker thread and wait  for it. Waiting for its turn and running are timed
        apart: a  job that never got its turn  is cancelled (it will not run later);  a job that runs
        too long has its Excel killed -- only  then, and only the running job's --  and the worker is
        retired, so later requests fall back to one Excel per step at once instead of queueing behind it.

        17 Nov 2026: the item belongs to the job of the thread that asks (`owner`). A job that was
        cancelled queues nothing more. When the worker lets the item go and that job was cancelled,
        the caller leaves with Cancelled, not with WorkerUnavailable:  a cancelled job must never fall
        back to an Excel of its own."""
        registry.check()
        try:
            return self._wait_for(fn, registry.current(), timeout, queue_timeout, kill_on_timeout)
        except WorkerUnavailable:
            registry.check()
            raise

    def _wait_for(self, fn: Callable[[], Any], owner: Any, timeout: float, queue_timeout: float, kill_on_timeout: bool) -> Any:
        if self._stopped:
            raise WorkerUnavailable("the Excel worker is stopped")
        fut: concurrent.futures.Future = concurrent.futures.Future()
        started = threading.Event()
        self._q.put((fn, fut, started, owner))
        if self._stopped and fut.cancel():
            # 17 Nov 2026: stopped or retired between the look above and the put (a Cancel of another
            # job): the item may lie behind the end marker, where it never runs and nobody lets it go --
            # its caller would wait QUEUE_TIMEOUT_S (half an hour), and its own Cancel could not reach it
            raise WorkerUnavailable("the Excel worker was stopped")
        if not started.wait(queue_timeout) and fut.cancel():
            raise WorkerUnavailable(f"the Excel worker was busy for {int(queue_timeout)} s")
        if fut.cancelled():
            # 17 Nov 2026: let go while it waited (a Cancel, or the worker was retired or stopped):
            # `started` was set for it,  so this caller returns at  once instead of after the timeout
            raise WorkerUnavailable("the Excel worker let this job go before its turn")
        try:
            return fut.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            if not kill_on_timeout:
                raise WorkerUnavailable(f"the Excel worker did not finish within {int(timeout)} s") from None
            self._retire()
            kill_excel(self._pid)  # the hung  call then fails  on the  worker thread, which drops everything
            raise WorkerUnavailable(f"Excel did not answer within {int(timeout)} s") from None

    def _let_go_of_what_waits(self) -> None:
        """What waits never runs, and its caller returns now (17 Nov 2026: `started` is set too;
        before, a waiting caller stayed blocked for QUEUE_TIMEOUT_S, half an hour)."""
        while True:
            try:
                item = self._q.get_nowait()
            except queue.Empty:
                break
            if item is not None:
                item[1].cancel()
                item[2].set()

    def _retire(self) -> None:
        """No more  jobs for this worker: cancel what waits, and let worker_for make  a new one."""
        self._stopped = self._retired = True
        with _LOCK:
            if WORKERS.get(self.sid) is self:
                WORKERS.pop(self.sid, None)
        self._let_go_of_what_waits()
        self._q.put(None)

    def cancel_jobs(self, cancelled: list[Any], free: bool = False) -> int:
        """Cancel (17 Nov 2026).  An item is "theirs" when its owner  is one of `cancelled`, or, with
        `free`, when it has no owner  (a check or a recalculation made ahead)  or its owner has ended.
        Every waiting item that is theirs never runs and its caller returns at once. When the item
        running now is theirs: the two steps of the time limit, retire then end this worker's Excel
        (the item then fails on the worker thread, which drops everything). An item of a job that is
        not being cancelled is never touched (but what  waits behind a running item that is stopped is
        let go with it, as on the time limit: its caller falls back to an Excel of its own).
        -> how many items were stopped that belong to no job in `cancelled`."""
        def named(owner: Any) -> bool:
            return owner is not None and any(owner is job for job in cancelled)

        def nobodys(owner: Any) -> bool:
            return free and (owner is None or owner.ended.is_set())

        stopped = 0
        with self._turn:  # the worker thread cannot move on to its next item meanwhile
            with self._q.mutex:
                waiting = [item for item in self._q.queue if item is not None]
            for _fn, fut, started, owner in waiting:
                if (named(owner) or nobodys(owner)) and fut.cancel():
                    started.set()
                    stopped += not named(owner)
            if self._busy and (named(self._owner) or nobodys(self._owner)):
                stopped += not named(self._owner)
                self._retire()
                kill_excel(self._pid)  # only an Excel this process started: kill_excel checks it
        return stopped

    # --- jobs (run on the worker thread) ------------------------------------------------------
    def _job_preopen(self, version: Path) -> dict[str, Any]:
        try:
            return self._preopen(Path(version))
        except Exception as exc:
            if excel_gone(exc):
                raise
            self._drop()  # nothing half-held: the next apply opens cold
            return {"available": False, "opens": None, "sheet_count": None, "error": str(exc)[:300]}

    def _preopen(self, version: Path) -> dict[str, Any]:
        self._drop()  # also deletes a recalculation made ahead that nobody took
        with perf_stage("verify_opens", file=version.name, worker=True):
            try:
                excel = self._ensure_excel()
            except Exception as exc:
                self._abandon()
                return {"available": False, "opens": None, "sheet_count": None, "error": str(exc)[:300]}
            stamp = _file_stamp(version)
            copy_path, sha = make_immutable_copy(version, self.apply_dir)
            info, wb = open_check(excel, copy_path)
            keep = wb is not None and info.get("opens") is True
            if keep:
                try:
                    keep = int(wb.Connections.Count) == 0  # a data connection could  refresh while held
                except Exception:
                    keep = False
            if keep and free_mb() < HOLD_MIN_FREE_MB:
                keep = False
            if keep:
                self._held = _Held(version, stamp, copy_path, sha, wb)
            else:
                close_quietly(wb)
                try:
                    copy_path.unlink()
                except OSError:
                    pass
                self._drop()
            wb = excel = None
        return info

    def _job_apply(self, source: Path, operations: list[dict[str, Any]], progress: Any, verify: Callable[[Path], dict[str, Any]]) -> dict[str, Any]:
        from .prep import apply_operations

        held = self._held
        warm = (
            held is not None
            and held.source == Path(source)
            and held.stamp == _file_stamp(source)
            and not any(o["op"] == "explicit_colors" for o in operations)
        )
        if not warm:
            # whatever is held  is not this version's copy: delete it,  then the old  path --  in an
            # Excel this worker started, so a hang can be killed
            self._drop()
            held = None
            try:
                excel = self._ensure_excel()
            except Exception:
                self._abandon()
                if self._stopped:
                    raise  # 17 Nov 2026: retired while Excel started -- nobody waits for this apply any more
                excel = None  # apply_operations starts its own Excel
            try:
                return apply_operations(Path(source), self.apply_dir, operations, progress=progress, defer_verify=True, verify=verify, excel=excel)
            finally:
                excel = None
                self._drop()
        self._discard_recalc()
        self._held = None
        run = ExcelRun(self._excel, held.wb, held.copy_path, held.source_sha256, verify)
        held.wb = None
        copy_path = held.copy_path
        failed: str | None = None
        gone = False
        cancelled = False
        res: dict[str, Any] = {}
        try:
            # 17 Nov 2026:  inside the try (it was above it) -- this first progress  call is where a
            # cancelled apply leaves,  and the open copy must then be  let go of like any other end
            if progress is not None:
                progress("apply_copy", "The copy was made and opened while the previous version was being analysed", None, done=0, total=len(operations))
            res = apply_operations(Path(source), self.apply_dir, operations, progress=progress, defer_verify=True, run=run)
        except registry.Cancelled:
            cancelled = True  # Excel lets go of the  half-written copy below; then the copy goes
            raise
        except Exception as exc:  # the  text only: the exception's frames hold COM proxies of this thread
            failed = f"{type(exc).__name__}: {str(exc)[:240]}"
            gone = excel_gone(exc)
        finally:
            run.wb = None
            run.excel = None
            run = held = None
            if gone:
                self._abandon()
            else:
                excel, self._excel, self._pid = self._excel, None, None  # every write starts from a freshly opened copy
                quit_excel(excel)
                excel = None
            if cancelled:
                try:
                    copy_path.unlink()  # so the next apply's copy takes this path again
                except OSError:
                    pass
        if failed is not None:
            try:
                copy_path.unlink()  # the old path makes this copy again, at this path
            except OSError:
                pass
            raise WorkerUnavailable(f"the apply in the open copy failed ({failed})")
        return res

    def _job_recalc_ahead(self) -> None:
        """After a Mind loop version, Do everything recalculates next: do it now, on the copy held open
        -- what POST /recalculate does (a full rebuild in a fresh  Excel with the add-in, saved, every
        error cell read), on a byte-identical copy, earlier. The copy then moves to where Recalculate
        would have put it."""
        from .recalc import recalculate_in

        if self._held is not None and free_mb() < HOLD_MIN_FREE_MB:
            self._drop()  # not now: Recalculate runs in its own fresh Excel
            return
        held, self._held = self._held, None
        if held is None:
            return
        try:
            res = recalculate_in(self._excel, held.copy_path, held.wb)
            held.wb = None
            target = next_copy_path(self.apply_dir.parent / "recalc", held.copy_path.name)
            held.copy_path.replace(target)
            self._recalc = (held.source, held.stamp, target, res)
        except Exception as exc:
            held.wb = None
            try:
                held.copy_path.unlink()
            except OSError:
                pass
            self._recalc = None
            if excel_gone(exc):
                raise
        finally:
            held = None
            excel, self._excel, self._pid = self._excel, None, None
            quit_excel(excel)
            excel = None

    def _job_take_recalc(self, version: Path) -> tuple[Path, dict[str, Any]] | None:
        found, self._recalc = self._recalc, None
        if found is None:
            return None
        source, stamp, path, res = found
        if source == Path(version) and stamp == _file_stamp(version) and path.is_file():
            return path, res
        return None

    # --- public (any thread) ------------------------------------------------------------------
    def preopen(self, version: Path, recalc_ahead: bool = False) -> dict[str, Any]:
        """The opens-in-Excel check of `version`, made by opening the next apply's copy of it.
        `recalc_ahead`: a Mind loop version -- recalculate the held copy right after (in the queue,
        before any later request), for the Recalculate step that follows."""
        try:
            info = self._submit(lambda: self._job_preopen(Path(version)))
        except WorkerUnavailable as exc:
            return {"available": False, "opens": None, "sheet_count": None, "error": str(exc)[:300]}
        if recalc_ahead and speculative_recalc() and info.get("opens") and not self._stopped:
            # no owner, on purpose (17 Nov 2026): it outlives the job  that made the version, and
            # only  a Cancel of everything, or of a  Recalculate, stops it
            self._q.put((self._job_recalc_ahead, concurrent.futures.Future(), threading.Event(), None))
        return info

    def take_recalc(self, version: Path) -> tuple[Path, dict[str, Any]] | None:
        """The  recalculation of  `version` made ahead (waits for it if it is still running), or None."""
        try:
            return self._submit(lambda: self._job_take_recalc(Path(version)))
        except WorkerUnavailable:
            return None

    def preopen_async(self, version: Path) -> None:
        if self._stopped:
            return
        # no owner, on purpose (17 Nov 2026): nobody waits for it, and it outlives the job that asked
        self._q.put((lambda: self._job_preopen(Path(version)), concurrent.futures.Future(), threading.Event(), None))

    def apply(self, source: Path, operations: list[dict[str, Any]], progress: Any = None, defer_verify: bool = False, recalc_ahead: bool = False) -> dict[str, Any]:
        """apply_operations on  `source`, in the copy held open when it is this version's.  The check
        that the result opens is  the open of the next copy (preopen).  Not deferred: done before
        returning, as apply_operations does. `recalc_ahead`: the result is  a Mind loop version."""
        verify = (lambda path: self.preopen(path, recalc_ahead=True)) if recalc_ahead else self.preopen
        res = self._submit(lambda: self._job_apply(Path(source), operations, progress, verify))
        if not defer_verify and "finish" in res:
            res.pop("finish")()  # a job of its  own, after the apply job released its Excel
        return res

    def release(self) -> None:
        """Let go of the held workbook and of Excel (before another Excel user, or a Recalculate)."""
        try:
            self._submit(self._drop, timeout=300, queue_timeout=120, kill_on_timeout=False)
        except WorkerUnavailable:
            pass  # busy: the caller goes ahead (memory only); a busy job is never killed for this

    def stop(self) -> None:
        if self._stopped:
            return
        self._stopped = True
        self._let_go_of_what_waits()  # what waits never runs
        self._q.put(None)
        self.thread.join(timeout=120)


WORKERS: dict[str, ExcelWorker] = {}
_LOCK = threading.Lock()


def worker_for(sid: str, apply_dir: Path) -> ExcelWorker | None:
    from .excel_com import com_available

    if not enabled() or not com_available():
        return None
    with _LOCK:
        w = WORKERS.get(sid)
        if w is None or w._stopped:
            w = WORKERS[sid] = ExcelWorker(sid, apply_dir)
        return w


def release(sid: str) -> None:
    w = WORKERS.get(sid)
    if w is not None and not w._stopped:
        w.release()
    registry.check()  # 17  Nov 2026:  the wait can be minutes -- a job cancelled meanwhile starts no Excel after it


def cancel(sid: str, jobs: list[Any], free: bool = False) -> int:
    """Cancel  was pressed  for the `jobs` (app.cancel.Job) of the workbook `sid`: let go of their items
    in its Excel worker (ExcelWorker.cancel_jobs). `free`: also the items that belong to no job, the
    recalculation made ahead among them.  -> how many items were stopped that belong to none of
    `jobs`.  Called by app.cancel.request, after the flags  are set."""
    w = WORKERS.get(sid)
    if w is None:
        return 0
    return w.cancel_jobs(list(jobs or []), free)


def stop_all_workers() -> None:
    with _LOCK:
        workers = list(WORKERS.values())
        WORKERS.clear()
    for w in workers:
        w.stop()


atexit.register(stop_all_workers)
