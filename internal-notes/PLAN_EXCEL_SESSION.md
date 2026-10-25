# Plan: one Excel per session (change #5, not built yet)

Written 2026-10-20 by a read-only design workflow (one mapper, two independent designs, one judge who re-read the code). Nothing here is implemented. Build it on branch `perf-excel-session`, prove it with `scripts/compare_final_workbooks.py` on Phoenix against `%LOCALAPPDATA%\MindReady\perf\night2\before_phoenix_try1_chain`.

## Change #5 (final): one Excel worker per session. While the re-analysis runs, it opens the next step's copy, and that open is the version's opens-in-Excel check. The Recalculate after the Mind loop is also moved behind the loop's re-analysis.

## Expected saving

Estimated from the Phoenix stages in before_phoenix_try1_timings.jsonl (one Prep round); not measured yet.

| Step | Today | With this plan | Why |
|---|---|---|---|
| Prep round 1 | 116.0 s | about 51 s | Excel start 4.9 s, open 58.1 s and the copy/hash move into the pre-open alongside the auto-fix scan (272 s) |
| Mind loop | 131.1 s | about 55 s | start 4.7 s, open 66.4 s and the copy/hash move alongside Prep's scan (365 s) |
| Recalculate | 114.6 s | a few seconds | the speculative recalculation runs alongside the loop's scan (383 s): pre-open about 70 s plus rebuild, save and read about 50 s |

Total: about 255 s off the roughly 1,134 s from the end of the auto-fix scan to the end of Recalculate, which is about 22%. Without the speculation it would be about 200 s (18%). Each extra Prep round saves another 60-65 s.

The two separate check Excels (77 s and 80 s) disappear. Each version is opened once instead of twice (5 opens become 3, all off the critical path), and Excel starts drop from 5 to 3 (plus the auto-fix's own). At most one Excel per session is alive at any time.

Saves stay at 3 (30-47 s each), because the re-analysis reads saved files. The Recalculate save is now hidden behind the scan.

## Approach

VERDICT. I re-read the code both designs depend on: excel_com.py (all), prep.py 1816-2010 and 2209-2289, inventory.py 195-219, recalc.py 183-260 and 415-428, server.py 159-206, 422-435, 759-813, 1093-1186 and 1232-1348, autofix.py 2411-2464, home.page.ts 600-703, and test_web_api.py 58-99. Test 528 of test_readiness_and_prep_fixes.py was also re-read. Both designs read the functions correctly. The pick is Design A (one worker per session, pre-open alongside the scan). It is the safer of the two:
- Every write still starts the way it does today: a freshly started Excel with the XLL, and one workbook opened from a byte-identical copy on disk. No in-memory state carries over between rounds: no calc chain, undo, pivot cache or MM_LOOP/MM_RESULT state inside the XLL.
- It needs no new endpoint and no front-end or night_chain changes.
- It speeds up the single buttons too.
- Every version keeps a real opens-in-Excel answer.

Design B's in-memory LiveWorkbook would also have:
- copied a workbook Excel holds open (not tested here);
- labelled the middle versions "unverified";
- needed a parallel orchestration path (CHAINS, new refusals in analyze, pivot-mode and autofix, a new runChainRest, night_chain changes) that would have to copy the callers' stop rules exactly.

How A meets the goal: today each version costs two opens on or beside the critical path (the write open and a second-Excel check). With this plan it costs one open, and that open sits off the critical path. The separate verify Excels are gone (no middle verifications as separate Excels, and at most one Excel per session). The open still happens once per version, but it runs while the 270-380 s re-analysis runs, so it no longer adds wall-clock time.

FLAWS IN A, FIXED HERE
1. COM apartment hazard (autofix.py:2426-2429). A's inline pre-open in the non-deferred case (reanalyze=false, test_web_api.py:78) quits and restarts Excel while apply_operations' frame still holds the old proxy (`run.excel`). That proxy is then released while the new Excel is running: exactly the "interface is unknown" failure. Fix: the worker always runs apply_operations with defer_verify=True and drops every proxy (gc) when the job ends. When the caller wanted non-deferred, worker.apply then calls res.pop('finish')() itself, which queues the pre-open as a separate job. The returned shape is the same as today's non-deferred result: verified flag set, change log written, no 'finish' key. The inline call() path is removed.
2. Orphan Excel. The worker thread is a daemon, and @app.on_event('shutdown') fires under TestClient only in `with` form. So add atexit.register(stop_all_workers), plus a conftest autouse fixture that stops the workers after each test. Otherwise pytest leaves EXCEL.EXE holding a held copy.
3. Recalculate's copy path. A left the recalculated copy in apply/vK, which moves every later version path along by one. Fix: after the warm recalculation closes the workbook, rename the copy to the path make_immutable_copy would have given in work_dir/recalc (the helper is pulled out as inventory.next_copy_path). apply/ and recalc/ paths are then the same as today.
4. _original_baseline stays unchanged (its own short Excel on the calling thread). Routing it through the worker gains nothing: after any recalculate job the worker quits its Excel because there is nothing left to hold, so the baseline's Excel is still the only one.
5. Hang handling is part of the design, not optional. The caller waits with future.result(timeout=MINDREADY_EXCEL_JOB_TIMEOUT_S, default 1800). On timeout the worker kills the pid first, then cleans up (from B's abandon). The COM call in flight then raises an RPC error, the job ends and the caller gets WorkerUnavailable.

GRAFTED FROM B
- Recalculate behind the loop's re-analysis. When the version just made has source 'loop', the worker queues a speculative-recalculation job right after the pre-open answers. It does what recalc._scan does today on the held copy: CalculateFullRebuild, save_in_place, error read, close. That is the same sequence in the same fresh Excel on the same byte-identical copy, only earlier. The result is cached under (version id, path, mtime_ns, size). POST /recalculate on that version takes the cached result, so the Recalculate step drops from ~115 s to the time of _recalc_out. Any other next step finds no held copy and opens cold, as today. MINDREADY_SPECULATIVE_RECALC=0 turns this off.
- Proxies are kept only as worker attributes. An owner-thread assertion stops any other thread touching them.
- abandon() kills the pid first.
- The 'unchanged' case: if Fix all issues ran but made no version, the worker pre-opens the current version so Prep round 1 still starts warm.
- To flag to Joseph, not fix: runAll decides the loop step from the loop plan loaded at session load (home.page.ts:217/322, 676), while night_chain re-reads it after Prep.

## Who owns Excel

Thread `excel-{sid}`, one per session, a daemon, created lazily by `_worker_of(s)`. It is the only thread that touches that session's Excel COM objects:
- it calls CoInitialize (STA) once;
- start_excel and quit_excel run only there;
- ctx.excel, ctx.pid, ctx.held.wb and the recalc cache live in a `_Ctx` that asserts the owner thread ident on every access.

Requests reach it through a queue.Queue of (fn, concurrent.futures.Future). The callers then wait with `Future.result(timeout)` and get back only dicts and Paths:
- FastAPI threadpool threads: /apply and /mind-loop-plan/apply through `_tracked_apply`, and /recalculate.
- The `verify-{sid}` thread: `res['finish']()` calls ExcelRun.verify, which is worker.preopen, a queued job.
- The `autofix-{sid}` thread: release and preopen_async.

The queue is FIFO:
- The pre-open (queued by the verify thread) and a speculative recalculation (queued by the worker itself right behind the pre-open) both run before any later apply or recalculate.
- `_apply_and_version` returns only after checker.join(), so the next request cannot overtake its own version's pre-open.

No job ever calls back into the queue from the worker thread. The non-deferred check is queued by the caller after the apply job has ended, so no proxy from one Excel survives in a frame while the next Excel starts (the rule at autofix.py:2426-2429). The progress callbacks only update `s.apply_status`, a plain dict, as today. The fallback path runs today's excel_session on the request thread, after the worker has dropped its Excel.

## When Excel dies or a write fails

EXCEL DIES.
- Detected by excel_gone(exc), by `excel_alive()` failing after a job, or by the job timeout (the pid is killed first, so a hung Excel cannot block).
- The worker drops everything: kill the pid if it is still alive, set proxies to None, gc, quit in a try, delete the held copy (a half-written copy is never registered, because `_add_version` runs only after apply_operations returns).
- The waiting caller gets WorkerUnavailable.

The server then runs today's code on the request thread:
- Apply: `apply_operations(s.current_path, s.work_dir/'apply', ops, progress, defer_verify)`. That is a fresh excel_session on the last saved version, with its own openpyxl fallback and today's verify_opens_in_excel, deferred or inline, with the 5 stages after a reset of `_apply_start`.
- Recalculate: make_immutable_copy plus recalculate.
The next job starts the worker's Excel again. The fallback is recorded in perf as 'excel_worker_fallback'.

A WRITE OR SAVE FAILS WHILE EXCEL IS ALIVE (NotSavedError, read-only open, a COM error at save).
raise_on_excel_error makes the job raise. The same fallback then gives today's result: a retry in a fresh Excel, then openpyxl with its warning. A single op that Excel refuses is not a failure: it goes into `failed` exactly as now.

THE PRE-OPEN FAILS.
- The copy does not open: opens=False, the label says 'FAILED TO OPEN' as today, and nothing is held.
- Excel will not start: opens=None, 'unverified' as today.
- The next job does a cold open inside the worker.

A HELD COPY IS STALE OR UNSUITABLE.
It is closed and its copy deleted, and the job opens cold. This covers: a version switch through /reanalyze or History, explicit_colors ops, data connections, low free RAM, the idle timeout, and release() before another Excel user.

THE SPECULATIVE RECALCULATION FAILS.
No cache entry is kept, and /recalculate runs as it would without speculation.

MINDREADY_EXCEL_WORKER=0 turns the worker off (today's code byte for byte). MINDREADY_SPECULATIVE_RECALC=0 turns off only the speculation.

## Steps

### `excel-upload-preparation/app/excel_com.py`

1. Split the body of excel_session into two functions; excel_session calls them inside its own CoInitialize/CoUninitialize, so behaviour is unchanged:
- `start_excel(mm_addin=True)`: DispatchEx, Visible/DisplayAlerts/AutomationSecurity=3/AskToUpdateLinks/EnableEvents, load_mm_addin, inside stage('excel_start').
- `quit_excel(excel)`: gc.collect, Quit in a try, gc.collect.
2. Add helpers:
- `excel_pid(excel)`: win32process.GetWindowThreadProcessId(excel.Hwnd)[1].
- `kill_excel(pid)`: errors swallowed.
- `excel_gone(exc)`: hresult in {-2147417848, -2147023174, -2147023170, -2147417851, -2146959355}.
- `excel_alive(excel, pid)`.
3. Add `open_check(excel, path) -> (dict, wb|None)`. It opens the file with open_for_write and returns the verify_opens_in_excel dict shape:
- an open that raises: opens=False and the error text;
- NotSavedError because Excel opened it read-only: opens=True, no workbook;
- Excel would not start: available=False, opens=None.
verify_opens_in_excel(path) is unchanged; tests patch it as `lambda path:`.

### `excel-upload-preparation/app/inventory.py`

Pull the path choice out of make_immutable_copy into `next_copy_path(work_dir, name) -> Path`: work_dir/name, else the first free vN/name, with the folder created. make_immutable_copy calls it, so its behaviour is the same.

### `excel-upload-preparation/app/excel_worker.py (new)`

1. Types:
- `WorkerUnavailable(RuntimeError)`.
- `HeldCopy(source_path, source_stamp, source_sha256, version_id, copy_path, wb)`.
- `ExcelRun(excel, held, verify, raise_on_excel_error=True)`.
2. `ExcelWorker(sid, apply_dir, idle_s)`:
- One daemon thread named `excel-{sid}`, a queue.Queue of (fn, Future), and a `_Ctx` (excel, pid, held, recalc_cache) that is used only on that thread.
- `_Ctx` records the owner thread ident and asserts it on every access.
3. Thread loop:
- CoInitialize once.
- `get(timeout=1)` plus PumpWaitingMessages.
- After idle_s with no job: release.
- Each job runs inside try/finally. The finally drops the job's locals and runs gc.
- After each job, if excel_gone(exc) or not excel_alive: abandon (kill the pid FIRST, set proxies to None, gc, quit_excel in a try, delete the held copy), and the Future gets WorkerUnavailable.
- Any other exception is passed through unchanged.
- `_submit(fn)` waits with future.result(timeout=MINDREADY_EXCEL_JOB_TIMEOUT_S). On timeout it kills the pid, which turns the hang into WorkerUnavailable.
4. Public methods take and return plain values only:
- `apply(source, ops, progress, defer_verify)`, as one job:
  - When any op is explicit_colors, or held does not match (source path, _file_stamp, version id): close held and delete its copy.
  - Start Excel if none is running.
  - Call prep.apply_operations(source, apply_dir, ops, progress=progress, defer_verify=True, run=ExcelRun(ctx.excel, held, self.preopen)).
  - Then set ctx.held = None and ExcelRun.excel = None.
  - After the job, on the caller's thread: if not defer_verify and 'finish' is in res, call res.pop('finish')(). That queues preopen as its own job, writes the change log and sets the flag: today's non-deferred shape.
- `preopen(version_path, version_source=None) -> dict`, the check, as one job:
  - close held, quit_excel, gc, start_excel;
  - copy, sha = make_immutable_copy(version_path, apply_dir);
  - info, wb = open_check(...).
  - Keep the workbook only if all of these hold: opens is True, wb is not None, wb.Connections.Count == 0, and free RAM (GlobalMemoryStatusEx) >= MINDREADY_EXCEL_HOLD_MIN_FREE_MB (default 3000).
  - Otherwise close it, delete the copy and quit Excel.
  - If version_source == 'loop' and speculative recalculation is on and the copy is held, queue `_speculative_recalc` right behind this job.
  - Return info.
- `_speculative_recalc()`, a job:
  - res = recalc.recalculate_in(ctx.excel, held.copy_path, wb=held.wb);
  - rename the copy to inventory.next_copy_path(work_dir/'recalc', name);
  - recalc_cache = (held.source key, res, new path);
  - held = None; quit Excel.
- `preopen_async(path)`.
- `recalculate(source, version_id, recalc_dir) -> (path, res)`:
  - A cached result for this source key: return it.
  - Held matches: recalculate_in on the held workbook, then the same rename.
  - Otherwise: close held, start a fresh Excel, make_immutable_copy(source, recalc_dir), recalculate_in.
  - In every case Excel is quit afterwards.
  - A non-death exception becomes _not_run('ERROR', f'Excel COM recalculation failed: {exc}'), today's text.
- `release()`, `stop()`.
5. Write the pids to %LOCALAPPDATA%/MindReady/excel_worker_pids.json and kill stale ones at server start.
6. Module-level `stop_all_workers()`, registered with atexit.

### `excel-upload-preparation/app/prep.py`

1. `_apply_with_excel(copy_path, operations, progress=None, *, excel=None, wb=None)`:
- `_run(excel)` uses the `wb` it is given instead of calling open_for_write (progress text: 'The copy is already open in Excel').
- Everything after the open is unchanged: calc mode, _write_one with retries, restore, save_in_place, and close_quietly in finally.
- If excel is None: today's `with excel_session()`.
2. `apply_operations(..., defer_verify=False, run=None)`:
- If run and run.held match: copy_path, source_sha256 = held.copy_path, held.source_sha256, still emitting the 'apply_copy' stage. Otherwise make_immutable_copy as now.
- With run=None the write call stays exactly `_apply_with_excel(copy_path, operations, progress)` (the patch at line 528 matches). With run: excel=run.excel, wb=held.wb.
- At 2235: when run.raise_on_excel_error is set, re-raise instead of falling back to openpyxl.
- `verify = run.verify if run else verify_opens_in_excel`, used at 2250 and 2279.
- The 'apply_verify' message when run is set: 'The next step's copy is opened in Excel alongside the re-analysis; that open is the check'.
- Hashes, entry, finish, stages and result shape are unchanged.

### `excel-upload-preparation/app/recalc.py`

1. `_scan(excel, copy_path, save=True, wb=None)`: uses the given wb, otherwise open_for_write. The rest is unchanged.
2. New `recalculate_in(excel, copy_path, save=True, wb=None)`: the same file checks and stage('recalculate') around _scan, with exceptions raised to the caller.
3. recalculate() is unchanged.

### `excel-upload-preparation/app/web/server.py`

1. `WORKERS` and `_worker_of(s)`: None when MINDREADY_EXCEL_WORKER=0 or not com_available(); otherwise created lazily with apply_dir=s.work_dir/'apply'.
2. `_tracked_apply`: `res = w.apply(s.current_path, ops, progress, defer_verify)`.
- `except WorkerUnavailable`: perf_record('excel_worker_fallback'), `_apply_start(s, len(ops), reanalyze)`, then today's apply_operations call on the request thread.
- `_apply_and_version` stays as it is, except that `finish` must call preopen with the version's source. So when res has 'finish', wrap it as `lambda: finish_with(source)`, where ExcelRun.verify is a functools.partial(preopen, version_source=source) built by w.apply from a `source` argument. That means w.apply takes `source` too.
3. `recalc` (1117): `copy_path, res = w.recalculate(s.current_path, s.current_version_id, s.work_dir/'recalc')`. On WorkerUnavailable, today's lines 1126-1127. RECALCULATING and _recalc_out are unchanged.
4. `start_autofix.worker()`:
- `w.release()` before compare_workbooks/run_autofix.
- After `_recalc_out(...)` and before `_run_scan` (1334): `w.preopen_async(output)`.
- When res has no output_path (nothing changed): `w.preopen_async(source)`.
- The fixer's own verify at autofix.py:2454 is unchanged.
5. `w.release()` before numbers_check (1229), POST /mind-loop run_loop (889), reports (1655) and _convert (528), for memory only.
6. Add on_event('shutdown') calling stop_all_workers.
7. `/api/health`: `excel_worker: bool`.

### `excel-upload-preparation/app/estimate.py`

After the first Excel step of a session, estimate_chain counts no Excel start or open for Prep rounds and the loop. When a loop is planned, it counts Recalculate as about 5 s. Saves stay on the critical path.

### `excel-upload-preparation/tests/conftest.py + tests/unit/test_excel_worker.py (new) + CHANGELOG.md (1.11.8), README, VERSION`

1. Add an autouse fixture that calls server.stop_all_workers() after each test.
2. Add the new tests listed below.
3. CHANGELOG: 'One Excel per session: the opens-in-Excel check is the open of the next step's copy, made in a fresh Excel alongside the re-analysis; after the Mind loop the recalculation runs there too; MINDREADY_EXCEL_WORKER=0 restores the old path'.
4. Replace the 1.11.4 'verify alongside' line with a pointer to this one.

## Why the results stay identical

1. Same input bytes. The workbook each step writes into is `make_immutable_copy(previous version, work_dir/'apply')`: the same function, directory and path rule as prep.py:2225 and inventory.py:195-219. Its recorded source hash is the one computed by that copy. A held copy is used only when the source path, mtime_ns, size and version id all still match. Version files are never written after they are registered.

2. Same Excel state at the first write. Each held copy is opened by `open_for_write` (the same `Workbooks.Open` arguments, excel_com.py:168), in an Excel freshly started by `start_excel`:
- the same settings and the same RegisterXLL, so `_xll.MM_LOOP` and MM_RESULT are written as now (test_web_api.py:674-704);
- with no other workbook open.
Between the open and the write the workbook only sits open. Workbooks with data connections are never held, so no refresh can run.

3. The same code after the open. `_apply_with_excel._run` (calculation mode, `_write_one` with retries, restore, `save_in_place`, `close_quietly`) and `recalc._scan` (`CalculateFullRebuild`, `save_in_place`, error read) are reused, not copied. Only the open call is skipped when a workbook is handed in.

4. Same version list.
- `_add_version` is called with the same arguments in the same order: ids, parents, major/minor, sources, change logs, and download names (`_register_file` is keyed by name).
- File paths: the pre-open takes apply/<name>, then apply/vN/<name>, exactly as the next apply would. A recalculated copy is renamed back to recalc/<name> or recalc/vN/<name>, and dropped copies are deleted, so later paths do not move.
- The labels keep 'verified / unverified / FAILED TO OPEN', filled in before `checker.join()` from an answer in the `verify_opens_in_excel` shape.
- The difference: the check is a read/write open of a byte-identical copy rather than a read-only open of the version file. Same bytes, same `Workbooks.Open` call, so a file that fails one fails the other. The one exception, a password to modify, gives opens=True with nothing held, which matches today's label.

5. Same re-analysis input. `_run_scan(s, output)` reads the file Excel saved and closed.

6. Recalculate. The speculative run is `_scan` itself, on a byte-identical copy of the same version, in a fresh Excel with the XLL. It is the same work as today's /recalculate, done earlier. `_recalc_out` is unchanged:
- no version is added after a major version (server.py:1173);
- the download is still the loop version;
- `s.recalc_path` is a recalculated copy in recalc/.
Time-volatile functions (NOW, RAND) differ between any two runs today as well.

7. Proof: the tests below run the chain with the worker off and on, and compare every cell formula and cached value, the defined names, calcPr, the recalculation error lists and the version lists.

## Risks

- Memory: the held Excel keeps the 179 MB workbook (about 1.5-3 GB of RAM) loaded during the scan, and until the next step or the idle timeout. Today's check Excel lives about 80 s. Mitigations: the free-RAM gate (MINDREADY_EXCEL_HOLD_MIN_FREE_MB), MINDREADY_EXCEL_IDLE_S (900 by default), and release() before other Excel users. With speculative recalculation after the loop, the workbook is let go about 2 minutes into the 380 s scan.
- The pre-open, and after the loop the recalculation (CalculateFullRebuild plus a 42 s save), compete with the openpyxl scan for CPU and disk. Compare the 'scan' timings before and after. If the scan slows by more than it saves, set MINDREADY_SPECULATIVE_RECALC=0.
- Wasted speculation: after a single-button Mind loop, the user may never press Recalculate, or may run another Prep round, which then opens cold as today. That costs about 50 s of hidden Excel work and one stray file in recalc/. The stray file only moves the disposable recalc path.
- explicit_colors (FMT-002) ops force a cold open, because the style-table zip edit (prep.py:2231) must happen before Excel opens the file. That round saves nothing. Do everything never sends them; the action is optional.
- Hung Excel: the job timeout kills the pid. If MINDREADY_EXCEL_JOB_TIMEOUT_S is too short, a legitimately slow 179 MB save could be killed. Default 1800 s, and log the stage reached when it fires.
- Orphan EXCEL.EXE after a hard kill of the server lasts until the next start, which kills the pids it recorded. A daemon thread at interpreter exit is covered by atexit, and pytest by the conftest fixture.
- The stage list at test_web_api.py:96 depends on the worker path emitting apply_copy, apply_write, apply_save, apply_verify and apply_log once each. On fallback, _apply_start resets the list first.
- In the non-deferred case (reanalyze=false), the check now runs as a second job after the apply job instead of inline. The response shape is the same, but the change log is written by finish(): check that test_web_api.py:78 and the change-log readers still see a written sidecar.
- Several sessions can each hold an Excel. That is one per session, but memory adds up, and the idle timeout limits it.
- Not fixed, flag to Joseph: runAll decides the loop step from the loop plan loaded with the session (home.page.ts:217/322/676). night_chain re-reads it after Prep.

## Tests

- Unit, fake COM, no Excel. Monkeypatch start_excel, open_check and quit_excel with recorders of threading.get_ident(), then check:
- every COM call happens on excel-{sid};
- jobs run in FIFO order, and the speculative recalculation is queued directly behind its pre-open;
- the idle timeout calls release;
- an RPC HRESULT gives WorkerUnavailable, the pid is killed before quit, the held copy is deleted, and the next job starts Excel again;
- a job timeout kills the pid.
- Unit: apply_operations(..., run=ExcelRun(fake, held, verify)):
- it uses held.copy_path and held.source_sha256 and never calls make_immutable_copy;
- it emits the 5 stages;
- it re-raises on an Excel error;
- finish() calls run.verify.
With run=None, the test_readiness_and_prep_fixes.py:528 monkeypatch and test_pivot_guard.py:115 / test_reference_args.py:121 behave as before.
- Unit: worker.apply with defer_verify=False returns res without 'finish', with verified_opens_in_excel set and the changelog sidecar written, and the pre-open ran as a separate job after the apply job's proxies were dropped (a recorder checks the fake proxy was released before start_excel was called again).
- Unit: _tracked_apply with a worker that raises WorkerUnavailable calls the plain apply_operations once, after _apply_start. The version list equals the no-worker run.
- Unit: a warm recalculation and a speculative one put the copy at the next_copy_path in recalc/, and the next apply takes the same apply/vN path as without the worker. /recalculate uses the cache only when the version id and stamp match.
- Existing tests stay green with MINDREADY_EXCEL_WORKER=1 and =0:
- test_web_api.py :58-80, :83-99, :506-524, :527-564, :567+, :674-704;
- test_prep_and_chat.py 118/348/439-455/647;
- test_theme_colors.py:191;
- test_change_apply.py:74.
- needs_excel, identity: on flagged_model_broken.xlsx (2 Prep rounds) and the xlsm loop fixture, run upload, Prep, Prep, the Mind loop and Recalculate through TestClient with the worker off and on. Assert equality of:
- every cell formula and cached value (openpyxl, data_only False and True);
- defined names, sheet order and visibility, calcPr;
- xl/vbaProject.bin and the `_xll.MM_LOOP` text;
- the recalculation result (status, formula_errors, addin_gap_errors, mmforexcel_loaded);
- READY-001;
- the version lists without sha256 and created_at;
- `compare_workbooks(final_off, final_on)` reporting 0 moved values.
- needs_excel, one Excel: during that chain, poll tasklist every 0.5 s. At most 1 EXCEL.EXE is alive per session, and every excel_start / excel_open perf record is on an excel-* thread (none on verify-*).
- needs_excel, fallback: kill the worker's Excel pid after a pre-open, then POST /apply. It returns 200 with the same applied ops, 'excel_worker_fallback' appears in perf, and the next apply is warm again.
- needs_excel: a workbook saved in manual calculation with iteration on keeps calcMode and iterate after a warm apply. A workbook with a data connection is not held.
- Acceptance on Phoenix (Joseph, at night): run scripts/night_chain.py with MINDREADY_EXCEL_WORKER=0 and =1.
- Diff the finals (formulas and values), the version labels and chain.json.
- In timings.jsonl, the pre-open excel_open records and the recalc_full_rebuild after the loop overlap the scans.
- apply_operations for Prep is about 51 s and for the loop about 55 s; recalc_endpoint is a few seconds.
