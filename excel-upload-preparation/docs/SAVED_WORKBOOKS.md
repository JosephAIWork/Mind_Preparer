# Saved workbooks (1.14.0)

The app keeps every workbook it was given, with every version and result, so a restart of the app or a
new browser window loses nothing. The main screen's **Recent workbooks** list opens one again where it
was.

## What the user sees

- Under the upload box, **Recent workbooks**: one line per saved workbook -- its name, when it was last
  opened, how many versions it has ("not checked yet" for an upload that was never scanned) -- with
  **Reopen** and **Remove**. The list is drawn only when at least one workbook is saved.
- **Reopen** loads the workbook with everything it had: every version (the current one stays current),
  the last check result, the plan, the last recalculation (Advanced > Recalculate shows its list without
  recalculating), the Mind loop and instances plans, the outside files and their bring-in result, next
  quarter's files, the downloads. The page then shows the same screens as after the workbook's last step.
  A workbook uploaded but never checked goes where an upload goes: the sheet choice for a big one, else
  straight to the scan. Nothing runs by itself: "Do everything right after the upload" is for uploads.
- **Remove** asks once, in place -- "Remove <name> from this computer?" Yes / No -- then deletes the
  workbook's folder (every version and result). Refused while a step still runs on that workbook
  (Run in Mind included).
- On start the app shows the list and waits: nothing reopens by itself. Any call about a workbook the
  restart closed (a tab still on `/?session=<id>`, a button pressed on an old page) answers "This
  workbook is not open any more (the app was restarted). Open it again from Recent workbooks on the
  main screen.": the page lets go of the workbook and the main screen shows that sentence above the
  Recent workbooks list, with a Show me to the list.
- A job that was running when the app was closed (Fix all issues, Bring in the outside files, next
  quarter's files, a scan) shows as stopped after the reopen, with "The app was closed while this ran."
  What it had finished before (a version it made) is there; the Fix all issues tile on the main screen
  says "Stopped: the app was closed while it ran" and names that version.
- A workbook uploaded but never checked is scanned after Reopen without starting Do everything,
  whatever "Do everything right after the upload" says (that box is for uploads).
- **Two copies of the app on the same data folder** (two ports, say 8600 and 8601): the copy that holds
  a workbook open keeps it; the other copy's list leaves it out, and its Reopen / Remove say "This
  workbook is open in another copy of the app (port 8600). Use it there, or close that copy of the app
  first." Once that copy is closed the workbook is listed again.

## Where the files are

`<data folder>\sessions\<session id>\` -- the data folder is `%LOCALAPPDATA%\MindReady` unless
`settings.yaml` says otherwise (`data_folder`). Inside, the same layout the %TEMP% folder had before
1.14.0: `upload\` (the original), `inputs\<nn>\` (the picked input files), `analysis\<version>\`,
`apply\`, `recalc\`, `outside_inputs\`, `instances\`, `reports\`, and so on. Next to them `state.json`
holds the session's state (below), `result.json` the last check result, `plan.json` the plan, and `open.lock` is locked by
the copy of the app that holds the session (the OS releases it when that process ends). Nothing is
written to %TEMP% any more; when the data folder cannot be written, the session falls back to a %TEMP%
folder for that run only and is not listed.

Remove deletes the folder. A file still held open by an Excel the app did not release leaves the folder
behind without its state file: it is never listed again and goes at the next Remove.

## The state file

`app/session_store.py` writes `state.json` after every step that changes the session: the upload, a
scan (analyze, reanalyze), an Apply (Prep, the assistant, a formula fix, the Mind loop, the Grid Namer),
the end of Fix all issues, the end of Bring in the outside files, the end of next quarter's files, a
recalculation, the pivot setting, the assistant's grid names, the model map and the instance finder once
computed, the instances build (its zip), the reports, added input files, and each reopen (last opened).

It holds everything the `Session` dataclass (`app/web/server.py`) can save: the versions, their files
and the downloads (paths relative to the session's folder, so the data folder may move), the current
version, the plan, the delta, the recalculation (and the version it belongs to), the
grid names, the size gate, the skipped sheets, the pivot setting, the Prep and Apply gauges, the
baselines (sets of (sheet, cell) as lists), the original's CELL/INFO values (tuple keys as lists), the
model map, the picked input files, what the macros open, the bring-in and next-quarter results, the Mind
inputs folders, the instance finder's result (and the assistant's block), the instances plan and zip,
the name shown in the list, when it was uploaded and last opened. The automatic fixer's state (kept next
to the sessions in the server) is saved with it.

The two big objects are not in the state file. The check result (`result`: the whole analysis, every
formula -- 19 MB of JSON on the anonymized IFRS 17 tool) goes to `result.json`, written only when a scan
replaced the result object; the plan (`plan`: every formula fix as an operation -- 61 MB on that tool)
goes to `plan.json`, written only when a scan or a re-plan (the pivot setting, the grid names) replaced
it. An Apply without a re-scan, a reopen, a recalculation rewrite the small state file alone (under
1 MB on that tool), and the Recent workbooks list reads that file alone. Both are read back on Reopen.
A state file of the first shape, with the result or the plan inside, still opens.

Locks, threads and live job state are not saved. A state saved while a job was running (`"running"`)
comes back as stopped (`"error"`, "The app was closed while this ran.").

A save never breaks the step it follows: a failure is printed and the step answers as before. The file
is written whole and then renamed, so a stop in the middle of a save leaves the previous file.

## Routes

- `GET /api/saved-workbooks` -- `{workbooks: [{id, name, uploaded_at, last_opened, versions,
  current_version, bytes, has_result, pending}]}`, the last opened first. Reads the state files; loads
  nothing.
- `POST /api/saved-workbooks/{id}/open` -- loads the session into memory under the same id (when it is
  not there already) and answers as `GET /api/sessions/{id}` does, plus `recalc` (the last
  recalculation, when it is the current version's). A workbook never checked answers in the shape of a
  deferred upload (`pending: true`). 404 when the id is not saved here; 409 when another copy of the
  app holds it.
- `DELETE /api/saved-workbooks/{id}` -- removes the folder and the session in memory; 409 while a step
  runs on it (Run in Mind too) or another copy of the app holds it; 404 when there is nothing of that id.
- The list leaves out a workbook another copy of the app holds (`session_store.held_elsewhere`: a try
  at the lock of `open.lock`, which that copy keeps locked while the session is in its memory).
- `GET /api/sessions/{id}` for an id saved on disk but not in memory answers 404 with the sentence above
  (nothing reopens by itself; the simplest choice -- a later version may open it instead).

The size gate and the expiry of a shared copy apply as before.

## Tests

`tests/unit/test_session_store.py`: a checked workbook (two versions) comes back after a restart --
versions, current version, report, plan, verdict, gauge, download, and a further Apply on it; Remove
deletes the folder and the session; Remove waits for a running step; a session saved while the fixer, a
bring-in and a scan ran shows them stopped; an upload never checked reopens as pending and scans; the
list order and a stray folder; a failed save never breaks the step; every odd shape of the state file
(sets, tuple keys, paths inside and outside the folder) round-trips; the check result has a file of its
own written only when it changes (and the first shape still opens); a workbook another copy of the app
holds is left alone; Remove waits for Run in Mind. Screens:
`MindReadyUI/src/app/core/recent-workbooks.spec.ts` (the list hidden when empty, Reopen loads, Remove
asks then removes; the main screen opens a saved workbook where it was, never starts Do everything
after a Reopen, says that a Fix all issues the restart cut was stopped), `core/restart.spec.ts` (a 404
for a workbook the restart closed lets go of it and keeps the sentence for the main screen) and
`show-me.spec.ts`.
