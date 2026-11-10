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
  workbook's folder (every version and result). Refused while a step still runs on that workbook.
- On start the app shows the list and waits: nothing reopens by itself. A browser tab still on
  `/?session=<id>` after a restart gets "This workbook is not open any more (the app was restarted).
  Open it again from Recent workbooks on the main screen." (with a Show me to the list) and goes back
  to the main screen.
- A job that was running when the app was closed (Fix all issues, Bring in the outside files, next
  quarter's files, a scan) shows as stopped after the reopen, with "The app was closed while this ran."
  What it had finished before (a version it made) is there.

## Where the files are

`<data folder>\sessions\<session id>\` -- the data folder is `%LOCALAPPDATA%\MindReady` unless
`settings.yaml` says otherwise (`data_folder`). Inside, the same layout the %TEMP% folder had before
1.14.0: `upload\` (the original), `inputs\<nn>\` (the picked input files), `analysis\<version>\`,
`apply\`, `recalc\`, `outside_inputs\`, `instances\`, `reports\`, and so on. Next to them `state.json`
holds the session's state (below). Nothing is written to %TEMP% any more; when the data folder cannot
be written, the session falls back to a %TEMP% folder for that run only and is not listed.

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
version, the check result, the plan, the delta, the recalculation (and the version it belongs to), the
grid names, the size gate, the skipped sheets, the pivot setting, the Prep and Apply gauges, the
baselines (sets of (sheet, cell) as lists), the original's CELL/INFO values (tuple keys as lists), the
model map, the picked input files, what the macros open, the bring-in and next-quarter results, the Mind
inputs folders, the instance finder's result (and the assistant's block), the instances plan and zip,
the name shown in the list, when it was uploaded and last opened. The automatic fixer's state (kept next
to the sessions in the server) is saved with it.

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
  deferred upload (`pending: true`). 404 when the id is not saved here.
- `DELETE /api/saved-workbooks/{id}` -- removes the folder and the session in memory; 409 while a step
  runs on it; 404 when there is nothing of that id.
- `GET /api/sessions/{id}` for an id saved on disk but not in memory answers 404 with the sentence above
  (nothing reopens by itself; the simplest choice -- a later version may open it instead).

The size gate and the expiry of a shared copy apply as before.

## Tests

`tests/unit/test_session_store.py`: a checked workbook (two versions) comes back after a restart --
versions, current version, report, plan, verdict, gauge, download, and a further Apply on it; Remove
deletes the folder and the session; Remove waits for a running step; a session saved while the fixer, a
bring-in and a scan ran shows them stopped; an upload never checked reopens as pending and scans; the
list order and a stray folder; a failed save never breaks the step; every odd shape of the state file
(sets, tuple keys, paths inside and outside the folder) round-trips. Screens:
`MindReadyUI/src/app/core/recent-workbooks.spec.ts` (the list hidden when empty, Reopen loads, Remove
asks then removes; the main screen opens a saved workbook where it was) and `show-me.spec.ts`.
