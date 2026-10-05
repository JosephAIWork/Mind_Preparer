# Auto-fix loop — unattended run log

Started **2026-10-05 15:46 (+03:00)**, hard stop **2026-10-06 03:46**. Branch `auto-fix-loop`
(from `fix-panel-progress-and-versions` @ dfd6ab2), pushed to `origin` (JosephAIWork/Mind_Preparer) only.

**If this run was interrupted: read this file top to bottom, then continue from "Next step".**

## The task (Joseph, 2026-10-05)

- Use cases: the four workbooks in
  `…\Tel Aviv Office - 060 R&D\70 GenAI Exploration\60 Mind Excel Prep\2026 09 - Models for testing`
  (RAROK is ignored). Local copies: `%LOCALAPPDATA%\MindReady\autofix\usecases\{horizon,pvfp,palermo,cnhi}.xlsm`.
  The originals are never touched.
- Run each one until **all the formulas come back clean**, with nobody helping.
- Edit the app as needed, document the edits, re-run the other use cases after any significant edit.
  Finish only when all 4 are tested on the latest version of the app (max 12 h).
- The feature: one click and the app goes through the formula errors by itself and picks the most
  appropriate fix. It is an option; the manual Fix panel stays.

## Decisions Joseph made before leaving

1. **Clean = zero errors at all** after a full Excel recalculation, including errors the original file already had.
2. **Never change a good number.** A fix that changes a cell that is fine today is undone; another way is tried;
   if nothing works the error stays and is listed for a person.
3. **Full size first.** Shrink a model only if one recalculation takes over ~20 minutes, and say what was cut.
4. **New branch, a commit after every step, pushed to the personal GitHub.** The team repo is not touched.

Also agreed: Excel only (nothing uploaded to Mind); port 8600 is left alone during the run (tests use 8601);
at the end 8600 is restarted on the final version.

## Baseline (original files, full Excel recalculation, macros off, links not updated)

| Use case | Size | Cells with a value | Formulas | Error cells | Full recalculation | Reading every cell |
|---|---|---|---|---|---|---|
| Horizon | 1.6 MB | 185,679 | 22,120 | 721 (623 #VALUE!, 54 #REF!, 44 #DIV/0!; 2 are constants) | 0.5 s | 3 s |
| PVFP | 3.8 MB | 156,954 | 152,324 | 65,883 (65,820 #REF!, 58 #N/A, 5 #VALUE!) | 3 s | 7 s |
| Palermo | 28 MB | 4,035,852 | 56,647 | 511 (507 #N/A, 4 #VALUE!) | 1 s | 24 s |
| CNHI | 37 MB | (pending) | | | | |

Every error above is already in the file as saved (none is created by recalculating here).
PVFP's errors come from a dead external link (`C:\Users\adam.senio\…\[PVFP Inputs_v4.0.xlsx]`), INDIRECT to
sheets that do not exist, and a 3-D reference `'LoB 1:>>'`.

## Design (app/autofix.py)

One Excel session per run: open a copy, recalculate, read every formula cell, then loop:
find the *root* error cells (an error whose own inputs are fine), choose a fix for each group of
same-formula roots, write the fixes, recalculate, and compare every other formula cell with its value
before. A fix that changed a good number is rolled back and the next strategy is tried.
Stops when no error is left, or nothing more can be fixed safely.

## Log

- 15:46 branch + local copies. 15:50-16:10 baselines of the four originals (table above; CNHI: 577,707 formulas, 0 errors).
- 16:10-16:25 `app/autofix.py` first version + tests; Horizon original: 721 -> 0 errors in 2 passes, 37 s.
- 16:25-17:20 PVFP original (65,883 errors). Found and built, one by one:
  recurrences that fail row after row are taken as a block; a dead range (`SUMIFS(#REF!, ...)`) is replaced by a
  value; **errors that a formula swallows on purpose** (`IFERROR(<error>, 0)`, "No GEP") cannot be fixed without
  changing that formula's value, so they stay, with the error cells between them and that formula pinned at once
  (was: one level per pass, hours); the walk back from a moved value gives an exact verdict to a fix that moved it
  alone; INDIRECT / OFFSET targets are asked from Excel through probe formulas in an empty column
  (Worksheet.Evaluate fails for INDIRECT inside a function; a scratch *sheet* breaks GET.WORKBOOK-style names).
  **PVFP original, strict rule: 65,883 -> 887 errors in 3 min 18 s, 21 passes, no good value changed.**
  The 887 left all feed 8 formulas that hide errors (LoB 1!C211, LoB 1!K139 "No GEP", Reporting_LoB 1!BU17, ...).
- 17:00 web API (`POST/GET /api/sessions/{id}/auto-fix`, `/stop`), `AutoFixPanel` on the Recalculate screen,
  `app/autofix_advisor.py` (the assistant orders the fallbacks and words the reason). Tests: 19 engine + 1 API.
- 17:18 Prep baseline through the app (old code, port 8601): Horizon 3 Prep rounds, 719 errors after;
  PVFP: analysis 14 min, Prep round 1 37 min (77k colour cells), 64,127 errors after Prep (65,883 before).
  The app's own analysis is the slow part, and one server process held 7.7 GB after two workbooks:
  the harness now starts a server per workbook and leaves a sheet that is >= 50 % of an oversized file out of the
  *scan* (Palermo `Data`, CNHI `Output`), as the upload screen offers.

- 17:20-17:30 engine on the other originals: **Palermo 511 -> 0 in one pass (56 s)**; PVFP with the rule
  lifted (`keep_good_values: false`): clean in 29 s, 421 hidden values changed and listed.
  UI click-through with Playwright on Horizon (`scripts/ui_autofix_check.py`): 721 -> 0, verdict green.
  It showed a race (Recalculate still running when the fixer was started -> verdict "not verified" on a
  clean version): Recalculate / Apply are now refused while the fixer runs, and the reverse.
- 17:30-18:00 **Prep changes numbers on PVFP.** `scripts/compare_versions.py` on the baseline's Prep output:
  2,462 computed values changed (1,781 into errors). Cause: title rows inserted on sheets read by position
  (3-D reference `'LoB 1:>>'!K65`; INDIRECT addresses built as text). Built: `app/numbers_check.py` +
  `POST /numbers-check` + "Check the numbers" on the Prep screen; `prep.insert_blocker` (no row insert on a
  sheet inside a 3-D range or named by an INDIRECT); the fixer compares with the original first and never
  gives a fallback to an error that was a value in the original.
  With the guards: PVFP after two Prep rounds (22 row inserts, 17 held back) -> **no computed value changed**.
  Horizon and Palermo (old Prep): no computed value changed either (Palermo's one difference was the
  workbook's own folder in a `CELL("filename")` cell -- now ignored).
- 18:10-18:25 `explicit_colors` by blocks (was ~10 COM calls a cell: 37 min on PVFP, >35 min on CNHI);
  dead formulas become `=IFERROR(NA(),0)` instead of a bare value (SKILL.md rule 6); version 1.8.0,
  CHANGELOG, README. Full unit suite under heavy load: 267 passed, 7 Excel-recalculation tests failed
  and passed when re-run alone (Excel COM under load) -- to be run again when the machine is quiet.
- 18:29 **full run 2** started on commit 128f9c3: `scripts/run_usecases.py --ports 8602,8603`
  (cnhi, palermo, pvfp, horizon) -> `%LOCALAPPDATA%\MindReadyutofixunsull_2`.

## Results so far (engine alone, on the original files, the owner's two rules)

| Use case | Errors | After | Passes | Time | Left for a person |
|---|---|---|---|---|---|
| Horizon | 721 | 0 | 2 | 37 s | - |
| Palermo | 511 | 0 | 1 | 56 s | - |
| CNHI | 0 | 0 | - | - | nothing to fix |
| PVFP | 65,883 | 887 | 21 | 3 min 18 s | 887 cells that feed 8 formulas hiding errors on purpose (`=IFERROR(K24/K12,"No GEP")`...) |
| PVFP, rule 2 lifted | 65,883 | 0 | 2 | 29 s | 421 hidden values changed, listed |

## Next step

1. Wait for full run 2 (`runs/full_2/run.log`, `summary.json`); `python scripts/usecase_report.py <run dir>`
   -> `docs/AUTOFIX_USE_CASES.md`.
2. Re-run the full unit suite with nothing else running.
3. Copy the final workbooks + reports to `excel-upload-preparation/runs/`, restart port 8600, final commit.
