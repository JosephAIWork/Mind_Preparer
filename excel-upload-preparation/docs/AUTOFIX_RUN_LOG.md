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

## Next step

1. Engine on Palermo original and on the Prep outputs (`runs/prep_baseline/*`), then the UI click-through.
2. Full runs of the 4 use cases with `scripts/run_usecases.py --ports 8602,8603` on the latest code.
3. CHANGELOG / README / VERSION 1.8.0, final runs, restart port 8600.
