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
PVFP's errors come from a dead external link (`C:\Users\<user>\…\[PVFP Inputs_v4.0.xlsx]`), INDIRECT to
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
- 18:30-18:45 the assistant may propose the repair itself (`=IF(BF36=0,0,BE36/BF36)`, `=C4+N(B5)`): written to the
  whole block, kept only if the error cells are cured, every working cell of the block returns exactly the same value
  and the numbers gate passes. Live on Horizon: 2 repairs kept (46 cells), 721 -> 0.
  Two full runs were started and stopped on purpose (18:29, 18:43): each time an edit that mattered came up while
  they were in their first minutes, and "tested on the latest version" has to be true for all four.
- 18:45-19:15 `repair_refused()`: what an assistant-written formula may contain (only what the original reads and
  calls, plus a short list of guards; no DDE, no other workbook).
  **A COM bug of mine**: 8 tests of the full suite failed (each passed alone) with "The interface is unknown" on the
  first recalculation after an auto-fix in the same thread. Cause: `run_autofix` still held the Excel proxy of its
  session while the next session (`verify_opens_in_excel`) ran. The session now lives in `_fix_in_excel()`.
  **Full unit suite: 281 passed** (247 before this run).
- 19:15 **final run (`full_3`)** started on commit 5f1dafa: `scripts/run_usecases.py --ports 8602,8603`
  (cnhi, palermo, pvfp, horizon) -> `%LOCALAPPDATA%\MindReady\autofix\runs\full_3`.

- 19:15-20:24 `full_3`, Palermo: 3 Prep rounds (345 changes), numbers check "the preparation changed no computed
  value (56,647 formula cells compared)", 511 errors -> **0** in 1 pass, confirmation recalculation PASS (v5).
- 20:45 **`full_3` cannot finish, and why.** The Prep action "Replace theme colours with explicit RGB" (FMT-002,
  ticked by default) never converges: Excel keeps the default text colour as a theme colour when the same colour
  is assigned back, so every round plans the same cells again (Palermo `Policy summary`: 11,807 -> 11,719 -> 11,720
  cells; 25 sheets in each of the 3 rounds), cell by cell where colours are mixed (~36 cells/s here). PVFP: 46 min
  for round 1, `LoB 1` (76,391 cells) again in round 2. CNHI: one sheet (`RAC Baseline`, 123,686 cells) took an
  hour, 60 more colour operations behind it, three rounds to go -- past the 12 hours.
  Decision: convert the colours where they live, in the workbook's style table (`xl/styles.xml`: a few hundred
  fonts and fills instead of hundreds of thousands of cells), then **run all four again** (`full_4`).
  I tried to stop the superseded runs (`full_3`, `full_3_allow`); the permission system refused ("interfere with
  workloads") and nobody is here to allow it, so they are left running and `full_4` uses other ports (8606, 8607).
  Their results are NOT the final ones (except as a second opinion on commit 5f1dafa).

- 20:50-21:13 **the colour action, done in the style table** (`app/theme_colors.py`, commit eecb47e): fonts and
  fills of `xl/styles.xml` converted on the fresh copy before Excel opens it. 0.3 s (Horizon) ... 8.6 s (CNHI).
  Excel's tint arithmetic is not the textbook one (float HLS is 1-2 off on more than half the shades): read 276
  colour/tint pairs back from Excel, tried variants, found it (whole-number HLS on Windows' 0..240 scale, the
  lightened luminance built from two parts cut separately) -- 276 of 276 exact. Tests: `test_theme_colors.py` (6).
- 21:13 `full_4` (all four, 3 Prep rounds, ports 8606/8607) started on eecb47e. **Horizon finished at 21:37**:
  169 + 6 + 1 Prep changes (no colour operation after round 1: it converged), "the preparation changed no computed
  value (22,120 formula cells compared)", 721 errors -> **0** in 2 passes, confirmation PASS, verdict
  **"Acceptable by Mind"**. Result: `runs/full_4/horizon`.
- 21:20 the machine is out of memory and CPU: the superseded runs I may not stop hold ~19 GB and three busy Excel /
  Python pairs (4 cores). Every step runs ~4x slower (Palermo analysis: 914 s instead of 241 s). With three Prep
  rounds CNHI and PVFP would end after the 12 hours. So, for the three left:
  **one Prep round** (with the colour fix the first round does nearly everything: Horizon's rounds 2 and 3 were 6
  and 1 changes) and **one comparison with the original** instead of two (`run_usecases.py --one-comparison`: the
  fixer compares a prepared version with the original itself). `full_4` stopped after Horizon (my own task).
- 21:37 `full_5`: three harness processes, one per workbook -- CNHI (8611), PVFP (8612), Palermo (8613):
  `scripts/run_usecases.py <wb> --ports <p> --prep-rounds 1 --one-comparison --out runs/full_5`
  (logs `run_cnhi.log`, `run_pvfp.log`, `run_palermo.log`). App code = eecb47e (`git diff eecb47e HEAD -- app` empty).

- 22:55 **CNHI done** (`runs/full_5/cnhi`): 3,449 Prep changes in one round (34 min with the re-analysis; the old
  colour step alone had taken 1 h 48 min), recalculation: 0 formula errors, "the preparation changed no computed
  value (577,707 formula cells compared)", verdict **"Acceptable by Mind"**.
- 23:02 **PVFP done** (`runs/full_5/pvfp`): 545 Prep changes, no computed value changed (152,324 formula cells),
  65,883 errors -> **887** in 21 passes (32,098 cells rewritten, 3,101 of them real repairs such as
  `=J446+N(I481)`; 20,369 fixes undone by the numbers gate), confirmation: 887. Exactly what the fixer gives on
  the original file. (The harness ends with exit code 1 when a workbook is not clean: not a failure.)
- 23:04 **Palermo done** (`runs/full_5/palermo`): 241 Prep changes, no computed value changed (56,647 formula
  cells), 511 -> **0** in 1 pass, confirmation PASS.
- 23:18 **PVFP with rule 2 lifted** (`runs/full_3_allow`, commit 5f1dafa, three Prep rounds): 65,883 -> **0** in
  2 passes, 23,820 cells rewritten, **421 values changed and listed**, confirmation PASS.
  **Full unit suite on the final code: 287 passed.**
- 23:20 the screens on the final code (`scripts/ui_autofix_check.py`, Horizon, port 8601): "Fix all 719 errors
  automatically" -> "Clean ... 721 -> 0 ... No value that was good before has changed", saved as v2, green banner
  "Acceptable by Mind -- verified". Screenshots in `runs/ui_final2`.
- 23:30 `docs/AUTOFIX_USE_CASES.md` written from the final runs; deliverables copied to
  `excel-upload-preparation/runs/autofix_2026-10-05/`; port 8600 restarted on the final version.

- 23:50 **one more edit, and so one more round.** The first Recalculate of PVFP takes ~25 minutes even on a calm
  machine: `recalc._scan` asked Excel three things about every error cell (Cells, Address, HasArray) -- 65,883
  cells, twice (the original is recalculated as the baseline). That wait sits right in front of the new button.
  Now the addresses are computed from the values already read and array formulas are asked about per sheet, then
  per row: **55 s** for the same 65,883 errors, same groups (commit 36929c3, `tests/unit/test_recalc_scan.py`).
- 00:03 **`full_6`**: all four again on 36929c3, four harness processes (Horizon 8621 with the default three Prep
  rounds; Palermo 8622, PVFP 8623, CNHI 8624 with `--prep-rounds 1 --one-comparison`). If it cannot finish before
  the 12 hours are over, commit 36929c3 is reverted and the results below (eecb47e) stand.

- 00:40-01:28 **`full_6` done, all four on the final code 36929c3**: Horizon 721 -> 0 ("Acceptable by Mind");
  Palermo 511 -> 0, confirmation PASS; PVFP 65,883 -> 887, its first Recalculate in **3 min 55 s** (28 min in the
  round before, on the same busy machine), confirmation 887; CNHI: 3,449 Prep changes written in 4 min 43 s
  (1 h 48 min with the old colour step), no formula error, no computed value changed (577,707 formula cells),
  "Acceptable by Mind". Same results as the round before.
- 01:07-01:12 **the superseded runs ended by themselves** (`full_3`, commit 5f1dafa, three Prep rounds): PVFP
  65,883 -> 887 (609 Prep changes), CNHI clean after 6,733 Prep changes (no computed value changed, "Acceptable by
  Mind"; it took 5 h 57 min, nearly all of it the old colour step). Nothing of mine is left running.
- 01:20 **full unit suite on the final code: 290 passed.** 01:24 the screens again on the final code
  (`runs/ui_final3`): "Fix all 719 errors automatically" -> clean in 1 min 38 s, green banner.
- 01:35 `docs/AUTOFIX_USE_CASES.md` rebuilt from `full_6`; deliverables in
  `excel-upload-preparation/runs/autofix_2026-10-05/`; port 8600 restarted on the final version.

## Results (through the app, final code 36929c3)

| Use case | Prep changes | Did Prep change a computed value? | Error cells | After "Fix all automatically" | Good values changed | App's verdict |
|---|---|---|---|---|---|---|
| Horizon | 176 (3 rounds) | no (22,120 formula cells) | 721 | **0** | 0 | Acceptable by Mind |
| Palermo | 241 | no (56,647) | 511 | **0** | 0 | blocked by 2 findings that are not formula errors |
| CNHI | 3,449 | no (577,707) | 0 | 0 (nothing to fix) | - | Acceptable by Mind |
| PVFP | 545 | no (152,324) | 65,883 | **887**, listed (they feed formulas that hide errors) | 0 | blocked: 3-D references, INDIRECT as a value, and the 887 |
| PVFP, rule 2 lifted (commit 5f1dafa) | 609 (3 rounds) | no | 65,883 | **0** | 421, listed | blocked: 3-D references, INDIRECT as a value |

The same results came out of the round before (eecb47e) and of the three-Prep-round runs on 5f1dafa.
The fixer alone on the original files (no Prep): Horizon 721 -> 0 in 37 s; Palermo 511 -> 0 in about a minute;
CNHI nothing to fix; PVFP 65,883 -> 887 in 3 min 18 s (rule lifted: 0 in 29 s, 421 values listed).
Details: `docs/AUTOFIX_USE_CASES.md`.

## What is left for a person

1. **PVFP: which way?** With the rule (default) 887 errors stay; with "Fix them too, and list every value that
   changes" the workbook is clean and 421 hidden values change. Both workbooks are in
   `runs/autofix_2026-10-05/` (`pvfp/`, `pvfp_rule_lifted/`).
2. **Read the reasons before trusting a fixed model.** Where an error became 0 the model computes with 0 there
   (PVFP's dead external link: tens of thousands of inputs are now 0). And the assistant's choice can differ
   from run to run: Horizon's running balance got `=IFERROR(...,0)` (the balance restarts at 0 on the failing
   rows) in most runs and `=C191+N(F192)-N(F191)` (the balance carries on) in one. Both keep every good value;
   which is right for the model is a person's call.
3. **What still blocks Mind**, by hand with the assistant: Palermo -- a name that does not exist
   (`Semi_dynamic_increase_rates_array`, 100 cells) and 10 array formulas wider than their grid; PVFP -- 10,404
   formulas with a 3-D reference and 4,410 that use INDIRECT as a value.
4. **Nothing was uploaded to Mind.** "Clean" is Excel's recalculation; the conversion in Mind is the next test.
5. **A PVFP file prepared with an earlier version of the app** has 2,462 computed values that differ from the
   original: prepare it again from the original.
6. **The branch** `auto-fix-loop` is on the personal GitHub only (`origin`). Nothing was pushed to the team repo.
7. Ideas noted on the way, not done: the analysis keeps every version's inventory in memory (7-8 GB for one
   session of PVFP or Palermo after two or three analyses -- three models side by side exhaust a 32 GB
   machine); on CNHI Prep still proposes ~1,500 optional grid titles after each Apply.
