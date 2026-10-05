# Fix all automatically — the four use cases

Run of 2026-10-05 (evening), app version 1.8.0, app code of commit `eecb47e`, branch `auto-fix-loop`.
The four workbooks are the client models in
`Tel Aviv Office - 060 R&D\70 GenAI Exploration\60 Mind Excel Prep\2026 09 - Models for testing`
(RAROK left out, as asked). The originals were never touched: every run works on copies.
Nothing was uploaded to Mind: "clean" here is Excel's own full recalculation.

## The result

| Workbook | Size | Formula cells compared | Prep changes | Did Prep change a computed value? | Error cells after Prep | After "Fix all automatically" | Good values changed | The fix took | App's verdict |
|---|---|---|---|---|---|---|---|---|---|
| **Horizon** | 1.6 MB | 22,120 | 176 | no | 719 | **clean** (confirmed by a separate recalculation) | 0 | 3 min 51 s | ready |
| **Palermo** | 28.0 MB | 56,647 | 241 | no | 511 | **clean** (confirmed by a separate recalculation) | 0 | 33 min 05 s | blocked |
| **CNHI** | 37.5 MB | 577,707 | 3,449 | no | 0 | **clean** (nothing to fix) | - | - | ready |
| **PVFP** | 3.8 MB | 152,324 | 545 | no | 65,883 | **887 left**, listed for a person | 0 | 27 min 07 s | blocked |

The fixer counts 721 on Horizon: the 719 formula errors and 2 cells where an error value had been typed in as data.

- **Horizon, Palermo, CNHI: zero error cells**, no value that was good before has changed, and the
  preparation changed no computed value. Horizon and CNHI end with the app's verdict
  **"Acceptable by Mind"**.
- **PVFP: 65,883 error cells → 887.** The 887 are left on purpose: each of them feeds a formula that
  hides errors, and giving it a value would change what that formula shows today. They are listed with
  the cell that would change and its value before / after. With that rule lifted for this workbook, PVFP
  comes out with **zero** error cells too, and the 421 values that moved are listed
  (see "PVFP with the rule lifted" below).
- "Clean" is not the fixer's own word for it: after the fix, the app's ordinary **Recalculate** was run
  again on the new version and found no formula error cell (887 for PVFP, the number the fixer reported).

## What was run

For each workbook, through the app's own web API — the same calls the buttons make
(`scripts/run_usecases.py`, a fresh server per workbook):

1. **Upload and analyze.** A workbook that unpacks to 100 MB or more is uploaded the way the upload
   screen offers: the one sheet that holds at least half of it is left out of the *scan* (Palermo `Data`,
   CNHI `Output`). The workbook stays whole — Prep, the recalculation and the fixer work on all of it.
2. **Standard Prep**: everything that is ticked by default, applied once. (Horizon: applied again while
   Prep still planned something — three rounds of 169, 6 and 1 changes.)
3. **Recalculate.**
4. **Fix all automatically**, with the two rules: clean means zero error cells, and a good number never
   changes. On a prepared version the fixer first compares the workbook with the original — *did Prep
   change what the model computes?* CNHI has no error to fix, so there the comparison was asked for
   directly ("Check the numbers").
5. **Recalculate again** — the independent confirmation.

The times in the tables are **not** what the app needs. Through most of the evening the machine was
running several models at once and was out of memory (see `AUTOFIX_RUN_LOG.md`); every step took two to
four times longer than alone. Alone, the fix itself takes 37 s on Horizon, about a minute on Palermo and
3 min 18 s on PVFP; the rest of each "auto-fix" line is the comparison with the original and the
re-analysis of the fixed version.

## What to know before reading the details

- **Where an error became 0, the model now computes with 0 there.** That is what "fixed" means, and it
  is why every fix carries its reason — in the app and below. PVFP is the extreme case: its inputs come
  from another workbook that is not there (`C:\Users\<user>\...\[PVFP Inputs_v4.0.xlsx]`). Almost all
  of its 65,883 errors are `#REF!` (65,820): thousands of cells read that dead link directly and the rest
  of the model reads them. Those cells are now `=IFERROR(<the link>, 0)`: the model recalculates without
  an error, on inputs that are zero until the real inputs are back.
- **PVFP cannot be made clean under rule 2, and that is the right answer.** The model *hides* some
  errors on purpose: `=IFERROR(K24/K12,"No GEP")` shows "No GEP" because K24 is an error. Give K24 a 0
  and that cell shows 0.5: a value that was fine has changed. The fixer undoes such a fix (20,369 fixes
  were undone on this model), tries empty text, and when that fails too it leaves the error — with the
  chain of error cells that leads to the hiding formula — and cleans everything else.
- **A real repair where one exists.** On Horizon `=+BE36/BF36` became `=IF(BF36=0,0,BE36/BF36)`; on PVFP
  `=J446+I481` became `=J446+N(I481)` in 1,632 cells, because I481 holds empty text. Such a rewrite is
  written to the whole block of cells that share the formula and kept only if every cell of the block
  that worked before still returns exactly the same value.
- **Prep used to change PVFP.** Before this version the standard Prep changed 2,462 computed values of
  PVFP (1,781 into errors): it inserted title rows on sheets that `=SUM('LoB 1:>>'!K65)` and
  `INDIRECT("'"&$E$2&"'!H"&n)` read by position. The comparison of step 4 is how it was found; Prep now
  refuses those inserts and says why. In this run Prep changed no computed value on any of the four
  (808,798 formula cells compared in all). **A PVFP file prepared with an earlier version of the app should be
  prepared again from the original.**
- **Prep's colour step no longer takes hours.** "Replace theme colours with explicit RGB" went through
  the cells one by one in Excel and came back at every Prep round: 46 minutes a round on PVFP, more than
  an hour for one sheet of CNHI. It is now done in the workbook's style table, in under ten seconds, once.

## What still stands between each workbook and Mind (the app's verdict)

Formula errors are one of the things the verdict looks at. What is left is by hand, with the assistant:

- **Horizon** — nothing. *Acceptable by Mind.*
- **CNHI** — nothing. *Acceptable by Mind.* (The spilled-range references Mind refuses, 38 formulas, are
  replaced by a Prep action that is ticked by default.)
- **Palermo** — two findings that are not formula errors: 100 cells use a name that does not exist
  (`Semi_dynamic_increase_rates_array`; the closest existing name is `Semi_dynamic_increase_rates` — Excel
  shows no error because those branches are not evaluated), and 10 array formulas are wider than their
  grid.
- **PVFP** — 10,404 formulas use a 3-D reference (`'LoB 1:>>'!K65`) and 4,410 use INDIRECT as a value;
  Mind supports neither. And the 887 errors above.

## Details, workbook by workbook

### Horizon

1.6 MB · result: **clean** -- every formula recalculates without an error · total time 22 min 51 s · versions: v1 → v2 → v3 → v4 → v5

| Step | Time | What happened |
|---|---|---|
| upload+analyze | 3 min 54 s | unpacks to 10.0 MB; every sheet scanned; verdict *blocked*, 1 blocking |
| prep round 1 | 5 min 07 s | 169 change(s) applied (fix_broken_refs 9, fix_broken_refs_whole 26, create_grid_titles 120, hide_marker 2, explicit_colors 12); verdict *unverified* |
| prep round 2 | 4 min 02 s | 6 change(s) applied (separate_merged_grids 3, create_grid_titles 3); verdict *unverified* |
| prep round 3 | 3 min 15 s | 1 change(s) applied (create_grid_titles 1); verdict *unverified* |
| numbers check (original vs prepared) | 34 s | The preparation changed no computed value (22,120 formula cells compared with the original). (719 cells are errors before and after; followed through 9 inserted row(s), 2 renamed sheet(s)) |
| recalculate | 1 min 38 s | 719 formula error cell(s): #N/A 54, #DIV/0! 44, #VALUE! 621 |
| auto-fix | 3 min 51 s | errors 721 → 0 in 2 pass(es); 704 cell(s) rewritten; 0 fix(es) undone by the numbers gate (time includes the comparison with the original and the re-analysis of the fixed version) |
| recalculate (confirmation) | 28 s | independent recalculation: 0 formula error cell(s) -- PASS |

**The app's verdict at the end:** Acceptable by Mind: no blocking problem and a clean Excel recalculation (13 optional remark(s))

**What the fixer wrote** (11 kinds of error; the largest):

| Cells | Where | Error | Reason | Before → after |
|---|---|---|---|---|
| 621 | `CALCUL!C192:C812` | #VALUE! | text where a number is expected: the formula is kept and returns 0 when it fails (This is a running numeric balance; a blank input should be treated as zero so the sum still works.) | `=C191+F192-F191` → `=IFERROR(C191+F192-F191,0)` |
| 23 | `CALCUL!BI35:BI57` | #DIV/0! | a division by zero: the formula is rewritten so that the error cannot occur, in every cell of its block; the 1 cell(s) of the block that already worked each still return(s) exactly what it did (Dividing two empty amounts should just give zero instead of an error.) | `=+BE36/BF36` → `=IF(BF36=0,0,BE36/BF36)` |
| 23 | `CALCUL!BJ35:BJ57` | #DIV/0! | a division by zero: the formula is rewritten so that the error cannot occur, in every cell of its block; the 1 cell(s) of the block that already worked each still return(s) exactly what it did (Dividing two empty amounts should just give zero instead of an error.) | `=+BG36/BH36` → `=IF(BH36=0,0,BG36/BH36)` |
| 23 | `C20 &&Hide!B8:C8` | #N/A | a value that is not available (a lookup that finds nothing): the formula could never compute again (its reference was deleted): it now reads "not available, shown as 0" (This is a numeric placeholder meant to be zero when no data exists.) | `=+NA()` → `=IFERROR(NA(),0)` |
| 4 | `C20 &&Hide!B9:C9` | #N/A | a value that is not available (a lookup that finds nothing): the formula is kept and returns 0 when it fails (This is a numeric placeholder meant to be zero when no data exists.) | `=+SUM(NA())` → `=IFERROR(+SUM(NA()),0)` |
| 3 | `C20 &&Hide!I39` | #N/A | a value that is not available (a lookup that finds nothing): the formula could never compute again (its reference was deleted): it now reads "not available, shown as 0" (This counts garanties, so a missing count should be treated as zero.) | `=NA()` → `=IFERROR(NA(),0)` |
| 2 | `OUTPUT!S3:T3` | #N/A | a value that is not available (a lookup that finds nothing): the formula is kept and returns 0 when it fails (This is a total amount, so a missing part should contribute zero.) | `=NA()-Q3` → `=IFERROR(NA()-Q3,0)` |
| 2 | `OUTPUT!I1` | #VALUE! | text where a number is expected: the cell held an error value typed or pasted as data: cleared | `#VALUE!` → `` |

### Palermo

28.0 MB · result: **clean** -- every formula recalculates without an error · total time 85 min 47 s · versions: v1 → v2 → v3

| Step | Time | What happened |
|---|---|---|
| upload+analyze | 15 min 48 s | unpacks to 188.7 MB; left out of the scan: Data; verdict *blocked*, 2 blocking |
| prep round 1 | 18 min 01 s | 241 change(s) applied (create_grid_titles 214, hide_marker 1, explicit_colors 26); verdict *blocked* |
| recalculate | 10 min 10 s | 511 formula error cell(s): #N/A 507, #VALUE! 4 |
| numbers check (by the fixer: original vs prepared) | inside the fix | The preparation changed no computed value (56,647 formula cells compared with the original). (511 cells are errors before and after; followed through 2 inserted row(s), 1 renamed sheet(s)) |
| auto-fix | 33 min 05 s | errors 511 → 0 in 1 pass(es); 507 cell(s) rewritten; 0 fix(es) undone by the numbers gate (time includes the comparison with the original and the re-analysis of the fixed version) |
| recalculate (confirmation) | 8 min 41 s | independent recalculation: 0 formula error cell(s) -- PASS |

**The app's verdict at the end:** Not acceptable by Mind yet: 2 blocking problem(s) (14 optional remark(s))

- `REF-001` (by hand, with the assistant): 100 cell(s) contain broken references -- name(s) that do not exist in the workbook (Excel shows #NAME?): Semi_dynamic_increase_rates_array (100 cell(s), closest existing name: Semi_dynamic_increase_rates). Mind's converter cannot compile them ('No function found ... not a function'). First: Policy summary!F6. Fix: a misspelt name is corrected to the existing one; a broken reference is replaced with NA() (an error sta
- `RSK-002` (by hand, with the assistant): 10 array formula(s) exceed their grid: Mind's converter refuses the model ('the array formula ... exceeds the grid size. Ensure that the column headers cover the entire array formula'). First: Temp!C5 spans C5:DZ5 but its grid Temporary output tab stops at F14 (first cell outside: G5). Fix: extend the header row over every column the array covers (a merged header counts for all the columns it spans), or shorten the a

**What the fixer wrote** (4 kinds of error; the largest):

| Cells | Where | Error | Reason | Before → after |
|---|---|---|---|---|
| 502 | `Decrements!M1004:M1505` | #N/A | a value that is not available (a lookup that finds nothing): the formula is kept and returns 0 when it fails (This is a mortality/decrement rate lookup that fails when the age isn't found in the table, so treating the missing rate as zero avoids distorting the calculation.) | `=1-(1-VLOOKUP(L1004,Decrements!$B$7:$D$125,IF(Policyholder_sex="M", 2,` → `=IFERROR(1-(1-VLOOKUP(L1004,Decrements!$B$7:$D$125,IF(Policyholder_sex="M", 2, 3` |
| 2 | `Non-economic assumptions!C58:D58` | #N/A | a value that is not available (a lookup that finds nothing): the formula is kept and returns 0 when it fails (This looks up a fund charge percentage and should be zero when the tariff isn't found, since a missing rate shouldn't be treated as a real charge.) | `=VLOOKUP(Policy_tariff,Assumptions_array,MATCH(C55,Assumption_headings` → `=IFERROR(VLOOKUP(Policy_tariff,Assumptions_array,MATCH(C55,Assumption_headings,0` |
| 2 | `Non-economic assumptions!F58:G58` | #N/A | a value that is not available (a lookup that finds nothing): the formula is kept and returns 0 when it fails (This looks up a duration figure in months and should be zero when the tariff isn't found, since a missing value shouldn't be counted as a real duration.) | `=VLOOKUP(Policy_tariff,Assumptions_array,MATCH(F57,Assumption_headings` → `=IFERROR(VLOOKUP(Policy_tariff,Assumptions_array,MATCH(F57,Assumption_headings,0` |
| 1 | `Saved down!C3` | #VALUE! | text where a number is expected: the formula is kept and returns empty text when it fails (This builds a text label from the file name, and when the expected underscore pattern isn't present it can't construct the text, so leaving it blank is correct since this is a label not a number.) | `=C2&"Run_version "&MID(C5,FIND("@",SUBSTITUTE(C5,"_","@",LEN(C5)-LEN(S` → `=IFERROR(C2&"Run_version "&MID(C5,FIND("@",SUBSTITUTE(C5,"_","@",LEN(C5)-LEN(SUB` |

### CNHI

37.5 MB · result: **clean** -- every formula recalculates without an error · total time 78 min 06 s · versions: v1 → v2

| Step | Time | What happened |
|---|---|---|
| upload+analyze | 23 min 50 s | unpacks to 364.2 MB; left out of the scan: Output; verdict *blocked*, 1 blocking |
| prep round 1 | 33 min 36 s | 3,449 change(s) applied (freeze_spill_refs 38, create_grid_titles 3376, explicit_colors 35); verdict *unverified* |
| recalculate | 5 min 39 s | 0 formula error cell(s) |
| numbers check (original vs prepared) | 14 min 52 s | The preparation changed no computed value (577,707 formula cells compared with the original). (0 cells are errors before and after; followed through 23 inserted row(s), 0 renamed sheet(s)) |

**The app's verdict at the end:** Acceptable by Mind: no blocking problem and a clean Excel recalculation (14 optional remark(s))

### PVFP

3.8 MB · result: **887 error cell(s) left** · total time 84 min 59 s · versions: v1 → v2 → v3

| Step | Time | What happened |
|---|---|---|
| upload+analyze | 12 min 40 s | unpacks to 33.2 MB; every sheet scanned; verdict *blocked*, 2 blocking |
| prep round 1 | 11 min 23 s | 545 change(s) applied (create_grid_titles 533, hide_marker 1, explicit_colors 11); verdict *blocked* |
| recalculate | 28 min 10 s | 65,883 formula error cell(s): #REF! 65,820, #N/A 58, #VALUE! 5 |
| numbers check (by the fixer: original vs prepared) | inside the fix | The preparation changed no computed value (152,324 formula cells compared with the original). (65,883 cells are errors before and after; followed through 8 inserted row(s), 1 renamed sheet(s)) |
| auto-fix | 27 min 07 s | errors 65,883 → 887 in 21 pass(es); 32,098 cell(s) rewritten; 20,369 fix(es) undone by the numbers gate (time includes the comparison with the original and the re-analysis of the fixed version) |
| recalculate (confirmation) | 5 min 38 s | independent recalculation: 887 formula error cell(s) -- ERROR |

**The app's verdict at the end:** Not acceptable by Mind yet: 2 blocking problem(s) (15 optional remark(s))

- `FRM-006` (by hand, with the assistant): 10404 formula(s) use a 3-D reference ('First:Last'!cell, every sheet between two tabs): Mind reads 'First:Last' as one sheet name and refuses the model ("Sheet 'LoB 1:>>' not found in workbook ... on compiling formula"). First: LoB_Total!H65 = =SUM('LoB 1:>>'!H65) spans 2 sheet(s) in tab order: LoB 1, >>; >> hold no grid and are left out of the proposal (Mind imports no empty sheet). Fix by hand or with the assistant
- `FRM-007` (by hand, with the assistant): 4410 formula(s) use the result of INDIRECT as a value (compared, multiplied, divided, negated or concatenated): Mind returns a reference for INDIRECT and cannot turn it into a value ('Run error: Unable to cast object of type AM.Models.AMReference to type System.IConvertible'). First: Inputs_PVFP!AA711 = =IF(INDIRECT("'"&$G711&"'!$Z$16")=0,0,66%*$Z988*INDIRECT("'"&$G711&"'!$Z$12")/INDIRECT("'" -> proposed =IF('LoB 1'!

**What the fixer wrote** (1,261 kinds of error; the largest):

| Cells | Where | Error | Reason | Before → after |
|---|---|---|---|---|
| 3,522 | `Inputs_PVFP!DA711:DL711` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (This looks up values in another LoB's sheet and should contribute nothing if the reference can't be resolved.) | `=IF(INDIRECT("'"&$G711&"'!$Z$16")=0,0,66%*$Z988*INDIRECT("'"&$G711&"'!` → `=IFERROR(IF(INDIRECT("'"&$G711&"'!$Z$16")=0,0,66%*$Z988*INDIRECT("'"&$G711&"'!$Z` |
| 3,143 | `Inputs_PVFP!AB152:CJ154` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (This links to an external file reference that is broken, so a missing number should count as zero in totals.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |
| 2,960 | `Inputs_PVFP!I549:CJ585` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (The external workbook link is broken; treating this missing rate as zero avoids skewing calculations.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |
| 2,000 | `Inputs_PVFP!I766:BF805` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (The source link to the external file is missing, so zero is the safe value for this amount.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |
| 2,000 | `Inputs_PVFP!I823:BF862` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (Broken external reference; zero keeps this quantity from distorting sums.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |
| 2,000 | `Inputs_PVFP!I878:BF917` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (The link to the external workbook is broken, so treat this as zero in calculations.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |
| 2,000 | `Inputs_PVFP!I1208:BF1247` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (This figure comes from a missing external link; zero avoids affecting totals wrongly.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |
| 1,850 | `Inputs_PVFP!I936:BF972` | #REF! | a reference that no longer exists: the formula is kept and returns 0 when it fails (The commission rate link is broken; zero is the neutral value for a rate that is multiplied elsewhere.) | `='C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs` → `=IFERROR('C:\Users\<user>\Documents\03. Projets\Software\Mind\[PVFP Inputs_v` |

**Left for a person** (887 cell(s) in 126 kinds; the largest):

| Cells | Where | Error | Why it was not fixed |
|---|---|---|---|
| 78 | `Inputs_PVFP!AA711` | #REF! | it feeds LoB 1!C211 and 62 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!C211 from False to True |
| 77 | `LoB 1!I78` | #VALUE! | it feeds Reporting_LoB 1!BU17 and 39 more good value(s), which only hold(s) because this error is there: fixing it would change Reporting_LoB 1!BU17 from 0 to 45600 |
| 76 | `LoB 1!J77` | #VALUE! | it feeds Reporting_LoB 1!BU17 and 39 more good value(s), which only hold(s) because this error is there: fixing it would change Reporting_LoB 1!BU17 from 0 to 45600 |
| 34 | `LoB_Total!K138` | #REF! | it feeds LoB 1!K139 and 89 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!K139 from 'No GEP' to 0.5 |
| 25 | `LoB 1!S498` | #REF! | it feeds LoB 1!C211 and 62 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!C211 from False to True |
| 23 | `LoB 1!T499` | #REF! | it feeds LoB 1!C211 and 62 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!C211 from False to True |
| 21 | `LoB 1!U500` | #REF! | it feeds LoB 1!C211 and 62 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!C211 from False to True |
| 19 | `LoB 1!V501` | #REF! | it feeds LoB 1!C211 and 62 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!C211 from False to True |
| 18 | `LoB 1!Z16` | #REF! | it feeds LoB 1!C211 and 62 more good value(s), which only hold(s) because this error is there: fixing it would change LoB 1!C211 from False to True |
| 18 | `LoB_Total!L73` | #VALUE! | it feeds Reporting_LoB 1!I30 and 16 more good value(s), which only hold(s) because this error is there: fixing it would change Reporting_LoB 1!I30 from 0 to 30723 |

## PVFP with the rule lifted

The button "Fix them too, and list every value that changes" (`keep_good_values: false`), run through
the app on PVFP after three Prep rounds (commit `5f1dafa`: the fixer's code is the same as in `eecb47e`,
only Prep's colour step differs):

| | With the rule (the default) | Rule lifted |
|---|---|---|
| Error cells | 65,883 → 887 | 65,883 → **0** |
| Passes | 21 | 2 |
| Cells rewritten | 32,098 | 23,820 |
| Fixes undone because a good value moved | 20,369 | - |
| Values that were good before and changed | **0** | 421 |
| The app's Recalculate afterwards | 887 formula error cells | no formula error |

The 421 are on `LoB 1` (210), `LoB_Total` (102), `Reporting_LoB 1` (51), `Reporting_LoB Total` (34) and
`Reporting_General` (24). For example `LoB 1!K25`, `=IFERROR(K24/K12,"No GEP")`: "No GEP" → 0.5; twelve
tests go from FALSE to TRUE; totals that were 0 become 12,160 or 82,080. Each one is in the result with
its formula, the value before and the value after.

This is the choice the app leaves to a person: with the rule, 887 errors stay and nothing the model
shows today changes; without it, the workbook is clean and the listed cells show a different value —
mostly a "No GEP" / 0 / FALSE that was only there because an input was an error.

## Other runs of the same evening

- **Palermo, three Prep rounds** (commit `5f1dafa`, before the colour fix): 345 Prep changes, no computed
  value changed (56,647 formula cells), 511 errors → 0 in one pass, confirmation PASS. Same result as
  above with one round.
- **The fixer alone on the original files** (no Prep; `python -m app.autofix`): Horizon 721 → 0 in 37 s,
  Palermo 511 → 0 in about a minute, PVFP 65,883 → 887 in 3 min 18 s (rule lifted: 0 in 29 s, 421
  values changed), CNHI nothing to fix.
- **The screens**, on the final code: `scripts/ui_autofix_check.py` clicks through the real pages on
  Horizon (upload, Findings, Recalculate, "Fix all 719 errors automatically", the result card): 721 → 0,
  "No value that was good before has changed", saved as v2, and the green banner
  "Acceptable by Mind — verified".

## How to run it again

```
cd excel-upload-preparation
python -X utf8 scripts/run_usecases.py <workbook.xlsm> [...] --ports 8611,8612 --out <run dir>
        [--prep-rounds 1] [--one-comparison] [--allow-handled] [--no-assistant]
python -X utf8 scripts/usecase_report.py <run dir> [...] --out report.md
python -X utf8 -m app.autofix <workbook> <work dir> [--assistant] [--allow-handled]     # the fixer alone
python -X utf8 scripts/compare_versions.py --original <original> --prepared <prepared> --changelog <apply logs>   # the numbers check alone
```

The final workbooks, the fixer's full result for each (every kind of fix with its cells, formula before
and after and the reason; everything left and why) and the run logs are in
`excel-upload-preparation/runs/autofix_2026-10-05/` (not in git: client workbooks are never committed).
