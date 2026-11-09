# Changelog

## 1.12.4 (9 November 2026): outside input files as Mind inputs (the Tokyo IFRS 17 tool)

- **A model and its input files, picked together.** The upload takes several files at once (File
  Explorer or drag and drop); the app finds the model (the workbook whose macros open other files)
  and lists the files its macro reads: matched, missing, and files it does not need. Missing files
  can be added later; a file that is not picked keeps the data the model already holds.
- **Bring in the outside files** (the first step of "Do everything", and a button). The app runs the
  model's OWN import macro on a copy, in an Excel of its own, with each path cell pointing at the
  picked file, and watches it through Excel's events: which file it opens and which cells it writes
  while that file is open. The values arrive exactly as the client's tool puts them. Message boxes
  are answered (No / Cancel on a question) and reported; a macro that stops early is reported, not
  hidden. The numbers check then compares later steps with this version.
- **Each outside file becomes a Mind input.** Its area moves to a sheet `IN_<sheet>` holding the grid
  `#IN_<sheet> /Input`; every formula that read the area reads the same cells there (a SUMIF whose
  ranges run past the side of the area is split in two: the same total); one CSV per input
  (`Mind inputs.zip`). Each quarter the Mind team imports new CSVs in Mind's **Inputs** window and
  runs -- no new upload of the model. An area some formula cannot follow (a 3-D reference over it,
  OFFSET from inside it, a range that covers only part of it) stays in the model and is reported.
- **Run in Mind uploads the input files** through the Inputs window after the template step.
- **New rule INP-010 "Outside input files"**: blocking until the outside files are brought in.
- **Only the workbook, no input files: "Go on without them".** The Outside files card lists every
  file as missing and "Do everything" stops at its first step, saying what to do: add the files
  (no new upload), or go on without them -- the model keeps the numbers its macro brought in last
  (Mind runs on them, it cannot swap them each quarter). INP-010 then warns instead of blocking and
  "Do everything" leaves the bring-in out; Undo, adding a file or starting a bring-in takes the
  choice back (`POST .../outside-inputs/skip`).
- Proven on the Tokyo tool (anonymized, invented input files): 6 of 6 files matched, import 11 s,
  6 Mind inputs, 46,800 formulas moved, no computed value changed; in Mind 6 of 6 inputs SUCCESS and
  Mind = Excel on the result grid (120 of 120 numbers); a "next quarter" (two inputs changed, model
  not re-uploaded) gives Excel's new answer too (120 of 120).
- **Hardened after an independent review** (4 reviewers, each finding challenged by a refuter on
  another model; 11 findings, none refuted, 10 fixed): a plain OK message box no longer counts as an
  early stop (Windows gives a lone OK button control id 2); a macro that clears every area first is
  paired correctly; whole-column clears of a file not picked are put back; the model's own
  Workbook_Open does not run and a box there cannot hang the job; when making the Mind inputs would
  change a value the import version stays current and nothing is offered for Mind; input files changed
  after a bring-in are noticed (fingerprint) and Run in Mind refuses stale CSVs; Do everything retries
  a failed bring-in; numbers that could not be checked stop the chain; the macro's message box reads
  correctly on a Windows of another language ("BBA_01のinput領域には93項目まで…" instead of mojibake).
- **RZS-006** leaves the Mind inputs made from outside files alone: their size is the model's import
  area, as the model's own macro keeps it, and their CSVs are written at exactly that size.
- Final check on Tokyo through the app: Acceptable by Mind (0 blocking), then the app's own Run in
  Mind: numbers match (1,703,906 cells), the 6 input files SUCCESS in Mind's Inputs window, run
  consistent with the audit trail, converged; 56 untitled grids in Mind (76 in 1.12.3).
- `scripts/night_chain.py` writes UTF-8 (a Japanese sheet name stopped it on a Hebrew console) and takes
  `--inputs <folder>`.

## 1.12.3 (8 November 2026): sheet names in any language (the Tokyo IFRS 17 tool)

- **Sheet and name text in any language.** A reference such as `=ロールフォワード!B10` was read as
  `=B10`: the reference reader only knew English letters, so on a workbook with Japanese sheet
  names the scan reported dozens of false circular references (a cell reading "itself") and the
  banner stayed on "blocked" while the recalculation was clean. Every reader of sheet names and
  defined names now takes letters of any alphabet (Japanese, Hebrew, accented, Cyrillic ...).
- **`CELL("filename", C6)` in C6 is not a circular reference.** CELL asks where the cell is, not
  what it holds (Excel computes it); it is read like ROW and COLUMN.
- **Text typed as `'=...` is text.** The scan took any text starting with "=" for a formula.
- **A CELL formula that shows "" is frozen as `=""`.** A formula whose saved value is empty text
  was read as "no value", so the fix left it "by hand"; and a CELL formula that was an error in
  the original and shows "" after Fix all issues is frozen to "" as well.
- **INDIRECT used as a range is rewritten too (FRM-007).** Checked in Mind: `VLOOKUP(x, INDIRECT($A5&"!A:B"), 2, FALSE)`
  is not computed there and the IFERROR around it showed "-" in 480 cells of the result sheet; the same
  lookup on `BBA_04!$A:$B` gives Excel's values. The Prep action ("Replace INDIRECT by the reference its
  text builds") now writes the direct reference wherever INDIRECT stands. An empty cell in the text is
  read as "" (a sheet chooser left blank: Excel's INDIRECT errs already); a missing target becomes NA(),
  or `INDEX($A$1,NA())` where a reference must stand (Excel refuses `OFFSET(NA(), ...)`).
- **A one-cell array formula over 255 characters is written with Formula2** (FormulaArray refuses longer
  text; the same one-cell array, the same value).
- **No title row is inserted inside a range a formula reads.** A row inserted inside `D29:H32` stretched it
  and `INDEX($D$29:$H$32, Scenario+1, 4)` read another row ("New" became "Old"); such an insert is now
  skipped (`prep.stretched_reader`).
- **A lone text cell read only through a whole column is not flagged (STR-008)**: the formula reads the
  sheet's table there, not the cell (Mind computed the lookup with it).
- **Mind's new "type of project" step.** Creating a project now first asks Milliman Mind or Milliman OpS;
  the automation picks Milliman Mind.
- `scripts/night_chain.py --no-loop`: the chain without a Mind loop.
- Checked on the Tokyo tool (anonymized): no computed value changed (110,175 formula cells), converted
  and ran in Mind at the first attempt, and Mind's results equal Excel's on the result sheet and a
  cash-flow grid.

## 1.12.2 (7 November 2026): the macro's loop through named ranges

- **The loop is found when the macro reaches its cells through `Names("X").RefersToRange`.** A macro
  that writes `Set cohortRng = ThisWorkbook.Names("CohortRng").RefersToRange` and then
  `cohortRng.Value = cohortYear` or `scenarioRng(1, 1).Value = i` had its loops read with no input
  cell, so no Mind loop was offered (the Tokyo IFRS 17 tool). The macro reader now writes these
  shortcuts out in full (`app/model_map.py`), including the loop limits read from them
  (`startcohortyear = startRng.Value`) and `Range("X").Cells(r, c)`.

## 1.12.1 (5 November 2026): what the Mind team saw when they tried it

- **The program no longer stops with "No module named win32timezone".** Excel hands back dates and
  times through a part of pywin32 that is loaded by name, which the compiled program left out; it is
  now built in (`tools/package/build.py`).
- **Run in Mind says which version it uploads: always the latest.** The screen names the version and
  its file before the click. While Fix all issues, a recalculation, a Prep round or an analysis is
  still changing the workbook, Run in Mind waits ("Wait a moment: ...") instead of uploading the version
  before that change (`GET /api/sessions/{id}/mind-loop/options`, 409 on `POST .../mind-loop`).
- **Choose the changes made before uploading.** Every change Run in Mind can make first is a tick box
  (ticked as before by default); an unticked change is never made, not even when Mind asks for it, and
  nothing ticked uploads the version as it is (`actions` in the `POST` body).
- **A clearer status while it runs, with the time.** The stages of a round (reading, changes, checks,
  opening Mind, upload and convert, run, results), the one going on now, the time spent, the time left
  and the clock time it should be done by. The estimate is kept on the long side on purpose (1.5 times
  the measured figures); past it, the screen says it is taking longer and is still working. The
  detailed log is one click away (`app/estimate.py` `estimate_mind_run`).
- **A row inserted for one grid no longer moves the grids beside it.** The empty row Prep inserts above
  a '#Title' trapped inside a grid (STR-001) or above a grid to make room for its title (STR-004) now
  shifts only that grid's columns down when another grid sits beside it at or below that row, so the
  neighbour stays where it is (before, the trapped-title fix was skipped there, and a title row moved
  the neighbour down, away from its own title). The columns that shift widen to take in anything that
  must move as one: a grid below, a merged range, a pivot table, a range a formula reads across them.
  When nothing sits beside the grid, or the sheet has an Excel table or a formula reading whole rows,
  a whole row is inserted as before. The numbers check and Run in Mind follow the moved cells
  (`app/prep.py` `plan_row_insert`). The grid names chosen by hand in Advanced (Grid Namer) insert their
  title rows the same way.
- **Sign in to Mind from the app.** Run in Mind has a "Mind sign-in" box: "Sign in to Mind" opens
  Edge with the app's own profile on Mind; once you are in, it saves the sign-in and closes by itself
  (the sign-in lasts about a week). "Check the sign-in" loads Mind without a window and says whether it
  still works. Before, signing in needed a command line, which the program does not have
  (`GET`/`POST /api/mind/sign-in`, `app/mind_client.py` `sign_in`). Run in Mind and the sign-in window
  wait for each other: they share the profile.

## 1.12.0 (27 October 2026): the hand-over to the Mind team

- **Send feedback to the Tel Aviv team.** When a step fails, a box under the header says in one plain
  sentence what stopped and why, with the button; the same button is in the header for anything else.
  The dialog prepares the message (the assistant, when a key is saved, writes its opening lines), shows
  it, then copies it and opens an email to the address in the settings (`feedback.email`); a file is
  saved only on request. The message is in codes -- the stage, the error's type, the app functions it
  went through -- with Excel's own code and the *shape* of the formulas involved (function names kept;
  references, text and numbers taken out), and never a file name, a sheet name, a cell address or a
  value. Each problem is kept in the data folder (`feedback\incidents.jsonl`, the last 300). The codes
  are hashes: the app holds no table of what they mean (`app/feedback.py`, `POST /api/feedback`).
- **No hidden Excel stays behind.** The app writes down every Excel it starts until it quits it. Closing
  the app's window ends its Excels first, and the next start ends any left by an app that was stopped
  another way (a crash, Task Manager). Only an Excel the app itself started is ever ended
  (`app/excel_com.py`).
- **The AI key is pasted in the app.** When no key is saved, the main screen and Advanced > Assistant
  show a box for it. The key is checked with the gateway (a key the gateway refuses is not saved), then
  saved encrypted for the Windows user (DPAPI) in the data folder; the page only ever shows its last 4
  characters. The old search for `secret.key` / `config.enc` in the folders around the app is gone, and
  with it the `cryptography` requirement (`app/ai_key.py`, `GET`/`POST /api/assistant-key`).
- **One settings file**, `settings.yaml`: the port, the data folder, the AI gateway and model, the Mind
  addresses, the Mind folder and project prefix the app works with, the MM for Excel add-in, and the
  API's own pages (off by default). A mistake in the file stops the app with a message naming the line.
  The app deletes in Mind only projects whose name starts with the prefix (`app/settings.py`).
- **CELL / INFO formulas (FRM-012)** are frozen to what the original upload shows, the values the user's
  own Excel saved, followed through the rows inserted since. A version saved by the app's Excel shows the
  app's working folder there instead (with the Windows user name in it), and such a value is never written.
- **Pivot tables (STR-009): every one blocks Mind's converter.** A workbook was refused with its pivot
  tables, refused again with only the one with a grouped field turned into values, and accepted with all
  of them as values. The pivot setting is values (the default) or formulas.
- **A workbook whose formulas read its pivots through GETPIVOTDATA is converted too** (it used to keep every
  pivot, and Mind refused it). Before a pivot becomes cells, each GETPIVOTDATA reading it is pointed at the
  cell it returns today (an error stays an error, as NA()), and each such cell is checked in Excel to show
  what it showed. A pivot's custom style is deleted once nothing uses it: Mind stopped on pivots styled
  "Results" with "Requested value 'Results' was not found" (`app/pivots.py`).
- **Broken references (REF-001):** a broken defined name written with its sheet (`Projn!IndxAsst4`) becomes
  `NA()` with the sheet, not `Projn!NA()`, which Excel refused.
- **Formulas Excel 365 stores as one-cell arrays are fixed like any other (FRM-008 to FRM-013).** Excel 365
  saves most formulas with INDEX or OFFSET as a "dynamic array" over its own cell, and the Mind-value fixes
  left every such cell "by hand": 512,796 cells of one model, so FRM-009 stayed blocking after Do everything.
  Only an array over several cells is left by hand now. A one-cell array is written back as one
  (`Formula2`: the old way added an `@` to the formula, which made it another formula).
- **A formula filled down is written as one fill.** The Mind-value fixes plan one change per block of cells
  holding the same formula before and after (for example `Output!BK133:BK64231`): 64,099 cells in about a
  second, where one write per cell took about 8 Excel calls each. The block's last cell is read back against
  the plan, and the block is put back as it was if it differs.
- **Lone text cells (STR-008).** A cell inside a defined name counts as read only when some formula reads that
  name (directly or through another name): 9 labels of one model were flagged only because a name over the
  whole sheet covered them, a name no formula uses (a macro's; Mind runs no macro). A cell with no room for
  a title above it used to be left by hand; now, if it sits just under a grid whose first column holds
  text labels, it gets a label beside it so the grid takes its row. Otherwise it moves (Excel's cut and
  paste, so every formula reading it follows) to a column of its own right of the sheet's contents, with
  its title above. A cell read through a range or a name over several cells keeps its place: the range
  would lose it.
- **"Do everything right after the upload"**, a box under the drop area (remembered in the browser): once the
  workbook is read, Do everything starts by itself, and a big workbook is read whole without stopping to ask
  which sheets to leave out. The computer must stay awake: when it sleeps, the page that drives the run waits.
- The header and the browser tab say **Beta**.
- The old Streamlit screen is removed; comments, messages and documents name no client model.

## Up to 1.11.10

The app reads an Excel model and checks it against the readiness rules mined from Mind's knowledge base
(`rules/`), with every check deterministic. It fixes what it can through Excel itself, on a copy:

- **Fix all issues**: finds the cells where formula errors start, gives each a value to fall back on,
  and undoes any fix that changed a value that was correct.
- **Fix the formulas** and the structure (Prep): every change Mind needs before it accepts the file --
  grid titles, flags, broken references, 3-D references, INDIRECT used as a value, external links, cells
  that hold `""`, empty cells compared with numbers, CELL / INFO, line breaks in text, pivot tables --
  each change reviewable before it is applied.
- **Build the Mind loop**: the run loop of the model's macro written as `MM_LOOP` / `MM_RESULT`.
- **Recalculate** in Excel with the MM for Excel add-in, and **check the numbers**: every formula cell
  of the original and the prepared version compared, through the rows inserted.
- **Do everything** runs those steps in order, each followed by its check.
- Under Advanced: the findings rule by rule, the Model map, the Grid Namer, the assistant, Run in Mind
  (upload, convert, run and read back, in a sandbox folder only), and every version.

Each rule that comes from a real Mind conversion says so in its source line. The development notes of
these versions are kept internally.
