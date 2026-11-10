# Changelog

## Unreleased (10 November 2026): next quarter's Mind input files

- **Make next quarter's Mind input files** (Outside files card, once the files are brought in;
  `app/next_quarter_inputs.py`; `POST` / `GET /api/sessions/{id}/next-quarter-inputs`). Joseph put a
  raw quarterly file (`2026_Q1 - BBA_01.csv`, 1,340 columns) in Mind's Inputs window and Mind refused
  it: its grid `IN_BBA_01` has 93 columns. Raw input files cannot go into Mind directly: Mind takes one
  file per input grid, shaped like the grid. Now the user picks the new quarter's raw files; the app
  runs the model's own import macro on a copy of the version the files were first brought into, reads
  each grid's block and writes it under the grid's name with the grid's own header. One zip:
  `README.md` (which file goes into which grid, and why raw files cannot) and the folder `Mind inputs`.
  No new model upload, no new version. A grid without a new file is left out (Mind keeps its data); a
  file already shaped like a grid goes in as it is; a raw file named like a grid is recognised; a file
  that fills more than its grid is refused with why; the macro's own message boxes are shown.
- **A raw `.csv` keeps every digit**: Excel opens a CSV keeping 15 significant digits. The app gives
  longer numbers their exact value in a `.xlsx` copy before the macro reads it. Joseph's six raw CSVs
  now give files byte for byte equal to the ones the bring-in made from the `.xlsx` files (before
  this: 15,011 values of IN_BBA_01 differed in the last digits).
- **Mind's refusal in words**: `mind_client.upload_inputs` reads Mind's reason for a refused file from
  the row's LOG ("CSV Logs") into `errors`, and for a raw file says what to do instead (`advice`).
- **Fix: Inputs window read too early.** `upload_inputs` read the rows while Mind's "Inputs import" box
  was still up; they show the last upload's status and file (the same names when files replace earlier
  ones), so a refused file could pass as SUCCESS. It now waits for the box to come and go.
- **Proven in real Mind** (`zzz_mindready_tokyo_inputs_1`, `docs/OUTSIDE_INPUTS.md`): the raw file
  refused, with the app's advice; set A (Joseph's raw CSVs) all six files SUCCESS (incl.
  `IN_VFA_01.txt`), Mind = Excel 120/120; set B (BBA_04 x 1.1) all SUCCESS, Mind = Excel 120/120, and
  16 numbers changed against set A, all in the BBA_04 row.

## Unreleased (10 November 2026): Mind instances and where instances could help

Built in October on a side branch as 1.13.0 and 1.14.0 and never released; brought onto this branch
and fitted to what it has since (outside inputs, stand-ins, Show me, plain codes, per-loop choices).

- **One macro's loops at a time** (review of 10 November 2026; `docs/MIND_LOOP.md`). Excel runs each
  run macro on its own, with the other macro's drivers fixed. The Mind loop plan wrote the loops of
  every macro into one model, so Mind ran them inside each other (the IFRS 17 tool: 6 x 4 x 6 x 6
  passes) and the values could differ from Excel. Now the plan builds one macro (`macro=` on
  `GET` / `POST .../mind-loop-plan`; a **Macro** dropdown on the Mind loop panel), lists the others'
  loops as not built, and names a driver of another macro that already holds a loop. Copies from loops
  take the loops of one macro (two macros: refused), pick that macro's loop plan, and no longer write
  another macro's loops into every copy. The Tokyo instances built before this fix
  (`zzz_mindready_tokyo_instances_1`) had the group and roll-forward loops inside every copy. Built
  again through the API (91 changes, no other macro's loops) as `zzz_mindready_tokyo_instances_2`:
  Mind's upload test 57 cells (a subset of the 63: rounding residues below 4e-6 and the 8 "TRUE"
  cells), run consistent with the audit trail, export against Excel 12,528 cells, 0 differences.
- **Loops told apart by procedure and loop variable**: two macros both using `i` no longer overwrite
  each other, and an output block is written by its own macro's loop.
- **Count cells**: a count cell that already holds the count this plan writes is accepted (the loop
  can be planned again after a write); the count's label never goes on another loop's count or label,
  a cell the plan writes, or a cell a formula reads.
- **A first value typed as a cell** counts as chosen, like a number: the header pass is kept.
- **Copies from loops** keep the clears of the macro's paste areas; copies' keys are always different
  (by position when two values would clash).
- **Prep's repairs with the instances** are compared with the plan's cells at their addresses before
  the plan's row and column inserts: no repair is dropped or overwritten because of another cell that
  only shares an address.
- **Screens**: the instances panel sends the Mind loop panel's choices (`loop_choices`, `macro`) and
  previews again when they change; a new preview drops the earlier write's download link.

- **Copies from loops** (`app/instance_builder.py`, the instances panel's "Copies from loops" boxes;
  `GET` / `POST .../mind-instances-plan` take `loops`, `results` and `loop_choices`). For a model whose
  macro runs the whole model once per scenario and cohort (or any counter or list loop), not one row
  of a table at a time: one copy per combination of the chosen loops' values. In the model: a split
  grid `#Instance values /InstanceSplit` two rows below the used range of the first loop's sheet (the
  key, then each loop's value, one row per copy), each chosen driver `=INDEX(<its column>,1)` (the one
  row Mind keeps in the copy; in Excel the first copy), the driver's dropdown removed, a cell with
  `MM_INSTANCEKEY`; the macro's other loops are written as `MM_LOOP`s, as the Mind loop plan has them,
  so they run inside every copy. The main workbook lists the keys and shows one block of the model for
  every copy -- the formula block of the output sheet by default, or the range asked for -- one row per
  copy and block row, `MM_INSTANCE(<cell>, <key>)`; no sum over the copies (it would add scenarios up).
  The notes say what Mind does with the model's `/Input` grids (a file named `<key>-<grid>.csv` changes
  one copy), that every copy computes (also the passes the macro skips), and name any cell that still
  uses a chosen loop in an `MM_` function. Every plan lists `loop_options` (each loop, and why copies
  cannot be made from it); a plan by the table that cannot be built points to them. The Mind loop
  plan's operations now say which loop they belong to (`loop`).
- **In Mind, the anonymized IFRS 17 tool as 24 copies** (scenario 0-3 x cohort 2021-2026; the group and
  roll-forward-year loops stay loops), built through the app's API from the prepped model:
  `zzz_mindready_tokyo_instances_1` converted in 129 s, its upload test shows the same 63 cells as the
  model without instances (rounding residues below 4e-6 and 8 "TRUE" text cells), the run took 3 min
  06 s, consistent with the audit trail. The main workbook's export after the run against Excel with
  each copy's cohort and scenario set in the driver cells: 24 x 6 x 87 = 12,528 cells, 0 differences.
- **`scripts/mind_upload.py`**: the log reader clicks "Hide Logs" only. It took "Close", which closes the
  upload window without adding the templates, so `--run` found no Run button.
- **Written with the formula repairs Prep still has**, as the Mind loop is now: a version built before
  Prep goes to Mind with CELL / INFO frozen, OFFSET ranges written out and the rest of the "Fix the
  formulas" set, so Mind computes what Excel does. A cell the builder writes keeps the builder's
  formula; a plan that is not ready writes nothing. The plan's notes list the repairs.
- **Checked again in Mind on 10 November 2026** with three small invented projects
  (`docs/MIND_INSTANCES.md`): each copy runs MM_LOOPINSTANCE to its own count; `MM_INSTANCE` reads one
  copy by key cell, literal key or index, and on a looped cell gives only the first pass -- the builder
  points it at the `MM_RESULT` cells, so it is right; a plain link is the sum over the copies. Nothing
  in the builder had to change.

### Mind instances (the builder)

- **Build the Mind instances** (Advanced > Mind loop, the Mind instances section below the loop). The
  macro's run loop turned into Mind instances: one copy of the model per value of a key column of the
  model-point table (a tariff, a product, a group), a loop inside each copy over that copy's own rows,
  and a second small workbook -- the main one, which Mind does not copy -- with the keys and the results.
  Preview, write it into a new version, then download both files as one zip and drop them together on
  an empty Mind project (`app/instance_builder.py`; `GET` / `POST /api/sessions/{id}/mind-instances-plan`
  `[/apply]`, `GET .../mind-instances/download`; `docs/MIND_INSTANCES.md`).
- **What changes in the model.** The model-point table becomes a split grid: a key column as its first
  column, the key of every row under the header `Instance key`, the title `#<Name> /InstanceSplit`
  directly above and an empty row above the title (the column and the rows inserted where needed). The
  loop's count cell gets `=COUNTA(<key column>)`, so each copy counts its own rows; a free cell holds
  the biggest count -- the loop size -- with a label; the driver cell gets
  `=MM_LOOPINSTANCE("<loop name>", <size cell>, <count cell>)` -- plus 1 when the range the loop
  indexes starts on the header row (the macro ran `2 To n`), so every pass reads a data row. Every
  formula of the block the macro copies out on each pass becomes `=MM_RESULT(<its source cell>)`, the
  sum over the copy's loop; the macro's paste area is emptied; a free cell shows the copy's key,
  `=MM_INSTANCEKEY("<first key>")`, with a label. Filled columns touching the table on its left are part
  of the grid in Mind's eyes, so the key column goes before them (and a headerless one gets a header).
  Cells that only read the paste area are left as they are and listed (they show 0 in a copy).
- **The main workbook** `<stem>_Main.xlsx`, written through Excel next to the model so its references
  resolve: sheet Instances with `#<key header> instances /InstanceKeys` (the model's file name as header,
  one row per key); sheet Summary with `#Results by <key header>` (one row per key, one column per row of
  the output block, each `=MM_INSTANCE(<model cell>, <key>)`) and `#All copies` (the whole output block
  as plain links to the model, its header row kept, with a Total row: in Mind a plain link to a copied
  workbook is the sum over the copies, checked in a real project). Keys stay what they are: a text
  key like `001` is written as text in the model's key column and in the main workbook alike. In Excel
  every grid shows the one copy Excel has.
- **The choices.** The key: a dropdown of the table's text columns with 2 to 200 distinct values, each
  with its row count. The default is the column the macro itself looks the loop row up in, else a header
  that says tariff, product, group, portfolio, segment, class, category, type, country, fund, treaty or
  cohort, else the column with the fewest values. "Every key of the list, even without rows" makes a copy
  per value of the list the macro reads, when the app finds it. The model is uploaded as
  `<stem>_instances` with its own extension; Mind finds the workbook to copy by that name, so the files
  keep their names.
- **A loop without a count cell** is not built: the plan says so in one sentence, and the loop panel's
  choice of a count cell comes first. The builder takes the first counter loop of the loop plan that has
  one.
- **A new change, `set_range_values`** (`app/prep.py`): one range written in one go, through Excel or
  openpyxl, counted in the change log like the other changes. The key column of a long table is one
  change, not one per row. After the write through Excel the last cell planned to hold something is read
  back (not the range's last cell, which is empty by design when the table's last rows have no key).
  `excel_text` now keeps every string a string: text Excel would turn into a number, a date or a
  boolean (`001`, `1E3`, `2024-01`, `TRUE`) gets the text prefix.
- **A refused change** (the key column, the driver, a result cell) leaves a model the plan does not
  describe: the version is kept, but the main workbook and the zip are not made, and the answer says
  which changes were refused. A paste-area cell that stayed does not stop the zip.
- **Two fixes on the way.** RZS-006 no longer counts an `/InstanceSplit` grid as resizable: a copy keeps a subset of the template's
  rows, never more, and Mind reads a fixed range over the whole key column as the copy's column -- so the
  readiness verdict no longer turns red right after the instances are built. The outcome card names the
  Mind loop and Mind instances steps.
- **In Mind:** the two files of the hand-built prototype this builder mirrors, dropped together on an
  empty project, converted, tested and ran as one (28 October 2026; the results-by-key grid matched the
  macro's sums by key). This builder's own two files, made through the API from the fixed model cut to
  the macro's 19 test policies (29 October 2026): converted and tested as one project, run in 3 s, the
  results-by-key grid (3 x 128) and the all-copies grid (128 x 10) equal to Excel's sums by tariff, 0
  differences, Total row equal.
- The feedback message names the step "Build the Mind instances". The guide for the Mind team has a
  section on it.

### Where instances could help (the finder)

- **Find where instances could help** (Advanced > Mind loop, the Mind instances section, a button
  below the builder). The app reads every formula of the model and its macro and points to what
  repeats: a column of the looped table that splits its rows into groups (one copy of the model per
  value), the same formulas side by side in one sheet (bands of columns or of rows), sheets with the
  same layout and formulas, and near copies that differ only in typed-in numbers or in a few cells of
  logic -- with the edits that would unify them. Each place is a card: one plain sentence, whether
  instances, a loop or leaving it as it is fit (yes, partly, no) and why, and **Use as the key**,
  which fills the builder's key column and previews the instances. The finder is blind -- no hint of
  what the key is -- and deterministic: pure Python over the scan the app already made, no Excel, no
  numpy, no model call; the same workbook gives the same answer. It writes nothing. With an AI key
  saved, the tick **Explain and rank with the assistant** adds a plain explanation and the assistant's
  order on top; the findings themselves never change (`app/instance_finder.py`,
  `app/finder_shapes.py`, `app/finder_patterns.py`, `app/finder_grouping.py`;
  `GET /api/sessions/{id}/instance-finder` `[?assistant=1][&refresh=1]`;
  `python -m app.instance_finder <workbook> [--json] [--assistant]`; `docs/INSTANCE_FINDER.md`).
- **Merged into dimensions** by the input the copies step through, never by labels. The column the
  macro looks the loop row up in, the list it matches the value into and the slots it pastes each
  pass into are one dimension; so are blocks on other sheets whose copies read consecutive rows or
  column slots of that area, including blocks that step by a computed index -- `OFFSET` with
  `ROW()`, `COLUMN()`, an expanding `COUNTA` or a `MATCH` on the key -- and two lists that hold the
  same values. Lookups keyed by the current row's value, also through a mapping table, and formulas
  that branch on it with `IF(<key>="...")` count as evidence; a header word is shown as a hint and
  weighs nothing. Chains (each copy reads the one before it) are set aside with the reason; time axes
  and fills that join nothing are counted, not reported.
- **Cells reading the macro's paste area, by class**, each with what it needs in Mind. Cells that add
  every slot up move to the main workbook as plain links to the copied model -- there a link is the
  sum over the copies; inside a copy they would see only that copy. Rows that each read one slot by a
  computed index become one row per copy in the main workbook. The block that reads the current key's
  slot and adds the pass to it stays as it is: in a copy it is the copy's own block, and `MM_RESULT`
  does the adding. Cells reading one fixed slot need `MM_INSTANCE(cell, key)`. The builder
  listed all of these as "left as they are"; the finder says what to do with each.
- **Checked on the use case** against the builder's own choice (`docs/INSTANCE_FINDER.md`, "The
  blind result on the 2c use case"): the finder runs blind on the same session as the builder's plan,
  and the two JSON answers are set side by side, fact by fact -- the table, the key column and its
  header, the evidence, the distinct values and rows, the other candidates, the cells reading the
  paste area by class, the time taken. `builder.matches_top` says in the JSON whether both pick the
  same column. On the use case it does: the finder's first place is the builder's key column on the
  19-row cut, on the full fixed model (26 groups over 94,932 rows) and on the raw file, with the
  macro's grouping as its first evidence; the finder adds what the builder leaves as it is -- which
  cells read the paste area and what each needs in Mind -- and the lookups and branches keyed by the
  column.
- **Limits.** Chains are set aside, not solved. Formula counts are not run time: Mind's speed and
  memory are not predicted (the biggest copy's row count is the nearest hint). The request is one
  long GET -- up to about two minutes on 350,000 formulas, no background job yet; a step the budget
  cut is written in `skipped` in plain words, nothing is skipped silently. The app keeps two workbooks
  open: after a third upload or an Apply the finder has to read the workbook again before it starts
  (minutes on a 28 MB model) and says so, with the seconds, in its answer. Sheet copies are matched at
  the same addresses, so a copy shifted by a row is not seen. Copies across separate workbooks are not
  covered. A block or sheet dimension is reported with its edits; the builder builds grouping columns
  only.

## Unreleased (9 November 2026): each Mind loop with its own choices

- **The Mind loop is written with the formula repairs Prep still has** for the version (CELL / INFO
  frozen, INDIRECT written out, ...: the "Fix the formulas" set). Found on Tokyo: the four loops were
  written before Prep, and `CELL("filename")` in B3 of the six calculation sheets -- the key of every
  lookup there -- is NaN in Mind (Mind answers with the workbook's name alone, no "]"), so Mind's test
  showed NaN on every formula of those sheets (14,737 cells). Not the loops: a probe gives the same NaN
  without any loop. The loop panel lists the repairs; `repairs: false` leaves them out.
- **New rule FRM-014 "math on a range that OFFSET returns"** and the Prep action **Write a fixed OFFSET
  range out where arithmetic uses it**: `SUMPRODUCT(1-OFFSET($H$2,2,0,3,1),OFFSET($O$2,2,0,3,1))` is
  -16,220 in Excel and -5,513,051 in Mind (probe workbook in real Mind); written as `1-$H$4:$H$6` Mind
  gives Excel's number. Tokyo: 84 formulas (14 per calculation sheet), 396 differences in Mind's test
  of the Do-everything version without loops.
- Proven in real Mind on Tokyo with its four loops, built through the app: Mind's test went from 14,737
  differences to 63 (Do-everything version + loops, zzz_mindready_tokyo_loops_2) and 75 (loop written
  on the unrepaired version), none a value: check cells that are 0 in Excel and a rounding residue below
  0.00002 in Mind, and the text "TRUE" read as 1 on Excel's side.
- **Advanced › Mind loop: every loop card has its own Name / First / Last boxes**, and a Count cell
  box when the loop asks for one (with "Use <free cell>"). First and last take a number or a cell
  (Sheet!A1). Sent as `loops` by driver cell (`GET .../mind-loop-plan?loops=<json>`, POST
  `{"loops": {...}}`); the old shorthands (`name`, `first`, `last`, `count_cell`) still work.
- **More loops are built from the workbook alone** (seen on the Tokyo tool):
  - a driver set to item i of a named list (`GroupGrid(i + 1, 1)`): MM_LOOP over the list, its end
    counted as the macro counts it (to the last filled cell);
  - a counter whose own dropdown lists its values (0, 1, 2, 3): MM_LOOP over that list (fixed range);
  - a bound reached through plain references (`=C10`, `=CalcYear`) reads the cell holding the number.
- **A workbook-level name wins over a sheet's own name of the same spelling** (the macros'
  `ThisWorkbook.Names("GroupName")`): the loop goes in the cell the macro writes, not a copy.
- **Count cells are safe**: never a filled cell, a driver, or another loop's count cell; a label
  ("<loop> count") is written to its left; a loop starting below 1 counts 1..n and shifts
  (`=MM_LOOP("Scenario",X1)-1`).

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
- **No input files? "Make them with AI".** When files are missing, the assistant reads the import
  macro and the model's sheets and describes each file; the app builds stand-ins from the numbers
  the model already holds, proves them with the model's own macro (every written cell equals the
  model's, then every computed value; the assistant tries again with what went wrong, 4 tries) and
  brings them in as usual. "Do everything" does it behind the scenes when files are missing.
  INP-010 then warns instead of blocking -- best results come with the actual input files -- and an
  actual file added later replaces its stand-in (`POST .../outside-inputs/run {"standIns": true}`,
  `app/stand_in_inputs.py`; `chat_completion(effort=)` for this one call).
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
- **A shared copy stops on a date** (`build.py --expires 2026-12-15 --contact "..."`, 15 December
  2026 by default, `--no-expiry` for none): "Works until 15 Dec 2026" in the header, a line in the
  last 7 days, then a plain message at start and a calm page instead of the app (410 on every API
  call); a clock moved back counts too (each build remembers its own latest date; a date that has
  not come yet, from a clock once set ahead, is put right by the internet's date at start); the
  developer version has no date (`app/expiry.py`, docs/SHARING.md). Build the screens first
  (`npx ng build`): build.py does not package a copy with a date from screens that cannot show it.
- **Plain English next to every code**: "FRM-009 · an empty cell is compared with a number" on the screens, in the
  texts and in the Excel report (`rules/plain_names.yaml`, required by a test); Excel's error values get their meaning
  once, Excel / VBA errors a plain sentence with Excel's words in brackets (`app/plain_codes.py`, docs/PLAIN_CODES.md).
- **"Show me" pointers.** Next to every instruction that sends the user somewhere (the banner's next
  step, "Apply the blocking repairs in Prep", "Check them under Advanced › Deliver", "on the Outside
  files card", "paste an AI key under Advanced › Assistant", a finding's "Use Bring in the outside
  files" or "The Prep action '…'", "the list at the top of this page") a quiet "Show me" link: one
  arrow with a short bubble at the next click ("Click Advanced", "Click Prep actions") until the place
  itself, outlined, with "Got it"; Esc or a click elsewhere ends it, and it never takes a click from
  the page. "Fix that under Advanced" now names the tab. `docs/SHOW_ME.md`.
- **"Show me", after an independent review** (8 findings, 7 fixed): in a window narrower than 66rem,
  where Carbon hides the header's Main and Advanced, the arrow uses the page's own links (it showed
  nothing for 7.5 s, then ended); a click on the way that cannot be shown ends the pointer at once; a
  double click goes on instead of ending it; a stopped "Fix all issues" names the right place (Deliver
  when its comparison did not run, none when it did not run at all: "Press Do everything again"); a
  clean recalculation names no list; the assistant's own answers get no link (only the app's words);
  the findings' messages and Prep's "Add to the list" note have theirs; the readiness banner, now one
  component for both screens, keeps its name for screen readers (the title carries Carbon's id).

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
