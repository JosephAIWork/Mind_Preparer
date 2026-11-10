# Outside input files

Some models get part of their data from other files: a macro opens each file and pastes its
numbers into the model (the Tokyo IFRS 17 tool's `calcCF` opens one `.xlsm` per product
group). Mind runs no VBA and opens no outside file, so the app brings those files in for it
and makes each one a **Mind input** that the Mind team can replace every quarter without
uploading the model again.

## What the user does

1. Pick the model **and** its input files together (several files in File Explorer, or drag
   them all onto the page). The app finds which one is the model.
2. The app lists the files the model's macro reads: matched, missing, and files it does not
   need. Missing files can be added later. A file that is not picked keeps the data the model
   already holds.
3. **Bring in the outside files** (the first step of "Do everything"): the app runs the
   model's own import macro on a copy, then makes the Mind inputs. The numbers check then
   compares later steps with this version, because new outside data changes values by design.
4. Download the model and **Mind inputs.zip** (one CSV per input). Run in Mind uploads both.
   Each quarter: **Make next quarter's Mind input files** from the new raw files (below), put
   the files it gives in Mind's **Inputs** window, then Run.

### Each quarter: next quarter's Mind input files (`app/next_quarter_inputs.py`)

**Raw input files cannot go into Mind directly.** Mind's Inputs window takes one file per input
grid, shaped like the grid: the same name, the same header, the same columns. A raw quarterly file
(the file the model's import macro opens) holds much more -- all its items and months, its own
header rows -- so Mind refuses it:

> Failed: Input control error: The file 2026_Q1 - BBA_01.csv doesn't have proper headers, there
> are 1340 columns while grid 'IN_BBA_01' of instance Original_Instance has 93 columns.

The grid `IN_BBA_01` holds only the 93-column block the macro takes out of that file (turned: items
across, months down). So, on the Outside files card, once the files are brought in:

1. **Make next quarter's Mind input files**: pick the new quarter's raw files (`.xlsx`, `.xlsm`,
   `.csv` ... as the client gets them; one, some or all).
2. The app runs the model's **own import macro** on a copy of the version the outside files were
   first brought into, with the new files (the same run as the bring-in: an Excel of its own,
   message boxes answered), then reads each grid's block and writes it under the grid's name with
   the grid's own header (`IN_BBA_01.csv`; `IN_VFA_01.txt`, TAB-separated, when a value or a header
   holds a comma -- Mind reads `.txt` as TAB and takes it).
3. **Download Next quarter Mind inputs.zip**: `README.md` (which file goes into which grid) and the
   folder `Mind inputs` with the files. Unzip; in Mind click **Inputs**, then **Browse** and pick the
   folder `Mind inputs` (or drag its files onto the window); wait for SUCCESS on every file; Run.
   No new model upload, no new version in Mind Ready.

What it says on the way:

- A grid without a new file is not in the zip: Mind keeps the data it has.
- A file already shaped like a grid (its header is the grid's) goes in as it is, under the grid's name.
  It must be UTF-8 text: one saved in another encoding (a CSV saved on a Japanese PC as "CSV (comma
  delimited)") is refused with what to do; the other files are still made. A `.txt` is read as
  TAB-separated, as Mind reads it: a comma-separated `.txt` shaped like a grid goes in as
  `<grid>.csv` (refused when it holds quotes or TABs).
- A raw file named like a grid (`IN_BBA_01.csv` with 1,340 columns) is recognised and read with the
  macro: "IN_BBA_01.csv has 1,340 columns; the Mind input grid IN_BBA_01 has 93. It looks like a raw
  input file, so the app reads it with the model's import macro."
- A new file that fills more than its grid holds is left out, with why (Mind's grid keeps the model's
  shape); the macro's own message boxes are shown when it refuses a file (too many items).
- Each file says how many values differ from the files made at the bring-in.
- The macro gets the raw files exactly as picked and opens each its own way (OpenText settings,
  `Local:=True`, line-by-line reads such as the Tokyo tool's `CountCSVRow`, checks of the name), as
  at the bring-in and in the client's tool. Excel keeps 15 significant digits when it opens a CSV
  (144686640.71679688 becomes 144686640.716796): the result says how many numbers of a CSV have
  more, and "values changed" does not count a number that differs only past its 15th digit
  (Joseph's raw CSVs against the bring-in's `.xlsx` files: 0 changed). A first version handed the
  macro a full-digit `.xlsx` copy of each CSV; the review of 10 November 2026 removed it, because it
  skipped the macro's own way of opening the file.
- While the files are being made, the bring-in, Run in Mind, the fixer, a recalculation and other
  workbook changes wait (one Excel job at a time). The card follows a running job again after a
  reload, and forgets the last job when another workbook is loaded.
- When Run in Mind (or `scripts/mind_upload.py`) hands Mind a raw file and Mind refuses it for its
  columns, the result says what to do instead (`mind_client.upload_inputs` -> `advice`).

**Proven in real Mind (10 November 2026, project `zzz_mindready_tokyo_inputs_1`, the anonymized IFRS 17
tool, through the app's own routes on a dev server):**

- Joseph's raw file, named like the grid (`IN_BBA_01.csv`, 1,340 columns): Mind refused it ("doesn't
  have proper headers, there are 1340 columns while grid 'IN_BBA_01' of instance Original_Instance
  has 93 columns"), read from the row's LOG; the app's result said what to do instead.
- Set A, made from Joseph's six raw `.csv` files (with the full-digit copy then in use): the six
  files are byte for byte the files the bring-in made from the same quarter's `.xlsx` files (0
  values differ). Without the copy (now), they differ only past the 15th digit: 0 values changed. All six SUCCESS in the
  Inputs window (5 `.csv` + `IN_VFA_01.txt`, TAB). Run consistent with the audit trail. Result_CF:
  Mind = Excel 120 of 120 numbers (Excel = the model with its own macro run on the same raw files).
- Set B, the same with BBA_04's numbers x 1.1 (one raw file changed, 62,019 cells): the app's file
  `IN_BBA_04.csv` differs in 26,456 values, the other five in none. All six SUCCESS; Mind = Excel
  120 of 120; against set A, 16 numbers changed, all in the BBA_04 row. No model upload.
- After the review (the raw files given to the macro as picked): set A made again in 84 s, 0 values
  changed against the bring-in; uploaded to the same project, all six SUCCESS, run consistent,
  Result_CF the same 120 numbers as above (Mind = Excel). The project holds set A again.
- Found on the way, fixed: `mind_client.upload_inputs` read the rows while Mind's "Inputs import"
  box was still up. Those rows show the last upload's status and file -- the same names when files
  replace earlier ones -- so a refused file could pass as SUCCESS. It now waits for the box to come
  and go (or 20 s when it is too quick to see), and reads Mind's reason from the row's LOG ("CSV Logs").

Routes: `POST /api/sessions/{id}/next-quarter-inputs` (the files; 202, background job) and
`GET /api/sessions/{id}/next-quarter-inputs` (the grids with the file name Mind expects, the job's
stage, its result: files, grid, source file, values changed, the zip).

### Without the input files: stand-ins made with the assistant (`app/stand_in_inputs.py`)

When files are missing, **Make them with AI** on the Outside files card (or in the box under Do
everything while it waits, below) has the assistant make them:

1. The assistant reads the import macro (the VBA of the import chain), the import table and the
   shape of every sheet the files fill (where its labels, constant numbers and formulas are).
2. It answers with a small recipe per file: the fixed cells the macro searches for ("Description")
   and the blocks to copy from the model (values only, transposed when the macro transposes), or
   a block of zeros where the model holds no data (its file was missing when the model last ran).
3. The app builds the files in an Excel of its own **from the numbers the model already holds**,
   runs the model's own macro on a copy with them, and checks every cell of every area the macro
   writes against the model (a 0 where the model is empty counts as equal), then every computed
   value. When something differs, the assistant is shown what happened -- the macro's message
   boxes, the cells that differ -- and tries again (4 tries, low thinking effort: ~1 min a try).
   Zeros that would change a result are left out: that block stays as the model holds it.
4. The proven files become the session's input files, marked **stand-in**, and the usual bring-in
   follows: Mind inputs, CSVs, the numbers check. INP-010 then warns (it does not block): best
   results come with the actual input files. Adding an actual file replaces its stand-in, and the
   files are brought in again.

Without an AI key the app says so and asks for the actual files.

### "Do everything" stops and asks for the input files (1.14.0)

Do everything never makes stand-ins by itself any more. When its first step finds that files the
macro opens were not picked, it **stops and asks**:

1. The server's bring-in job is started with `waitForInputs` and, with files missing, goes into the
   state `waiting_for_inputs` instead of running: no Excel, no version, nothing changed. `GET
   /api/sessions/{id}/outside-inputs` then says `waiting: true` and lists `needed`: every file the
   macro opens, by the name the model expects (else its key, else its cell), whether a picked file is
   there for it, and the sheet it fills; `message` is "Waiting for 2 input files: BBA_05, VFA_01".
2. The main screen shows **one box** under Do everything: "Do everything is waiting for 2 input
   files: BBA_05, VFA_01. Pick them, or let the assistant make stand-ins." with the list, **Pick the
   files**, **Go on** (greyed out until every file is there), **Make them with AI** (only when an AI
   key is saved) and **Stop**. The Outside files card says "Waiting for ... · See the box under Do
   everything" with a Show me. Picking files uses the usual `POST .../inputs`; the list follows.
3. **Go on** -> `POST .../outside-inputs/continue` (409 while a file is still missing, naming it):
   the job starts at the import itself -- the stages after the wait, and only those (`stages` in the
   GET answer: `start, import, scan, mind_inputs, numbers, rescan, done`, never `stand_ins`). The chain
   in the browser was held at this step the whole time, so the steps after it follow as usual.
   **Make them with AI** -> `continue` with `{"makeStandIns": true}`: the stand-in path above, with
   its warning (best results come with the actual input files); 409 without an AI key.
4. **Stop** -> `POST .../outside-inputs/stop`: the job ends (`state: stopped`, "Stopped while waiting
   for the input files. Nothing was changed: pick the files, then press Do everything again."), Do
   everything stops at its first step. A bring-in already running cannot be stopped (409).
5. A page reloaded or a session reopened while it waited: the box comes back (Go on then starts Do
   everything again, which goes straight through the step); a workbook saved while it waited and
   opened after a restart of the app shows the job as stopped ("The app was closed while Do
   everything waited for the input files. Pick them, then press Do everything again.").

The card's own button ("Bring in the 4 matched files") still brings in the matched files alone, and
`POST .../outside-inputs/run` without `waitForInputs` behaves as before. Tests:
`tests/unit/test_outside_inputs_api.py` (section 8) and `MindReadyUI/src/app/core/outside-inputs.spec.ts`.

## How it works (`app/outside_inputs.py`)

- **Finding the import.** The VBA is read for calls that open outside files
  (`Workbooks.Open`, `OpenText`, `Open … For Input`, `QueryTables.Add`, `GetObject`). The
  import macro is the procedure without parameters nearest to such a call (`calcCF`, not
  `OneScenarioRun`, which only adds a message box).
- **The import table.** Cells whose formula builds a file name (`…&".xlsm"`) or hold a typed
  path. For each row: the same-row cells its formula reads (the row's key, e.g. `BBA_01`) and
  the sheet it fills. The exact file name is worked out by Excel from the formula's part after
  the last `\`.
- **Matching.** Exact name first, else the row's key as a whole word of the file name. Names
  compare after Unicode NFKC and case-folding (full-width `（１）` = `(1)`), so any language works.
- **Running the import.** The model's own macro runs on a copy in an Excel of the app's own
  (never the user's open Excel), each path cell pointing at the picked file. Excel's events show
  which file the macro opens and which cells it writes while that file is open: that range is
  the file's **import area**. Message boxes are answered and reported. The values therefore
  arrive exactly as the client's tool puts them.
- **Making the Mind inputs.** For each import area:
  - a new sheet `IN_<sheet>` with the grid `#IN_<sheet> /Input`: header row (the names above
    the area, or C1…Cn), then the data, values only;
  - every formula that read the area now reads the same cells of that grid;
  - a SUMIF/SUMIFS whose ranges run past the side of the area is split in two (the columns
    outside plus the columns inside: the same total);
  - the old area is cleared;
  - one CSV per grid (UTF-8, decimals written in full, never an exponent; TAB `.txt` when a
    value holds a comma).
  An area whose readers cannot all be moved (a range that only partly covers it in another way,
  OFFSET from inside it, a whole column across it, an array formula over it) stays inside the
  model and is reported.

## Proven in real Mind (2026-11-08)

- A grid titled `#CF_Test /Input` gives the project an **Inputs** button (left of Run).
- The Inputs window takes a folder (`.csv`, `.txt`, `.zip`, `.parquet`), matches each file to
  the grid of the same name (Japanese headers fine), shows SUCCESS, needs no save, and the next
  run (even in a new browser session) uses the imported numbers.
- Tokyo (anonymized, fake inputs): 6 of 6 files matched, import 11 s, 6 Mind inputs,
  46,800 formulas moved (10,800 through the SUMIF split), no computed value changed.
