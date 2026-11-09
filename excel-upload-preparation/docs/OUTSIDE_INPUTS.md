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
   Each quarter: put the new CSVs in Mind's **Inputs** window, then Run.

### Without the input files: stand-ins made with the assistant (`app/stand_in_inputs.py`)

When files are missing, **Make them with AI** on the Outside files card (and "Do everything",
behind the scenes, when the assistant is set up) has the assistant make them:

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
