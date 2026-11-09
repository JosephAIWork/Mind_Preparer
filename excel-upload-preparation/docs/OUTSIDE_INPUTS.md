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

### Only the workbook, no input files

The app still finds what the macro reads and lists every file as missing; INP-010 blocks and
"Do everything" stops at its first step, saying what to do. Two ways on:

- **Add files** on the Outside files card (or drop them there): no new upload of the model.
- **Go on without them** (same card): the model keeps the numbers its macro brought in last
  (they are saved in the workbook as plain values). Mind runs on those numbers but cannot swap
  them each quarter. INP-010 becomes a warning, the banner no longer blocks on it, and "Do
  everything" leaves the bring-in out. **Undo** on the card takes the choice back; adding a
  file or starting a bring-in takes it back too. Offered only while the version never went
  through a bring-in (`POST /api/sessions/{id}/outside-inputs/skip`, body `{"skip": true|false}`).

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
