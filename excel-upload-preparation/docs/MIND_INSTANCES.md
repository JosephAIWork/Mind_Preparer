# Mind instances

The macro's run loop turned into Mind instances: one copy of the model per value of a key column of
the model-point table (a tariff, a product, a group), a loop inside each copy over that copy's own
rows, and a second small workbook -- the main one -- with the keys and the results. Added in 1.12.5.
`app/instance_builder.py`, `GET /api/sessions/{id}/mind-instances-plan?key=&keys=present|all` (the
plan, nothing written), `POST /api/sessions/{id}/mind-instances-plan/apply` (written as a new version,
the main workbook beside it, both in one zip), `GET /api/sessions/{id}/mind-instances/download` (the
zip), and the **Mind instances** section under Advanced > Mind loop, below the loop panel.

## Why instances here

The Mind loop (`docs/MIND_LOOP.md`) covers a macro that runs one row of a table per pass and writes
that pass's results to a block: `MM_LOOP` in the driver cell, `MM_RESULT` row by row. It does not
cover a macro that runs one policy at a time and *adds* each policy's results to its group's block --
one block per tariff, placed by a counter of the macro's own. One loop over every policy gives one sum
over all of them, and the groups are lost; the loop builder lists such a block as not built.

Instances keep the groups. Mind makes one copy of the whole model per key; each copy sees only its own
rows of the table; the loop inside the copy runs those rows and `MM_RESULT` adds them up, so the copy's
block is the group's block, as the macro built it. A second workbook, which Mind does not copy, holds
the list of keys and reads every copy's results side by side. The price is Mind's: one copy per key,
each with the whole model, computed side by side (Mind's own note: instances use more memory and
processor than loops).

## How Mind instances work

From Mind's documentation, checked in a real Mind project on 22 October 2026 (`docs/MIND_VS_EXCEL.md`,
"Instances demo") and, on 28 October 2026, with the two files of the hand-built prototype this builder
mirrors: they converted, ran, and the results-by-key grid matched the macro's sums by key. On
29 October 2026 this builder's own zip (the fixed model cut to the macro's 19 test policies, 3 keys)
went through the same check: both files converted and tested as one project, the run took 3 seconds,
and both grids of the main workbook equalled Excel's sums by tariff -- 384 cells of the results-by-key
grid and 1,280 cells of the all-copies grid, 0 differences, Total row equal.

- **Instances are copies of one workbook.** Mind makes them from a grid titled `#<Name> /InstanceKeys`
  in the **main workbook**, the one workbook Mind never copies. Each column header of that grid is the
  file name, without its extension, of a workbook to copy; each row under it is one key, and one copy.
- **The split.** In the copied workbook, a grid titled `#<Name> /InstanceSplit` is cut by its **first
  column**: each copy keeps only the rows whose first column holds that copy's key.
- **A loop whose size differs per copy:** `MM_LOOPINSTANCE(Name, MaxSizeRange, InstanceSizeRange[,
  Index, IndexStart])`. With two single cells it runs from 1 to the smaller of the two: the first cell
  holds the biggest count of any copy, the second this copy's count -- `=COUNTA(<the split grid's key
  column>)`, which in Mind counts only the copy's rows. In Excel the function returns 1.
- **Reading the results.** `MM_RESULT(cell)` with no loop index is the sum over the loop -- what the
  macro accumulates. `MM_INSTANCE(cell, key or index)` reads one copy. A plain reference from another
  workbook to a cell of the copied workbook is the sum over the copies (checked in a real Mind project
  on 29 October 2026: the prototype's all-copies grid, 1,280 cells and its Total row, equal to the
  macro's sums). `MM_INSTANCEKEY("default")` is this copy's key; the text given is what Excel shows.
- **Where the results show.** A run's values are visible in the browser session that ran it, or after
  the project is saved (the floppy icon top left, then Save). A page reload without a save shows the
  uploaded workbook's Excel values again -- not an error of the model (it cost the prototype six
  rounds).
- **Grids.** A title starts with `#` and sits directly above the grid's top-left header cell; header
  cells are text; text cells touching each other form one grid, so the row above a title must be empty
  across the grid's width, or the title merges into the grid above it. Without the `#` the title is read
  as a header row and every flag in it is lost: for `/InstanceKeys` that means no copies at all, and no
  error.
- **Between workbooks** a reference is an ordinary Excel external reference to a file in the same
  folder: `'[model.xlsm]Sheet'!A1`. Excel stores add-in calls as `_xll.MM_...`; Mind reads them as they
  are, and the app's rules strip the prefix.

## Checked again in Mind on 10 November 2026 (small probes)

Three small invented projects (`zzz_mindready_inst_probe_1` to `_3`, notes in
`%LOCALAPPDATA%/MindReady/tokyo_instances_2026-11-10/probe/PROBE_NOTES.md`). Every number Mind gave
matched the one worked out by hand. Nothing in them contradicts the builder:

- `MM_LOOPINSTANCE(<loop size cell>, <count cell>)` with the count `=COUNTA(<split grid key column>)`
  runs each copy to its own count; the count grew by itself when an input file added a row.
- `MM_INSTANCE(cell, key)` gives the same value with a key cell, a literal key or the key's index.
  On a looped cell it gives only the **first pass**. The builder's `#Results by <key>` grid reads the
  output block, whose cells are `MM_RESULT` sums, so it gets the whole loop. Keep it that way: never
  point `MM_INSTANCE` at a calculation cell the loop changes.
- A plain reference from the main workbook to a copied cell is the sum over the copies (the
  `#All copies` grid).
- **Inputs per copy.** An `/Input` grid in the copied model takes one file per copy, named
  `<key>-<grid>.csv`; a file named `<grid>.csv` alone is not matched (Mind asks which grid). A split
  `/InstanceSplit /Input` grid takes one file with every key's rows when Mind's **Split between
  instances** switch is on. An `/Input` grid in the main workbook is read by every copy: one file
  updates them all (checked for a single cell, not yet for a range).
- Mind's **Export grid** gives only the selected copy unless **Export all loaded instances** is on,
  and the CSV has no key column.

## What the builder changes in the model

The plan comes from the analysis, the model map and the loop plan (`build_instance_plan`). It takes the
first loop of the loop plan that is built and counted: a counter loop whose last value is a cell. The
table is the largest INDEX / OFFSET range whose row argument is the driver (by its defined name or its
address). Every change is listed before it is written, with the address it has at the moment it runs
(the inserts come first), and written through Excel into a new version like any Prep change.

1. **The table becomes a split grid.** A key column becomes its first column, holding the key of every
   row (one write for the whole column, the new `set_range_values` change), under the header
   `Instance key`; the title `#<Name> /InstanceSplit` goes directly above it, with an empty row above
   the title. The column is inserted before the table unless its first column already is the key
   column; 0, 1 or 2 rows are inserted so that the title row and the row above it are empty across the
   table's width. Filled columns touching the table on its left (a row number only the macro used,
   say) are part of the grid in Mind's eyes: the table is taken to start there, the key column goes
   before them, and a column without a header text gets `Column <letter>` (a grid's first row is its
   header).
2. **The count cell** -- the cell the macro reads as the loop's last value -- gets
   `=COUNTA(<key column>)`: each copy counts its own rows.
3. **A free cell gets the biggest count**, the loop size, with a label beside it.
4. **The driver cell** gets `=MM_LOOPINSTANCE("<loop name>", <size cell>, <count cell>)`. When the
   range the loop indexes starts on the table's header row (the macro ran `2 To n`), `+1` is added:
   `MM_LOOPINSTANCE` counts from 1 and `INDEX(range, 1)` would read the header, so the driver reads
   the data rows. A macro that starts further into the table is noted; every row of the table joins
   its copy's loop.
5. **The output block** -- every formula cell of the block the macro copies out on each pass, the loop
   plan's "from" block, where a formula adds the pass's value from the calculation sheet to what the
   paste area holds (`=<calculation cell>+OFFSET(<paste area>, ...)`) -- becomes
   `=MM_RESULT(<the first single cell of another sheet the formula reads>)`: the sum over the copy's
   loop, the macro's accumulation.
6. **The paste area** the macro wrote into is emptied (the loop plan's clears).
7. **A free cell shows the copy's key:** `=MM_INSTANCEKEY("<first key>")`, with a label.

Cells that only read the paste area are left as they are and listed in the notes: in a copy they show 0.

**The key.** The candidates are the table's text columns with 2 to 200 distinct values (the first
5,000 rows are sampled for the candidates; the chosen column is read whole). The default is (a) the
column the macro itself looks the loop row up in -- `WorksheetFunction.VLookup(<loop variable>,
Range("<name>"), k, ...)` over a name that covers the table: its k-th column -- else (b) the first
candidate whose header says tariff, product, group, portfolio, segment, class, category, type, country,
fund, treaty or cohort, else (c) the candidate with the fewest distinct values. A key given by the user
wins. `keys=present` (the default) makes one copy per key value that has rows; `keys=all` one per value
of the list the macro reads, when the app finds that list (else present).

**The model's name.** Mind finds the workbook to copy by its file name, so the model is written as
`<stem>_instances` with its own extension (the stem reduced to letters, digits and underscores), and
that name is the header of the `/InstanceKeys` grid. Do not rename the files after the download.

## What the main workbook holds

`<stem>_Main.xlsx`, written through Excel with the MM for Excel add-in, next to the model, so its
references to the model resolve. Mind never copies it.

| Sheet | Grid | What it holds |
|---|---|---|
| Instances | `#<key header> instances /InstanceKeys` | header = the model's file name without extension; one row per key |
| Summary | `#Results by <key header>` | one row per key, one column per row of the output block; each cell `=MM_INSTANCE('[<model file>]<sheet>'!<last column of that block row>, <key cell>)` |
| Summary | `#All copies` | the whole output block as plain links to the model (`='[<model file>]<sheet>'!<cell>`), with its header row (the row above the block, when it holds text) and a Total row (`=SUM`). In Mind a plain reference to a copied workbook is the sum over the copies |

A grid that would hold more than 30,000 formulas is left out of the main workbook (the plan says so in a
note). In Excel the main workbook shows the one copy Excel has: every `MM_INSTANCE` cell and every plain
link the model's own value. The copies' values come only from a run in Mind, and show in the session
that ran or after a save. Keys are written as they are: a text key like `001` stays text,
in the model's key column and in the main workbook alike (`excel_text` prefixes what Excel would turn
into a number, a date or a boolean).

## Uploading the two files to one Mind project

1. On the Mind instances section: keep the key found or pick another; tick **Every key of the list,
   even without rows** if every value of the macro's list should get a copy; click **Preview**; read
   the cards (Key, Loop, Output block, Main workbook) and the notes; then **Write it into a new
   version**.
2. Click **Download both files (zip)**: `<stem>_instances.zip` holds the model and the main workbook,
   side by side.
3. In Mind, make an empty project and drop **both files on it together**, in one upload. The two files
   of the hand-built prototype converted and tested as one project on 28 October 2026, this builder's
   own zip on 29 October 2026. Keep the file names: the `/InstanceKeys` header is the model's file
   name, and a renamed model has no copies -- Mind gives no error for it.

## What to expect in Mind

- **The upload check shows 0 for the copies after the first.** Mind's Test step compares each cell with
  the values the workbook was saved with, before any run and without the copies' own values; an
  `MM_INSTANCE` cell for the second copy and beyond reads 0 there. It is a false alarm (seen, and run
  through, on 22 October 2026): the values come with the run. Do not fix a formula for it.
- **After a run** the Results grid of the main workbook shows one row per key and the All copies grid
  the sum over the copies -- where Excel showed one copy and one loop pass. These differences from
  Excel are the design, not errors. Read them in the session that ran, or save the project first: a
  reload without a save shows the Excel values again.
- **The run against the macro's values** (29 October 2026, this builder's zip of the fixed model cut
  to the macro's 19 test policies): the results-by-key grid, 3 keys x 128 rows, 0 differences from
  Excel's sums by tariff; the all-copies grid, 128 x 10, 0 differences; its Total row equal
  (2,356.99 / 3,488.60 / 5,845.58 for the three non-zero fees). Conversion 78 s, run 3 s.
- **Cost.** One copy per key, each with the whole model. The copy count and the run time of the full
  portfolio are Mind's cost, which the builder does not estimate.

## What the builder does not do yet

- Cells that only read the paste area are left as they are: the area is empty in a copy, so they show
  0. They are listed in the plan's notes.
- A loop without a count cell (`For i = 1 To 5`) is not built: the plan says `needs_decision` and why.
  Choose a count cell on the loop panel first, then build the instances.
- The copy count and the run time of the full portfolio are Mind's cost; the builder states the keys
  and their row counts, nothing more.
- One loop -- the first counter loop of the loop plan -- and one key column; nested loops and a second
  key are not built. The instance formulas are written, not checked: no rule reads `MM_INSTANCE` yet.
- The Mind check covered the fixed model cut to the macro's 19 test policies (3 keys). The full
  portfolio (94,932 policies, 26 keys, a loop of 39,477 passes in the biggest copy) has not been through
  Mind: the copy count and the run time are Mind's cost.
- The source must be Mind-ready first: the raw use-case file failed Mind's conversion on its own
  (an array formula wider than its grid, a defined name) before any instances work. Run the fixes,
  then build the instances.
