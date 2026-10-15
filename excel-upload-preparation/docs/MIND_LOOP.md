# Mind loop

The macro's run loop written as Mind formulas, from the Model map. `app/loop_builder.py`,
`GET /api/sessions/{id}/mind-loop-plan` (the plan, nothing written),
`POST /api/sessions/{id}/mind-loop-plan/apply` (written as a new version), the **Mind loop**
section of the Model map screen, `python -m app.loop_builder <workbook> [--json]`.

## What is written

| The macro | In Mind |
|---|---|
| copies row *n* of a table into a driver row, *n* from a first cell to a last cell (Horizon) | every driver cell: `=MM_LOOP("Name", 'Table'!Xfirst:Xlast)`, its own column of the table |
| walks a list with `Range("start").Offset(0, i - 1)` (CNHI's scenarios, clusters) | the driver cell: `=MM_LOOP("Name", the stretch of the list)` |
| sets the driver to the counter itself, from a first cell to a last cell (Palermo) | `=MM_LOOP("Name", LastCell)` when the first value is 1; `=MM_LOOP("Name", LastCell, first, first, LastCell)` otherwise (the IndexStart / IndexEnd form; verified in Mind: instances first..last, MM_RESULT index 1 = the first) |
| computes the driver from the counter: `proj_year + i - 1` (CNHI's year) | the same expression with `MM_LOOP("Name", CountCell)` in the counter's place; the count cell is named by the user when the workbook has none |
| writes a source block into `Target!A<n>:J<n>` on every pass (Horizon) | row by row: `=MM_RESULT('Source'!$AY$35, "Name", k)` for loop value k |
| clears a results area before a run | listed with what an earlier run left there; emptied only when asked |

The loop name is the driver's defined name (`MP_ROW_SELECT`, `active_scenario`), else the
source sheet's name (`BASE_Polices`); the bounds are the bound cells' numbers (a formula's
stored value counts). All three can be changed before writing.

## What is not written, and said

- a block the macro places by a counter of its own (Palermo's `ByTariffPaster.Offset(0,
  10 * (Temp2 - 1))`): in Mind the results stay in the source block, one value per loop
  value; `=MM_RESULT(cell, "Name", k)` reads one;
- a block that depends on several nested loops (CNHI's paste blocks per scenario x year x
  cluster): `MM_RESULT` with one loop would sum over the others; read the source block
  with every loop named;
- a counted loop with no cell holding the count (CNHI's `For i_year_rac = 1 To 5`): name a
  free cell and the plan writes the count there;
- a values copy made once after the loop (Palermo's `Summary Output Pasted`): nothing to
  write, the sheet it copies from holds the results.

## The four models

- **Horizon**: `BASE_Polices`, 6 values (rows 1 to 6 of `BASE Polices`; row 1 is the
  header row, the macro runs it too -- set first to 2 to leave it out); 42 `MM_LOOP` cells
  on `CALCUL!A2:AP2`; `OUTPUT!A1:J6` from `CALCUL!AY35:BH35`, 60 `MM_RESULT` cells.
  Checked in Excel against `Bouton2_Cliquer`: 60 of 60 cells the same.
- **Palermo**: `MP_ROW_SELECT` on `Projection!C3`, 1 to 19 from `Macro!B2`; the per-policy
  block is placed by the tariff (not built), the summary paste is a values copy.
- **CNHI**: `active_scenario` over `'Model Inputs'!E8:R8`, `active_cluster` over
  `'Validation list'!B7:B26`; `active_year` needs a count cell; the 30 paste blocks depend
  on all three loops (not built).
- **PVFP**: no run loop (one sheet per LoB): nothing to build.

## Checking it

    python scripts/check_loop_plan.py <workbook.xlsm> --macro <Name>

runs the macro in Excel on a copy (macros allowed for that copy only), writes the plan into
another copy through the app's own Apply, plays every loop value into the driver cells and
compares the source cells of every `MM_RESULT` block with the macro's block. MM for Excel
(loaded into the app's Excel since 1.10.1) evaluates one pass only -- `MM_LOOP` gives the
first loop value, `MM_DIMSIZE` says 1 -- so the check plays the values itself: it proves
the wiring (which column, which row, which loop value); the run in Mind proves Mind.

## Verified in real Mind (2026-10-15)

Horizon, the Do-everything version (Fix all issues, Fix the formulas, the loop over rows 2 to 6,
recalculated clean), in a sandbox project:

- Upload / Convert / Test complete in about 90 s; the Test step counts Stochastic 0, Deterministic
  22,212 cells with 0 differences, Assumptions 163,741. The Deterministic test evaluates ONE loop
  instance (the workbook's cached values are the first pass, so they agree), so the Deterministic
  Differences log never shows the other passes: the loop is checked after the run, not there.
- Add template, Run: completed in 1.8 s, consistent with the audit trail.
- How to see a loop in Mind: the template selector in the project bar (the workbook's name) lists the
  sheets; a sheet lists its grids `Untitled(row,col)`; a grid opens a panel. The panel's play icon
  (top left, "Partial run") recomputes that grid; the panel then shows the dimension name
  (`BASE_Polices`), a "Display all dimensions" control and one row per loop instance.
- Every OUTPUT cell matched the macro run in Excel on the same version: 10 columns x 5 passes, pass k =
  loop value k + 1. Mind's own AI Assistant, asked in the project, answered "Defined loop BASE_Polices,
  size 5 ... read by 50 MM_RESULT cells".
- Mind reads a grid's first row as its header. Where the row above the block had empty cells, Mind
  split the block into four grids and showed pass 1 as the header of two of them (values right,
  display wrong). The loop builder now completes that header row from the source block's own
  captions (`_complete_header`).
- The loop written on the untouched original is refused by Mind: 15 formula compilation errors, all
  `SUMIFS('BASE Polices'!#REF!, ...)` on sheet C20, the broken references the app's REF-001 fix
  repairs. The loop needs the fixed workbook, which is what the main screen's buttons produce.

Scripts: `scripts/mind_loop_probe.py` (upload, logs, run, parts), `scripts/mind_loop_compare.py`
(Mind's rows vs the macro's values), `scripts/mind_ask.py` (Mind's AI Assistant),
`scripts/night_chain.py` (the Do-everything chain through the API).

## The two MM_LOOP forms, verified in Mind (2026-10-15, a tiny probe workbook)

`scripts/build_loop_forms_probe.py` writes a workbook with `K = MM_LOOP("K", last, 3, first, last)` (first =
3, last = 7) and `J = MM_LOOP("J", last)`, plus `MM_RESULT(cell, "K", k)` / `("J", k)` for k = 1..7. In Mind:

- `MM_DIMSIZE("K")` = 5 and `MM_DIMSIZE("J")` = 7; the first instance of K is 3, of J is 1.
- `MM_RESULT(x10, "K", k)` = 30, 40, 50, 60, 70 for k = 1..5 (k = 1 is loop value 3) and 0 for k = 6, 7;
  `MM_RESULT(x10, "J", k)` = 10..70. So the builder's `k = value - first + 1` is what Mind expects, and an
  index past the loop gives 0, not an error.
- `MM_RESULT` needs a cell as its first argument: with an expression (`$B$4*10`) Mind's test step fails at
  runtime with "trying to access null reference". The builder always passes a cell.

## The ranges follow the input table (1.11.1)

A table loop is written as `'Table'!X<first>:INDEX('Table'!X:X, MM_LASTROW('Table'!<key><first>))` in every
driver cell, the key column being the first column of the table filled on every loop row. Mind reads
`MM_LASTROW` from the table's grid, so a row added to the table (Input Manager, a new upload) joins the
loop; Excel's add-in stops at the key column's first empty cell, so the local values are today's rows.
Verified in Mind on Horizon (2026-10-15): the dimension spinner on an OUTPUT grid runs 1..5, and 1..6 after
a seventh policy row was appended through `scripts/add_policy_row.py`; Mind's test counts 0 deterministic
differences; the Excel check gives 50 of 50 cells against the macro.

Why not `MM_COLUMN(X2)` alone: in Excel it stops at the first empty cell of X, and when a column has an
empty cell inside the loop rows (Horizon's Q, Z, AA, AD) it collapses to one cell, which `MM_LOOP` then
reads as a count: wrong local values and an error text, and 895 deterministic differences in Mind's test
(Mind itself computed the right values). A list the macro walks to its last item is written with
`MM_COLUMN` / `MM_ROW` from its first item (lists have no gaps).

A column of the table that holds nothing in the loop rows gets no loop, and a loop written there by an
earlier build is cleared (it would be a definition with another length once the table grows). Every
definition of a loop must keep the same length, so a new row should fill every part of the table (Mind cuts
a table into grids at empty columns). `--fixed-ranges` on the command line, `fixed_ranges` on the API and
the box on the Advanced loop panel give the old `X2:X6` form. LOOP-003 treats table-following definitions
as equal by construction and flags a mix with fixed ones.

## Defaults and notes (1.11.1, after the design review)

- When the macro starts on the table's text header row, the loop starts one row (column) further: in Mind
  the header would be instance 1. Choosing first = the header row keeps the macro's header pass.
- A loop cut short (last before the end of the table) keeps fixed ranges: rows added later do not join it.
- Every table-following loop says that Excel may differ from Mind (the add-in stops at the key column's
  first empty cell), which growth reaches it (a new upload always; Mind's Input Manager only imports into a
  grid titled with /Input), and that the result block is written for the loop values present now: Mind
  computes every instance, the block lists those of today.
