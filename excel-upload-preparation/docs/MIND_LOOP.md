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
| sets the driver to the counter itself, from a first cell to a last cell (Palermo) | `=MM_LOOP("Name", LastCell)` when the first value is 1; `=MM_LOOP("Name", LastCell, first, first, LastCell)` otherwise (the documented IndexStart / IndexEnd form, to be checked on the first run) |
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
