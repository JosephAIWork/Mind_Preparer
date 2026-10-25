# Model map

Where a model's inputs are, which cell the run loop sets, which settings stay
for a run but change between runs (and how they are toggled), and where the
outputs land. `app/model_map.py`, the route `GET /api/sessions/{id}/model-map`,
the screen **Model map** (Tools), and `python -m app.model_map <workbook>`.

Everything is read off the workbook itself. Nothing is guessed and no model is
asked:

| Source | What it gives |
|---|---|
| VBA (olevba) | `For` loops: the cell set on every pass (driver), the first / last value cells (bounds), the ranges results are pasted into (outputs), the named cells the macro tests (switches), areas cleared before a run, save paths, which macro is the entry point |
| Defined names | what a range is for; a single-cell name on a constant that formulas read is a setting |
| Data validations | a list dropdown is a switch, with its options |
| Sheet names | input / data / assumptions -> inputs; output / results / reporting / summary -> outputs |
| Formulas | the constants read most by absolute reference (the key of a model without a loop) |

Each item carries its evidence (the macro statement, the name, the validation)
so a person can check it; **Show** on the screen opens the cells around it.

## Four client models (2026-10-13)

**UC-1** (a macro loop over policies): the loop copies row *n* of the policy
table into one calculation row, *n* from a start cell ("first row") to an end
cell ("last row", itself a formula), and writes a block of ten results to row
*n* of the output sheet; the output area is cleared first. Inputs: one row per
policy plus loan, person and tariff tables and label | value parameters (e.g.
the medical selection abatements). No dropdown.

**UC-4** (no run loop): a macro copies a template sheet once per line of
business and writes the line's name into one cell of the copy: one sheet per
item, keyed by that cell. Switches: eight dropdowns on the input sheet (model
choices for DAC, loss ratio / UPR, the data provided, reinsurance cost), one
on the cost-of-capital sheet (SCR / MSR); fifty TRUE / FALSE dropdowns, one per
line and item; named settings such as the last triangle year (6,785 readers),
a threshold and the first cell of the yield curve. Outputs: the reporting
sheets.

**UC-2** (an entry-point macro that also calls a paste macro and e-mail
macros): the loop sets the policy-run cell from a start cell to an end cell on
the macro sheet; per policy it pastes a per-tariff result block into an area
shifted by tariff; after the loop a summary block is pasted as values into
another sheet, which a macro saves to a path held in a cell. Switches: three
Yes / No dropdowns the macro tests (clear a range, full run, send e-mails); a
table says which tariffs run. Status cells the macro writes: error detection,
an error flag, timers. Named settings: the valuation date, a lapse stress, fee
terms. Inputs: a data table of about 95,000 rows and the assumption sheets
(lapse, mortality, tariff mapping, economic / non-economic assumptions).

**UC-3** (one macro calling three): three nested drivers on the inputs sheet:
the active scenario (a dropdown), the active year (projection year + 0..4) and
the active cluster (a dropdown over a list); scenarios from a start cell to an
end cell. Outputs: per cluster, a block copied from the baseline cost sheet;
category blocks; one 236-row block per scenario x year x cluster pasted into
the output sheet from a working sheet; two more output blocks per scenario;
the output areas are cleared before a run. Named settings: the projection
year, the active year. Inputs: assumption settings, model inputs, budget
inputs, cost and population data, the list sheet.

## Limits

- A loop driven by something other than `For ... To` (Do / While, a
  Worksheet_Change event) is not read; the key cells then stand in.
- A per-item target written through `.Offset` of a variable that is not the
  loop variable (UC-2's tariff offset, UC-3's tracker counter) is reported as
  "same place, overwritten" though it moves.
- Sheet roles come from names: a sheet called "... summary" is listed under
  outputs because of "summary" though in one client model it is the calculation sheet.
- `.xlsb` macros are read too; a workbook without macros has no loop section
  beyond the key cells.
