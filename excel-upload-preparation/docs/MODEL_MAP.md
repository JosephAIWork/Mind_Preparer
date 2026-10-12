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

## The four models (2026-10-13)

**Horizon** (`Bouton2_Cliquer`): the loop copies row *n* of `BASE Polices`
(A:NI) into `CALCUL!A2:NI2`, *n* from `CALCUL!T5` ("Ligne debut") to
`CALCUL!T6` ("ligne fin", holds `=N5`), and writes `CALCUL!AY35:BH35` to
`OUTPUT!A{n}:J{n}`; `OUTPUT!A{first}:I100000` is cleared first. Inputs: BASE
Polices (one row per policy), Base Pret, BASE Personne, BASE Majo, Tarifs CRD,
Tarifs CI, Criteres, Parametres (label | value pairs, e.g. the medical
selection abatements). No dropdown.

**PVFP**: no run loop. `Macro1` copies `'LoB 1'` once per line of business and
writes the LoB name into `G2` (and `E2` of a reporting sheet that no longer
exists under that name): one sheet per item, keyed by `G2`. Switches: eight
dropdowns on `Inputs_PVFP` (`DAC_Model` H71, `LossRatio_UPR_Model` H72,
`GWP_Provided_Data` H73, H74, `Cost_of_Reinsurance_Model` H78 / `_Choice` I78,
`UPR_Model` H79) and `CoC!F2` (SCR / MSR); fifty `LoBn_Boolean` VRAI/FAUX
dropdowns (G51:G60, K, O, S, W); named settings `Last_Year_Triangle` G69
(6,785 readers), `Threshold` G65, `Yield_Curve_First_Cell` H82. Inputs:
Inputs_PVFP, Inputs_CoC, Yield_Curve. Outputs: the Reporting_* sheets.

**Palermo** (`UndiscountedCashflow`, entry point; calls QuickPaste and the
e-mail macros): the loop sets `Projection!C3` (MP_ROW_SELECT, "Policy run")
from `Macro!B1` (MacroStart_UDC2) to `Macro!B2` (MacroEnd_UDC2); per policy
it pastes `Macro!A11:J138` (AggregateOutputByTariff) into `Macro!V13:AE140`
(ByTariffPaster, shifted by tariff); after the loop `'Summary Output'!M9:EK48`
is pasted as values into `'Summary Output Pasted'!A1:DY40`; `SaveTotal` saves
that sheet to the path in `'Saved down'!C9`. Switches on `Macro!B3:B5`
(Yes/No dropdowns the macro tests): `Del_range2`, `Full_Run?`, `SendEmails`;
`'Tariffs to run'!C2:D28` (TarRun) says which tariffs run. Status cells the
macro writes: `Error_detection`, `Error_Flag`, timers. Named settings:
`Valuation_date` Control!C3, `Lapse_stress` Control!C7, `ME_applicable_months`,
the TBR and RPM fee terms. Inputs: Data (TarRange A5:C94940), the assumption
sheets (Lapse & PUP rates, Mortality rates, TBR inputs, Tariff mapping, Dynamic
increase rates, Economic / Non-economic assumptions, Securis % receivable).

**CNHI** (`combined_macro` -> baseline, baseline_projections,
specialised_services): three nested drivers on `Model Inputs`: `C8`
active_scenario (dropdown over E8:R8), `C173` active_year (proj_year + 0..4),
`C174` active_cluster (dropdown over the `cluster` list, Validation list!B7
down); scenarios from `Output!G2` (start_1) to `Output!G3` (end_1). Outputs:
`'RAC Baseline'!V12:AJ117` from `'Baseline Cost'!AB18:AP123` per cluster, the
CatA..D blocks from `RAC Calculation`, `Output!B9:U...` (paste_output, one
236-row block per scenario x year x cluster) from `'Output working'!AB13:AU248`,
the SS Output and Inbound Output blocks per scenario; `Output!B9:U330408` and
`X9:AO1408` are cleared before a run. Named settings: `proj_year`
('Assumption Setting'!D8), `active_year`. Inputs: Assumption Setting, Model
Inputs, Budget Inputs, Cost data, IRAT data, Population, Validation list.

## Limits

- A loop driven by something other than `For ... To` (Do / While, a
  Worksheet_Change event) is not read; the key cells then stand in.
- A per-item target written through `.Offset` of a variable that is not the
  loop variable (Palermo's tariff offset, CNHI's `tracker`) is reported as
  "same place, overwritten" though it moves.
- Sheet roles come from names: "Policy summary" is listed under outputs
  because of "summary" though it is the calculation sheet.
- `.xlsb` macros are read too; a workbook without macros has no loop section
  beyond the key cells.
