# UC-5 -- Phoenix TVOG (170 MB, pivot tables)

What the app learned from this use case (moved out of the app's comments and messages;
the app itself names no client). Grouped by the app file the insight belongs to.

## app/estimate.py

- **estimate calibration**: the time estimate's constants were fitted on four Do everything runs: Palermo (27 MB) and Phoenix TVOG (170 MB, one per pivot setting); FALLBACK_SCAN_S_PER_FORMULA = 0.0035 s/formula measured on Palermo and CNHI.

## app/validators/structure.py

- **STR-009 origin**: Phoenix TVOG (2026-10-18): rejected with its 8 pivot tables ("The method or operation is not implemented", a grouped field); accepted with them as values

## autofix.py

- **multi-area writes**: Phoenix: clearing 10,002 error constants one Excel call each took most of the fixer's 977 s of writing -> multi-area ranges.

## excel_worker.py

- **Excel per step**: Opening the 179 MB Phoenix version in Excel took a minute per step of Do everything; one held Excel per session (excel_worker) removed those opens.

## inventory.py

- **pivot grouped fields**: Phoenix: a grouped pivot field is what Mind's converter could not read.

## pivots.py

- **pivot tables**: Phoenix TVOG: Mind's converter stopped with "The method or operation is not implemented"; with its pivot tables turned into values the same file was accepted.
