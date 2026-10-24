# UC-2 -- Palermo (27 MB, macro loop over model points `MP_ROW_SELECT`)

What the app learned from this use case (moved out of the app's comments and messages;
the app itself names no client). Grouped by the app file the insight belongs to.

## app/estimate.py

- **estimate calibration**: the time estimate's constants were fitted on four Do everything runs: Palermo (27 MB) and Phoenix TVOG (170 MB, one per pivot setting); FALLBACK_SCAN_S_PER_FORMULA = 0.0035 s/formula measured on Palermo and CNHI.

## app/theme_colors.py

- **theme colours cell by cell**: the Excel cell-by-cell recolouring never converged on Palermo (11,807 -> 11,719 -> 11,720 cells on one sheet: Excel keeps the theme reference for the default text colour) and took about an hour for one 123,686-cell sheet of CNHI (60 more sheets behind it); hence the style-table conversion.

## app/validators/mind_values.py

- **FRM-012 origin**: Palermo 'Saved down'!C5 = MID(CELL("filename",'Summary Output'!A1),FIND("[",CELL("filename",...))+1,...) stopped the Mind run: "startIndex ('-1') must be a non-negative value" (2026-10-08)
- **defined names as lookup ranges**: Palermo: VLOOKUP(Policy_tariff,Assumptions_array,MATCH(B67,Assumption_headings,0),FALSE) -- FRM-009 must follow defined names that stand for a range (1.8.12)

## fast_xlsx.py

- **XML parser**: Palermo (27 MB, 5.7 million used cells): half of the load time was defusedxml's Python handlers; the data_only second load cost another six minutes.

## formula_utils.py

- **mask_strings**: Palermo: 4.4 million mask_strings calls for 57,537 formulas; most hold no quote, so the early return skips them.

## inventory.py

- **blank cells**: Palermo: ws.iter_rows() creates 5.4 million blank cells for the used rectangle -> iterate ws._cells.
- **style proxies**: Palermo: reading cell.font/fill/number_format built 13 million style proxies (two minutes) -> styles worked out once per (font, fill, numFmt).

## loop_builder.py

- **counter driver**: Palermo: Projection!C3 = polProcess is the loop counter itself.
- **stale paste area**: Palermo: Macro!A13 = ... + OFFSET(V13, ...) over the per-tariff results the macro pastes (Macro!V13:JU140, 728 stale values from an old run); the macro clears them first, so the loop plan clears them too.

## model_map.py

- **labelled settings**: 'label | value' pairs live on Horizon's Parametres (label on one row, number on the next), Palermo's Macro and Control sheets, CNHI's Model Inputs.

## recalc.py

- **error cells**: Palermo: reading the whole used range (values + formulas) to find error cells took 47 s on 5.7 million cells at every recalculation -> SUMPRODUCT count + SpecialCells.
