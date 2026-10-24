# UC-1 -- Horizon (French loan-insurance model, macro loop over policies)

What the app learned from this use case (moved out of the app's comments and messages;
the app itself names no client). Grouped by the app file the insight belongs to.

## app/validators/mind_values.py

- **FRM-008 origin**: Horizon (2026-10-06): =IFERROR(1+INT((D192-1)/12),"") blank in Excel, a value in Mind on 1,242 cells (CALCUL E192:E812, F192:F812); =IF(D192="","",IFERROR(...)) fixes both
- **FRM-009 origin**: Horizon: H9 = VLOOKUP(B2,'Base Pret'!A:K,11,0) landed on an empty cell (deferral period empty for that loan); IF(D35<=$H$9,0,1) returned 0 on 157 live rows, reserves wrong; C20!C15 = +OUTPUT!P4 compared TRUE in Excel, 0 in Mind
- **FRM-010 origin**: Horizon: =1+($C$10=1)*IFERROR(VLOOKUP(F192,Criteres!$N$11:$O$33,2,0),...) came back NaN on 621 rows -- Mind matched the empty header cell of the table Prep had titled

## fast_xlsx.py

- **parser benchmark**: Horizon: stdlib expat iterparse 5.4 s vs defusedxml 7.4 s vs lxml 10.6 s.

## formula_utils.py

- **find_calls_detailed**: Horizon: about a million find-call questions per scan; a formula that does not contain the function name cannot call it (ASCII shortcut).
- **reference-only arguments**: 2026-09-24: an assistant fix rewrote SUMIFS('BASE Polices'!#REF!, ...) to SUMIFS(NA(), ...) in 17 cells and Excel refused every one (COM 0x800A03EC) -> REFERENCE_ARGS list.

## model_map.py

- **labelled settings**: 'label | value' pairs live on Horizon's Parametres (label on one row, number on the next), Palermo's Macro and Control sheets, CNHI's Model Inputs.
- **per-item driver**: Horizon: driver CALCUL!A2:NI2 <- BASE Polices row n.
