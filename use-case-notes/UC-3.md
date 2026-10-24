# UC-3 -- CNHI (37 MB, health-cost model, switch cells year / cluster / scenario)

What the app learned from this use case (moved out of the app's comments and messages;
the app itself names no client). Grouped by the app file the insight belongs to.

## app/theme_colors.py

- **theme colours cell by cell**: the Excel cell-by-cell recolouring never converged on Palermo (11,807 -> 11,719 -> 11,720 cells on one sheet: Excel keeps the theme reference for the default text colour) and took about an hour for one 123,686-cell sheet of CNHI (60 more sheets behind it); hence the style-table conversion.

## app/validators/mind_values.py

- **FRM-011 origin**: CNHI Wave!P38: =BB62&"<line break>/ Beneficiary - "&D8 came back from Mind as another text; =BB62&CHAR(10)&"x" came back the same (diagnostic sheet, 2026-10-08)

## formula_utils.py

- **scan memo**: CNHI: 1.7 million cell_refs_in_formula calls for 245,847 formulas in one scan, one rule after another -> formula_memo works each answer out once per formula text. The mask cache (65,536) is far smaller than CNHI's formula count.

## inventory.py

- **MM_ formulas**: CNHI: the rules ask for one MM_ function after another (38 passes per scan) over 245,847 formulas -> _mm_formulas filters once.
- **bounded workbook cache**: CNHI (36 MB): unbounded per-version workbook caches took the web server past 13 GB and the machine out of memory -> _BoundedCache.

## loop_builder.py

- **derived driver**: CNHI: active_year = proj_year + i - 1.
- **dropdown drops MM_LOOP**: 2026-10-16: Mind dropped the MM_LOOP written into CNHI's two driver cells because they carry a list dropdown (also seen in a probe workbook) -> remove_validation op.

## model_map.py

- **labelled settings**: 'label | value' pairs live on Horizon's Parametres (label on one row, number on the next), Palermo's Macro and Control sheets, CNHI's Model Inputs.
