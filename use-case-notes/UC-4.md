# UC-4 -- PVFP (lines of business `LoB n`, external links)

What the app learned from this use case (moved out of the app's comments and messages;
the app itself names no client). Grouped by the app file the insight belongs to.

## app/prep.py

- **FRM-007 example**: INDIRECT("'"&$G710&"'!$Z$16") with G710 = "LoB 1" became 'LoB 1'!$Z$16 (the caution text and docstring now use a generic $G10 / 'Line 1').
- **INDIRECT sheet-name patterns**: text pieces like "'CoC "&n&"'!F14" name every sheet starting with 'CoC '.

## app/prep.py, app/numbers_check.py

- **row inserts vs 3-D refs and INDIRECT**: an ordinary Prep changed 2,462 computed values, 1,781 of them into errors: a 3-D reference =SUM('LoB 1:>>'!K65) reads the same cell on every sheet by position (a row inserted on one sheet shifts that sheet alone), and INDIRECT("'"&$E$2&"'!H"&n) builds an address as text ($E$2 holds 'LoB 1').

## app/validators/mind_values.py

- **AVERAGE/COUNT/MIN/MAX skip empty**: PVFP Reporting_LoB Total!H30: Mind skips an empty cell where Excel saw a 0 (1.8.9)
- **FRM-013 origin**: PVFP 'LoB 1'!AB78 = AB77-AB84 with AB84 = "": #VALUE! in Excel (kept by the fixer), 0 in Mind; the IFERROR cells fed by it ("No GEP" in Excel) became numbers in Mind
- **IF(ISNUMBER(..)) guard**: PVFP Reporting_General!I23 = IF(ISNUMBER('LoB 1'!H171*NOT(LoB3_Boolean)),...,"") -- Excel takes the fallback on a "" cell, Mind computes with "" as 0
- **guarded rewrite on an error cell**: PVFP: replacing IFERROR by IF(cell="",...) where the read cell holds an error changed 19 values, 9 into errors -> such rows keep their formula

## app/validators/references.py

- **FRM-006 origin**: PVFP (2026-09-30): 'First:Last'!A1 3-D reference -> "Sheet '...' not found in workbook ... on compiling formula"; divider tabs '>> Reporting', '>>>' without grids -> "Spreadsheet not found error" (2026-10-01)
- **FRM-007 origin**: PVFP (2026-10-01): INDIRECT used as a value -> "Unable to cast object of type 'AM.Models.AMReference' to type 'System.IConvertible'"
- **external links**: PVFP: 20,934 formulas read an external file on a colleague's drive -> converter reported '20,000 deterministic errors'
- **lookup tables are not cycles**: PVFP: once INDIRECT became direct references, =INDEX(Inputs!$I$152:$CZ$10240, ...) reached a cell reading the lookup's own sheet; Excel reported no circular reference (1.8.4)

## app/validators/structure.py

- **STR-008 origin**: PVFP: 'LoB 1'!G2 = "LoB 1" stood alone once Prep moved the label beside it into a title; "GWP_Retained#"&$G$2 came back as "GWP_Retained#" in Mind and 20,000 cells followed; test workbook 5 (2026-10-08) confirmed numbers alone / text under a '#' title / next to a label are imported

## external_links.py

- **external links**: PVFP: the link points to a file on a colleague's machine; 20,934 formulas live on the external-link cache (example formula ='[1]General Assumptions'!$D$13*H21).

## model_map.py

- **key cell without macro**: PVFP has no macro loop; its key is 'LoB 1'!G2; dropdown over a block of named flags LoB1_Boolean ... in G51:G60.
