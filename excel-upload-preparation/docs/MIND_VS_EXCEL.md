# Where Mind computes another value than Excel

What is known, and how it is known. Everything here was observed in Mind on
2026-10-06; nothing is taken from documentation or guessed. Mind's own values
were not available -- only its difference list (which cells differ from Excel,
and the kind: "Mind returns 0", "Excel blank, Mind has value", "Mind returns
NaN", "Values differ"). Where a sentence below says what Mind *does* with a
value, it is the reading that fits every cell of those lists.

## The case

Horizon, after "Fix all automatically" (1.8.0): zero error cells in Excel,
verdict "Acceptable by Mind". Uploaded to Mind: **4,352 cells differ**, in 22
groups. They come down to two causes.

| | Where | Cells | Cause |
|---|---|---|---|
| 1 | CALCUL rows 192-812 (after the loan ends), columns E, F, then C, O, P | 3,105 | empty text in arithmetic (FRM-008) |
| 2 | CALCUL rows 35-191 (the live months), columns H, I, J, then G, R, S, Z, AA and eight totals; C20!C18 | 1,247 | an empty cell compared with a number (FRM-009) |

1. `E192 = IFERROR(1+INT((D192-1)/12),"")` and `F192 = IFERROR($C$6+A192-YEAR($H$6),"")`,
   where D192 and A192 hold `""`. Excel: blank. Mind: a value. C, O and P read F.
2. `H9 = VLOOKUP(B2,'Base Pret'!A:K,11,0)`; column K (the deferral period) is
   empty for this loan, so Excel shows 0. `I35 = IF(D35<=$H$9,0,1)*SUMIFS(...)`
   is 145,000 in Excel and 0 in Mind; J the same; H is multiplied by `(I35>0)`.
   `C20!C15 = +OUTPUT!P4` (an empty cell) and `C18 = +C3+C7-C11=C15`: TRUE in
   Excel, 0 in Mind.

The first explanation for cause 2 was wrong, and the test workbook is what
showed it: the SUMIFS with `">="&DATE(...)` criteria was suspected, a helper
column was built to avoid it, Mind still returned 0 -- and the same SUMIFS
matched in the test workbook. H9 had been ruled out because it was not in the
difference list; neither was C20!C15, which does behave differently. **A cell
that is absent from Mind's difference list can still be the cause**: an empty
cell that Excel shows as 0 is not reported as a difference.

After the rewrite of both (`Horizon_v4`): Mind reports no difference.

## The test workbook

`scripts/build_mind_test_workbook.py` builds it (invented data, 49 formulas on
the sheet `Probe`, Excel's value typed next to each). "Mind" below is its
difference list: *same* = not in the list.

### A. SUMIFS with a criterion built from a date

| Cell | What it tests | Formula | Excel | Mind |
|---|---|---|---|---|
| C2 | Horizon's formula: whole columns, criterion ">="&DATE() | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!C:C,">="&DATE(Params!$B$2,1,1),Loans!C:C,"<"&DATE(Params!$B$2,12,31))` | 60000 | same |
| C3 | the same on the table's rows only | `=SUMIFS(Loans!$D$2:$D$13,Loans!$A$2:$A$13,Params!$B$1,Loans!$C$2:$C$13,">="&DATE(Params!$B$2,1,1),Loans!$C$2:$C$13,"<"&DATE(Params!$B$2,12,31))` | 60000 | same |
| C4 | control: loan id only | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1)` | 255000 | same |
| C5 | date criterion alone: ">="&DATE() | `=SUMIFS(Loans!D:D,Loans!C:C,">="&DATE(2023,1,1))` | 620000 | same |
| C6 | date criterion typed as a number: ">=44927" | `=SUMIFS(Loans!D:D,Loans!C:C,">=44927")` | 620000 | same |
| C7 | date criterion with equals: DATE() | `=SUMIFS(Loans!D:D,Loans!C:C,DATE(2023,1,1))` | 150000 | same |
| C8 | the text the criterion becomes | `=">="&DATE(Params!$B$2,1,1)` | >=44927 | same |
| C9 | comparison criterion on a column of plain numbers | `=SUMIFS(Loans!D:D,Loans!F:F,">="&2023)` | 620000 | same |
| C10 | ">="&DATE() against the dates held as plain numbers (column G) | `=SUMIFS(Loans!D:D,Loans!G:G,">="&DATE(Params!$B$2,1,1))` | 620000 | same |
| C11 | candidate 1: helper year column, equals criteria only | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!F:F,Params!$B$2)` | 60000 | same |
| C12 | candidate 2: SUMPRODUCT on the table's rows | `=SUMPRODUCT((Loans!$A$2:$A$13=Params!$B$1)*(Loans!$C$2:$C$13>=DATE(Params!$B$2,1,1))*(Loans!$C$2:$C$13<DATE(Params!$B$2,12,31))*Loans!$D$2:$D$13)` | 60000 | same |
| C13 | candidate 3: the bounds typed as numbers in cells | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!C:C,">="&Params!$B$3,Loans!C:C,"<"&Params!$B$4)` | 60000 | same |
| C14 | candidate 3b: the bounds in cells holding =DATE(), shown as numbers | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!C:C,">="&Params!$B$5,Loans!C:C,"<"&Params!$B$6)` | 60000 | same |
| C15 | candidate 4: ">="&(DATE()+0) | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!C:C,">="&(DATE(Params!$B$2,1,1)+0),Loans!C:C,"<"&(DATE(Params!$B$2,12,31)+0))` | 60000 | same |
| C16 | candidate 5: ">="&INT(DATE()) | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!C:C,">="&INT(DATE(Params!$B$2,1,1)),Loans!C:C,"<"&INT(DATE(Params!$B$2,12,31)))` | 60000 | same |
| C17 | candidate 1 on a row after the end (the year is "") | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!F:F,Params!$B$7)` | 0 | same |
| C18 | candidate 1 on a row after the end (""+1) | `=SUMIFS(Loans!D:D,Loans!A:A,Params!$B$1,Loans!F:F,Params!$B$7+1)` | 0 | same |

### B. A cell whose formula returns "" (empty text)

| Cell | What it tests | Formula | Excel | Mind |
|---|---|---|---|---|
| C19 | Horizon column E: IFERROR(1+INT((""-1)/12),"") | `=IFERROR(1+INT((Params!$B$7-1)/12),"")` | (blank) | **different** |
| C20 | Horizon column F: IFERROR(51+""-2022,"") | `=IFERROR(51+Params!$B$7-2022,"")` | (blank) | **different** |
| C21 | candidate: IF(x="","",...) in front | `=IF(Params!$B$7="","",IFERROR(1+INT((Params!$B$7-1)/12),""))` | (blank) | same |
| C22 | N("") | `=N(Params!$B$7)` | 0 | same |
| C23 | MIN(90,"") | `=MIN(90,Params!$B$7)` | 90 | same |
| C24 | Horizon column C: 14+N("")-N(64) | `=14+N(Params!$B$7)-N(Params!$B$9)` | -50 | same |
| C25 | candidate for column C: IF(x="",0,x) | `=14+IF(Params!$B$7="",0,Params!$B$7)-IF(Params!$B$9="",0,Params!$B$9)` | -50 | same |
| C26 | candidate for column O: IF(x="",90,MIN(90,x)) | `=IF(Params!$B$7="",90,MIN(90,Params!$B$7))` | 90 | same |
| C27 | "">=156 inside IF | `=IF(Params!$B$7>=156,0,1)` | 0 | same |
| C28 | ""<=0 inside IF | `=IF(Params!$B$7<=0,0,1)` | 1 | same |
| C29 | (""<=156)*1 | `=(Params!$B$7<=156)*1` | 0 | same |
| C30 | x="" | `=Params!$B$7=""` | TRUE | same |
| C31 | ISNUMBER("") | `=ISNUMBER(Params!$B$7)` | FALSE | same |

### C. A reference to an empty cell, compared

| Cell | What it tests | Formula | Excel | Mind |
|---|---|---|---|---|
| C32 | Horizon C20!C15: =+(an empty cell) | `=+Params!$B$8` | 0 | same |
| C33 | Horizon C20!C18: 2311+0-2311 = that cell | `=2311+0-2311=C32` | TRUE | **different** |
| C34 | 0 = (an empty cell), directly | `=0=Params!$B$8` | TRUE | **different** |
| C35 | candidate a: N(empty cell) | `=N(Params!$B$8)` | 0 | same |
| C36 | the check against candidate a | `=2311+0-2311=C35` | TRUE | same |
| C37 | candidate b: empty cell + 0 | `=Params!$B$8+0` | 0 | same |
| C38 | the check against candidate b | `=2311+0-2311=C37` | TRUE | same |
| C39 | candidate c: IF(cell="",0,cell) | `=IF(Params!$B$8="",0,Params!$B$8)` | 0 | same |
| C40 | the check against candidate c | `=2311+0-2311=C39` | TRUE | same |
| C41 | (empty cell)="" | `=Params!$B$8=""` | TRUE | same |
| C42 | candidate d: N() on the check's side | `=2311+0-2311=N(C32)` | TRUE | same |

### D. A lookup that lands on an empty cell, compared

| Cell | What it tests | Formula | Excel | Mind |
|---|---|---|---|---|
| C43 | Horizon H9: VLOOKUP that lands on an empty cell | `=VLOOKUP(Params!$B$1,Loans!A:H,8,0)` | 0 | same |
| C44 | Horizon column I: IF(1<=that cell,0,1)*60000 | `=IF(1<=C43,0,1)*60000` | 60000 | **different** |
| C45 | candidate: N(VLOOKUP(...)) | `=N(VLOOKUP(Params!$B$1,Loans!A:H,8,0))` | 0 | same |
| C46 | the same IF against the N() cell | `=IF(1<=C45,0,1)*60000` | 60000 | same |
| C47 | candidate: VLOOKUP(...)+0 | `=VLOOKUP(Params!$B$1,Loans!A:H,8,0)+0` | 0 | same |
| C48 | the same IF against the +0 cell | `=IF(1<=C47,0,1)*60000` | 60000 | same |
| C49 | IF(1<=(an empty cell),0,1), directly | `=IF(1<=Params!$B$8,0,1)` | 1 | **different** |
| C50 | IF(1<=N(an empty cell),0,1) | `=IF(1<=N(Params!$B$8),0,1)` | 1 | same |

## What follows from it

Different in Mind:

- **Arithmetic on `""`** does not stop: `IFERROR(""-1 ..., fallback)` returns
  a number, not the fallback (C19, C20). Testing for `""` first gives the same
  in both (C21).
- **An empty cell compared with a number is not 0**: `0 = empty` is false
  (C34), `1 <= empty` is true (C49). The same through a plain reference
  (C32 -> C33) and through a VLOOKUP that lands on an empty cell (C43 -> C44).
  The cell that carries the empty value (C32, C43) is itself *not* reported.
- `N()` or `+0` around the reading makes it 0 in both (C35-C40, C42, C45-C48,
  C50), and so does `IF(cell="",0,cell)`.

The same in Mind, so there is nothing to rewrite:

- SUMIFS over whole columns, with a comparison criterion built from a date
  (`">="&DATE()`), from a number typed as text, with `+0` or `INT()`, with the
  bounds held in cells; SUMPRODUCT over a range (C2-C18).
- `N("")`, `MIN(90,"")`, `""` compared with a number, `x=""`, `ISNUMBER("")`
  (C22-C31).
- `cell=""` on an empty cell (C41), and arithmetic on an empty cell (Horizon:
  `C11 = C3+C7-C15` matched).

## The third difference: a lookup with a "" key, after Prep's titles

Mind reported no difference for the file above (the v2 rewritten by a script).
The same model **through the app** -- Prep with its grid titles, the fixer, Prep
again -- came back with one group: `CALCUL!T192:T812`, "Mind returns NaN", on
the rows after the loan ends, where the key F is `""`:

    T192 = 1+($C$10=1)*IFERROR(VLOOKUP(F192,Criteres!$N$11:$O$33,2,0),IF(F192<50,Criteres!$O$12,Criteres!$O$33))

The formula is identical in both files. Four more uploads found the cause:

- Taking back Prep's 29 title cells on `CALCUL`: still NaN. Taking back the 29
  on `Criteres`: clean. Testing for `""` before the lookup (column T only): clean.
- A sheet `Diag` with each part of the formula in its own cell, in three
  versions of the file (D: as the app made it; E: the two titles around the
  table T looks up taken back -- `Criteres!N8`, `O10`; F: the other 27 taken back):

| Cell | What it shows | Excel | D: the app's output | E: N8, O10 taken back | F: the other 27 taken back |
|---|---|---|---|---|---|
| C2 | the key: CALCUL!F192 ("" in Excel) | (blank) | same | same | same |
| C3 | is the key ""? | TRUE | same | same | same |
| C4 | the lookup with that key | not found | **Mind: text** | same | **Mind: text** |
| C5 | the fallback alone | 1.15 | same | same | same |
| C6 | lookup, else fallback (what T multiplies by 0) | 1.15 | **Mind: text** | same | **Mind: text** |
| C7 | 0 times that | 0 | **Mind: NaN** | same | **Mind: NaN** |
| C8 | the lookup with the number 51 | 1.0942 | same | same | same |
| C9 | the lookup with the number 71 (last row) | 1.15 | same | same | same |
| C10 | the lookup with CALCUL!F35 (51) | 1.0942 | same | same | same |
| C11 | Criteres!O12 | 1.0879 | same | same | same |
| C12 | Criteres!O33 | 1.15 | same | same | same |
| C13 | Criteres!N12 | 50 | same | same | same |
| C14 | is Criteres!O33 a number? | TRUE | same | same | same |
| C15 | is Criteres!N11 (no header) ""? | TRUE | same | same | same |
| C16 | Criteres!O11 (the header) | AT - Âge atteint | same | same | same |
| C17 | the lookup with "" typed in the formula | not found | **Mind: text** | same | **Mind: text** |
| C18 | the lookup with that key, table without its header row | not found | same | same | same |
| C19 | the same key on the other table (Criteres E:F) | not found | **Mind: text** | **Mind: text** | same |
| C20 | CALCUL!T192 | 100% | **Mind: NaN** | same | **Mind: NaN** |
| C21 | CALCUL!T35 | 100% | same | same | same |
| C22 | T192's formula rebuilt here | 1 | **Mind: NaN** | same | **Mind: NaN** |
| C23 | does the key match anything in the first column? | not found | **Mind: differs** | same | **Mind: differs** |
| C24 | is the lookup's result a number? | FALSE | same | same | same |
| C25 | is the lookup's result text? | FALSE | **Mind: differs** | same | **Mind: differs** |

So: **a titled table whose first header cell is empty** (`Criteres!N11`; the
header `AT - Âge atteint` sits in O11 only) makes Mind match a `""` key to
that empty header cell. The lookup returns the header text (C4, C6, C17, C25),
`0 * text` is NaN (C7, C20, C22), and MATCH returns a row (C23). Without the
title, and in Excel, `""` finds nothing. Number keys are not affected (C8-C10).
The same happens on the other table of that sheet (C19, titles `E8`, `F10`).
A small test workbook with the same layout on invented data did not
reproduce it (`Mind_test_workbook_3.xlsx`): what in Mind's grid detection
makes the difference is not known.

Rule **FRM-010** flags an exact VLOOKUP / HLOOKUP / MATCH inside IFERROR whose
key is a cell that holds `""` and whose looked-up column or row has an empty
cell; the Prep action puts the test for `""` in front, as for FRM-008.

## CNHI: a diagnostic sheet on the app's output (2026-10-08)

CNHI converted and ran in Mind, with value differences on nine sheets. A
`Diag` sheet with 57 cells, each one part of the formulas concerned, was
added to the app's output and uploaded:

| Cell | What it shows | Excel | Mind |
|---|---|---|---|
| C2 | RAC Baseline!H15 as it is | 4.322152 | same |
| C3 | its numerator: SUMPRODUCT(row, weights, --(labels=H7)) | 13688209 | same |
| C4 | its denominator: SUMPRODUCT(weights, --(labels=H7)) | 3166989 | same |
| C5 | the same with (labels=H7)*1 | 3166989 | same |
| C6 | the same as one product inside SUMPRODUCT | 3166989 | **differs** |
| C7 | the same with SUMIFS | 3166989 | **0** |
| C8 | SUMPRODUCT of the weights alone | 15774369 | same |
| C9 | how many labels equal H7 (COUNTIF) | 20 | **NaN** |
| C10 | how many labels equal H7 (SUMPRODUCT --) | 20 | **differs** |
| C11 | the label H7 | AF | same |
| C12 | SUMPRODUCT of two rows (no condition) | 3.42E+11 | same |
| C13 | RAC Baseline!E15 as it is | 4.083767 | **differs** |
| C14 | RAC Baseline!I11 (a weight) | 3166989 | same |
| C15 | YEAR of the projection midpoint | 2026 | same |
| C16 | proj_year | 2026 | same |
| C17 | the row offset (year difference x 117) | 0 | same |
| C18 | OFFSET(RAC Baseline!H15, 0, 0) | 4.322152 | **differs** |
| C19 | OFFSET(RAC Baseline!H15, 117, 0) | 3.989483 | same |
| C20 | RAC Baseline!H132 (117 rows below H15) | 3.989483 | same |
| C21 | RAC Calculation!E15 as it is | 4.322152 | **differs** |
| C22 | Baseline!P76 as it is (a column offset) | 0 | same |
| C23 | MATCH of the chosen column | 1 | same |
| C24 | MATCH of the first column | 1 | same |
| C25 | OFFSET(RAC Baseline!E19, 0, P76) | 36.70534 | **differs** |
| C26 | OFFSET(RAC Baseline!E19, 0, 0) | 36.70534 | **differs** |
| C27 | RAC Baseline!E19 as it is | 36.70534 | **differs** |
| C28 | Baseline!O80 as it is | 36.70534 | **differs** |
| C29 | the chosen column label | Combined All | same |
| C30 | RAC Baseline!E12 | Combined All | same |
| C31 | Validation list!D16 (holds 80+) | 80+ | **text** |
| C32 | is it text? | TRUE | same |
| C33 | its length | 3 | same |
| C34 | the text 80+ typed in the formula | 80+ | **text** |
| C35 | INDEX(age_band, 9) | 80+ | **text** |
| C36 | MATCH(80, age limits, 1) | 9 | same |
| C37 | MATCH(79, age limits, 1) | 8 | same |
| C38 | INDEX(age_band, MATCH(80,...,1)) | 80+ | **text** |
| C39 | INDEX(age_band, MATCH(79,...,1)) | 70-79 | same |
| C40 | Age-Gender Mix!J16 (holds 80+) | 80+ | **text** |
| C41 | Age-Gender Mix!J15 (holds 70-79) | 70-79 | same |
| C42 | Exhibits!Z24 (holds 80+) | 80+ | **text** |
| C43 | IRAT data!K6385 as it is | 80+ | **text** |
| C44 | Validation list!D16 & "" | 80+ | **text** |
| C45 | Wave!BB62 & a line break & text | HBP Category
x | same |
| C46 | Wave!BB62 & "/x" | HBP Category/x | same |
| C47 | Wave!P38 as it is | HBP Category
/ Beneficiary - Combined All | **text** |
| C48 | Baseline Cost!G151 as it is | 0 | **differs** |
| C49 | the SUMIFS inside it | 1.75E+10 | same |
| C50 | COUNTIFS on the first criterion | 231 | same |
| C51 | COUNTIFS on the second criterion | 32 | same |
| C52 | the first criterion | Category A | same |
| C53 | the second criterion | Wave 2 | same |
| C54 | SUM of the whole H column | 1.19E+11 | same |
| C55 | the SUMIFS with the criterion typed in | 9.01E+10 | same |
| C56 | COUNTA of the F column | 674 | same |
| C57 | Baseline Cost!G150 as it is | 1.75E+10 | same |

What it settles:

- The model's `SUMPRODUCT(row, weights, --(labels=H$7))` computes the same
  (C2-C5). `COUNTIF` and `SUMIFS` over that 330-column row range do not
  (C9: NaN, C7: 0), nor do single-argument SUMPRODUCT forms (C6, C10) -- the
  model uses none of them.
- The text `80+` is reported as different even typed into a formula (C34),
  while `70-79` is not: not a model matter.
- A line break typed inside a string literal gives another text (C47);
  written as `CHAR(10)` it gives the same (C45) -- rule FRM-011.
- `OFFSET(cell,0,0)` differs from the cell itself (C18 vs C2), and the
  weighted averages `E15`, `E19` on `RAC Baseline` differ while every input
  tested matches: not explained by the parts. Every sheet with numeric
  differences is one where Prep inserted title rows; a run with the two
  row-inserting actions off is the next test.

## PVFP: a lone text cell, and what came after (2026-10-08)

PVFP with its external links frozen (1.8.6) converts and runs in Mind -- with
20,000 deterministic differences, all from one cell: `'LoB 1'!A122 =
"GWP_Retained#"&$G$2` came back as `GWP_Retained#`, Mind's `$G$2` being
empty though it holds `LoB 1`. Prep had moved the label beside it (`E2`,
`Nom Onglet :`) into a grid title, leaving `G2` alone. Test workbook 5:

| Cell | What it tests | Excel | Mind |
|---|---|---|---|
| C2 | a text cell standing alone (Lone!C3 = 'LoB 1'), read by a formula | x#LoB 1 | **empty (x#)** |
| C3 | the same cell, is it empty? | FALSE | **TRUE** |
| C4 | a number cell standing alone (Lone!C6 = 42) | 84 | same |
| C5 | a text cell under a '#' title with no header (Lone!C9 = 'LoB 2', title at C8) | x#LoB 2 | same |
| C6 | a text cell inside a titled 2-column grid (Grid!B4 = 'LoB 3') | x#LoB 3 | same |
| C7 | a label next to a value, no title (Lone!F3 'Name', G3 'LoB 4') | x#LoB 4 | same |
| C8 | a label above a value, no title (Lone!F6 'Name', F7 'LoB 5') | x#LoB 5 | same |
| C9 | a text cell in the header row of a titled grid (Grid!B3 'Name') | x#Name | same |

So Mind imports grids and loose *numbers*, but not a loose text cell. Rule
**STR-008** flags a lone text cell that a formula reads; the Prep action
writes a `#` title above it (a one-cell grid), before the title action runs.
With that one title: 20,000 -> 1,151 differences. The rest, read with Mind's
values: `AVERAGE` over a cell that lands on an empty cell (FRM-009 now covers
every reader, `N()` at the origin); the error cells kept on purpose
(`AB78 = AB77-AB84`, `AB84` holding `""`) computing in Mind (FRM-013,
`IF(AB84="",NA(),...)`); the comparison's own boolean rendering (1 vs True,
no longer counted). With those: 1,151 -> 8 real cells, all from
`IFERROR(SUMPRODUCT(I53:DL53*...),"")` -- a `""` inside a *range* in array
arithmetic, test workbook 6.

Test workbook 6 (Data!C3:H3 = 1..6 with E3 = `IF(1=1,"",0)`, C6:H6 all
numbers, C9:H9 with E9 empty, C12:H12 the weights):

| Cell | What it tests | Excel | Mind |
|---|---|---|---|
| C2 | `IFERROR(SUMPRODUCT(row*weights),"")` on the row holding a `""` cell | "" | **820** |
| C3 | the same on the row of numbers only | 910 | same |
| C4 | the same on the row holding an empty cell | 820 | same |
| C5 | guard `IF(SUMPRODUCT(ISTEXT(row)*1)>0,"",...)` on the `""` row | "" | **820** |
| C8 | `SUMPRODUCT(ISTEXT(row)*1)` on the `""` row | 1 | **0** |
| C10 | `COUNTIF(row,"")` on the `""` row | 1 | same |
| C11 | `COUNTIF(row,"")` on the row with an empty cell | 1 | **0** |
| C13 | `COUNTA(row)-COUNT(row)` on the `""` row | 1 | **0** |
| C15 | guard `IF(COUNTA(row)-COUNT(row)>0,"",...)` on the `""` row | "" | **820** |
| C17 | `SUMPRODUCT(row, weights)`, two arguments, on the `""` row | 820 | same |

So in Mind a `""` cell inside a range is a 0 to every range test, and only
a single-cell test (`E3=""`, test workbook 2) sees it. FRM-008 (1.8.10)
names the `""` cells of the range one by one:
`IF(OR(K53="",L53="",...),"",IFERROR(SUMPRODUCT(I53:DL53*...),""))`.

With that: 68 lines, 59 the boolean rendering and 9 real. Test workbook 7
(Data!C3 = `IF(1=1,"",0)`, Data!C4 = 5; Other!C5 = 7 inside a titled grid,
Other!H2 = `Other` standing alone; Probe!C5 = 99 on the probe's own sheet):

| Cell | What it tests | Excel | Mind |
|---|---|---|---|
| L2 | `IF(ISNUMBER(C3*1),C3*1,"")` on the `""` cell (PVFP, Reporting_General!I23) | "" | **0** |
| L3 | guard `IF(C3="","",IF(ISNUMBER(C3*1),C3*1,""))` | "" | same |
| L4 | the ISNUMBER form on a number | 5 | same |
| L5 | `ISNUMBER(C3*1)` alone | FALSE | **TRUE** |
| L6 | `IF(ISERROR(C3*1),"",C3*1)` | "" | **0** |
| L7 | guard on the ISERROR form | "" | same |
| L8 | `OFFSET(INDIRECT("'"&Data!C7&"'!C5"),,0)`, the sheet name from a cell under a label | 7 | same |
| L9 | the same with the sheet name from a LONE text cell (Other!H2) | 7 | **99** (Probe!C5) |
| L10 | `OFFSET(INDIRECT("'Other'!C5"),,0)` written out | 7 | same |
| L11 | the direct reference `OFFSET('Other'!C5,,0)` | 7 | same |
| L12 | `INDIRECT("''!C5")`, an empty sheet name | #REF! | **99** (the formula's own sheet) |
| L13 | `SUM(INDIRECT(a):INDIRECT(b))` with the lone cell's sheet name | 6 | **0** |
| L14 | `OFFSET('Other'!H4,,0)-0`, an empty cell outside any grid | 0 | same |
| L15 | `'Other'!D5-0`, an empty cell inside a grid | 0 | same |

So (1) ISNUMBER / ISERROR hide the `""` arithmetic like IFERROR does, and
FRM-008 (1.8.11) guards those forms too; (2) a lone text cell that feeds
INDIRECT's sheet name is empty in Mind, and an empty sheet name means the
formula's own sheet there (`Reporting_LoB Total!E32` read its own H31,
36,209, where Excel read `LoB_Total!H31`, empty): STR-008's title above F2
fixes it. A sheet that holds only lone cells has no grid and does not exist
in Mind at all ("Sheet not found" on compile). The PVFP 1.8.11 trial: 59
lines, all the boolean rendering -- no real difference left.

## Palermo and CNHI through the app (2026-10-08)

**CNHI** (title actions unticked, so no title rows inserted): 5,043 lines,
all artefacts -- `80+` read as 80 by the comparison's Excel side (Exhibits,
Age-Gender Mix, IRAT data), five cells of rounding noise on Baseline Cost
row 128. The 1.8.2 numeric differences (RAC Baseline, `--($V$7:$MU$7=H$7)`
TRUE on 178 empty cells) do not appear without the title rows.

**Palermo**: the run stopped on `'Policy summary'!Z6` ("Index was outside
the bounds of the array"). Test workbook 8 (Data!C3 = 0 as the lookup text,
Data!C4 = 0 as the switch):

| Cell | What it tests | Excel | Mind |
|---|---|---|---|
| C2 | `IF(1=1,0,INDEX(INDIRECT(cell holding 0),1,1))` | 0 | same |
| C3..C8 | unreached INDEX(INDIRECT("zzz")), INDEX(NA()), INDEX out of bounds, 1/0, NA(), a VLOOKUP that finds nothing | 0 | same |
| C9, C10 | the same reached and caught by IFERROR | 0 | same |
| C11 | `IF(C4=0,0,INDEX(INDIRECT(C3),1,1))`, the switch from a cell | 0 | same |
| C12..C14 | unreached INDEX over a valid INDIRECT out of bounds, MATCH over INDIRECT(0), the NA() stand-in | 0 | same |
| C15 | `INDEX(INDIRECT("Data!C5:D6"),2,2)` reached | 4 | same |

So Mind does not evaluate the branch an IF does not take, and the run
completed. The cause was the switch itself: `ME_Table` names
`'Policy summary'!C36 = ME_charge_table_ME_lookup`, which names
`'Non-economic assumptions'!D67 = VLOOKUP(Policy_tariff,Assumptions_array,
MATCH(B67,Assumption_headings,0),FALSE)`, and that lookup lands on the
empty cell P7: 0 in Excel, empty in Mind, so `ME_Table=0` is FALSE in Mind
and Z6 computes `INDEX(INDIRECT(<empty>),...)`. The 1.8.9 scan missed it
because every argument is a defined name; 1.8.12 reads them. With `N()`
around 21 such lookups (no computed value changed), Palermo converts and
runs; the five cells left were the `Saved down` file paths fed by the
one-cell array formula `{=LEFT(CELL("filename",A1),...)}`, now frozen by
FRM-012 as well.

## Instances demo: a grid title without '#' (2026-10-22)

**The case.** Three hand-built workbooks to show how instances work:
`1_Central` (an `/InstanceKeys` grid listing Product A, B, C under the header
`2_Product`, plus a data table), `2_Product` (the workbook to copy, its name
from `MM_INSTANCEKEY`) and `3_Summary` (a plain link to `2_Product`, and one
row per copy with `MM_INSTANCE(cell, 1 / 2 / 3)`). Uploaded to Mind as
**testy1** (folder MindReady_Sandbox). Mind's log:

    3_Summary > Summary > D12: Excel value: 1000 Mind value: 0
    3_Summary > Summary > D13: Excel value: 1000 Mind value: 0

**What Mind did** (read with the Mind MCP, read-only): every grid of Central was
"Untitled"; the title text `Instances /InstanceKeys` had become the *header* of
a one-column grid B4:B8, so the model had no flags at all;
`get_specific_grid(Instances)` answered "This model doesnt contains Instances
grid", and `2_Product` had only `Original_Instance`. The title over the data
table split it in two (D4:D8 under the header "Product data", E5:F8 apart).

**Why 0 -- two causes, found one after the other.** First upload: copies 2 and
3 did not exist (the title problem below). Second upload, titles fixed: the
three copies existed and the log was *the same*. The 0 comes from Mind's upload
check, not from the model: after a run in Mind (2026-10-22) the grid showed

    No.  Product    Total premium
    1    Product A  1000
    2    Product B  1200
    3    Product C   800

which is exactly what the model should give.

**Lessons.**

1. A grid title must **start with `#`** and sit directly above the grid's
   top-left cell (`#Instances /InstanceKeys`). Without the `#` the title is read
   as the header row and every `/Flag` in it is lost. For `/InstanceKeys` that
   means no instances at all, and the upload gives no error.
2. Before uploading a workbook built by hand, read it with the app's own grid
   detection (`app/grids.py` `detect_grids`, the same reading the scan uses):
   it gives each grid's name and flags. Here it would have shown no names and
   no flags. After the fix it shows `Instances` with `instancekeys`,
   `Product data`, `This product`, `All products`, `Each product`, and the
   notes as lone text cells.
3. Text cells touching each other form a grid: a sheet title over a subtitle,
   a column of notes. Keep notes in single cells with empty cells around them.
   A header cell that starts with `#` reads like a title; `No.` was used instead.
4. In the upload log, "Mind value: 0" on an `MM_INSTANCE(cell, 2 / 3 ...)` cell
   does **not** prove the instance is missing. After the fix (re-upload
   2026-10-22 11:39) Mind had the three copies Product A, B, C, and the upload
   check still reported D12 and D13 as 0. The plain link C6 was *not* reported
   (1,000 = 1,000, no sum of the copies), and the project had not been run
   ("Project not run - no output available"; every copy still showed Product A,
   1,000). Reading: the upload check compares each cell before any run, without
   the instances' own values -- copy 1 gives the Excel value, a copy number
   above 1 gives 0. So check the model's instance list first (Mind MCP
   `get_project_structure`); the values themselves come only from a run.
   **Confirmed by the run above: this log line is a false alarm of Mind's
   upload check.** Do not "fix" a formula for it.
5. When the instances do exist, cells built on them differ from Excel by
   design: Excel has only one copy. Seen in Mind after the run: D12 1,200 and
   D13 800 (Excel 1,000), C12 / C13 Product B / Product C. Expected but not
   yet looked at: the plain link C6 at 3,000 (Excel 1,000). These are not
   errors.

The demo files are in `PyCodes\MindInstancesDemo` (outside the app).

## In the app

- Rules **FRM-008**, **FRM-009** and **FRM-010** (`app/validators/mind_values.py`), all
  REQUIRED. On the files above they flag exactly what Mind listed: Horizon v2
  -- 1,242 cells for FRM-008 (E192:E812, F192:F812), H9 and C20!C15 for
  FRM-009; Horizon v4 -- nothing; the test workbook -- C19, C20, and C33, C34,
  C44, C49 (through C32 and C43).
- Prep actions **Test for empty text before the arithmetic or lookup** (FRM-008,
  FRM-010) and **Read empty cells as 0 where they are compared** (FRM-009) write
  the rewrites; Excel's values stay as they are.

## Not covered

- A `""` key in a lookup that is not inside IFERROR (in Excel that cell shows
  #N/A -- the recalculation step catches it), or in XLOOKUP / INDEX-MATCH forms
  other than MATCH(key, range, 0).

- `IF(ISERROR(x), ...)` used instead of IFERROR.
- INDEX/MATCH, XLOOKUP, OFFSET or INDIRECT landing on an empty cell: only
  plain references and exact VLOOKUP / HLOOKUP are worked out from the file.
- Text other than `""` in arithmetic; an empty cell compared with another
  empty cell or with TRUE/FALSE -- not tested in Mind.
- A cell that is empty or `""` only in another scenario than the one saved in
  the file: the rules read the values Excel stored.
