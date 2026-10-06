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
