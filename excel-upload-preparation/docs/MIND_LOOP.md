# Mind loop

The macro's run loop written as Mind formulas, from the Model map. `app/loop_builder.py`,
`GET /api/sessions/{id}/mind-loop-plan` (the plan, nothing written),
`POST /api/sessions/{id}/mind-loop-plan/apply` (written as a new version), the **Mind loop**
section of the Model map screen, `python -m app.loop_builder <workbook> [--json]`.

## What is written

| The macro | In Mind |
|---|---|
| copies row *n* of a table into a driver row, *n* from a first cell to a last cell (UC-1) | every driver cell: `=MM_LOOP("Name", 'Table'!Xfirst:Xlast)`, its own column of the table |
| walks a list with `Range("start").Offset(0, i - 1)` (UC-3's scenarios, clusters) | the driver cell: `=MM_LOOP("Name", the stretch of the list)` |
| sets the driver to the counter itself, from a first cell to a last cell (UC-2) | `=MM_LOOP("Name", LastCell)` when the first value is 1; `=MM_LOOP("Name", LastCell, first, first, LastCell)` otherwise (the IndexStart / IndexEnd form; verified in Mind: instances first..last, MM_RESULT index 1 = the first) |
| computes the driver from the counter: `start_year + i - 1` (UC-3's year) | the same expression with `MM_LOOP("Name", CountCell)` in the counter's place; the count cell is named by the user when the workbook has none |
| writes a source block into `Target!A<n>:J<n>` on every pass (UC-1) | row by row: `=MM_RESULT('Source'!$AY$35, "Name", k)` for loop value k |
| clears a results area before a run | listed with what an earlier run left there; emptied only when asked |

The loop name is the driver's defined name (e.g. `RUN_ROW`, `ACTIVE_SCENARIO`), else the
source sheet's name (e.g. `Policies`); the bounds are the bound cells' numbers (a formula's
stored value counts). All three can be changed before writing.

## One macro at a time (review, 10 November 2026)

A workbook can have two run macros, each with its own loops (the IFRS 17 tool: `AllScenarioRun` over
cohort x scenario, `EndemicCalculationRun` over group x roll-forward year). Excel runs each macro on
its own, with the other macro's drivers fixed at their saved values. Written into one model, Mind
would run all four loops inside each other (6 x 4 x 6 x 6 passes) and `MM_RESULT` with one loop would
add up the passes of the others: the values would differ from Excel.

- The plan builds the loops of one macro (`macro=`, default: the first macro with a loop; the Mind
  loop panel shows a **Macro** dropdown when there are two or more). The plan lists the macros and
  their loops (`macros`). The other macros' loops are listed as `not built`, with why.
- A driver of another macro that already holds an `MM_LOOP` (written earlier) is named in the notes:
  build on a version without it.
- A macro's outputs and cleared areas are its own: another macro's block is not written.
- Inside one macro, the loops are told apart by procedure and loop variable, so two procedures that
  both use `i` keep their own loop.
- A count cell that already holds the count this plan writes (an earlier write with the same choices)
  is accepted. The label `<loop> count` goes to the left of the count only when that cell is free:
  empty, not another loop's count or label, not a cell the plan writes, and read by no formula.
- A first value typed as a cell (`Sheet!A1`) counts as chosen, like a number: the header pass is kept.

## What is not written, and said

- a block the macro places by a counter of its own (UC-2's `.Offset(0,
  10 * (tariff - 1))`): in Mind the results stay in the source block, one value per loop
  value; `=MM_RESULT(cell, "Name", k)` reads one;
- a block that depends on several nested loops (UC-3's paste blocks per scenario x year x
  cluster): `MM_RESULT` with one loop would sum over the others; read the source block
  with every loop named;
- a counted loop with no cell holding the count (UC-3's `For i_year = 1 To 5`): name a
  free cell and the plan writes the count there;
- a values copy made once after the loop (UC-2's pasted summary sheet): nothing to
  write, the sheet it copies from holds the results.

## The four client models

- **UC-1**: a loop named after the policy table, 6 values (rows 1 to 6 of the table; row 1 is the
  header row, the macro runs it too -- set first to 2 to leave it out); 42 `MM_LOOP` cells
  on the calculation row; the output block (6 rows x 10 columns) from the calculation sheet's
  result block, 60 `MM_RESULT` cells. Checked in Excel against the macro: 60 of 60 cells the same.
- **UC-2**: the policy-run cell, 1 to 19 from the count on the macro sheet; the per-policy
  block is placed by the tariff (not built), the summary paste is a values copy.
- **UC-3**: the scenario over a row of the inputs sheet, the cluster over a list; the year
  needs a count cell; the 30 paste blocks depend on all three loops (not built).
- **UC-4**: no run loop (one sheet per line of business): nothing to build.

## Checking it

    python scripts/check_loop_plan.py <workbook.xlsm> --macro <Name>

runs the macro in Excel on a copy (macros allowed for that copy only), writes the plan into
another copy through the app's own Apply, plays every loop value into the driver cells and
compares the source cells of every `MM_RESULT` block with the macro's block. MM for Excel
(loaded into the app's Excel since 1.10.1) evaluates one pass only -- `MM_LOOP` gives the
first loop value, `MM_DIMSIZE` says 1 -- so the check plays the values itself: it proves
the wiring (which column, which row, which loop value); the run in Mind proves Mind.

## Verified in real Mind (2026-10-15)

UC-1, the Do-everything version (Fix all issues, Fix the formulas, the loop over rows 2 to 6,
recalculated clean), in a sandbox project:

- Upload / Convert / Test complete in about 90 s; the Test step counts Stochastic 0, Deterministic
  22,212 cells with 0 differences, Assumptions 163,741. The Deterministic test evaluates ONE loop
  instance (the workbook's cached values are the first pass, so they agree), so the Deterministic
  Differences log never shows the other passes: the loop is checked after the run, not there.
- Add template, Run: completed in 1.8 s, consistent with the audit trail.
- How to see a loop in Mind: the template selector in the project bar (the workbook's name) lists the
  sheets; a sheet lists its grids `Untitled(row,col)`; a grid opens a panel. The panel's play icon
  (top left, "Partial run") recomputes that grid; the panel then shows the dimension name
  (the loop's name), a "Display all dimensions" control and one row per loop instance.
- Every cell of the output block matched the macro run in Excel on the same version: 10 columns x 5 passes, pass k =
  loop value k + 1. Mind's own AI Assistant, asked in the project, answered "Defined loop <name>,
  size 5 ... read by 50 MM_RESULT cells".
- Mind reads a grid's first row as its header. Where the row above the block had empty cells, Mind
  split the block into four grids and showed pass 1 as the header of two of them (values right,
  display wrong). The loop builder now completes that header row from the source block's own
  captions (`_complete_header`).
- The loop written on the untouched original is refused by Mind: 15 formula compilation errors, all
  `SUMIFS('<policy table>'!#REF!, ...)` on a check sheet, the broken references the app's REF-001 fix
  repairs. The loop needs the fixed workbook, which is what the main screen's buttons produce.

Scripts: `scripts/mind_loop_probe.py` (upload, logs, run, parts), `scripts/mind_loop_compare.py`
(Mind's rows vs the macro's values), `scripts/mind_ask.py` (Mind's AI Assistant),
`scripts/night_chain.py` (the Do-everything chain through the API).

## The two MM_LOOP forms, verified in Mind (2026-10-15, a tiny probe workbook)

`scripts/build_loop_forms_probe.py` writes a workbook with `K = MM_LOOP("K", last, 3, first, last)` (first =
3, last = 7) and `J = MM_LOOP("J", last)`, plus `MM_RESULT(cell, "K", k)` / `("J", k)` for k = 1..7. In Mind:

- `MM_DIMSIZE("K")` = 5 and `MM_DIMSIZE("J")` = 7; the first instance of K is 3, of J is 1.
- `MM_RESULT(x10, "K", k)` = 30, 40, 50, 60, 70 for k = 1..5 (k = 1 is loop value 3) and 0 for k = 6, 7;
  `MM_RESULT(x10, "J", k)` = 10..70. So the builder's `k = value - first + 1` is what Mind expects, and an
  index past the loop gives 0, not an error.
- `MM_RESULT` needs a cell as its first argument: with an expression (`$B$4*10`) Mind's test step fails at
  runtime with "trying to access null reference". The builder always passes a cell.

## The ranges follow the input table (1.11.1)

A table loop is written as `'Table'!X<first>:INDEX('Table'!X:X, MM_LASTROW('Table'!<key><first>))` in every
driver cell, the key column being the first column of the table filled on every loop row. Mind reads
`MM_LASTROW` from the table's grid, so a row added to the table (Input Manager, a new upload) joins the
loop; Excel's add-in stops at the key column's first empty cell, so the local values are today's rows.
Verified in Mind on UC-1 (2026-10-15): the dimension spinner on an output grid runs 1..5, and 1..6 after
a seventh policy row was appended through `scripts/add_policy_row.py`; Mind's test counts 0 deterministic
differences; the Excel check gives 50 of 50 cells against the macro.

Why not `MM_COLUMN(X2)` alone: in Excel it stops at the first empty cell of X, and when a column has an
empty cell inside the loop rows (UC-1's columns Q, Z, AA, AD) it collapses to one cell, which `MM_LOOP` then
reads as a count: wrong local values and an error text, and 895 deterministic differences in Mind's test
(Mind itself computed the right values). A list the macro walks to its last item is written with
`MM_COLUMN` / `MM_ROW` from its first item (lists have no gaps).

A column of the table that holds nothing in the loop rows gets no loop, and a loop written there by an
earlier build is cleared (it would be a definition with another length once the table grows). Every
definition of a loop must keep the same length, so a new row should fill every part of the table (Mind cuts
a table into grids at empty columns). `--fixed-ranges` on the command line, `fixed_ranges` on the API and
the box on the Advanced loop panel give the old `X2:X6` form. LOOP-003 treats table-following definitions
as equal by construction and flags a mix with fixed ones.

## Defaults and notes (1.11.1, after the design review)

- When the macro starts on the table's text header row, the loop starts one row (column) further: in Mind
  the header would be instance 1. Choosing first = the header row keeps the macro's header pass.
- A loop cut short (last before the end of the table) keeps fixed ranges: rows added later do not join it.
- Every table-following loop says that Excel may differ from Mind (the add-in stops at the key column's
  first empty cell), which growth reaches it (a new upload always; Mind's Input Manager only imports into a
  grid titled with /Input), and that the result block is written for the loop values present now: Mind
  computes every instance, the block lists those of today.

## Checking the loop through Mind's MCP server (optional, 1.11.2)

With a Mind MCP token (env `MIND_MCP_TOKEN`, or `%LOCALAPPDATA%\MindReady\mind_mcp_token`; issued per user by
the Mind team, 24 h) the app reads a project as data (`app/mind_mcp.py`):

    python scripts/mind_mcp_loop_check.py <project link> --source "Calc!AY35:BH35" --check check.json --expect-size 6

starts a classic run (the browser runs leave no saved results), reports it (result, consistency, and on a
failure the log's error lines with every Mind cell mapped to its Excel address), reads the loop's size and
every instance of the source block, and compares instance k with the macro's pass k. A cell the loop does
not reach (a constant) holds one value, which counts for every instance. `scripts/mind_loop_probe.py` does
the upload with the browser and then this check through MCP when a token is there, and falls back to the
browser (`--browser` forces it) when there is none or it expired. Upload, convert, add template and delete
have no MCP tool: they stay in the browser. Verified 2026-10-15 on UC-1 with a seventh policy row: run
Success / Consistent in 6.5 s, a loop of size 6, 50 of 50 cells the same as the macro.

## The loop goes to Mind with Prep's formula repairs (1.12.5)

A loop written on a version Prep has not repaired yet would go to Mind with formulas Mind computes
otherwise. On the Tokyo tool (2026-11-09) that made every formula of six calculation sheets NaN: their
B3 reads the sheet name through `CELL("filename")`, which Mind answers with the workbook's name alone
(see MIND_VS_EXCEL.md). The loop plan therefore takes Prep's plan for the same version and writes the
formula repairs it still has with the loop (`loop_builder.FORMULA_REPAIRS`: the main screen's "Fix the
formulas" set, among them FRM-012's CELL / INFO freeze and FRM-014's OFFSET ranges). A cell the loop
writes keeps the loop. The plan lists them (`repairs`, a note, a line on the loop panel);
`repairs: false` on the API leaves them out and the note says Mind will compute other values. In Do
everything the Prep step runs first, so nothing is left for the loop to add.

**What the loop does not write, and says (10 November 2026).** Prep's blocking items that are not
formulas -- a lone text cell to title (STR-008), a pivot table (STR-009), a grid title trapped inside a
grid (STR-001), an array formula past its grid (RSK-002) -- are structural: the loop and the instances
leave them to Prep. A plan written while Prep still has such items carries one line, "Prep still has N
blocking items that are not formulas, e.g. a formula reads a lone text cell (STR-008). Run Fix all
issues before uploading." (`prep_blockers` on the plan: id, title, count, rule ids;
`loop_builder.prep_blockers`). Every ready plan also carries one line on the two kinds of cell Mind's
upload test lists after a build and that are not errors (rounding residues on check cells, the text
TRUE / FALSE): see MIND_VS_EXCEL.md, "Mind's upload test after a loop or instances build".

Proven in real Mind (2026-11-09), Mind's test after the upload:
- the loop written on the unrepaired version through the app (6 CELL freezes and 84 OFFSET ranges
  written with it; zzz_mindready_tokyo_loops_probe_7): 14,737 differences before, 75 after, none a value:
  67 check cells that are 0 in Excel and a rounding residue below 0.00002 in Mind (B47-B46 of two equal
  totals near a billion), and 8 cells holding the text "TRUE" that the test reads as 1 on Excel's side;
- the Do-everything version plus the four loops through the app (84 OFFSET ranges written with it;
  zzz_mindready_tokyo_loops_2): 63 differences, the same two kinds (55 residues below 0.000004, 8 "TRUE").

## The macro's text in its own code page (10 November 2026)

The VBE stores module text in the ANSI code page of the PC that saved the workbook (cp932 on a
Japanese PC, cp1255 on a Hebrew one), and Excel 365 records 0 ("the system's") as the project's code
page. olevba answers 0 with UTF-8 and replacement characters, so the Japanese comments and string
literals of the Tokyo tool -- `Find("年月")`, `MsgBox "完了しました"` -- arrived as mojibake, and the
stand-in assistant and the message-box matching could not read them. `model_map.read_vba` now decodes
the module bytes itself (`decode_vba_source`): the stored code page when there is one; else UTF-8,
cp932 when a quarter of what it gives is kana, cp949 when it is mostly Hangul, the PC's own code
page, cp936 / cp950 when they give ideographs, cp1252 -- each only when it decodes every byte, cp1252
with replacement characters as the last resort. The loops are found as before (the fixture
`tests/unit/fixtures/vba_japanese_cp932.xlsm`: a row-per-record model whose module is cp932 with code
page 0, as a workbook from a Japanese PC has). The limit: a Western file on a Hebrew PC with no code
page recorded reads its accented letters as Hebrew ones (no single-byte code page can be told from
another); Chinese is read right from the stored or the PC's own code page.

## A driver cell must not have a dropdown (1.11.3)

Mind reads a cell with a list dropdown (data validation) as an input and drops its formula: an MM_LOOP
written there is lost, and the run has no loop. Seen on UC-3 (both drivers have a dropdown; Mind made
neither loop) and confirmed with a probe workbook on 2026-10-16: the same MM_LOOP in four driver cells,
plain, with a defined name, with a dropdown, with both; Mind made loops of the first two only, and of all
four once the dropdowns were removed. The loop builder therefore adds a `remove_validation` operation for
a driver cell with a dropdown (the note says why). The defined name on a driver is harmless.
