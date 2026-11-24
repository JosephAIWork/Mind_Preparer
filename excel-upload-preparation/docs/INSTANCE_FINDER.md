# Where instances could help: the instance finder

The finder reads a workbook's formulas and its macro and points to where Mind instances, or a Mind
loop, would remove repetition or compute every choice in one run. It reports and ranks; it writes
nothing. Added in 1.12.5. `app/instance_finder.py` (with `app/finder_shapes.py`,
`app/finder_patterns.py` and `app/finder_grouping.py`), `GET /api/sessions/{id}/instance-finder`
(`?assistant=1` adds the assistant's explanation, `?refresh=1` computes the result again for the current
version), the fold **Not sure?
Let the app look** on the Instances tab under Advanced (since 1.13.0; before that the section **Where
instances could help** below the builder), and `python -m app.instance_finder <workbook> [--json]
[--assistant]`.

The finder is blind: it gets no hint of what the key is -- no column name, no word such as tariff or
product. It is deterministic: pure Python over the scan the app already made, no Excel, no numpy, no
model call; the same workbook gives the same answer. The assistant pass on top is optional and never
changes the findings.

## What it finds

Four patterns. Each one found is a *dimension*: something the model repeats over.

- **A grouping column of the looped table (P1).** The macro's loop steps through the rows of a
  table; one of its columns splits the rows into groups. One copy of the model per value is what the
  instances builder makes (`docs/MIND_INSTANCES.md`). The finder names the column and says why.
- **Repeated formula blocks (P2).** The same formulas side by side in one sheet -- bands of columns
  or of rows -- each computing the same thing from a different input. One block and a loop, or one
  copy per block.
- **Repeated sheets (P3).** Sheets with the same layout and the same formulas at the same addresses.
  One sheet and instances.
- **Near copies (P4).** Copies that differ only in typed-in numbers, in a number inside a formula
  (`=C7*0.1` against `=C7*0.12`) or in a few cells of different logic, with the edits that would
  unify them.

Also reported, so that nothing is hidden:

- **Chains, set aside.** Copies where each one reads the one before it (year sheets, period blocks).
  Neither a loop nor instances expresses a sequence; the finder lists them with the reason and gives
  them no rank.
- **Time axes, ignored.** A filled column or row that reads the previous one (a projection) is a time
  axis, never a dimension. Counted in `ignored.time_axes`.
- **Filled areas that join nothing.** One formula filled down or across is a fill. A fill alone is
  never a dimension; it becomes part of one only when it steps through something a finding points at.
  The rest are counted (`ignored.filled_areas`, `ignored.small_fills`) and the biggest ones listed.
- **The macro loops the Mind loop step already handles** (`macro_loops`, from the loop plan). Report
  only.
- **Switch cells** -- a typed-in cell many lookups use as their key -- from the model map, each with
  `readers`: the formulas that write its defined name, or its address whole (`B3`, `$B$3`, `B$3`; never
  `B30` or `AB3`) on its own sheet or with the sheet in front from any sheet. Report only, never a
  dimension on their own.

## How it decides

### Evidence for a grouping column, in order of weight

The table is the one the macro's counter loop indexes (`INDEX` / `OFFSET` with the loop's driver as
the row), taken from the loop plan and widened by the filled columns touching it on the left. Every
column is swept once and whole, using the values a formula cell was saved with, no sampling: distinct
values, rows per value, the biggest group. A candidate is a text column filled in at least 80 % of
the rows with 2 to 200 distinct values -- the builder's own rule -- or any column the evidence below
points at, even a column of numbers.

- **(a) The macro groups its output by the column.** The strongest evidence. The finder reads the
  macro's text (comments cut, continuation lines joined) and follows the chain: a `VLookup` of the
  loop row in column *k* of the table, a `Match` of that value in a list, and a paste statement that
  puts each pass's block into slot *n* of an anchor with `Offset(0, w * (n - 1))` (or along the
  rows). The `Offset` is parsed with balanced parentheses, so `10 * (k - 1)` gives the width
  10. Any other `VLookup` of the matched value is a table keyed by the value.
- **(b) Lookups keyed by the column's current-row value.** The current-row cells are the `INDEX` /
  `OFFSET` hits of the loop. Their aliases (`=<cell>`, followed to a depth of 3) count too. A cell
  whose whole formula maps a current-row cell through a table (`VLOOKUP(key, table, k, 0)`,
  `XLOOKUP`, `INDEX(col, MATCH(key, col, 0))`) is a *mapped key*, one level deep. Every formula
  whose key argument -- the first argument of `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `MATCH`, `LOOKUP`,
  `CHOOSE`, `SWITCH` or `INDIRECT`, the row and column of `INDEX` and `OFFSET`, the criteria of
  `SUMIF`, `SUMIFS` and `COUNTIFS` -- names one of these cells, by address or by defined name, is a
  keyed lookup; the tables it reads are "one set of assumptions per value".
- **(b') Formulas that branch on the value.** `IF(<key> = "text", ...)`, also inside `IFS`, `OR`,
  `AND`, `SWITCH` and `CHOOSE`, with `<key>` a current-row cell, an alias or a mapped key. Reported
  from 5 formulas up, with how many of the compared texts are values of the column (or of the mapping
  table's output column), and the edit "move the values chosen by these IFs into a table with one row
  per copy".
- **(c) Counts.** Distinct values, rows, the biggest group and its share, and the balance (1 means
  every group has the same size).
- **(d) Header words.** A header that says tariff, product, group and the like is shown as a hint
  and weighs nothing. Labels are never proof.

Evidence (a), (b) and (b') is structural: what the macro and the formulas read. The builder's own
key rule is run on the same sweep and reported beside the finder's choice (`builder`, with `source:
"rule"`). In the app the route puts the builder's real plan there instead (`source: "plan"`) and says
whether both pick the same column (`matches_top`).

### Repeated blocks, sheets and near copies

Every formula is turned into its R1C1 form (`app/finder_shapes.py`), so a formula filled down or
copied across is the same text everywhere; two coarser forms mask the references, then the typed-in
numbers. Bands: for every period from 2 to 40, positions whose formulas match the ones a period away
in at least 90 % of the cells form a run; a run whose first block holds at least 10 formulas is a band
of *k* copies. Sheet copies: two sheets are copies when, over the cells either of them holds, at
least 75 % hold the same formula at the same address (the smallest similarity inside a group is
reported, so a near miss is visible). A uniform area -- every row equal and every column equal -- is
one fill with two axes, never two findings.

Where copies differ, every cell is sorted into: identical / reads a different input (the loop index)
/ differs only in a typed-in number (a table row per copy) / different logic. The edits list the
constant cells and the formula numbers that differ, with their values per copy, and the cells of
different logic with the rule that fits: an `IF` on the copy's key when the cell exists in every copy
with at most 3 forms, else an override kept per copy. Cells present in some copies only mean copies
of different size: instances, with `MM_LOOPINSTANCE` for the loop inside each copy (Mind's
documentation: a named loop sized per instance; `docs/MIND_INSTANCES.md`). A plain loop would have
to run the biggest size and pad the shorter copies.

A chain is tested on what the copies read: a copy is a chain member when a relative reference of its
cells lands in an earlier copy. A fixed cell every copy reads -- an absolute reference, a whole row or
column -- is a parameter, not a chain. Sheets are a chain when every sheet after the first reads the
one before it; scattered reads between sheets are `cross_reads`, an override need.

### Merged into dimensions by the input they step through

Findings are never merged on labels: a header is only what the author typed. Each finding gives
*anchors* -- a table column, a list range, an area cut into slots of width *w*, a stepped-through
area, a band, a fill, a group of sheets -- and *edges* join them; one connected group is one
dimension.

- **The macro chain** joins the column, the list it matches the value into, the slot area it pastes
  into (widened to the area the macro clears before a run) and every table keyed by the value.
- **Stepping.** Blocks whose copies read consecutive rows, or consecutive column slots, of one area
  step through that area. Blocks whose cells are identical in R1C1 and step by a *computed index* --
  `OFFSET(anchor, ROW()-n, (ROW()-m)*w)`, the same with `COLUMN()`, `w*(COUNTA($X$1:$X1)-1)` over an
  expanding range, or `w*(MATCH(key, list, 0)-1)` -- are read from the formula itself: the anchor is
  resolved through the defined names, the width and the axis come from the index. The `MATCH` form is
  not stepping but a keyed read (below); it joins the list and, when the key is a current-row cell,
  the column.
- **Equality.** Two findings with the same anchor merge. Two lists merge when they hold the same set
  of texts; a list merges with a key column when its texts equal, or cover at least 90 % of, the
  column's distinct values (a list may hold keys without rows). Slot anchors that lie inside the same
  cleared area with the same width merge; so do two slot anchors of one sheet with the same width
  whose areas overlap.

The dimension's name is the column's header, else the list's defined name, else a header hint, else
the address -- a label only; `label_from` says which.

### Cells that read the macro's paste area, by class

The area the macro clears before a run, cut into slots of width *w*. Every formula is tested with its
defined names expanded and the anchors of its `OFFSET` / `INDEX` calls resolved, so a cell that
reaches the area through a name is seen. Each class needs a different thing in Mind:

| Class | What the cells do | In Mind |
|---|---|---|
| `all_slots` (the adders) | add every slot up | In a copy they see only that copy's own area, which is empty after the builder. Move the adding row to the main workbook as plain links to the copied model: there a link is the sum over the copies (Mind's documentation; checked in a real Mind project on 29 October 2026, `docs/MIND_INSTANCES.md`). Or read the main workbook's All copies grid. Counted in `cells_changed`. |
| `fixed_slot` | read one slot, or several but not all | One slot: one copy's value, `MM_INSTANCE(cell, key)` in the main workbook. Several: `cross_reads`, one `MM_INSTANCE` per slot. A formula with several `MM_INSTANCE` calls returns n times one value: the team's own observation in its Mind projects of 22-29 October 2026, not a sentence of Mind's documentation. |
| `computed_slot` | one row (or column) per slot, by `ROW()`, `COLUMN()` or `COUNTA` | One row per copy in the main workbook: a plain link per copy, or `MM_INSTANCE`. These are the `formulas_saved` of a grouping column. |
| `keyed_slot` | read the current key's slot and add the pass to it (`MATCH` on the key) | Each pass is added to the key's own slot. In a copy it is the copy's own block, and `MM_RESULT` does the adding. Nothing to write; reported, never counted as saved. By the rule of (b) these cells are keyed lookups too; the lookups line says how many of its count are this block, so nothing is counted twice in silence. |
| `follow_adders` | read the adders' own cells | They move with the adders. Text only. |

### Score, rank and fit

Scores are integers. One strong structural source beats any count.

| Evidence | Points |
|---|---|
| (a) the macro groups its output by the column | 50 |
| (b) keyed lookups | 10, plus up to 20 growing with the count |
| (b') branch formulas | 5, plus up to 10 growing with the count |
| (c) counts | 0 with fewer than 2 or more than 200 values; else 5, plus up to 10 for balance (halved outside 3 to 100 values) |
| (d) a header word | 0 |
| a column with two values only | -10 |
| a block or sheet dimension | up to 30 for the formulas saved; +15 when not a chain; +10 when no cell differs in logic; +5 when numbers differ (a table row per copy is the natural unification) |
| a fill alone | no score, never a dimension |
| each extra kind of evidence the merge joins (stepping, keyed, list, bands, sheets) | +20, at most +40 |

Ties go to the dimension that saves more formulas, then to the label. Chains have no rank; they are
a separate list with the reason.

Per dimension: `copies` (the distinct values, or the *k* copies, or the list's length when only a
list is known); `biggest_copy` (the rows of the biggest group -- Mind's run time is about the biggest
copy's passes -- or the formulas per copy); `formulas_saved` (`(copies - 1) x per copy` for blocks and
sheets, plus the computed-slot readers that become one row per copy); `cells_changed` (numbers and
formula numbers to move, logic cells, adders, fixed-slot readers, and the branch formulas when that
edit is taken).

Fit, for each of the three options, is `yes`, `partly` or `no`, with the reasons:

- **Instances** fit when a key column or a list exists, there are at most 200 copies and no chain.
  Partly with `cross_reads` (one `MM_INSTANCE(cell, key)` per formula in the main workbook) or above
  50 copies: many copies make a run slower, and Mind's documentation suggests `/InstanceSelect`
  groups for models with lots of instances. (That instances use more memory and processor than
  loops is the team's own note from its Mind projects of 22-29 October 2026, not a sentence of Mind's
  documentation.) When no column or list to split on exists, instances fit partly: the finder asks for
  a table with one row per copy, its first column the key. Chains carry no fit: they are set aside
  before the fit is computed. A column with two values: partly, "an IF or a loop of 2 does the same".
  Adders add the reason "N adding cells move to the main workbook" -- never "already the sum".
- **A loop** fits a block or sheet dimension with no logic cells and no size difference: one block
  and `MM_LOOP` over the copies; the cells that read a different input read the loop index instead.
  Partly with logic cells (each needs `IF(MM_DIMINDEX(...))`) or with sizes that differ (instances
  with `MM_LOOPINSTANCE` inside fit better). For a grouping column the loop over every row is the
  Mind loop step itself: one loop of all the rows; instances cut it to the biggest copy's passes.
- **Leave as is** fits when there is little to gain (fewer than 200 formulas saved); else it fits
  partly, with the formulas kept to maintain as the reason.

Thresholds: sheets alike at 75 %; bands at 90 %; at least 10 formulas in a block; periods up to 40; a
chain from 10 % of the sampled reads; a key column with 2 to 200 distinct values filled at 80 %; at
most 200 copies for instances; at least 5 branch formulas; a fill of at least 500 cells before it is
handed to the merge; fewer than 200 formulas saved is "little to gain".

## The blind result on the 2c use case

The check: the model of the 2c use case -- the macro runs one policy at a time and adds each policy's
results to its group's block -- is uploaded to a fresh server and scanned; then the builder's plan
(`GET .../mind-instances-plan`, no key given) and the finder (`GET .../instance-finder`) are read from
the same session, so that the analysed copy is still in the app's two-workbook cache
(`counts.reload_seconds` must be 0; if not, the table says so). The table is filled by the check's
script from the two JSON answers, never by hand, and holds no key value and no name from the workbook.

"Same?" is yes when the finder's top grouping column has the builder's column and header, the builder
chose it from the macro and the finder's kinds hold `macro`; `builder.matches_top` with `source:
"plan"` says it in the JSON too. A line the builder leaves as it is, or does not see, is marked
"finder only". When the top proposal is another column of the table, both are built and run in Mind
and their result grids go side by side here; a block or sheet dimension, which the builder cannot
build, is reported with its edits list.

| Fact | Finder (blind) | Builder | Same? |
|---|---|---|---|
| file checked (`counts.formulas`, `counts.sheets`) | 57,537 formulas, 33 sheets | the same uploaded version | -- |
| date of the check | 29 October 2026 | -- | -- |
| time taken (`counts.seconds`; `counts.reload_seconds`) | 27.9 s; reload 0.0 s (the scan before it: 245.1 s) | the plan's own time: 49.2 s | -- |
| the table it splits (`key.sheet`, `key.table`) | the model-point sheet, `A5:AQ94936` | `plan.table.sheet`, `plan.table.ref`: the same sheet, `A5:AQ94936` (header row 4) | yes |
| key column and header (`key.column`, `key.header`) | column C, header the word tariff | `plan.key.column`, `plan.key.header`, `plan.key.chosen_by`: column C, the same header, macro | yes |
| evidence (`kinds`: macro, lookups, branches, stepping, keyed, list) | macro, lookups, branches, stepping, keyed, list | `chosen_by` is all the builder states: macro | yes: the builder's reason is the finder's first kind |
| distinct values and rows (`key.distinct`, `key.rows`) | 26 values, 94,932 rows | `plan.key.distinct`, `plan.table.rows`: 26 values, 94,932 rows | yes |
| biggest group (`key.biggest.rows`) | 39,477 rows | the largest of `plan.key.counts`: 39,477 rows | yes |
| other candidates, in order (`dimensions[1..].key.column`, with the reason) | column E (68 values; 601 lookups); a 2-block band on an assumptions sheet (10 formulas per block); column W (2 values; 830 branch formulas); column T (2 values; 1,500 branch formulas); column J (2 values; 2,096 branch formulas) | `plan.key.candidates`, sampled from the first 5,000 rows: column C (26 values), column J (2 values), column T (2 values), column W (2 values) | the builder's other text candidates are all among the finder's |
| lookups keyed by the column, also through the mapping table (`key.keyed_formulas`) | 1,448 (of which 1,280 are the keyed block below) | -- (the builder does not look) | finder only |
| formulas that branch on the key (`key.branch_formulas`) | 4,877, most of them on the key mapped through a second table | -- (left as they are) | finder only |
| paste area: cells adding every slot (`readers.all_slots.count`) | 1,280: move to the main workbook as plain links | listed among the cells left as they are | finder only |
| paste area: rows reading a computed slot (`readers.computed_slot.count`) | 3,562 cells in two blocks of 26 rows (`COUNTA` and `(ROW()-n)*10`): one row per copy in the main workbook | not seen: they reach the area through defined names | finder only |
| paste area: the keyed block (`readers.keyed_slot.count`) | 1,280: the block itself; nothing to write | counted as a reader, left as it is | finder only |
| paste area: fixed-slot readers (`readers.fixed_slot.count`, `cross_reads`) | 0, 0 cross reads | in the note | finder only |
| cells reading the adding block (`readers.follow_adders.count`) | 1,280: they move with the adders | -- | finder only |
| the list the macro matches into (`evidence[kind=list].count`, and the list equal to it by value) | 26 values; a second list on another sheet holds the same values | `plan.main.keys`: 26 keys (`keys=present`; every value of the list has rows) | yes |
| the builder's agreement in the JSON (`builder.matches_top`, `builder.source`) | `True`, `plan` | -- | -- |
| what else (block and sheet dimensions, chains with their count, filled areas; `counts.dimensions`, `counts.chains`, `counts.time_axes`) | 6 dimensions in all: 5 columns of the table and 1 block dimension; 1 chain set aside (3 copies of 1,500 formulas, each reading the one before); 1 filled area that joins nothing; 12 time axes | -- | finder only |
| the assistant pass (`assistant.available`, its order, its extra candidates) | available; its order 1, 2, 3, 4, 5, 6 (the finder's rank numbers); 1 extra candidate (flagged as not checked); 14.2 s; it put the finder's #1 first | -- | -- |

The other two files, checked the same way on the same day, give the same top dimension:

- **The 19-row cut (the file whose instances zip was proven in Mind on 29 October 2026)** (57,537 formulas; the scan 149.2 s): top dimension column C with macro, lookups, branches, stepping, keyed, list; 3 values over 19 rows, biggest group 10 rows; readers 1,280 / 0 / 3,562 / 1,280 / 1,280 (adding / fixed slot / computed slot / keyed block / following the adders); the builder's plan: column C by macro; `builder.matches_top` True; the finder 31.2 s, reload 0.0 s, the plan 25.0 s; the assistant put the finder's #1 first.
- **The raw file as the team sent it, before any fix** (57,536 formulas; the scan 162.9 s): top dimension column C with macro, lookups, branches, stepping, keyed, list; 26 values over 94,932 rows, biggest group 39,477 rows; readers 1,280 / 0 / 3,562 / 1,280 / 1,280 (adding / fixed slot / computed slot / keyed block / following the adders); the builder's plan: column C by macro; `builder.matches_top` True; the finder 23.7 s, reload 0.0 s, the plan 51.4 s; the assistant put the finder's #1 first.

## The assistant pass

Off by default. The tick **Explain and rank with the assistant** shows only when an AI key is saved
(`?assistant=1` on the route). One request to the gateway, 90 seconds at most, the pattern of the
auto-fix advisor. What goes out, at most 12,000 characters, is the finder's own result and nothing
else: for each dimension (up to 8) the headline, the kinds, the counts (copies, biggest copy, formulas
saved, cells changed), the evidence texts with their example cells and the macro's lines (up to 6),
the reader-class texts, the key's top 12 values with their row counts, up to 3 blocks (sheet, where,
copies, formulas per copy, the index), the edit counts with up to 3 cells of different logic and their
formulas per copy, the branch count and the three fit lines with their reasons; then the chains, one
line each, and the builder's choice. No sample formulas, no labels near cells, no list values. The
key's values do go to the gateway, as the app's other assistant features already send cell texts.

The answer asked for is JSON only: `{"ranking": [...], "explanations": {"1": "..."}, "edits": {"2":
["..."]}, "extra": [{"headline", "why", "cells"}]}` -- the assistant's order, two plain sentences per
dimension for a non-programmer, the edits that would unify near copies, and candidates the tool did
not find. The first balanced `{...}` of the reply is parsed; the rest is checked: a ranking names
every dimension once or is dropped, texts are clipped, `extra` is capped at 5 and every item is
flagged "suggested by the assistant, not checked".

The deterministic result is never changed: `dimensions`, `rank`, `score`, `evidence` and `fit` stay
as computed. The assistant's order shows as `assistant_rank` on each card, its sentences as
`assistant_text`, its edits as `assistant_edits`, the whole in the `assistant` block. Any failure is
the deterministic result with a message, never an error: no key saved ("No AI key is saved yet: paste
one under Advanced > Assistant."), a timeout or gateway error, prose instead of JSON, an empty
answer, or JSON with none of the four fields filled in ("The assistant gave no usable answer"). The
route answers 200 in every one of these cases. An answer is kept per version beside the plain result,
so asking again costs no second gateway call, and `assistant=0` after `assistant=1` is the plain
result again. A no-key or failed answer is not kept: the next tick asks again. **Read again**
(`refresh=1`) rebuilds the result and forgets the assistant's answer.

## Limits

- **No numpy, no Excel, no model call** in the finder itself: pure Python over the analysis the scan
  already made. The same workbook gives the same answer.
- **Every list is bounded and every cut announced.** At most 12 dimensions (`counts.dimensions`
  holds the full count), 8 evidence lines per dimension, 6 cells per line, 12 key values, 40 edits
  per kind, 10 chains, 6 macro lines; `truncated` says what was cut. The two-minute budget (started
  after the workbook is in memory) can skip a step on a very big model; the step is then written in
  `skipped` in plain words ("repeated blocks on sheet X not checked: the time budget ran out"). Nothing
  is skipped silently: a paste area cut into more than 1,000 slots, a list of more than 2,000 cells,
  cells reading the paste area that no key column, list or stepping block joins (a macro that pastes by
  its loop counter), and a workbook copy that could not be opened are each one note in `skipped`.
- **The app keeps two workbooks open** (`WORKBOOK_CACHE_LIMIT` in `app/inventory.py`). A third
  upload, another session's scan or an Apply drops the analysed copy; the finder then has to read the
  workbook again before it starts -- minutes on a 28 MB model -- and says so: `counts.reload_seconds`
  and a note in `skipped` ("the workbook had to be read again (N s): the app keeps two workbooks open
  and this one had been dropped"). No step is skipped because of it; the "up to two minutes" of the
  panel is then "plus the time to read the workbook again".
- **One long GET.** No background job in 1.12.5: the request takes up to about two minutes on
  350,000 formulas, plus the reload above, plus the builder's plan once per version when the panel has
  not computed it yet. The finder runs only on its button, never when the panel opens.
- **Copies at the same addresses only.** Sheet copies are aligned on the absolute address: a sheet
  copy shifted by a row or a column is not seen as a copy. Bands repeat with a period of at most 40
  positions; a wider block is not seen.
- **Labels are never proof.** The name of a dimension is the header or the defined name, shown to
  help the reader; the evidence is what the macro and the formulas read. The finder cannot tell a
  list of items from a list of periods by the words; a time axis is seen by its formulas (each
  position reads the previous one) or a header row of numbers.
- **Run time is not predicted.** Formula counts are not run time; Mind's speed and memory are not
  estimated. `biggest_copy` -- the rows of the biggest group -- is the nearest hint of a run's length.
- **One table, one key.** The table the first counter loop indexes; a mapped key is followed one
  level, an alias three. Copies across separate workbooks are not covered. Nested loops and a second
  key are not looked for.
- **The builder builds grouping columns only.** A block or sheet dimension (P2, P3) is reported with
  its edits list; turning it into one block and a loop, or one sheet and instances, is done by hand.
  Chains are set aside, not solved.
- **The assistant's extra candidates are not checked**; the flag says so on each.

## Running it

- **In the app:** Advanced > Instances > **Not sure? Let the app look** (the finder runs when the fold
  is first opened; the fold is there with or without a run loop, and when the plan could not be read --
  the finder reads the formulas, not the macro). The top suggestion by default, with **Use it**, and **Show all N** for the rest; each card has the headline, the three fit chips,
  **Use as the key** (fills the builder's key column and previews the instances; a column the
  dropdown cannot show is named beside it) and **Show why** (the evidence, the reader classes, the
  reasons, the edits, the assistant's sentence; one card open at a time). **Read again** rebuilds the
  result for the current version.
- **The route:** `GET /api/sessions/{id}/instance-finder?assistant=0|1&refresh=0|1` answers
  `{"version_id", "finder"}`; the result is cached per version (a new version after an Apply
  invalidates it). 404 for an unknown session, 409 before the analysis, 500 `instance finder failed:
  <Type>: <message>`; the assistant never turns into an error. `GET /api/health` reports
  `"finder": true`.
- **The answer** (`schema: "instance-finder/1"`): `counts` (sheets, formulas, dimensions, chains,
  time_axes, seconds, reload_seconds), `dimensions` (each with `key`, `copies`, `biggest_copy`,
  `formulas_saved`, `cells_changed`, `evidence`, `blocks`, `readers`, `cross_reads`, `edits`, `fit`),
  `chains`, `macro_loops`, `switches`, `builder`, `ignored`, `ignored_detail`, `skipped`,
  `truncated`, `assistant`. Only plain JSON types; `json.dumps` passes without a default.
- **From the command line**, from the repo folder: `python -m app.instance_finder <workbook>` prints
  the text of `describe()` -- one line per dimension with its evidence, the reader classes and the
  three fit lines, then the chains, the macro loops, the filled areas and whether the builder's rule
  agrees; `--json` prints the result; `--assistant` adds the assistant pass when a key is saved.
- **Tests** (no Excel needed, invented workbooks only): `python -m pytest
  tests/unit/test_finder_shapes.py tests/unit/test_instance_finder.py tests/unit/test_web_api.py -q`.
  The echo of the blind check, `test_uc2_blind_top_dimension_matches_the_builders_key`, reads its
  expectations from a local file that is never committed and skips elsewhere.
