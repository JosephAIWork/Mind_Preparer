# Changelog

## 1.8.2

**FRM-010 -- empty text as lookup key.** The Horizon file the app made in
1.8.1 went to Mind and came back with one group left: `CALCUL!T192:T812`,
NaN, where `T = 1+($C$10=1)*IFERROR(VLOOKUP(F192,Criteres!$N$11:$O$33,2,0),...)`
and F192 holds `""`. Four diagnostic uploads (`docs/MIND_VS_EXCEL.md`) showed
the mechanism: once Prep has given the looked-up table a title, Mind matches
the `""` key to the table's empty first header cell (`Criteres!N11`) and the
lookup returns the header text; `0 * text` is NaN. In Excel, and in Mind
without the title, `""` finds nothing and IFERROR returns its fallback. Number
keys are not affected. The rule flags an exact VLOOKUP / HLOOKUP / MATCH
inside IFERROR whose key holds `""` over a column or row with an empty cell;
the Prep action **Test for empty text before the arithmetic or lookup** (now
serving FRM-008 and FRM-010) puts `IF(key="", fallback, ...)` in front, down
the whole filled-down block. On the app's Horizon output it flags exactly
T35:T812; on the version Mind cleared, nothing. 102 rules.

**A correction found by PVFP.** The rewrite goes down the whole filled-down
block, and on PVFP one row of such a block read a cell that holds an error:
`IF(#REF!="", 0, ...)` is `#REF!` where `IFERROR(..., 0)` returned 0 -- 19
computed values changed, 9 into errors, and the numbers check after the
fixer's Prep said so. A row whose tested cell holds an error now keeps its
formula (Excel and Mind both stop on the error and fall back), listed in the
action's skip list. On PVFP: 120 rewrites, 7 held back, no computed value
changed.

## 1.8.1

**Two places where Mind computes another value than Excel**, found the only
way they can be found: by uploading. Horizon came out of 1.8.0 with zero
error cells in Excel and "Acceptable by Mind"; in Mind, 4,352 cells held
another value than in Excel. Two causes, each confirmed in Mind with a
49-cell test workbook on invented data (`scripts/build_mind_test_workbook.py`;
what was tested and how it came out: `docs/MIND_VS_EXCEL.md`). With both
rewritten, Mind reports **no difference** for Horizon.

- **FRM-008 -- empty text in arithmetic.** `=IFERROR(1+INT((D192-1)/12),"")`
  with D192 = `""`: Excel stops on the `""` (#VALUE!) and returns the IFERROR
  fallback; Mind counts `""` as 0, computes a number and never reaches the
  fallback. 1,242 cells were blank in Excel and held a value in Mind, and
  what read them went wrong after them (two columns came out NaN in Mind).
  Prep action **Test for empty text before the arithmetic**:
  `=IF(D192="","",IFERROR(1+INT((D192-1)/12),""))`, down the whole filled-down
  block so a column keeps one formula.
- **FRM-009 -- an empty cell compared with a number.** Excel reads an empty
  cell as 0; Mind does not (`0 = empty` is false there, `1 <= empty` is true).
  The dangerous form is the one nobody sees: a formula that *lands* on an
  empty cell and shows 0 in Excel -- a plain reference (`=+OUTPUT!P4`) or an
  exact lookup (`H9 = VLOOKUP(B2,'Base Pret'!A:K,11,0)`, the loan had no
  deferral period). `IF(D35<=$H$9,0,1)` returned 0 in Mind on the 157 live
  months of the loan and the reserves came out wrong. Prep action **Read empty
  cells as 0 where they are compared**: `=N(VLOOKUP(...))` on the cell that
  lands (one rewrite serves all 942 comparisons that read it), or `N(ref)` where
  the empty cell is compared directly.

Both rules are REQUIRED (a model that computes wrong in Mind is not
acceptable), both actions are on by default and leave Excel's values as they
are -- "Check the numbers" proves it on every formula cell. They read the
values Excel stored in the file, so a workbook Excel never calculated is
reported as not checked.

What the rules do **not** cover, so that they flag only what Mind confirmed:
`IF(ISERROR(x),...)` in place of IFERROR; INDEX/MATCH, XLOOKUP and OFFSET
landing on an empty cell (plain references and exact VLOOKUP / HLOOKUP are
worked out; the others would need Excel to say where they land); a cell that
is empty in another scenario than the one saved in the file.

**The automatic fixer and FRM-008.** "Fix all automatically" answers a
#VALUE! with `=IFERROR(formula, 0)`. When the #VALUE! came from `""` in
arithmetic, that is exactly the form Mind computes differently. It is caught:
the analysis after the fix raises FRM-008 and the verdict sends the user back
to Prep for the rewrite above. On Horizon the fixer happened to choose the
other repair (`N()`), which Mind computes like Excel.

**Tested so far.** Horizon through the whole app flow on this code: Prep
wrote the 1,557 rewrites and changed no computed value, the fixer took 721
error cells to 0, the Prep round after the fixer rewrote the 665 formulas the
fixer had left in a form Mind computes differently, verdict "Acceptable by
Mind". That final file has not been uploaded to Mind yet (the one Mind
cleared was rewritten by a script with the same forms). **Palermo, PVFP and
CNHI have not been re-run on 1.8.1**: the run was stopped when the machine
ran out of memory. `scripts/run_usecases.py` now applies the blocking Prep
actions once more after the fixer, as the verdict asks.

Also: `app/validators/mind_values.py` (the scan, shared by the two rules and
the two actions), `tests/unit/test_mind_values.py` (4 tests), 101 rules.

## 1.8.0

Built and tested in one unattended run (2026-10-05/06) on four client models:
Horizon, PVFP, Palermo, CNHI. The run's own log, with every dead end, is
`docs/AUTOFIX_RUN_LOG.md`; the use-case results are in
`docs/AUTOFIX_USE_CASES.md`.

Result through the app, on the final code: **Horizon 721 error cells -> 0,
Palermo 511 -> 0, CNHI none to begin with** -- no value that was good before
changed, and Prep changed no computed value (808,798 formula cells compared
with the originals). **PVFP 65,883 -> 887**: the errors its own formulas hide
on purpose, each listed with the value it would change; 0 with the owner's
rule lifted for it, and the 421 values that moved listed.

### Fix all automatically (the Recalculate step)

Until now every formula error went through the Fix panel by hand: open the
error, ask the assistant, read the proposal, Apply, recalculate. One model
had 65,883 error cells. The Recalculate screen now has one button, **Fix all
automatically**, and the app goes through the errors by itself. The manual
way is unchanged and stays right below it.

Two rules, set by the tool's owner:

1. **Clean means zero error cells**, the ones the original workbook already
   had included.
2. **A good number never changes.** A fix that changes a value that was fine
   is undone and another way is tried; when none works the error stays and is
   listed for a person, with the reason.

How it works (`app/autofix.py`), in ONE Excel session:

- **Read**: full recalculation, then every formula cell in R1C1 notation --
  cells filled down or across share one R1C1 text, so a block of thousands
  of cells is one group, one decision and one write.
- **Find where errors start.** An error whose own inputs are fine is a
  *root*; everything downstream only repeats it and clears by itself. Only
  roots are fixed. What INDIRECT and OFFSET point at is asked from Excel
  (probe formulas down an empty column: one write, one recalculation, one
  read -- `Worksheet.Evaluate` answers `#REF!` for INDIRECT inside another
  function, and a scratch *sheet* makes Excel re-evaluate GET.WORKBOOK-style
  names, which fails with macros off).
- **Fix**: the formula is kept and given a value to return when it fails,
  `=IFERROR(<formula>, 0)` -- `""` where the same formula returns text
  elsewhere, `FALSE` for a test. A formula that can never compute again (only
  `NA()` / `#REF!` is left of it, or a `SUMIFS(#REF!, ...)` Excel will not
  even store) becomes `=IFERROR(NA(), 0)`: still a formula (SKILL.md rule 6),
  and one that says what it is. An error value sitting in a data cell is
  cleared. Nothing else is rewritten.
- **Or a real repair, when the assistant sees one**: `=+BE36/BF36` ->
  `=IF(BF36=0,0,BE36/BF36)`, `=C4+B5` -> `=C4+N(B5)` where B holds a dash typed
  as text (the balance is then carried forward instead of reset to 0). Such a
  rewrite is written to the **whole block** -- every cell sharing the formula
  -- and kept only if the error cells are cured, **every cell of the block
  that works today still returns exactly the same value**, and the numbers
  gate passes. Otherwise it is undone and the wrap is used. A language model's
  text is about to go into a client's workbook, so before anything is written
  `repair_refused()` checks that the rewrite only reads what the original
  reads (same sheets, same names, no other workbook), only calls what the
  original calls plus a short list of guards (IF, IFERROR, IFNA, N, VALUE,
  ISNUMBER...), and holds no DDE link.
- **The numbers gate**: recalculate, re-read every formula cell, and compare
  with the workbook as it was opened. A formula cell that had a valid value
  must still have exactly that value. The walk back from a moved value names
  the fix that moved it; a fix that moved a value *alone* has failed without
  a round of its own, shared suspects are tried again by halves.
- **Pass after pass** until no error is left or nothing more can be fixed
  safely; then a full rebuild as confirmation, save, verify the file opens.

What it meets in real models, and what it does:

- *A recurrence that fails row after row* (`=C191+F192-F191`, 621 rows, each
  failing on its own text input): fixing the first only makes the second the
  next root -- one pass per row. Once a formula block has had a fix kept and
  shows a root again, its cells that only wait on each other go together.
- *An error a formula swallows on purpose*: `=IFERROR(K24/K12,"No GEP")`
  shows "No GEP" because K24 is an error. Give K24 a 0 and that cell shows
  0.5 -- a good value changed. Such a fix is undone; `""` is tried (text
  breaks arithmetic downstream, so the handler still fires); when that fails
  too the error stays, together with the error cells between it and that
  formula (pinned at once, not found out one level per pass), and cleaning
  goes on from the cells that read them. On PVFP: 65,883 errors -> 887 in
  3 min 18 s, the 887 all feeding 8 formulas that hide errors.
- *The option to lift rule 2*, explicit and off by default: "Fix them too,
  and list every value that changes" (`keep_good_values: false`). PVFP then
  comes out clean in 29 s, and the 421 values that moved are listed.
- *The assistant* (`app/autofix_advisor.py`) says, per kind of error, which
  value fits -- 0, empty text or FALSE -- from the cell's labels, its formula
  and what it reads, may propose the repair itself (above), and words the
  reason shown to the user ("Dividing two blank cells can't produce a rate,
  so treat it as zero."). It only proposes: the block check and the numbers
  gate have the last word, and no answer means the built-in order, never an
  error.

Web API: `POST /api/sessions/{id}/auto-fix` (`use_assistant`,
`keep_good_values`, `time_budget_min`) runs in a background thread;
`GET .../auto-fix` reports the pass and the errors left, `?full=1` the whole
result (every kind of fix with its cells, before/after and reason; what is
left and why; the run's own recalculation; the analysis of the new version);
`POST .../auto-fix/stop` stops after the current round and keeps what is
fixed. The result is a new **major** version ("v3 -- Auto-fix: 702 cell(s)
fixed · recalculated clean") whose recalculation is the fixer's own, so
READY-001 is PASS and the verdict is green without another Recalculate. No
Recalculate or Apply is accepted while the fixer runs, and the reverse.

Front-end: `components/AutoFixPanel.tsx` on the Recalculate screen -- the
button, a live line (pass, errors left of how many, elapsed, Stop), then what
was fixed (cells, error, reason, before -> after), what is left for a person
(with "Fix by hand" opening the Fix panel on that cell), and the values that
changed when rule 2 was lifted. The error table is capped at 300 rows (65,883
rows froze the page).

### The numbers check: did the preparation change what the model computes?

Found while testing the fixer on PVFP: after the app's standard Prep the
model computed **2,462 different values -- 1,781 of them had become errors**
(`LoB_Total!K8`: 2,280,000 -> 0; `Inputs_CoC!G20`: 0 -> #REF!). Nothing in
the app said so, and a fixer that puts a 0 over an error would have hidden it.
Cause: Prep inserted title rows on sheets that are read *by position*:

- a 3-D reference, `=SUM('LoB 1:>>'!K65)`, reads K65 on every sheet from
  'LoB 1' to '>>'. A row inserted on 'LoB 1' alone shifts that sheet only;
- an address built as text, `=INDIRECT("'"&$E$2&"'!H"&n)`, is not a
  reference at all: Excel moves nothing.

Three changes:

- **`app/numbers_check.py`** recalculates the original and the current
  version in Excel and compares every formula cell, following each cell
  through the rows / columns inserted and the sheets renamed since (the
  change log of every Apply; versions now record their `parent`). Verdict in
  one sentence, counts (good -> error, good -> another value, error -> good),
  per sheet, with examples. `POST /api/sessions/{id}/numbers-check`;
  **Check the numbers** on the Prep screen once a version other than the
  original is current; `scripts/compare_versions.py` from the command line.
- **Prep no longer inserts a row there.** `prep.reference_index` knows the
  sheets inside a 3-D range (`positional`) and the sheets an INDIRECT names
  (`text_addressed`: a sheet named in one of its text pieces, or by a cell it
  reads; a cell holding some other text names a sheet that does not exist and
  puts nothing at stake; nothing readable at all means every sheet).
  `insert_blocker()` gives the reason in words; `create_grid_titles`,
  `separate_merged_grids` and the Grid Namer skip the insert and say why.
- **The fixer compares first** when it runs on a prepared version: an error
  that was a value in the original is the preparation's doing. It is never
  given a fallback; it stays, listed with what it was ("it was 1140000 in
  the original workbook"), and the result says what Prep changed.

### Also

- `scripts/run_usecases.py` drives the web API through the whole flow for a
  list of workbooks (upload -> standard Prep -> numbers check -> Recalculate
  -> Fix all automatically -> Recalculate again), with a fresh server per
  workbook (`--ports 8602,8603` runs them side by side: one server process
  held 7.7 GB after two analyses). `--prep-rounds N` (default 3: Prep is
  applied again while it still plans something) and `--one-comparison` (the
  fixer's own comparison with the original is used instead of asking twice:
  on a 37 MB model each comparison is two full recalculations).
  `scripts/usecase_report.py` turns one or several runs into the tables of
  `docs/AUTOFIX_USE_CASES.md`. `scripts/ui_autofix_check.py` clicks through
  the real screens with Playwright.
- **Theme colours are replaced in the workbook's style table**
  (`app/theme_colors.py`; FMT-002 "Replace theme colours with explicit RGB",
  ticked by default). The action went through the cells in Excel: read the
  colour, assign it back. Two faults on real models. It never finished its
  job: Excel keeps the default text colour as a theme colour when a cell is
  assigned the colour it already shows, so every Prep round planned the same
  cells again (one sheet of Palermo: 11,807 -> 11,719 -> 11,720 cells; 25
  sheets in each of three rounds). And it was slow, ~10 COM calls a cell
  where colours are mixed: 46 minutes for one round on PVFP, an hour for
  *one* sheet of CNHI (123,686 cells) with 60 more colour operations behind
  it -- that run could not finish. A cell does not carry a colour, it points
  at a font and a fill; now the fonts and fills of `xl/styles.xml` are
  converted on the fresh copy before Excel opens it. Every sheet at once:
  0.3 s (Horizon) to 9 s (CNHI), and nothing is left to plan at the next
  scan. The RGB written is the one Excel shows. Its tint arithmetic is not
  the textbook one (whole-number HLS on Windows' 0..240 scale; the
  floating-point formula is 1 or 2 off on more than half the shades): found
  by reading 276 colour / tint pairs back from Excel, all exact now.
  Going through the cells stays as the fallback for a package whose style
  table cannot be read -- by **blocks** of one colour since this version
  instead of cell by cell (`tests/unit/test_explicit_colors.py` compares
  both ways).
- **Recalculate no longer asks Excel about every error cell**
  (`app/recalc.py`). The scan read values and formulas in bulk, then asked
  three things about each error cell (Cells, Address, HasArray). On a model
  with 65,883 error cells that is ten minutes a scan, and the first
  Recalculate scans twice (the original is recalculated as the baseline):
  about 25 minutes before the "Fix all automatically" button could even
  appear. The addresses are now computed from what was already read, and
  `Range.HasArray` is asked for the sheet's used range (False = no array
  formula at all: one question), then per row, then per cell only on rows
  that hold one. Same errors, same arrays, same groups; 55 s for that model
  alone, 4 minutes for its first Recalculate in the app on a busy machine
  (it was 28). `tests/unit/test_recalc_scan.py`.
- A COM lesson: the fixer's Excel session lives in a function of its own
  (`_fix_in_excel`). While the Excel proxy of that session was still held by
  `run_autofix`'s frame during the next session (`verify_opens_in_excel`),
  its late release left the thread's COM apartment unusable: the next
  recalculation in the same process failed with "The interface is unknown"
  (8 tests of the full suite, each passing when run alone).
- Tests: `tests/unit/test_autofix.py` (23: R1C1 reading, range index, the
  ladder, and the engine in a real Excel -- roots, rollback, layers, arrays,
  swallowed errors, pinned channels, INDIRECT / OFFSET, the assistant's
  fallbacks and repairs), `tests/unit/test_numbers_check.py` (4), three more
  in `test_prep_reference_safety.py`, two in `test_web_api.py`,
  `test_explicit_colors.py` (2), `test_theme_colors.py` (6: the tint against
  what Excel showed, the style table, the package, and an Apply through
  Excel that leaves no theme colour and the same look). `test_recalc_scan.py` (3). Full suite: 290 passed.

## 1.7.4

Merge of `main` with Yoav's branch `readiness-verdict-and-prep-fixes`
(his 1.7.2 and 1.7.3 below). The changes in this entry were made on `main`
in parallel, under the number 1.7.2; they are listed here unchanged.

**Merge decision (REF-001).** Both branches fixed `SUMIFS('Sheet'!#REF!,...)`:
`main` by collapsing the call to `NA()` inside the piecewise `fix_broken_refs`,
the branch by a separate `fix_broken_refs_whole` action (own checkbox, caution
note) that replaces the whole formula with `=NA()`. Kept: the separate action,
now filled with the collapsed formula -- only the call that needed the range
becomes `NA()`, so `=IFERROR(SUMIF(#REF!,1),0)` -> `=IFERROR(NA(),0)` keeps its
0 instead of turning into an error. `_com_message` keeps the plain-English
0x800A03EC text and the branch's transient-COM retry.

### Formulas Excel refuses are caught before Apply (17 failed fixes on a real workbook)

On `_20240118 Horizon PM PRC 122024 - Anonymisé V2.xlsm` the assistant
proposed 21 fixes for `'BASE Polices'!#REF!` on sheet C20; 17 failed at
Apply with a raw `(-2147352567, 'Exception occurred.', (..., -2146827284), None)`.
Cause: it rewrote `SUMIFS('BASE Polices'!#REF!, ...)` to `SUMIFS(NA(), ...)`
(the REF-001 `#REF!` -> `NA()` idea), and Excel refuses to store any formula
with a value where a function needs a cell range (COM 0x800A03EC, nothing is
written). The 4 that worked were `SUM(NA())` -- SUM takes values.

- **`formula_utils.reference_arg_problems`** knows the reference-only
  argument positions of SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF(S),
  MAXIFS, MINIFS, COUNTBLANK, OFFSET, ROW, COLUMN, AREAS, CELL and SUBTOTAL,
  and flags an argument there only when it is certainly not a reference
  (NA(), #N/A, a number, text, TRUE, an array constant, arithmetic,
  IFERROR/FILTER/...). A range, a name, `#REF!` and INDEX / OFFSET /
  INDIRECT / IF / IFS / CHOOSE / SWITCH / XLOOKUP / LET pass. 49 cases were
  checked one by one in Excel and are pinned in `tests/unit/test_reference_args.py`.
- **`collapse_na_reference_calls`**: a call holding NA() in such a position
  becomes `NA()` itself, innermost first (`=+SUMIFS(NA(),J:J,"x")` ->
  `=+NA()`; `=IFERROR(SUMIF(NA(),1),0)` -> `=IFERROR(NA(),0)`). Exact: the
  call held a broken range, so it could only ever produce an error.
- **Assistant proposals** (`validate_proposal`): a `set_formula` /
  `set_array_formula` Excel would refuse is rejected with the concrete repair
  (`Use =+NA() instead ...`), so the assistant's repair round fixes it
  before the user sees a proposal.
- **REF-001 prep action** (`fix_broken_refs`): now also rewrites
  sheet-qualified broken references (`'BASE Polices'!#REF!`; they used to be
  "left for review") and collapses calls the NA() swap would make invalid.
  The REF-001 recommendation text tells the assistant the same.
- **Plain-English apply errors** (`_com_message`): 0x800A03EC now reads
  "Excel refused this formula: SUMIFS argument 1 must be a cell range, but it
  is NA() ..." instead of the raw COM tuple; the Assistant screen lists the
  failed cells grouped by reason (the Fix panel already listed them).

### Faster assistant: Sonnet 5 by default, and replies stream as they are written

A "Propose a fix" turn could sit on bouncing dots for minutes: one turn is up
to five model calls in a row (the answer, up to 2 ```lookup round-trips, up to
2 repairs of a rejected ```changes block), each sent without streaming.

- **Default model `claude-sonnet-5`** (`app/llm.py` `DEFAULT_MODEL_ID`; the
  grid namer uses the same constant). Measured through the APIM gateway on
  2026-09-24, same ~150-word answer, one run each: Sonnet 4.6 7.7 s, Sonnet 5
  6.2 s, Opus 4.8 6.7 s, GPT-4.1 3.7 s; the gateway still has no Haiku.
  `set MIND_READY_MODEL=claude-sonnet-4-6` switches back.
- **Sonnet 5 rejects `temperature`** (400 "`temperature` is deprecated for
  this model"). `chat_completion` now sends it only to the 4.x models
  (`_takes_temperature`) and, for any other model that answers such a 400,
  retries once without it.
- **Sonnet 5 thinks before it answers by default** -- and its thinking is
  not streamed through the gateway. On a real chat turn (3k input tokens,
  `max_tokens` 1200) it spent ~885 tokens / ~9 s thinking and then ran out
  of room: `stop_reason: max_tokens`, answer cut off or EMPTY (shown as
  "(no answer: None)"). Newer models now get `output_config.effort: "low"`
  (env `MIND_READY_EFFORT`; empty = the model's default) plus
  `THINKING_HEADROOM_TOKENS` (3000) on top of `max_tokens`; a 400 naming
  `output_config` drops it and retries. A reply that is cut off now ends
  with "_(The answer was cut off at the length limit.)_" and an empty one
  says "The model returned no text (stop_reason ...)" -- never silent.
  Live, same turn on a small workbook: Sonnet 5 first text 1.4-3.6 s, done
  in 5-8 s with a valid proposal; Sonnet 4.6 first text 1.9-3.5 s, done in
  12-15 s.
- **Streaming**: `chat_completion(..., on_delta=)` posts `stream: true` and
  parses the gateway's SSE (`_sse_data` / `_parse_claude_stream`, ported
  from IFRS_DataScraper/app/apim.py: bytes-level lines decoded as UTF-8; an
  `error` event or a stream without `message_stop` is a failed call, never a
  partial answer). A gateway that ignores `stream` and answers JSON is read
  the old way. `MIND_READY_STREAM=0` turns streaming off.
- `answer_question(..., on_event=)` reports a `phase` event
  (`answer` / `lookup` / `repair`) before every model call and a `delta` for
  each chunk of text.
- **Web API**: `POST /api/sessions/{id}/chat/stream` runs the same turn as
  `/chat` and answers NDJSON lines (`phase`, `delta`..., then `done` with the
  `/chat` payload, or `error`); 404/409/422 stay plain HTTP errors. `/chat` is
  unchanged (the body is now shared through `_chat_turn` / `_chat_payload`).
  `/api/health` reports `model` and `streaming`.
- **Front-end**: `api.chatStream` (falls back to `api.chat` with mocks or on
  an older backend); `components/StreamingDraft.tsx` shows the reply as it
  arrives in the Assistant screen and the Fix panel, with "Looking up cells in
  the workbook…" / "Checking the proposed change…" between calls and
  "Preparing the change…" while the model writes its ```changes block (that
  block is never shown). The proposal card appears when the turn is done, as
  before.
- Tests: SSE parsing, cut-off and error streams, JSON fallback, the
  temperature rule and retry, the phase events of the lookup and repair
  scenarios, and `/chat/stream` (same `done` payload as `/chat`, in-band
  error). `test_chat_completion_payload_shape` updated for the top-level
  `system` param of the 2026-09-16 route switch.

## 1.7.3

### An Apply that says what it did -- and that really writes

On a workbook saved with "read-only recommended", Apply ran through every
stage, reported "38 changes applied, verified" and changed nothing: the
blocking problem was still there after each round, with no sign of why.

- **The save went to the wrong place.** With alerts off, Excel answers the
  "open as read-only?" prompt of such a workbook with yes, whatever
  `ReadOnly=False` says; `Save()` on a read-only workbook then writes a copy
  into the user's default folder (their Documents) and returns normally. The
  working copy stayed byte-identical to its source, the re-analysis read the
  same file, and a stray copy of the model appeared in Documents.
  `open_workbook` now passes `IgnoreReadOnlyRecommended`; every in-place
  write (Prep apply, formula edits, recalculation, report workbook) goes
  through `open_for_write` -- which refuses a workbook Excel opened read-only
  all the same (locked by another program, password to modify) -- and
  `save_in_place`, which wants the file itself written.
- **Nothing is "applied" that is not in the file.** `apply_operations`
  compares the output with its source: identical means nothing was written,
  every operation is reported as refused with that reason, status `ERROR`.
- **Apply outcome** (`app/readiness.apply_outcome`, field `outcome` of
  `POST .../apply` and `.../grid-namer`). Repair by repair: changes sent,
  written, refused; each rule's status and cell count before and after; what
  the new analysis still plans and what it leaves by hand; a verdict
  (`resolved` / `partial` / `unchanged` / `failed` / `written`) and one
  sentence. Also the repairs the new analysis plans that the previous one
  did not (`appeared`), the rules that got worse (`regressed`), and the
  blocking problems before and after with who acts next on each.
- **Applied-changes gauge** (`prep_progress.applies` / `.written`, kept per
  session in `Session.apply_history`). One entry per Apply with the changes
  really written into the file, by level, the refusals and the blocking
  problems it resolved; the Prep gauge shows *written / still planned* for
  blocking and optional changes and the figure of every Apply, the repair
  gauge (Fix panel, Findings) the total written.
- **FRM-002 says what the Mind documentation says, no more.** The KB page
  "Supported Excel formulas" (last updated 2021-11-17, read again on
  2026-09-28) names the functions Mind supports and says nothing about the
  others -- and FILTER, which it omits, converts in Mind. A native function
  missing from the page was reported as an ERROR and a blocker, and the
  assistant and the report told the user to replace it. It is now a WARNING
  ("neither documented as supported nor as refused; a conversion in Mind is
  the proof"); only an MM_ name missing from the MM_ registry stays an
  ERROR. Functions Mind accepts beyond the page live in
  `references/mind-confirmed-functions.yaml`, each with its evidence
  (`MIND_CONVERSION` or `USER_PROVIDED`): FILTER, and IFNA / INDIRECT /
  XLOOKUP as stated by the tool's owner. The assistant no longer proposes a
  replacement for an unlisted function unless the user asks for one.
- **Two refusals Mind reported that the app had let through** (a real model,
  2026-09-29). *"The array formula on sheet Temp, cell F5 exceeds the grid
  size. Ensure that the column headers cover the entire array formula"*: the
  header was one cell merged over B4:E4 and the arrays ran to DZ; the app read
  the merge as one cell, split the block into two grids, and RSK-002 saw
  nothing. Grid detection now lets a merged cell occupy every cell it covers
  (`grids.MergedInto`), so the header row is as wide as Mind reads it and
  RSK-002 names the first cell outside the grid, in Mind's words. *"No
  function found ... (not a function : Semi_dynamic_increase_rates_array)"*:
  the formula referred to a name the workbook does not define (a misspelling
  of `Semi_dynamic_increase_rates`; Excel shows #NAME?). REF-001 now also
  flags names with no definition -- table names, names of any alphabet,
  names with '?', sheet-qualified references and LET variables excluded --
  and proposes the closest existing name.
- **Two more refusals, from Mind's runner and compiler** (real models,
  2026-09-30), detected before upload, with no automatic repair by design:
  **FRM-005 circular references** -- "Run error: Circular reference found".
  The dependency graph of every formula cell (references, ranges, defined
  names; INDIRECT/OFFSET targets cannot be followed and are counted;
  ROW/COLUMN arguments and OFFSET's base are coordinates, not dependencies;
  a cell reading its own address on a run-time sheet -- INDIRECT with
  ADDRESS(ROW(),COLUMN()) -- is the pattern Mind stopped on and is an error) is searched for
  cycles and each one is reported as its chain of cells; Excel's iterative
  calculation setting is reported too. Ranges are shared nodes, so a model of
  57k formulas is searched in under 30 s; a work budget makes the rule say
  NOT_SUPPORTED rather than run forever. Fix by hand: the Mind mechanism for
  a value feeding the next round is MM_ITERATIONS with /iterationinput and
  /iterationoutput. **FRM-006 3-D references** -- `'First:Last'!cell`:
  Mind reads `First:Last` as one sheet name ("Sheet ... not found in workbook
  on compiling formula"). The sheets the reference spans are listed in tab
  order and the explicit per-sheet formula is proposed for the assistant or
  the user to apply -- without the sheets that hold no grid: Mind creates no
  spreadsheet for a divider tab ('Spreadsheet not found error' on the same
  model, 2026-10-01), and an empty sheet adds nothing to the sum. Any formula
  reading such a sheet is flagged too.
- **Excel executor**: a change Excel refuses with RPC_E_CALL_REJECTED (busy for
  an instant) is repeated up to five times before it counts as refused.
- **FRM-007 INDIRECT used as a value** -- compared, multiplied, divided,
  negated, concatenated or passed to a value-only function: Mind returns a
  reference for INDIRECT and cannot turn it into a value ("Run error: Unable
  to cast object of type 'AM.Models.AMReference' to type
  'System.IConvertible'", model PVFP, 2026-10-01). Where the text INDIRECT
  builds can be worked out from the workbook's constants (string literals,
  cells holding text, ROW()/COLUMN(), simple arithmetic, LEFT/RIGHT/LEN/IF),
  the direct reference is proposed; a target that does not exist becomes
  NA() in its place (never the whole formula: the call may sit in a branch
  never taken). Verified on PVFP: 4,410 formulas rewritten through the app.
- **Excel executor**: calculation is set to manual while the changes are
  written and restored before the save -- in automatic mode Excel
  recalculated every dependent after each write, and 12,678 writes into a
  70k-formula model ran for hours; they take 23 minutes now.
- **Front-end.** Prep and the Fix panel show that outcome after every Apply
  ("What the last Apply did"); a blocking repair that resolved its problem
  stays in the Blocking block as *Resolved* instead of vanishing; a repair
  that is back says so ("new since the last Apply", "back after the last
  Apply -- it had no effect").

## 1.7.2

### One verdict, one vocabulary, and no more endless Prep rounds

A real model prepared with 1.7.1 went through seven Prep rounds (v8 to v14)
without anything blocking being fixed, while the two problems Mind actually
refused the file for had no automatic repair and were never singled out.
This release answers the only question that matters on every screen, tells
a change Mind *needs* from one it merely benefits from, and removes the two
causes of the endless rounds.

- **Readiness verdict** (`app/readiness.py`, `GET /api/sessions/{id}/readiness`,
  and the `readiness` field on every analysis, apply and recalculation
  response). Three states: `blocked` (a REQUIRED rule is in ERROR or waits for
  user input -- Mind's converter refuses the file or computes it wrong),
  `unverified` (nothing blocks, but this version was not recalculated in
  Excel, or the recalculation found genuine formula errors), `ready` (nothing
  blocks and a clean recalculation of *this* version). Each blocking finding
  says how it can be fixed (`prep` / `prep_skipped` / `assistant`), and
  `next_step` names the single next screen. The front-end shows it as a
  banner under the top bar, with the five workflow steps.
- **Levels everywhere.** Findings carry the rule's `priority` (models.Finding,
  `run_rule`); every Prep action carries `level` (`blocking` when its rule is
  REQUIRED, `optional` otherwise), an optional `level_note` and a `caution`.
  Blocking actions are on by default -- `fix_broken_refs` was off before,
  which left the one blocking repair unticked while cosmetic ones were ticked.
  Findings gained a Level column, a "Blocking only" filter and blocking-first
  ordering; Prep is split into *Blocking / By hand / Optional*.
- **Prep gauge** (`prep_progress`, kept per session in `Session.plan_history`):
  blocking problems, blocking and optional operations planned, the planned
  operations of every analysed version, and a `stalled` flag when three
  analyses in a row leave work planned without a net decrease.
- **Endless round 1 fixed.** `plan_separate_merged_grids` inserted an empty
  row above a '#Title' that sat on its grid's *first* row (next to a label or
  a header); the block moved down and the next scan proposed the same insert.
  Such a title is now left for review with the reason ("shares the grid's
  first row with other content -- move it by hand").
- **Endless round 2 fixed.** A title written into a non-first cell of a merged
  range is silently ignored by Excel; Prep reported "applied, verified" and
  re-proposed it forever. `plan_actions` now drops such targets into the skip
  list (`merged_ranges` / `merged_conflict`, from the inventory's
  `merged_cells`), and the Excel executor refuses them (`_merged_guard`) and
  reads every cell back after writing (`_read_back_guard`): a silent no-op is
  a failure, never an applied change.
- **Broken ranges.** `fix_broken_refs` only patched a #REF! that stood for a
  value; a #REF! standing for a range (a deleted column inside SUMIFS) was
  left for review with no way forward. New action `fix_broken_refs_whole`
  replaces the whole formula with `=NA()` (the cell already returned an
  error), with its own checkbox and caution.
- **Recalculation errors the original already had are the model's own.**
  The first recalculation of a session also recalculates a copy of the
  ORIGINAL upload once (`Session.original_baseline`); every error cell of a
  prepared version that the recalculated original also has (same
  sheet+address, or same formula text on the same sheet when rows were
  inserted, or a formula that already contained #REF!) is marked
  `preexisting` (`recalc.classify_errors`). Cached values are never the
  baseline -- they can be stale (a number saved before the column a formula
  pointed to was deleted). The verdict only counts *new* errors, so a model
  with legitimate #DIV/0! cells can still turn green, and a wrong assistant
  fix (VALUE("") where NUMBERVALUE("") returned 0) shows up as new errors
  before anyone ships the file. The recalculate response carries
  `preexisting_errors`, `new_errors` and `compared_with_original`; the
  assistant is told to keep a replacement's behaviour on blanks.
- **Merged first cell.** Excel refuses `ClearContents` / a value on the single
  first cell of a merged range ("We can't do that to a merged cell"); the
  executor now writes to the merge area instead, so a caption in a merged
  band can still become a title.
- **Style actions no longer run in 5000-cell batches**
  (`inventory.MAX_STYLE_CELL_REFS` 5000 -> 250 000).
- **Assistant proposals** get 4000 output tokens (`chat_context.PROPOSAL_MAX_TOKENS`)
  instead of the 1200 default that truncated a ```changes block at ~35
  operations with nothing said; the Fix panel warns when a proposal covers
  fewer cells than the finding counts.
- **History: "Restore as current"** re-analyses an earlier version and makes
  it current (later versions stay). Prep shows what the last Apply did and
  what is left; the upload screen no longer hard-codes a rule count.

Follow-ups from the first test sessions on real models:

- **Spilled-range references are blocking and repaired.** Mind's own
  validation refuses every `INDEX(A1#, ...)` with "Unsupported formula:
  ANCHORARRAY()" (the token Excel stores for `A1#`). FRM-003 now fails with
  ERROR on such references (a bare `@` stays a WARNING), and the new blocking
  Prep action `freeze_spill_refs` replaces each `A1#` / `ANCHORARRAY(A1)` by
  the fixed range the spill covers today, taken from the anchor's array
  extent in the file, on the same sheet or across sheets; a consumer that
  spills itself becomes a fixed-size array formula over its current range.
  References whose anchor is not an array formula are left for review.
  Verified through Excel on real dynamic arrays: values unchanged, FRM-003
  passes afterwards.
- **The assistant answers "is X supported?" from the tool's list, not from
  memory.** Any Excel or MM_ function named in a question gets a
  "## Function support" block in the retrieval: on the Mind list / not on it
  (FRM-002 flags it) / registered MM_ function, with the workbook's usage and
  a supported equivalent the tool vouches for (`NUMBERVALUE`, `IFNA`,
  `XLOOKUP`, `CONCAT`, `TEXTJOIN`, `SWITCH`, `ANCHORARRAY`). The system prompt
  forbids hedging ("most platforms...") and explains that `_xlfn.` is Excel's
  storage prefix, not a defect. Motivating case: the assistant called 54
  `IFNA` calls "a non-issue" while FRM-002 flags IFNA (IFERROR is on the
  list, IFNA is not).
- **Formulas are shown the way Excel shows them**: the cell window, the
  FRM-002 call sites and the sites quoted to the assistant drop `_xlfn.`,
  `_xll.` and the other storage prefixes.
- **Findings routes each row like the banner**: "Repair in Prep (N)" when a
  Prep action holds the change, BY HAND when Prep left every site for
  review, "Fix with assistant" otherwise.
- **A repair gauge in the Fix panel and on Findings**: blocking problems and
  open findings per version, with the trend and the verdict, so a run of
  assistant fixes reads as a curve going down like Prep does. Each analysis
  records `open_findings` and `status_counts` in the gauge history.
- **Launcher** uses the project's `.venv` interpreter when it exists.
- **Apply says what it is doing.** The Prep bar used to walk through four
  phase names on a timer and then sit on "Change log" for as long as Excel
  worked. `apply_operations(progress=...)` now reports its stages
  (`progress.APPLY_STAGES`: copy, write, save, verify, change log) and, while
  writing, every operation (`done` / `total`) -- and the cells done inside a
  colour operation, which is one operation over thousands of cells. The
  server keeps it per session and returns it in `GET /status` (field
  `apply`); Prep and the Fix panel show the step ("Step 2 of 6"), the changes
  written, the elapsed time, then the re-analysis stages.
- **Prep gauge redrawn**: the four steps to Mind with "you are here", one bar
  per kind of work (blocking problems, blocking repairs, optional repairs:
  done / left, measured against the most the session ever planned), and the
  figures of every analysed version in words instead of 4-pixel columns.
- **An automatic repair is never the only way.** A finding Prep holds a
  repair for only offered "Repair in Prep"; it now also offers "By hand"
  (Findings, the banner) and "Fix by hand" (the Prep action), which open the
  Fix panel on the same cells.
- **Launcher** opens the browser once the server answers; the check goes to
  127.0.0.1 (through `localhost` the first request of a process can spend 2 s
  on IPv6, longer than the check's timeout).

## 1.7.1

### One backend again: the Shlomo copy's size gate and scan status folded into the main app

The Shlomo copy (`excel-upload-preparation-shlomo`, branched at 1.6.5) had
grown a large-workbook size gate with sheet skipping and a live scan status
(its own "1.6.6", 2026-09-02) that the main app never received, while the
main app went on to 1.6.7-1.7.0 (grid naming, the run-in-Mind loop, the Grid
Namer). The shared front-end (`../FigmaOutput`) had been built against the
union of both and so needed both sets of endpoints -- running it on the Shlomo
copy made every **Run in Mind** click fail with `404 Not Found` on
`POST /api/sessions/{id}/mind-loop`, before any browser was opened. This
release merges the Shlomo copy's commit into the main app (three-way, common
base 1.6.5), so one backend serves everything the front-end calls.

- **Size gate** (`app/sizing.py`): every upload is inspected from the zip
  package alone -- size on disk, decompressed size, every sheet with the
  decompressed size of its part (`xl/workbook.xml`, or the BIFF12
  `xl/workbook.bin` for an `.xlsb`). Threshold `upload_size_threshold_mb: 25`
  in config/default.yaml, env override `MIND_READY_SIZE_THRESHOLD_MB`.
- **Sheets the scan ignores** are never parsed (`inventory.load_workbook_selective`,
  `build_analysis(..., ignore_sheets=)`): they come back as empty placeholders
  at their original index, are listed under `workbooks[0].ignored_sheets` and
  nowhere else, add a `SHEETS_IGNORED` risk, and the assistant is told not to
  touch them.
- **Scan progress** (`app/progress.py`): every mode's `run(...)` takes a
  `progress(stage, message, fraction, **facts)` callback (copy, load, inventory,
  names, rules, report, plan; convert first for an `.xlsb`).
- **Web API**: `POST /api/sessions` takes `defer=1` and `ignore_sheets`;
  `POST /api/sessions/{id}/analyze`; `GET /api/sessions/{id}/status`;
  `/api/health` advertises `upload_gate` and `size_threshold_mb`. The apply,
  reanalyze and Grid Namer re-analyses all go through the same tracked scan.
  The 1.7.0 grid names (`s.grid_names`) still feed the plan after a tracked scan.
- **Streamlit UI**: the same gate and an `st.status` progress block.
- Tests: `tests/unit/test_sizing.py` and the two web-API gate tests join the
  suite (146 total). The Shlomo copy is superseded by this release; run
  `excel-upload-preparation\run_mind_ready_web.bat`.
- **`scripts/run_after_login.py <workbook>`** (new): for the day the Mind
  session has expired and nobody is at the keyboard. It opens the headed Edge
  login window on the dedicated MindReady profile, waits (up to 24 h) for a
  person to sign in, then uploads the workbook to the running app and drives
  `POST /api/sessions/{id}/mind-loop` to the end, echoing every event to
  `runs/after_login_<stamp>/run_after_login.log` and leaving `summary.txt` +
  `loop_report.json` there. Its defaults are the "get it into Mind" ones:
  numbers gate off (`--check-numbers` turns it on), sandbox projects kept
  (`--delete-projects`). The Mind session (Auth0 login at
  `login.milliman-mind.com`) lasts about a week; `python -m app.mind_client
  check` tells you whether it is still valid.

## 1.7.0

### Grid Namer: the user draws the grid names, the conventions do the rest

A new **Grid Namer** screen (sidebar), for the person who knows what a block
of cells *means*: the whole current version of the workbook is shown sheet by
sheet (calculated values where the file carries them, else the formula text;
detected grids outlined and colour-coded — title cells, named grids, untitled
grids, areas queued for naming). Drag over an area, give it a name, tick the
documented flags (`/Input`, `/Export`, …, from `references/mind-flags.yaml`;
flags with arguments typed as text), and **Submit** writes the
`#Name /Flags` titles into a new Excel-verified version. Everything the user
did not touch is titled in the same apply by the existing conventions
(captions, labels, headers, assistant names) — optional, on by default.

- **Naming labels Mind's detection, it does not redraw it**: each selection is
  resolved to the one detected grid it touches. A selection across several
  grids is refused with their refs ("Mind reads them separately"); the one
  exception is the documented caption-above-a-table pair, which is exactly
  what a user selects as "one table" — the caption becomes the title and the
  pair becomes one named grid.
- **Same reference-safety rules as the automatic titler** (`plan_named_areas`
  in `app/prep.py`): a title cell, above-cell or row insert that some formula
  reads is refused with the reader named; a refused area is also excluded from
  the automatic pass (retrying would only repeat the refusal). Renaming a
  titled grid releases its old name; user names are reserved so the automatic
  pass uniquifies against them (`plan_create_grid_titles` grew
  `exclude` / `reserved` / `pre_titled` / `pre_inserted` for this — one
  combined apply, no collisions, shared row inserts).
- **Endpoints**: `GET /api/sessions/{id}/sheet-cells` (one sheet's used area,
  value-or-formula per cell, from the last recalculation when there is one),
  `GET /api/mind-flags` (documented title flags with their meaning),
  `POST /api/sessions/{id}/grid-namer` (plan manual + automatic titles, apply
  via Excel, new version `source: "grid_namer"`, re-analysis, refusals with
  reasons in `skipped`).
- **Browser tab** now reads `MindPrep v<version>` from the running backend
  (`document.title` after `/api/health`; static fallback "MindPrep" in the
  built page).
- **One half of an untitled caption pair is extended to the pair** (found by
  the live smoke test on the caption layout): naming only the table half
  would leave the caption column to the automatic pass, whose separate title
  lands adjacent and makes the re-detected grids merge wrongly. The UI's
  selection panel mirrors the extension ("+ its table, named as one grid").
- Tests: `tests/unit/test_grid_namer.py` — 11 new (resolution, caption pair,
  pair extension from one half, canonical flag casing, unknown-flag refusal,
  read-title refusal, old-name release, reserved-name uniquifying, shared row
  inserts, and the three endpoints end-to-end through the FastAPI
  TestClient). 135 total.

## 1.6.8

### The app runs itself against real Milliman Mind

`app/mind_loop.py` (new) + `scripts/run_in_mind.py` + `POST/GET
/api/sessions/{id}/mind-loop` + the **Mind** screen. One call takes a raw model
and, unattended, repeats *prepare a fresh copy → local gates → upload →
convert → add template → run → read what Mind shows → decide* until every gate
passes or it can say exactly why not. No person and no assistant is consulted
while it runs (the assistant names grids once, before the first iteration, when
reachable). Runbook: `docs/RUN_IN_MIND_LOOP.md` (CLI, app, reading the report,
every `stuck` reason, the numbers gate, cleanup); short form in
`docs/RUN_IN_MIND.md`, Part D.

- **Gates** (all must pass; a gate that cannot be checked is a failure, never a
  pass): prepared · numbers · structure · names · convert · run (Mind: *consistent
  with the audit trail*) · counts (Mind's `Untitled(r,c)` count == the app's
  predicted untitled count).
- **Numbers gate** (`numbers_gate` / `compare_values`): fresh copies of source and
  prepared workbook are fully recalculated by Excel and compared cell for cell,
  mapping coordinates back through the plan's row inserts (`row_map`, keyed by the
  plan's source sheet names so `&&Hide` renames still resolve). Cells the plan
  wrote as content (titles, moved labels) are skipped; cells whose *formula* the
  plan rewrote (`fix_broken_refs`) are compared under the conservative rule --
  unchanged, or error → error, never a valid value into anything else. Volatile
  cells (`NOW()`, `TODAY()`, `RAND*`) and their transitive dependents are skipped;
  chart sheets are ignored.
- **One change per retry, always from the original source**: enable the opt-in
  fix Mind's Convert log asks for (`MIND_ERROR_TO_ACTION`); disable the action to
  blame for a moved number (`actions_touching` walks the changed cell's
  precedents -- references and defined names -- back to a cell an action wrote,
  so a formula rewrite outranks a row insert on the same sheet); disable the
  insert-heaviest action when the grid count balloons. Default-on actions are
  as disableable as opt-ins (`active_actions`). No rule → `stuck` with the reason.
- **Mind operations** folded into `app/mind_client.py` from the proven scratch
  flows: `open_manager`, `ensure_folder`, `create_blank_project`, `open_project`,
  `upload_and_convert` (wizard statuses + error log), `add_template`, `run_model`,
  `template_names` / `untitled_entries` (the `<stem> ▾` selector, decoded to
  cells), `export_summary`, `delete_project` (sandbox names only; the tile's own
  menu via DOM ancestry, and the confirmation modal must name the exact project
  or the delete is refused). `.xlsb` sources are converted through Excel first.
- **A title may never change a value** (`prep.reference_index`, `referenced_by`,
  `whole_reference_to`). The first unattended run's numbers gate caught the
  titling action changing **790 cells** of the real Shlomo model: a title written
  into column A of the policy table was counted by `=COUNTA(CoverageDetails!A:A)`
  and the model picked a *different policy*; titles inside `Parameters!C36:C119`
  were returned by `INDEX` lookups; the company name in `Information!A1` was
  "moved" into a title while five sheets read `=Information!A1`. Every write
  site of `create_grid_titles` / `separate_merged_grids` now asks what the
  workbook's formulas and defined names read: no title into a cell any formula
  reads, no row insert on a sheet a formula counts or indexes over whole columns,
  no caption rewritten when something reads it, and a label that is read stays
  in place (its text still names the grid). Each refusal is listed with the
  reading formula. On Shlomo: 109 safe titles remain, 41 refused.
  **Last night's converted model carried this corruption unnoticed -- Mind's
  "consistent with the audit trail" only checks Mind against the workbook's own
  cached values.**
- **Counts gate is reported, not enforced**: Mind's `Untitled(r,c)` entries are
  matched cell by cell against the app's untitled anchors (`untitled_counts`);
  Mind reads some blocks differently from the app (on Shlomo, 17 grids Mind
  starts one column later than the app does, e.g. `D13` vs `C13`). That is a
  detection gap no prep action can change, so it is written to
  `detection_gaps` in the report instead of blocking convergence.
- Adversarial review before the first unattended run found and fixed 11
  defects, four of which would have produced a false *converged*: rewritten
  cells never compared, a numbers gate that could not run counted as passed, a
  Mind failure after Convert falling through, default-on actions never
  disableable. Regression tests in `tests/unit/test_mind_loop_review.py`.

**Reference run (2026-08-30, raw Shlomo model, unattended):** run 1 stopped
itself twice -- numbers gate (790 changed cells, titling disabled), then the
counts mismatch -- and led to the two changes above; run 2 **converged in one
iteration**: 136 operations applied, 63,174 cells compared / 0 differ, grids
181 → 155, untitled 181 → 75 (55 refused with the reading formula named),
Upload/Convert/Test complete, model run consistent with the audit trail, Mind
listing 51 `Untitled` (8 Mind-only, 32 app-only detection gaps). Output and
report under `runs/shlomo_20260830_converged/`.

124 unit tests (the offline loop test preps through Excel, recalculates both
copies and converges on the local gates alone).

## 1.6.7

### Grid naming: error values are not titles, headings become names, and the assistant can name the rest

Driven by a grid-by-grid audit of the real Shlomo model in Milliman Mind, where
the template list showed `Untitled(60,8)` blocks beside meaningless names like
`Cashflows C4` and `I`.

- **`#N/A` is no longer read as a grid title** (`grids.looks_like_title`, new
  `ERROR_LITERALS`). Excel stores error *values* as text starting with '#', so
  five `#N/A` cells inside `CoverageDetails!A6:CF416` (a 411x84 policy data
  table) were reported as titles "trapped" in the grid (STR-001) -- and the
  default-on `separate_merged_grids` action offered to insert whole rows
  straight through that table to free them. Applying the old plan shattered the
  workbook from 156 grids to 258 in one pass (and to 648 over three). With the
  fix the same plan moves it 156 -> 160. **This was a data-structure-destroying
  default; anyone who ran prep on a model containing error values should
  re-check the result.**
- **Section headings above a block are now used as its name**
  (`prep._section_header_name`). Real models write a section number beside a
  section title on the row directly above the block it heads, so the cell above
  is neither blank nor a loose label and every earlier heuristic missed it. The
  Shlomo model now names blocks "Demographic Assumptions", "Expense Cashflow
  Calculation", "All Cashflows - Direct" ... instead of "Cashflows C4".
- **A grid whose cell above is occupied is no longer skipped**: the title
  action inserts a row (only when no other grid spans it) and writes the title
  there, the way it already did for a grid starting on row 1. Unnamed grids in
  the Shlomo model drop 60 -> 39 in a single pass with no fragmentation.
  Post-insert coordinates are resolved centrally in `prep._ordered`
  (`_resolve_insert_at`), so a title still lands correctly when other actions
  in the same approved set insert rows above it.
- **Context-aware names from the APIM assistant** (`app/grid_naming.py`, new):
  `build_grid_context` collects the cells above, left and inside a grid (with
  formulas shown as `formula => computed value`), and `suggest_names` asks for
  short names in one batched call per 25 grids. `POST /api/sessions/{id}/grid-names`
  returns deterministic-vs-suggested for review and folds the accepted names
  into the prep plan. Names are advisory and never bypass review: they only
  change *which* name `create_grid_titles` proposes, never where it writes or
  whether the edit is safe, and they are used only where the deterministic name
  is weak (`prep.is_weak_name`). With the gateway unavailable the deterministic
  name stands.

- **Runbook + CLI**: `docs/RUN_IN_MIND_NAMING.md` records how to reproduce this
  audit by hand, and `scripts/audit_grid_names.py` is the whole read-only audit in
  one command (`--context` dumps the cells around each unnamed grid, `--suggest`
  adds the assistant's names).

Rule set unchanged: 105 rules, 96 active. 98 unit tests.

## 1.6.6

### Broken-#REF! handling, FILTER supported, value reconciliation (from the real-Mind Shlomo loop)

Found by an autonomous loop that uploaded a real model (Shlomo_IFRS) to live
Milliman Mind, read what failed, and fixed the app until Mind converts it.

- **REF-001** (`validators.structure.broken_defined_name_refs`): flags cells that
  reference a broken (#REF!) defined name OR contain a literal #REF! -- Mind's
  converter cannot compile these ("not a function" / "Formula compilation error").
  A new **opt-in** prep action ("Replace broken (#REF!) references with NA()")
  rewrites every standalone broken reference to `NA()`: an error stays an error, so
  no valid result changes (IFERROR fallbacks and live branches keep their value);
  array formulas are handled; a range-endpoint #REF! is left for review.
- **FILTER is Mind-supported** (`validators.formula.MIND_CONFIRMED_SUPPORTED`): the
  KB "Supported Excel formulas" scrape omits FILTER but real Mind conversion accepts
  it, so FRM-002 no longer flags it. Extensible as more functions are confirmed.
- **REP-001** (`validators.structure.totals_reconcile`): value reconciliation --
  flags any total (`=SUM(range)` or `=a+b+c`) whose value does not equal the sum of
  its addends. Report-only (never changes numbers).

Rule set: 105 rules, 96 active (REF-001 + REP-001 added). 87 unit tests. Confirmed
end-to-end: the app's fixes make the raw Shlomo model convert cleanly in real Mind.

## 1.6.5

### Aligned with Mind's converter: reject /ProjectSettings setting names Mind doesn't recognise

Driven by the real-Mind closed loop: uploading a workbook to Milliman Mind and
reading its verdict showed that declaring an `EnableCheckFormatInputManager` or
`EnableDebugMode` setting in a `/ProjectSettings` grid makes Mind's **Convert**
step fail ("The <name> setting does not exist") -- yet the app was *recommending*
the former (INP-006) and *crediting* the latter (PRJ-005). A minimal workbook
without them converts cleanly.

- **New rule PRJ-006** (`validators.project.prj_006`): flags any `/ProjectSettings`
  setting name Mind's converter rejects (`MIND_INVALID_SETTINGS`, seeded with the
  two confirmed names) as ERROR -- "uploading this workbook fails the Convert step".
- **INP-006** no longer tells users to add an `EnableCheckFormatInputManager`
  setting; it explains that format checking is toggled per grid in Mind's Input
  Manager at import time (not a workbook setting) and points to PRJ-006.
- **PRJ-005** no longer credits a Mind-invalid name as a configured debug setting.
- Fixtures: the well-formed fixture drops the two invalid settings; the broken
  fixture declares one so PRJ-006 is exercised.

Rule set: 103 rules, **94 active** (PRJ-006 added; FORMULA-001 was removed in 1.6.4).
Also confirmed against real Mind: the upload accepts **XLSX/XLSM only** (not .xlsb).

## 1.6.4

### Hidden readiness sheets, a workbook view for every error, FORMULA-001 gone, Download in the sidebar

- **Readiness sheets are hidden.** `Mind_Readiness_Report` and
  `..._Summary` in the "with report" copy are written as hidden sheets (both
  the Excel and the openpyxl writer; a visible sheet stays active), so the
  copy can be uploaded to Mind without the report being read as data.
- **What is going on in the workbook.** Every Fix panel (finding sites,
  recalculation errors, root-cause groups) shows an *In the workbook* grid
  around the cell -- the cells' contents as Excel last stored them, with
  the error cell highlighted -- and a **Show formulas** toggle that swaps
  in the formulas (`{...}` marks an array formula; members show the
  anchor's formula). Click a cell in the list to move the view. Backend:
  `GET /api/sessions/{id}/cells?sheet=&cell=&rows=&cols=`
  (`inventory.cell_window`); values come from the last recalculation of the
  current version when there is one, else from the analysis copy.
- **FORMULA-001 removed** ("a target version is configured but no
  per-version compatibility matrix is implemented"): it could never be
  anything but NOT_SUPPORTED and only ever blocked the overall status.
  Rule set: 102 rules, 93 active (function support is still checked by
  FRM-002 against the KB list).
- **Download area** in the sidebar: the latest version (label, verified
  mark, file name) with a one-click download; a compact icon when the
  sidebar is collapsed.

Tests: 84 (was 83): the cells endpoint (contents, formulas, array members
after a recalculation) and hidden-sheet assertions on both report writers.

## 1.6.3

### Every error can be fixed; errors with one root cause are fixed together

- Recalculation errors are **grouped by root cause** (`app/recalc.py::
  group_errors`): `#NAME?` from the same unknown function(s); any error from
  formulas of the same *shape* (`formula_signature`: references -> REF,
  numbers -> N, strings masked); add-in-gap `#NAME?` by the MM_ function(s)
  involved. The recalculation result carries `groups` (biggest first).
- Recalculate screen: a "Shared root causes -- fix all at once" section
  with one **Fix all N** button per group; every formula-error row has a
  **Fix** button plus a "shared xN" badge that opens the group; every
  add-in-gap line has **Ask / Fix**. Nothing is left without a fix option.
- Fix panel: a **group mode** (all cells listed, "Propose one fix for all N
  cells", quick fixes = planned changes touching any of the cells, apply,
  then "Recalculate now" reporting how many of the N still fail); the
  single-error mode shows "Same root cause in K other cells -- Fix all N"
  and switches to the group; a **Propose a fix** button is always present
  (findings with several sites ask for one fix covering all sites; add-in
  gaps ask the assistant to confirm none is needed).
- Assistant: a group focus tells the model the cells share one cause and
  asks for operations that fix EVERY listed cell (one per cell, or a range
  when the same formula applies) and to say so; a single-error focus lists
  its siblings so the answer says they can all be fixed together; the
  `<recalculation>` context now starts with the root causes.

- **Array formulas are changed as a unit.** Excel refuses to change part of
  a Ctrl+Shift+Enter array, and the old executor silently rewrote the *whole*
  array whenever any of its cells got `set_formula`. Now: recalculation
  errors carry `array` (e.g. `C7:C9`) and the group cause names it; when the
  operations cover every cell of the array it is dismantled first and each
  cell gets exactly the proposed content; a fix that covers only part of it
  fails with a clear message (nothing half-applied) unless it addresses the
  top-left cell alone, which replaces the whole array formula and is
  recorded as such; a new assistant operation `set_array_formula`
  `{sheet, range, formula}` writes one array over a range. The assistant is
  told which cells form an array (including the members that are not in
  error) and to fix the whole of it; `validate_proposal` refuses any
  operation that covers only part of a known array, naming the uncovered
  cells, and a rejected proposal is sent back to the model once with the
  validation errors so it can repair it (`proposal_retries`). Validation
  also refuses **no-op operations** (a value/formula the cell already has,
  the same array formula re-entered, clearing an empty cell) so a "fix"
  that changes nothing is bounced back to the model instead of being
  applied and reported as a change. For an array formula the model also gets
  the decisive fact -- the array's size versus the ranges it references --
  with the concrete fix when it overflows its source (`array_size_hint`:
  "set_array_formula on C5:C6 with the same formula and clear_cell C7:C9"),
  both in the focus prefix and in the rejection; up to two repair rounds.
- Fix panel and assistant note now report **failed operations** ("1 change
  could not be applied: Arr!C9 — …") instead of only counting successes;
  COM errors show Excel's own message.

Tests: 83 (was 76): grouping (signature normalisation, unknown-function,
add-in and shape groups, ordering), the group / sibling / array chat focus,
partial-array proposal refusal, the proposal retry, and the Excel-backed
array-unit executor test.

## 1.6.2

### Ask / Fix on recalculation errors

- Recalculate screen: every genuine formula error row has an **Ask / Fix**
  button (the error value itself is clickable too) and every add-in-gap
  `#NAME?` line has **Ask**; both open the same Fix panel used by Findings,
  now with a *recalculation* mode: the error and the formula Excel
  evaluated, what the error means, quick fixes = planned changes that
  target that exact cell, and the mini chat focused on the error
  ("Explain this error", "Propose a fix", "Which cells does this formula
  depend on?"). After applying a change the panel offers **Recalculate
  now** and reports whether the cell still fails; the Recalculate screen
  keeps the latest result (store-backed) and flags it as stale when the
  workbook version moved on.
- Backend: the session remembers the last recalculation
  (`recalc.version_id` tells which version it ran on); `POST .../chat`
  accepts a recalc `focus {kind: "recalc", sheet, cell, error, formula,
  addin_gap}` and always appends the last recalculation (status, genuine
  error cells with formulas, add-in-gap cells) to the assistant's context
  (`app/chat_context.py::recalculation_context`). The system prompt tells
  the assistant to fix genuine errors from the cell's formula and never to
  "fix" an add-in-gap `#NAME?` by rewriting an MM_ formula.
- **Fixed**: a recalculation run after an Excel apply in the same process
  failed with "The interface is unknown" (RPC_S_UNKNOWN_IF) -- the exact
  sequence the Fix panel performs (apply, then "Recalculate now").
  `app/recalc.py` had its own COM lifecycle that released Range/Worksheet
  proxies *after* `CoUninitialize`, poisoning the apartment for the next
  session. It is now built on `app/excel_com.py::excel_session` with every
  proxy scoped inside an inner function (`_scan`), like every other Excel
  path in the app; the stray "Windows fatal exception 0x80010108" noise is
  gone too.
- Recalculation results carry `ran`: false when Excel could not run at all
  (COM failure / pywin32 missing), so an errored run is never mistaken for
  "no errors" -- the Recalculate screen shows it as a failure and the Fix
  panel's "Recalculate now" reports "could not run" instead of "fixed".

Tests: 76 (was 73): recalc-focused chat gets the error prefix, the cell
contents and the `<recalculation>` context; the recalculate endpoint
records the version; regression for recalculate -> Excel apply ->
recalculate in one process (skipped without Excel).

## 1.6.1

### Fixed: a fix did not show up after re-analysis (web app)

The engine was right (re-analysing v2 after an apply dropped the fixed
rules); the web front-end lost the information on the way:
- the Prep apply did not re-analyse, so Findings kept showing the previous
  version until a manual re-analysis;
- a rule that became PASS simply *disappeared* from the table (PASS is
  filtered out by default), so a fixed blocker looked "not fixed";
- a partially fixed rule (e.g. FRM-002 with several call sites) stayed ERROR
  with no sign of progress.

Now every apply (Prep, Assistant, Fix panel) re-analyses the new version
in the same request, and the backend returns a **delta** against the
previous analysis (`fixed` / `improved` / `regressed` rule ids and the
status counts side by side). The Findings screen shows a "Since the
previous analysis" banner, keeps fixed rules visible with a FIXED badge
("show fixed rows"), and the Prep result card says what changed.
Re-analyse always targets the current (latest) version and refreshes the
version list. `correction_available` on each finding now means "the
current prep plan has an operation for it" (it used to reflect a YAML flag
that is true for only three rules). FRM-002 reports up to 25 call sites
per function instead of one.

### Fix panel -- click a finding, fix it there

- Clicking a status pill or the **Fix / Ask** button on any non-PASS row
  of the Findings table, or a grid card in the **Workbook Map**, opens a
  side panel for that finding: rule, status, location, message, every
  `Sheet!Cell` site from the finding's observed data, **quick fixes** (the
  prep-plan actions that cover the rule, applied with one click) and a
  **mini chat** focused on the finding ("Explain this finding", "Propose a
  fix", or free text). Proposals from the mini chat are applied right
  there; after any apply the panel re-reads the new analysis and shows
  "<rule> is now PASS" (FIXED badge) or the remaining status.
- Backend: `POST .../chat` accepts `focus {rule_id, sheet, cell}`, which is
  prefixed to the question so the retrieval pulls that finding's observed
  data and the cell contents into the turn.
- Workbook Map cards now show only the findings located *inside* that grid
  (or on its title cell) -- previously every card on a sheet showed the
  sheet's worst finding -- plus their rule ids; sheet-level findings appear
  as chips above the cards.
- Workbook screen: "Load another workbook" starts a new session.

Verified in a real browser: Fix on PAR-002 -> quick fix -> panel says
"PAR-002 is now PASS" with FIXED badge, banner "fixed 2 -- PAR-002,
PRJ-002", fixed row visible; grid card click opens the panel; mini chat on
TRN-001 explains the finding via the live gateway, proposes the header fix
and applies it -> "TRN-001 is now PASS".

Tests: 73 (was 71): delta + fixability after apply/re-analyse (incl.
honest regression when going back to v1) and finding-focused chat.

## 1.6.0

The Figma-built "Mind Ready" front-end (`Mind Copilot Skill/FigmaOutput`,
React 19 + Vite + Tailwind, generated from docs/FIGMA_UI_PROMPT.md) is now
wired to the engine through a FastAPI backend. The Streamlit UI is unchanged
and still works.

### Backend -- `app/web/server.py`

- One endpoint per function of the front-end's `src/services/api.ts`:
  `POST /api/sessions` (upload + mode, `.xlsb` converted through Excel
  first), `POST /api/sessions/{id}/reanalyze`, `.../apply` (operations ->
  new verified version, optional re-analysis in the same call),
  `.../suggest`, `.../chat`, `.../recalculate`, `.../reports`,
  `GET .../versions`, `GET .../files/{name}`, `GET /api/health`.
- Per-session **version lineage**: v1 original upload (v2 converted for
  `.xlsb`), then one version per apply -- prep, assistant or formula fix --
  each a separate file written by Excel, verified to open, with its change
  log; download names are unique within the session.
- The engine's dicts are mapped onto the TypeScript contracts
  (`WorkbookSummary` built from the inventory: sheets, grids with titles/
  flags/headers/inner titles, standalone text, protection; `ValidationReport`
  and `PrepAction` pass through; `ApplyResult`, `ChatReply` with validated
  proposals, `RecalcResult`, `ReportBuild`, `Version`).
- Serves the built front-end (`FigmaOutput/dist`, or `MIND_READY_DIST`)
  with an SPA fallback, so `run_mind_ready_web.bat` -> http://localhost:8600
  is the whole app. `requirements.txt`: fastapi, uvicorn, python-multipart.

### Front-end wiring (`FigmaOutput/src`)

- `services/api.ts` talks to `/api` (mocks remain behind
  `VITE_USE_MOCKS=true`); `vite.config.ts` proxies `/api` to :8600 in dev.
- Fixes to what only worked with mocks: the upload screen now selects a real
  file and runs on "Run analysis" (it used to post an empty fake file);
  download links use the real session id; "Re-analyze" (top bar and prep
  result card) refreshes summary, report and plan; the prep selection resets
  when a new plan arrives; applying an assistant proposal updates the store
  from the backend's re-analysis and offers the changed file; errors surface
  inline instead of failing silently.

Verified in a real browser against the running server: upload -> findings
(table + workbook map) -> prep apply (Excel, verified) -> download ->
re-analyze -> history lineage -> assistant change via the live gateway
(rename sheet) -> re-analyzed -> reports generated and downloaded.

Tests: 71 (was 65) -- `tests/unit/test_web_api.py` covers health, upload
contracts, mode filtering and input validation, apply + version lineage +
download + re-apply, chat with a validated proposal, reports and
recalculation.

## 1.5.1

- **Fixed**: applying changes a second time (to the file the previous apply
  produced) crashed with `shutil.copy2` -> `CopyFile2` -- the copy helper
  targeted `work_dir/<name>`, which *was* the source. `make_immutable_copy`
  now never copies a file onto itself or over a previous output: when the
  target exists it uses a versioned sub-folder (`v2/`, `v3/`, ...). Every
  earlier output stays intact. Regression test added.
- UI restyled, behaviour unchanged: the upload / mode / run workflow moved
  to a sidebar with a "current file" card (name, status pill, sheet/grid/
  formula counts) and environment notes; status pills and KPI tiles replace
  emoji; sections are a segmented control (still session-keyed, so a long
  Excel apply no longer hides its own result); prep actions, apply results,
  proposals and report downloads sit in bordered cards; the findings table
  has fixed column widths.

## 1.5.0

The assistant can now change the workbook on request, and the grid-title
prep action names grids from their surroundings.

### Ask the assistant -- "tell it to change anything"

- `app/chat_context.py`: the system prompt defines a change protocol. The
  assistant can (a) request cell contents / grid rows / sheet listings /
  function call sites it hasn't seen with a ```lookup block -- the app
  answers within the same turn (up to 2 lookups) -- and (b) end its answer
  with a ```changes block: `set_value`, `set_formula` (single cells or
  ranges; relative references adjust like fill-down), `clear_cell`,
  `rename_sheet`, `insert_row`, `insert_column`, `set_sheet_visibility`,
  `unprotect_sheet`.
- `app/prep.py::validate_proposal` checks every proposed operation against
  the real workbook (sheet exists -- case-insensitive match --, cell/range
  parses, formulas start with '=', sheet names legal and unique, row/column
  valid) and looks up the current cell value for the before/after table;
  malformed operations are reported, never applied.
- UI: the proposal appears under the conversation as a before/after table
  with **Apply** / **Discard**. Apply writes it through Excel to a fresh
  copy, verifies the file opens, **re-analyzes it automatically** (so the
  next question is answered against the changed workbook), notes the
  outcome in the conversation and offers the changed file for download.
  The model's text is never written into the workbook by itself.
- Conversation notes (analysis / apply outcomes) are now part of the
  history the model sees.

### Prep workbook -- context-aware grid titles

- `create_grid_titles` (replaces `title_untitled_grids`, now on by default)
  gives every untitled grid exactly one `#Name` title:
  - **caption above a table**: a text cell whose right neighbour is empty
    while the row below starts a wider block -- Mind (and the app's
    detection) read the caption as a one-column grid that swallows the
    table's first column and the rest of the table as a second grid ("two
    headers for one range"); the caption gets a `#` prefix and the table
    becomes one grid;
  - **empty cell above**: the name comes from a standalone label above
    (blank row between) or to the left -- the label is *moved* into the
    title (cleared afterwards), so no orphan text remains -- else from the
    text header row, else sheet + position;
  - **grid on row 1**: a row is inserted first (only when that cuts no
    other grid), then the title is written at the post-insert coordinate.
  - Names are made unique across the workbook; a grid that already has a
    title, or contains a trapped title (STR-001), is never given another;
    a standalone label that is not next to any grid is reported in the
    action's skip notes rather than touched.
- Executor: new `clear_cell`, `insert_column`, `set_sheet_visibility`
  operations; range targets for value/formula ops; operations flagged
  `after_inserts` run after the row/column inserts (post-insert
  coordinates); openpyxl fallback handles cell ops on ranges and refuses
  the structural ones.

Tests: 64 (was 59): context titles on a fixture with every pattern
(label above, caption, label left, header-row name, row-1 grid) plus an
Excel-backed apply that re-analyzes to all grids titled and STR-004 PASS;
proposal validation (good and bad ops); lookup round-trip and proposal
extraction with a mocked gateway; applying an assistant proposal.

## 1.4.0

Two user-facing features on top of 1.3.0's verified outputs: the app now
**applies real preparation changes** to the workbook (on approval) and has a
**grounded, multi-turn assistant** for questions about the workbook.

### Prep workbook -- actual changes, reviewed cell by cell

- `app/prep.py` (new): `plan_actions()` turns findings into concrete
  *operations* (op, sheet, cell/row, before, after, note) without touching
  any file; `apply_operations()` copies the source (hash first), applies the
  approved operations in **one Excel COM session**, saves, re-opens the
  result in Excel to verify it loads, and appends the change-log sidecar.
  openpyxl remains a fallback for value/formula operations only -- it
  refuses structural ones (renames, row inserts, colours, protection)
  because it cannot keep references/styles intact.
- Actions (rule -> change; default-on unless noted):
  - STR-007: hidden sheets renamed with the `&&Hide` marker (Excel updates
    every formula reference on rename).
  - LOOP-002 + RES-002: loop-name capitalisation normalised in MM_LOOP,
    MM_RESULT, MM_DIMSIZE, MM_DIMINDEX, MM_LOOPLABELS string arguments.
  - INP-005: `/Input` added to `/Reorder` grid titles.
  - FLG-001: undocumented flags that are an unambiguous near-miss of a
    documented one corrected (`/Inpt` -> `/Input`); ambiguous ones skipped
    with a reason.
  - EXP-003 / PRJ-002 / PAR-002 / INP-003: near-miss headers of the special
    grids corrected to the documented column names (`Kind` -> `Type`,
    `Setting` -> `Name`).
  - STR-001: an empty row inserted above a `#Title` trapped inside a grid --
    only when no other grid on the sheet spans that row (otherwise skipped
    with the reason, never a blind insert).
  - FMT-002: theme colours rewritten as explicit RGB (same look).
  - Off by default (they change cell content or security): STR-004 generic
    `#Grid_<sheet>_<anchor>` titles for untitled grids, FMT-003 apostrophe +
    space in styled empty cells (KB recipe), FMT-005 sheet unprotect.
- **Replace an incompatible formula** (FRM-002 / RSK-004 / FORMULA-002):
  pick a flagged cell, optionally ask for an AI suggestion (the suggested
  `=...` line pre-fills the box), edit the replacement yourself, apply. The
  model's text is never written on its own -- only what you approve
  (SKILL.md rule #6 still holds: formulas are never converted to values).
- Every apply offers the prepared file for download and a one-click
  re-analysis of it; the change log records method, verification and every
  applied/failed operation.
- Inventory now keeps the coordinates of theme-coloured and styled-empty
  cells (capped at 5,000 per sheet) so those actions can target exact cells.

### Ask about this workbook -- grounded assistant with memory

- `app/chat_context.py` (new): each turn sends a bounded *context pack*
  (workbook facts, every sheet with its grids/flags/headers, formula and MM_
  statistics, every non-PASS finding in full, the PASS rule ids, and the
  prep actions the app can apply) plus *retrieved detail* for whatever the
  question names -- cell references (formula/value there, enclosing grid),
  sheet and grid names (their rows), rule ids (the finding's observed
  data), MM_ function names (call sites) -- and the conversation so far
  (last 12 turns). The system prompt forbids inventing facts, forbids
  claiming readiness, and routes fixes to "Prep workbook".
- `app/llm.py`: generic `chat_completion()` over the same APIM Claude
  gateway (system as leading message -- the proven payload shape);
  `suggest_formula_fix` now uses it and asks for the replacement formula on
  its own `=` line; `extract_formula()` picks it out for the UI.

### UI

- Sections: Findings / Prep workbook / Ask the assistant / Recalculate /
  Reports, switched with a session-keyed radio rather than `st.tabs` --
  tabs snap back to the first tab after every rerun, which hid the result
  of a long Excel apply (caught by the Playwright check). The old
  per-finding "Apply fix" buttons are superseded by the prep plan
  (FILE-002 conversion still happens automatically for `.xlsb`).
- Verified with Playwright on the broken synthetic model: 7 prep changes
  applied through Excel and verified, prepared file downloaded with the
  corrections in place, assistant answered a grounded question and routed
  the fix to "Prep workbook".

Tests: 59 (was 50) -- prep planning on the well-formed/broken/messy
fixtures, openpyxl apply with re-validation, structural ops refused without
Excel, one Excel-backed apply (rename + unprotect, verified), context pack
and retrieval grounding, chat history threading and gateway payload shape
(mocked).

## 1.3.0

Two things drove this release: hands-on proof that the app hands back a
*working* Excel file, and implementing every rule against the Milliman Mind
knowledge base (`CompleteMindDocn.docx`, a page-by-page scrape of
kb.milliman-mind.com, 431 pages) instead of leaving 65 of 86 active rules
at `NOT_SUPPORTED`.

### Output workbooks are written by Excel, and verified

- **Bug found and fixed**: the 1.2.0 report workbook did not open in Excel
  when built from a real Mind workbook. `openpyxl` silently drops package
  parts it doesn't model -- the test workbook carries a 13 MB Power Pivot
  data model (`xl/model/item.data`) and 20 `customXml` parts -- and Excel
  refused the result ("Open method of Workbooks class failed"), for `.xlsm`
  and `.xlsx` alike. The output had shrunk from 13.7 MB to 0.5 MB.
- `app/excel_com.py` (new): one shared Excel COM session helper (hidden,
  alerts off, macros force-disabled, links never updated) plus
  `verify_opens_in_excel()`, which is now the app's definition of "a real
  working Excel file".
- `app/excel_report.py`: `build_report_workbook` copies the analyzed file and
  lets **Excel itself** append the `Mind_Readiness_Report` /
  `_Summary` sheets (`Range.Value2` bulk write, formatting via COM), then
  re-opens the result in Excel. Returns a `ReportBuildResult` (path,
  method, `verified_opens_in_excel`, warnings). The openpyxl path is kept
  only as a fallback for machines without Excel and is labelled reduced
  fidelity. New `build_standalone_report()` writes a fresh Summary +
  Findings `.xlsx` that is valid regardless of the source workbook.
- `app/change_apply.py`: `fix_loop_case_mismatch` now plans its edits from
  the analysis (`plan_loop_case_fix`) and writes them through Excel
  (`Range.Formula` / `FormulaArray`), verifying the output opens; the change
  log records `method` and `verified_opens_in_excel`. openpyxl fallback as
  above. `convert_output_format` uses the shared session.
- UI: "Generate reports" produces both downloads (standalone report always;
  workbook copy + report only offered when Excel re-opened it) with a
  visible verification line; fixed-workbook downloads show method and
  verification; proper MIME types. Verified end-to-end with Playwright on
  the real 13.7 MB `.xlsm`: uploaded, analyzed, both files downloaded, both
  re-opened by Excel COM, all package parts intact (13,742,301 bytes).
- COM proxies are scoped inside inner functions and collected before
  `Excel.Quit()` (a proxy collected afterwards raises RPC_E_DISCONNECTED in
  the GC, which faulthandler prints as a "Windows fatal exception").

### Every active rule now has a real validator (94 of 94)

- `tools/mine_kb_docx.py` (new, deterministic, zipfile + regex, no
  python-docx): mines the KB scrape into
  `references/mm-function-registry-kb.yaml` (117 MM_ functions, 100 with a
  dedicated article: syntax, category, description), `references/
  supported-excel-functions.yaml` (the KB's "Supported Excel formulas" list,
  217 native functions -- the source the README said was missing) and
  `references/mind-flags.yaml` (49 documented grid/header flags with their
  argument shape and uniqueness, 6 markers; every entry is verified to occur
  in the KB text or the script fails). Scrape artefacts (hyphenated /
  concatenated tokens) are repaired.
- `app/grids.py` (new): Mind's documented grid detection ("General
  structure and guidelines": scan left-to-right/top-to-bottom, extend right
  and down to the first empty cell, ignore alone text cells, text-only first
  row = headers, `#Name /Flag1 /Flag2` title cell above the grid), with
  `/Flag.(arg).x.y` parsing and header markers (`/HideRows`,
  `/InstanceSelect`, `&&HideColumn`, `&hide`).
- `app/formula_utils.py`: string literals and quoted sheet names are masked
  before tokenizing; OOXML storage prefixes are stripped (`_xlfn.`,
  `_xludf.`, `_xlpm.`, and `_xll.` -- Excel stores MMForExcel calls as
  `_xll.MM_LOOP`, seen on the real workbook and previously a false
  FORMULA-002 error); A1-reference parsing (`cell_refs_in_formula`,
  `parse_ref`), literal helpers, call spans.
- `app/inventory.py`: single-pass inventory now also captures array
  formulas (openpyxl returns `ArrayFormula` objects, not `=` strings -- 103
  in the real workbook were invisible to every formula check before),
  data-table formulas, number-format categories vs the KB's preserved list,
  theme colours, styled-empty cells, outlines/hidden rows+columns, data
  validations, conditional formatting, sheet protection, tab colour,
  `docProps/app.xml` Application, VBA part size, package parts only Excel
  can preserve; plus a process-wide cache of the opened workbook
  (`open_cached` / `cell_value`) for validators that read raw cells.
  Real workbook: 3.6 s inventory, 1.3 s for all 94 validators.
- **Fixed**: locked-cell counting (FMT-005) counted every cell, because
  Excel locks cells by default; per the KB, locking only bites on a
  *protected* sheet, so only those are counted now.
- New validator modules: `iteration.py` (CAL-001..005: /CalculationSteps,
  MM_ITERATIONS once per workbook, /iterationinput+/iterationoutput
  pairing), `performance.py` (DBG-001..006: whole-column / >50k-cell
  ranges, lookup review, profiler/debug/F9 recommendations), `io.py`
  (INP-001..006, EXP-001..004: /Input naming, /InputSettings columns and
  `{GridName}`-style patterns, /Reorder prerequisites, format-check
  setting, /Export, unique /ExportSettings with documented columns,
  booleans, extensions and `{yyyyMMdd}`-style tags), `lookup.py`
  (LKP-001..005: MM_READTABLE range starts on a header row, Header names
  exactly one column, duplicate key combinations, ≤6 comparisons,
  not-found text vs NaN), `environment.py` (MMX-001..003: add-in
  requirement and `_xludf.`/`_xll.` prefixes, saving application from
  app.xml, MM_ function classification), `parameter.py` (PAR-001..004:
  Label|Type|PossibleValues|Values, the nine types, `min|max|step` and
  pipe-list syntax), `project.py` (PRJ-001..005: unique /ProjectSettings,
  Name|Value|Locked|Hidden, hidden settings, stochastic configuration incl.
  `#NbSimulations`, debug settings), `kb.py` (see new rules below).
- Extended validators: `structure.py` (STR-001 merged-grid detection via a
  trapped `#Title`, STR-002 alone text cells, STR-003 header recognition
  incl. ERROR on header-matched flagged grids, STR-004 untitled grids,
  STR-005/006 default sheet/workbook names, GRID-004 on real flags),
  `format.py` (FMT-001 number formats outside Standard/Number/Text/Boolean/
  Date, FMT-002 theme colours, FMT-003 styled empty cells, FMT-006
  outlines, FORMAT-001 preservation inventory), `formula.py` (FRM-001
  inventory; **FRM-002 now checks native functions against the KB list**
  and MM_ functions against both registries -- MM_LOOPLABELS is no longer
  a false error; FRM-003 `@` implicit intersection, spill references,
  array formulas; FORMULA-002 VBA UDF detection with a read-only search of
  `vbaProject.bin` for the name), `loop.py` (LOOP-004 MM_LOOPINSTANCE;
  RES-001..005 MM_RESULT: loop names must exist with exact case, no loop
  named twice, omitted dimensions are summed, SIM usage; RZS-001..007:
  MM_SETSIZE `+` syntax / nesting / SetSize-of-SetSize, destination block
  empty or holding the add-in's own copies, resize flags need a driver,
  named groups need two members, fixed ranges that stop on a resizable
  grid's edge, MM_GETRANGE), `risk.py` (RSK-001 two resize drivers in one
  grid, RSK-002 array formulas past the grid boundary).
- `rules/kb-rules.yaml` (new, category `kb`, 8 rules the readiness
  standard doesn't have but the KB states): FLG-001 documented flags with
  argument shape, GRP-001 `/Group.(Name).x.y`, BKP-001 backup
  source/destination pairing + MM_BACKUPBUTTON placement, HID-001
  /HideRows values, RNG-001 MM_RANGE alone in a single-cell table, UNQ-001
  unique special grids, TRN-001 ISO 639-1 translation headers, SIM-001
  `#NbSimulations` content. Rule set 1.2.0: 103 rules, 94 active, 9 draft
  (INS-*/LNK-* multi-workbook, unchanged).
- `app/validators/excel_functions.py`: catalog of native Excel function
  names, used only to tell "native but not Mind-supported" (FRM-002) from
  "not an Excel function at all" (FORMULA-002).
- Report: `FIX_HINTS` for every new rule; `config/default.yaml`
  `large_range_cells: 50000` (DBG-003).
- Tests: 50 unit tests (was 30). `tests/kb_fixtures.py` builds a
  well-formed Mind-style model (every KB-backed validator PASSes, or
  WARNINGs where the KB only asks for review) and a deliberately broken one
  (each validator fails on its own defect), plus array-formula / `@` /
  XLOOKUP / UDF fixtures; pure tests for grid detection and the tokenizer;
  the miner is re-run against the .docx when it is present and must
  reproduce the checked-in files; two Excel-backed tests (skipped without
  Excel) prove the written files open in Excel.

Honesty notes: grid boundaries are this app's implementation of the
documented algorithm, so findings that depend on them carry `INFERENCE`
evidence; the KB's supported-function page is version-agnostic, so
FORMULA-001 (per-version matrix) remains `NOT_SUPPORTED`; READY-001 still
requires an explicit recalculation for PASS; nothing new is auto-applied.

## 1.2.0

UI, Excel-format report output, and real auto-fix/recalculation, in response
to hands-on testing of 1.1.0 against a real workbook. Scope note: the mined
rule set (rules/*.yaml) marks almost nothing `correction.automatic: true` by
its own design (SKILL.md non-negotiable rules #4/#6) -- "auto-fix" here means
genuinely automatic where a rule says it's safe, plus a one-click
human-approved apply for everything else, not blind automatic correction of
arbitrary findings.

- `config/target-version.yaml`: `target_mind_version` defaults to `latest`
  (was `null`). `FORMULA-001` no longer asks `REQUIRES_USER_INPUT` for a
  missing version; it reports `NOT_SUPPORTED` (no compatibility matrix
  exists yet, but a version *is* configured).
- `app/excel_report.py` + `scripts/generate_report_workbook.py`: the primary
  human-facing output is now an Excel workbook (a copy of the analyzed file
  plus an appended `Mind_Readiness_Report` sheet: formatted findings table
  with a deterministic "Recommended Fix" column, plus a summary sheet).
  JSON (`validate_workbook.py`) remains available for scripting.
- `app/recalc.py` (phase 9, real): Excel COM automation (win32com), macros
  force-disabled (`AutomationSecurity`), scans for genuine formula errors
  after `CalculateFullRebuild`. **Important finding from testing on this
  machine**: no MMForExcel add-in is registered, so every `MM_` function call
  shows `#NAME?` when recalculated here -- the adapter detects this (checks
  COMAddIns/AddIns) and reports it as `NOT_SUPPORTED` with an explanation,
  never as a false `ERROR`, separately from genuine formula errors.
  Recalculation is a deliberate, separate action (`scripts/recalculate_workbook.py`,
  the UI's Recalculate button) -- not run automatically during every
  ANALYZE/VALIDATE pass (`modes/plan-mode.md`: "recalculation is optional").
- `app/change_apply.py` (phase 8, narrow real scope): `convert_output_format`
  (Excel COM Save-As, also used to auto-convert an uploaded `.xlsb` before
  analysis -- openpyxl can't read `.xlsb` at all) and
  `fix_loop_case_mismatch` (rewrites a case-inconsistent loop name to its
  most-used spelling across every `MM_`-function string argument that uses
  it). Every apply works on a fresh copy and writes an append-only JSON
  change-log sidecar (`<file>.changelog.json`), never into the workbook.
- `app/llm.py` (phase 7, narrow real scope): one on-demand "suggest a fix"
  call for formula findings (Claude via the shared Milliman APIM gateway,
  same `secret.key`/`config.enc` discovery convention as
  `reserve_narrator/utils/key_loader.py`). Recommendation only, tagged
  `RECOMMENDATION`, never auto-applied.
- `app/ui/streamlit_app.py` + `run_excel_upload_prep.bat`: local web UI --
  upload, run a mode, review findings, apply/suggest fixes per finding,
  recalculate, download the Excel report. Verified end-to-end with a
  real browser (Playwright) against this machine's Streamlit process,
  including a real Claude-via-APIM call and a real Excel COM recalculation.
- `app/validators/formula.py::unsupported_functions` (FRM-002) now locates
  the actual sheet/cell/formula for each unregistered `MM_` call (previously
  only the function name), so the report's Location column and the LLM
  suggestion have real context instead of a bare name.
- `tests/unit/test_excel_report.py`, `tests/unit/test_change_apply.py`: new.
  `app/recalc.py`/`convert_output_format`'s Excel COM paths are not
  unit-tested (require driving real Excel) -- verified manually/via
  Playwright instead.

## 1.1.0

MVP implementation of CLAUDE_CODE_PROMPT.md phases 1-6 (models/schemas, rule
registry, workbook inventory, deterministic validators, mode workflows).
Phases 7-9 (LLM abstraction, change-set application, recalculation adapter)
are **not** built yet -- `READY-001` always reports `NOT_SUPPORTED`, so
`PASS` is unreachable in this version by design, per non-negotiable rule #5.

- Mined `01_Mind_Readiness_Standard.md` (sibling doc, ~90 rules) into
  `rules/*.yaml` via `tools/mine_readiness_standard.py`, and
  `02_MM_Function_Registry.md` (35 MM functions) into
  `references/mm-function-registry.yaml` via `tools/mine_function_registry.py`.
  Both scripts are deterministic parsers, safe to re-run against updated
  source docs.
- **ID collision resolved**: the old hand-authored `loop-rules.yaml`
  (`LOOP-001..003`) overlapped with the doc's own `LOOP-001..004`. The doc's
  IDs are now canonical; the one rule with no doc equivalent (MM_LOOPLABELS
  size matching) was kept, renamed to `LBL-001`.
- **Category strings normalized** so every rule's `category` matches the
  middle segment of its own `validation.implementation` path (`FORMAT-001`
  was `formatting`, `GRID-001`/`GRID-004` were `grid`, `READY-001` was
  `readiness` -- all fixed; a regression test guards this).
- New `app/` package: domain models (`models.py`), rule loader
  (`rules_engine.py`), workbook inventory (`inventory.py`, openpyxl +
  zipfile, never executes VBA/links/OLE), preservation/package-diff
  (`preservation.py`), readiness aggregation (`report.py`,
  `NOT_SUPPORTED > ERROR > REQUIRES_USER_INPUT > WARNING > PASS`), and
  validators/modes packages.
- Implemented validators (everything else in the mined rule set is
  registered but reports its own `on_unevaluable` status, never silently
  skipped): `FILE-001/002`, `STR-001/004/007`, `GRID-001/004`,
  `FMT-004/005/007`, `FRM-002/004`, `FORMULA-001`, `LOOP-001/002/003`,
  `LBL-001`, `RSK-003/004/005`, `READY-001` (always `NOT_SUPPORTED`).
- `INS-*`/`LNK-*` (instance/interlink -- inherently cross-workbook) mined as
  `status: draft`, since `config/default.yaml` has
  `multiple_workbook_context.enabled: false`.
- `scripts/inventory_workbook.py`, `scripts/validate_workbook.py`,
  `scripts/compare_packages.py` are now real (schema-valid JSON output).
  `scripts/apply_changes.py`, `scripts/recalculate_workbook.py`,
  `scripts/verify_recalculation.py` remain honest `NOT_SUPPORTED` stubs
  (phases 8/9, deferred) but now emit schema-correct shapes.
- `tests/unit/`: rule-schema conformance, inventory correctness, validator
  PASS/FAIL behavior on synthetic fixtures (`tests/conftest.py`, built with
  openpyxl at test time, clearly labeled per `tests/README.md`), and a full
  `PLAN_MODE` run validated against `schemas/processing-result.schema.json`.

## 1.0.0

- Initial architecture package.
