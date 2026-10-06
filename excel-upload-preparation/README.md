# Excel Upload Preparation

Provider-neutral, Python-oriented instruction and rule package for preparing Excel workbooks for Milliman Mind upload.

## Modes

- Plan mode
- Prep Mind Loops
- Fix Incompatible Formulas
- Structure Fix

## Core guarantees

- Never modifies the source workbook.
- Produces a new `.xlsx` or `.xlsm` copy.
- Preserves workbook content where technically possible.
- Stores the change log in the application, not in the workbook.
- Requires successful recalculation before upload readiness can be `PASS`.
- Keeps LLM reasoning separate from deterministic workbook processing.

## Statuses

`PASS`, `WARNING`, `ERROR`, `REQUIRES_USER_INPUT`, `NOT_SUPPORTED`.

## Current status (1.8.1)

All 12 phases have a real implementation, and **every active rule has a
real validator** -- nothing falls through to `NOT_SUPPORTED` any more except
the two honest cases below.

- **Same numbers in Mind as in Excel** (1.8.1): two rules for formulas that
  recalculate clean in Excel and still come out different in Mind -- found by
  uploading Horizon, each confirmed in Mind with a small test workbook
  (`docs/MIND_VS_EXCEL.md`). **FRM-008**: arithmetic on a cell that holds `""`
  inside IFERROR (Excel returns the fallback, Mind computes a number).
  **FRM-009**: a number compared with an empty cell, directly or through a
  reference / VLOOKUP that lands on one (0 in Excel, not 0 in Mind). Each has
  a Prep action that rewrites the formulas without changing Excel's values;
  after both, Mind reported no difference for Horizon.

- **Fix all automatically** (1.8.0): one button on the Recalculate screen and
  the app goes through the formula errors by itself (`app/autofix.py`,
  `POST /api/sessions/{id}/auto-fix`). In one Excel session it finds the
  cells where errors *start*, gives each a value to fall back on
  (`=IFERROR(<formula>, 0)` -- the assistant says which value fits and why),
  recalculates, and **undoes any fix that changed a value that was good**;
  what cannot be fixed that way stays, listed for a person with the reason.
  The manual Fix panel is unchanged. On the four test models: Horizon (721
  errors) and Palermo (511) come out clean with no good value changed, CNHI
  has no error to begin with; PVFP (65,883 errors from a dead external link)
  goes down to the 887 its own formulas hide on purpose, and comes out clean
  when the owner's rule is lifted for it -- see `docs/AUTOFIX_USE_CASES.md`.

- **The numbers check** (1.8.0): *did the preparation change what the model
  computes?* `app/numbers_check.py` recalculates the original and the current
  version and compares every formula cell through the rows Prep inserted
  (`POST /api/sessions/{id}/numbers-check`, "Check the numbers" on the Prep
  screen, `scripts/compare_versions.py`). It exists because an ordinary Prep
  changed 2,462 computed values on one test model: title rows inserted on
  sheets that a 3-D reference (`=SUM('LoB 1:>>'!K65)`) or an INDIRECT address
  reads by position. Prep now refuses those inserts and says why.

- **Theme colours in seconds** (1.8.0): "Replace theme colours with explicit
  RGB" (FMT-002) is done in the workbook's style table (`app/theme_colors.py`)
  instead of cell by cell in Excel. It used to take an hour for one sheet of a
  large model and come back at every Prep round; it now takes under ten
  seconds for the whole workbook, once, with the exact RGB Excel shows.

- **Recalculate on a model full of errors** (1.8.0): the recalculation no
  longer asks Excel about each error cell (three calls a cell, twice). The
  first Recalculate of a model with 65,883 error cells took about 25 minutes;
  it takes a few minutes now.

- **One verdict** (1.7.2): every analysis, apply and recalculation answers
  *is this version acceptable by Mind?* -- `blocked` (a REQUIRED rule fails),
  `unverified` (nothing blocks, no clean Excel recalculation of this version
  yet) or `ready` -- with what stands in the way and the next step
  (`app/readiness.py`, `GET /api/sessions/{id}/readiness`). Findings carry
  their rule's priority and Prep actions a `level`: **blocking** (Mind
  refuses the file without it) or **optional** (Mind reads the file as it
  is). A Prep gauge tracks the planned repairs per version and flags a
  stalled run. See CHANGELOG 1.7.2 for the two endless-round causes fixed.

- **Assistant speed** (1.7.2): the assistant runs on `claude-sonnet-5`
  (env `MIND_READY_MODEL` overrides) at `effort: low` (env
  `MIND_READY_EFFORT`; Sonnet 5's default thinking made answers slow and cut
  them off) and its replies stream -- the web app
  shows the text as it is written, with a status line during lookups and
  proposal repairs (`POST /api/sessions/{id}/chat/stream`, NDJSON;
  `MIND_READY_STREAM=0` turns streaming off).

- **Large workbooks** (1.6.6): every upload is first inspected from the zip
  package alone (`app/sizing.py` -- size on disk, *decompressed* size, and
  every sheet with the unpacked size of its part; sheet names are read from
  `xl/workbook.xml` or, for an `.xlsb`, from the binary `xl/workbook.bin`).
  Above `upload_size_threshold_mb` (25, config/default.yaml; env
  `MIND_READY_SIZE_THRESHOLD_MB`) the app asks whether some sheets should be
  ignored before it scans anything; ignored sheets are never parsed
  (`inventory.load_workbook_selective`), appear in no finding, and stay in
  every output untouched. While a scan runs, a **status indicator** shows
  the stage (convert / copy / read sheet i of N / scan sheet i of N / names /
  rule i of M / report / plan), overall progress and elapsed time -- web app
  via `GET /api/sessions/{id}/status`, Streamlit via `st.status`.

- **Analysis & validation**: `app/inventory.py` (single-pass, openpyxl +
  zipfile, no execution of VBA/links) + `app/grids.py` (Mind's documented
  grid detection and `#Name /Flag` parsing) + `app/validators/` (structure,
  format, formula, loop/resize/result, risk, lookup, io, parameter,
  project, iteration, performance, environment, kb) + `app/rules_engine.py`.
  Rule set 1.2.0: 105 rules, 96 active, 9 draft (INS-*/LNK-* multi-workbook); FORMULA-001 removed in 1.6.4, PRJ-006 added in 1.6.5, REF-001 (broken #REF! references) + REP-001 (value reconciliation) added in 1.6.6.
  Grid titles are read with `grids.looks_like_title()`, which excludes Excel
  error values (`#N/A`, `#REF!`, ...) -- they begin with '#' but are results,
  not titles (1.6.7).
- **Grid naming**: `app/prep.py` names an untitled grid from the cells around
  it -- including the *section heading directly above it*, the layout real
  models use -- and inserts a row when the cell above is occupied, so a block
  with a heading still gets a title. When the surroundings say nothing the
  heuristic falls back to `"<Sheet> <Anchor>"`; `app/grid_naming.py` can then
  ask the APIM assistant for a context-aware name instead (one batched call,
  `POST /api/sessions/{id}/grid-names`). Assistant names are advisory: they
  only change *which* name a reviewable prep operation writes (1.6.7).
- **Grid Namer** (1.7.0): a screen where the *user* names grids by hand --
  the whole current version is shown sheet by sheet with the detected grids
  outlined; drag over an area, type a name, tick documented flags, Submit.
  `plan_named_areas` (app/prep.py) resolves each selection to the grid Mind's
  detection sees there (or the caption-above-a-table pair), writes the
  `#Name /Flags` title under the same reference-safety rules, and the
  automatic titler names everything else in the same apply without colliding
  (user names reserved, user grids excluded, row inserts shared).
- **Runs itself in Milliman Mind** (1.6.8): `app/mind_loop.py` /
  `scripts/run_in_mind.py` / the **Mind** screen take a raw model and, with
  nobody in the loop, prepare → verify the numbers (full Excel recalculation
  of source and prepared copy, compared cell for cell) → upload → convert →
  run → read Mind's verdict and template list → change one thing → repeat,
  until every gate passes or it reports exactly why it is stuck. Real Mind is
  driven through `app/mind_client.py` (Playwright + Edge, sandbox projects
  only). Runbook: `docs/RUN_IN_MIND_LOOP.md` (short form in
  `docs/RUN_IN_MIND.md`, Part D).
- **Knowledge-base backing**: `tools/mine_kb_docx.py` mines
  `CompleteMindDocn.docx` (the kb.milliman-mind.com scrape) into
  `references/mm-function-registry-kb.yaml` (117 MM_ functions),
  `references/supported-excel-functions.yaml` (217 native functions -- the
  list FRM-002 now checks against) and `references/mind-flags.yaml` (49
  documented flags). Re-run it against an updated document; never hand-edit.
- **Outputs are real Excel files** (this is what 1.3.0 is about):
  `scripts/generate_report_workbook.py` / `app/excel_report.py` write the
  workbook copy + `Mind_Readiness_Report` sheets **through Excel (COM)** and
  re-open the result in Excel to verify it loads. A pure-Python save of a
  real Mind workbook (13 MB Power Pivot data model, customXml parts)
  produced a file Excel refused to open -- so openpyxl is only a labelled
  fallback for machines without Excel. A standalone report `.xlsx`
  (Summary + Findings) is always produced as well. JSON
  (`scripts/validate_workbook.py`) is still available for scripting.
- **Recalculation** (real): `scripts/recalculate_workbook.py` / `app/recalc.py`
  drives your installed Excel via COM. **On a machine without the
  MMForExcel add-in installed, every `MM_` function call will show
  `#NAME?`** -- the adapter detects this and reports `NOT_SUPPORTED` with an
  explanation rather than a false `ERROR`. It's a deliberate, separate step
  (not run automatically during analysis).
- **Auto-fix** (narrow, real): `app/change_apply.py` -- `.xlsb`/format
  conversion and loop-name case normalization (LOOP-002), both written by
  Excel and verified to open; everything else the rules flag needs a human
  decision by the rules' own design.
- **AI-assisted suggestions** (narrow, real): `app/llm.py`, one on-demand
  "suggest a fix" call for formula findings via the shared Milliman APIM
  Claude gateway. Recommendation only, never auto-applied.
- **Prep workbook** (real changes): `app/prep.py` plans concrete
  operations from the findings (rename hidden sheets with `&&Hide`, fix
  loop-name casing across MM_LOOP/MM_RESULT/..., add `/Input` to `/Reorder`
  grids, correct misspelt flags and special-grid headers, separate merged
  grids, give every untitled grid one `#Name` title taken from the cells
  around it -- captions, nearby labels, header rows -- with unique names,
  make theme colours explicit; optional: keep styles on empty cells,
  unprotect sheets) plus user-edited formula replacements. You tick what
  you approve; Excel applies it to a fresh copy in one session, the copy is
  verified to open, and a change log is written.
- **Ask about this workbook**: `app/chat_context.py` -- a multi-turn
  assistant grounded in the analysis and findings (sheets, grids, flags,
  formulas, every rule result, retrieved cell/grid/rule detail; it can look
  up cells it hasn't seen). Ask questions, or tell it to change anything:
  it proposes the exact operations, you click Apply, Excel writes them to
  a fresh copy, the file is verified and re-analyzed, and the conversation
  continues on the changed workbook. Needs the shared gateway key
  (secret.key / config.enc).
- **Web app** (1.6.0): the Figma-built React front-end in
  `../FigmaOutput` served by `app/web/server.py` (FastAPI) --
  `run_mind_ready_web.bat` -> http://localhost:8600. Same engine, same
  guarantees, plus a per-session version lineage (every apply is a new
  Excel-verified file with its change log), a "since the previous
  analysis" delta after every apply, and a Fix panel on every finding and on every
  recalculation error (quick fixes + focused mini chat, recalculate to
  confirm; errors sharing one root cause are fixed together). Build the front-end once with
  `npm install && npm run build` in `FigmaOutput`.
- **UI**: `run_excel_upload_prep.bat` launches `app/ui/streamlit_app.py`:
  sidebar (upload, mode, run, current-file card) and sections Findings /
  Prep workbook / Ask the assistant / Recalculate / Reports (report
  downloads carry a visible "verified to open in Excel" line). Verified
  end-to-end with Playwright on a real 13.7 MB `.xlsm`.

Still honestly `NOT_SUPPORTED`: `READY-001` until a real recalculation has
been run (`FORMULA-001`, which could never pass, was removed in 1.6.4). `PASS` still requires a clean, real
recalculation -- it's not handed out for free, and on a machine without
MMForExcel installed it will only ever be `NOT_SUPPORTED`, `WARNING`, or
`ERROR`, never a false `PASS`.

Setup: `pip install -r requirements.txt`, then `python -m pytest tests/unit`
(146 tests; the Excel-backed ones skip without Excel).
Run the UI: `run_excel_upload_prep.bat` (or `streamlit run app/ui/streamlit_app.py`).

## Implementation note

Rules are mined from `01_Mind_Readiness_Standard.md` and
`02_MM_Function_Registry.md` (sibling docs) via `tools/mine_readiness_standard.py`
and `tools/mine_function_registry.py`, and the KB references from
`CompleteMindDocn.docx` via `tools/mine_kb_docx.py` -- deterministic parsers,
safe to re-run against updated source docs. `rules/kb-rules.yaml` is the one
hand-authored rule file (8 rules citing their KB articles). See CHANGELOG.md for what's implemented
vs. registered-but-`NOT_SUPPORTED`.
