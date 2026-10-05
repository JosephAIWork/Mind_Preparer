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

## Current status (1.7.2)

All 12 phases have a real implementation, and **every active rule (94 of
94) has a real validator** -- nothing falls through to `NOT_SUPPORTED` any
more except the two honest cases below.

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
