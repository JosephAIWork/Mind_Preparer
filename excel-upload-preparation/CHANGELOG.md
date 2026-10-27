# Changelog

## 1.12.0 (27 October 2026): the hand-over to the Mind team

- **Send feedback to the Tel Aviv team.** When a step fails, a box under the header says in one plain
  sentence what stopped and why, with the button; the same button is in the header for anything else.
  The dialog prepares the message (the assistant, when a key is saved, writes its opening lines), shows
  it, then copies it and opens an email to the address in the settings (`feedback.email`); a file is
  saved only on request. The message is in codes -- the stage, the error's type, the app functions it
  went through -- with Excel's own code and the *shape* of the formulas involved (function names kept;
  references, text and numbers taken out), and never a file name, a sheet name, a cell address or a
  value. Each problem is kept in the data folder (`feedback\incidents.jsonl`, the last 300). The codes
  are hashes: the app holds no table of what they mean (`app/feedback.py`, `POST /api/feedback`).
- **The AI key is pasted in the app.** When no key is saved, the main screen and Advanced > Assistant
  show a box for it. The key is checked with the gateway (a key the gateway refuses is not saved), then
  saved encrypted for the Windows user (DPAPI) in the data folder; the page only ever shows its last 4
  characters. The old search for `secret.key` / `config.enc` in the folders around the app is gone, and
  with it the `cryptography` requirement (`app/ai_key.py`, `GET`/`POST /api/assistant-key`).
- **One settings file**, `settings.yaml`: the port, the data folder, the AI gateway and model, the Mind
  addresses, the Mind folder and project prefix the app works with, the MM for Excel add-in, and the
  API's own pages (off by default). A mistake in the file stops the app with a message naming the line.
  The app deletes in Mind only projects whose name starts with the prefix (`app/settings.py`).
- **CELL / INFO formulas (FRM-012)** are frozen to what the original upload shows, the values the user's
  own Excel saved, followed through the rows inserted since. A version saved by the app's Excel shows the
  app's working folder there instead (with the Windows user name in it), and such a value is never written.
- **Pivot tables (STR-009): every one blocks Mind's converter.** A workbook was refused with its pivot
  tables, refused again with only the one with a grouped field turned into values, and accepted with all
  of them as values. The pivot setting is values (the default) or formulas.
- **A workbook whose formulas read its pivots through GETPIVOTDATA is converted too** (it used to keep every
  pivot, and Mind refused it). Before a pivot becomes cells, each GETPIVOTDATA reading it is pointed at the
  cell it returns today (an error stays an error, as NA()), and each such cell is checked in Excel to show
  what it showed. A pivot's custom style is deleted once nothing uses it: Mind stopped on pivots styled
  "Results" with "Requested value 'Results' was not found" (`app/pivots.py`).
- **Broken references (REF-001):** a broken defined name written with its sheet (`Projn!IndxAsst4`) becomes
  `NA()` with the sheet, not `Projn!NA()`, which Excel refused.
- **Formulas Excel 365 stores as one-cell arrays are fixed like any other (FRM-008 to FRM-013).** Excel 365
  saves most formulas with INDEX or OFFSET as a "dynamic array" over its own cell, and the Mind-value fixes
  left every such cell "by hand": 512,796 cells of one model, so FRM-009 stayed blocking after Do everything.
  Only an array over several cells is left by hand now. A one-cell array is written back as one
  (`Formula2`: the old way added an `@` to the formula, which made it another formula).
- **A formula filled down is written as one fill.** The Mind-value fixes plan one change per block of cells
  holding the same formula before and after (for example `Output!BK133:BK64231`): 64,099 cells in about a
  second, where one write per cell took about 8 Excel calls each. The block's last cell is read back against
  the plan, and the block is put back as it was if it differs.
- **Lone text cells (STR-008).** A cell inside a defined name counts as read only when some formula reads that
  name (directly or through another name): 9 labels of one model were flagged only because a name over the
  whole sheet covered them, a name no formula uses (a macro's; Mind runs no macro). A cell with no room for
  a title above it used to be left by hand; now, if it sits just under a grid whose first column holds
  text labels, it gets a label beside it so the grid takes its row. Otherwise it moves (Excel's cut and
  paste, so every formula reading it follows) to a column of its own right of the sheet's contents, with
  its title above. A cell read through a range or a name over several cells keeps its place: the range
  would lose it.
- **"Do everything right after the upload"**, a box under the drop area (remembered in the browser): once the
  workbook is read, Do everything starts by itself, and a big workbook is read whole without stopping to ask
  which sheets to leave out. The computer must stay awake: when it sleeps, the page that drives the run waits.
- The header and the browser tab say **Beta**.
- The old Streamlit screen is removed; comments, messages and documents name no client model.

## Up to 1.11.10

The app reads an Excel model and checks it against the readiness rules mined from Mind's knowledge base
(`rules/`), with every check deterministic. It fixes what it can through Excel itself, on a copy:

- **Fix all issues**: finds the cells where formula errors start, gives each a value to fall back on,
  and undoes any fix that changed a value that was correct.
- **Fix the formulas** and the structure (Prep): every change Mind needs before it accepts the file --
  grid titles, flags, broken references, 3-D references, INDIRECT used as a value, external links, cells
  that hold `""`, empty cells compared with numbers, CELL / INFO, line breaks in text, pivot tables --
  each change reviewable before it is applied.
- **Build the Mind loop**: the run loop of the model's macro written as `MM_LOOP` / `MM_RESULT`.
- **Recalculate** in Excel with the MM for Excel add-in, and **check the numbers**: every formula cell
  of the original and the prepared version compared, through the rows inserted.
- **Do everything** runs those steps in order, each followed by its check.
- Under Advanced: the findings rule by rule, the Model map, the Grid Namer, the assistant, Run in Mind
  (upload, convert, run and read back, in a sandbox folder only), and every version.

Each rule that comes from a real Mind conversion says so in its source line. The development notes of
these versions are kept internally.
