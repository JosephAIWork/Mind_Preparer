# Changelog

## Unreleased (after 1.11.10): the hand-over to the Mind team

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
  of them as values. The pivot setting is now formulas (the default) or values.
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
