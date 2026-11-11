# Runbook — audit and fix grid names, and check them in Milliman Mind

Companion to [RUN_IN_MIND.md](RUN_IN_MIND.md), which covers getting a workbook
to convert and run. This one covers the next question: **does every grid have a
name a human recognises, once it is in Mind?**

It reproduces, step by step, the audit that produced Mind Ready **1.6.7**.

> Since **1.7.0** there is also a hands-on alternative: the **Grid Namer**
> screen (sidebar). It shows the whole workbook sheet by sheet with the
> detected grids outlined; drag over an area, type the name you want, tick
> flags, Submit — everything you did not touch is named by the conventions
> this runbook audits. Use this runbook when you want the *automatic* names
> checked; use the Grid Namer when you already know what a block is called.

Why it matters: Mind lists every grid by name in the `work ▾` template selector.
A grid with no `#Title` above it is listed as **`Untitled(60,8)`** — the row and
column of its first cell. A grid the app could only call `Calc C4` is no
better. Neither stops the model converting or running; both make it unreadable.

---

## Before you start

| Need | Check |
|---|---|
| App running, 1.6.7 or later | <http://localhost:8600/api/health> → `"version":"1.6.7"` |
| Excel installed | the same endpoint → `"excel":true` (grid titles are written through Excel) |
| Assistant reachable (optional) | the same endpoint → `"assistant":true` (the AI key is saved in the web app) |
| Workbook is `.xlsx` / `.xlsm` | Mind rejects `.xlsb`; *Save As* → Excel Macro-Enabled Workbook first |

Start the app with `run_mind_ready_web.bat` in `excel-upload-preparation`.

Everything below reads an **immutable copy**. Your source workbook is never
touched, and nothing is written to it at any point.

---

## Part A — Audit the names (read-only, no UI)

`scripts/audit_grid_names.py` is the whole audit in one command.

```bash
cd excel-upload-preparation
python scripts/audit_grid_names.py "path\to\YourModel.xlsm"
```

It prints the five numbers that matter:

```json
{
  "grids": 156,
  "unnamed": 60,
  "weak": 76,
  "title_operations": 50,
  "blocked": 19
}
```

Read them like this:

- **`unnamed`** — grids with no `#Title` at all. **This is your `Untitled(r,c)`
  count in Mind.**
- **`weak`** — grids whose name carries no meaning: the `<Sheet> <Anchor>`
  fallback (`Calc C4`), or a scrap of a header row (`I`, `I I ...`).
  Judged by `prep.is_weak_name()`.
- **`title_operations`** — how many titles the prep plan can write on its own.
- **`blocked`** — grids it *cannot* title, each with the reason, printed below
  the totals. Read these: they are the ones needing a human decision.

Add `--output audit.json` for the full per-grid table (current name, what the
heuristics propose, and where that name came from).

### A2. Read the surrounding cells yourself

To judge a name you need to see what the grid actually holds:

```bash
python scripts/audit_grid_names.py "YourModel.xlsm" --context context.txt
```

`context.txt` gets one block per unnamed/weak grid — the cells above it, the
cells to its left, and its first rows, with **formulas shown as
`formula => computed value`** so a column of `=INDEX(...)` is readable:

```
--- Calc!C27:E39  (13 rows x 3 cols)  current=None
 ABOVE:
   B25=#Calc B24 | B26==MAX($B$3:B25)+1 => 2 | C26=Demographic Assumptions
 DATA:
   Gender | =INDEX(Policies!A:CE, => ... | =IF(D27="Male","M","F") => M
   Smoker | =INDEX(Policies!A:CE, => ... | =IF(D28="Smoker","S","NS") => NS
```

This is exactly what a person (or the assistant) reads to name a grid. Naming
from this file, before looking at any suggestion, is how you keep your own
judgement independent — do that if you intend to compare candidates.

### A3. Ask the assistant for names

```bash
python scripts/audit_grid_names.py "YourModel.xlsm" --suggest --output audit.json
```

Each weak grid gains a `"suggested"` field. The assistant is asked in **batches
of 25** through the Milliman APIM gateway (`app/grid_naming.py`), so a whole
workbook costs a handful of calls. If no AI key is saved it says so on
stderr and the deterministic names stand — naming never hard-depends on the
network.

---

## Part B — Apply the names in the app

1. **Upload** your `.xlsm` on the Workbook screen (<http://localhost:8600>).
2. Go to **Prep**. The first panel is **“Name grids from their context.”**
3. Click **Suggest names**. The app sends only the grids whose deterministic
   name is meaningless, and shows a review table:

   | Grid | From the cells around it | Assistant |
   |---|---|---|
   | `Calc!C43:D47` `5x2` | *Calc C43* (sheet name + position) | Sum Assured And Inflation |

   Names in italics are the meaningless fallback. Accepted names are folded
   into the plan below — **nothing is written yet.**
4. Tick **“Give every untitled grid one '#Name' title…”** and review its
   operations cell by cell (**Show operations**). Read its **SKIPPED** list too
   — that is the `blocked` set from Part A.
5. Click **Apply**. Excel writes the titles to a fresh copy, the copy is
   verified to open, and a change log is written beside it.
6. **Download** the prepared file from the **Download** area in the left sidebar.

> The assistant only changes *which* name a title operation writes. It never
> decides *where* a title goes or whether the edit is safe — those stay
> deterministic, and every operation is still reviewed and approved by you.

### Confirm it improved

Re-run Part A against the file you just downloaded:

```bash
python scripts/audit_grid_names.py "YourModel_v2.xlsm"
```

`unnamed` should drop sharply (60 → 39 on one client model) and `grids` should
stay roughly flat. **If the grid count jumps, stop** — see “When it goes wrong”.

---

## Part C — Verify in Milliman Mind

1. Upload the prepared file to a Mind project and convert it — steps 6–10 of
   [RUN_IN_MIND.md](RUN_IN_MIND.md).
2. After **Add template to Milliman Mind**, click the **`work ▾`** selector in
   the middle of the top bar. It lists every template and grid by name.
3. Scan the list for **`Untitled(row,col)`** entries and for names like
   `Calc C4`. Each one is a grid Part A predicted; the row/column in
   `Untitled(60,8)` maps to row 60, column 8 (= `H60`).

That comparison — Mind's own list against the audit — is what proves the audit
matches what Mind actually shows.

---

## When it goes wrong

**Grid count jumps after applying.** The workbook is being fragmented, not
tidied. Do not upload it; go back to the previous version.

This is what the 1.6.7 investigation found: Excel stores error *values* as text
starting with `#`, and they were being read as `#Title` markers. Five `#N/A`
cells inside a 411×84 policy table looked like titles “trapped” in the grid
(STR-001), so the default-on **“Insert an empty row before a '#Title' trapped
inside a grid”** action offered to insert rows straight through that table.
Applying it took the model from **156 grids to 258 in one pass, and 648 over
three**.

Fixed in 1.6.7 (`grids.looks_like_title()` ignores Excel error literals). To
check any workbook prepped by an **earlier** version, compare grid counts before
and after:

```bash
python scripts/audit_grid_names.py "before.xlsm"   # note "grids"
python scripts/audit_grid_names.py "after.xlsm"    # should be close, not multiplied
```

A large increase means that version's prep sliced the workbook — re-prep the
original with 1.6.7.

**Do not loop the prep to chase zero.** Applying prep to its own output
repeatedly keeps inserting rows and slowly grows the workbook for ever fewer
new names. Run **one** pass, then handle what is left by hand.

---

## What the app cannot name, and why

The reasons printed under `blocked` are worth understanding — they are design
limits, not bugs. Since **1.14.0 a title never moves a cell**: up to 1.13.0 the
app inserted a row (or cells over the grid's own columns) when the cell above
was taken, and on the IFRS 17 tool's `Result_CF` that pushed the key column
(read as a grid of its own) one row down while the block beside it stayed —
every key then sat beside the wrong row. Now the title goes into the one cell
Mind reads a title from, directly above the grid's top-left cell, or nowhere:

- **“the cell above (C2) holds 'Demographic Assumptions' (part of the grid
  …) — a title needs that cell and nothing is moved to make room”** — a
  title must sit directly above its grid, and that cell belongs to another
  grid (a section heading row, the table right above). Leave an empty row
  above the grid by hand, then name it again. On one client model this is most
  of its assumption lookup tables (mortality, lapses, discount curves,
  tariffs …): their **header row and their data are read as two separate
  grids**, because the label column beside the data is empty so the block is
  not rectangular. Merging header and data into one block is a modelling
  decision for a human — the app will not restructure a table to win a name.
- **“starts on row 1, so there is no cell above it for a title”** — insert an
  empty row above it by hand, then name it again.
- **A caption that is a formula is replaced (1.14.1, Joseph's decision of 11
  November 2026)** — a key column with a caption on top and the block beside
  it are one table; its title goes into the caption cell ('#' + the caption's
  own text). When that caption is a formula (`Result_CF!A1` on the IFRS 17
  tool: `="実行中のシナリオ　["&Ctrl_01!$C$27&"]"`), the cell loses its
  formula and gets '#' + the text the formula last showed (the value Excel
  stored for it) plus the flags, for the automatic namer, the Grid Namer's
  name and the assistant's name alike. The action's note says so: *“this cell
  held a formula (…); it is replaced by the title. Open issue: if the caption
  must stay a formula, leave an empty row above the table instead.”* The
  table is no longer listed under "could not name". When the formula's last
  value is empty or an error, the title is made up from the surroundings as
  for an empty cell. A caption cell something reads is still refused (“the
  caption above it (A1) is read by …”), formula or not; nothing is ever
  inserted. **Open issue**: the title is a frozen copy of what the formula
  showed; if the caption must keep following its formula, leave an empty row
  above the table by hand and name it again.
  Only a text caption counts: a year, a date or TRUE above a key column is a
  value something may read, so the two stay apart and nothing is written
  over it.
- **“a title in B3 would touch B2 ('…') right above it, and Mind would read
  the two as one grid”** — a lone text or a grid ends right above the title
  cell, and Mind reads on from it through the title into the grid. Leave an
  empty row between them by hand. The same beside the title cell, on its
  left (“… would touch B2 ('…') beside it, which is read by … / is part of
  the grid …”): Mind reads left to right and would run from that cell into
  the title. A lone label there that nothing reads is moved into the title
  when the title's name is its own text (the automatic pass naming the grid
  from it); the Grid Namer's name, or the assistant's, leaves the label where
  it is and reports the grid ("... which is a label of its own"). The old namer wrote such titles: on the
  anonymized IFRS 17 tool its own re-scan then found 85 grids with a
  `#Title` trapped inside; the 1.14.0 namer's output has none of its own.
- **“contains a trapped '#Title' (STR-001) — separate the grids first”** — a
  `#Title` sits inside the grid. Deal with that deliberately; do not chain the
  separation action blindly (see above).

---

## Reference

| Piece | Where |
|---|---|
| Audit CLI | `scripts/audit_grid_names.py` |
| Grid detection, title vs error value | `app/grids.py` — `looks_like_title()`, `ERROR_LITERALS` |
| Deterministic naming | `app/prep.py` — `_section_header_name()`, `_context_name()`, `is_weak_name()` |
| Title placement (nothing moved) | `app/prep.py` — `_title_site()`, `plan_create_grid_titles()`; `app/grids.py` — `_merge_key_columns()` |
| Assistant naming | `app/grid_naming.py` — `build_grid_context()`, `suggest_names()` |
| API | `POST /api/sessions/{id}/grid-names` (`{"apply": true}`) |
| UI | Prep screen → “Name grids from their context” |
| Rules | STR-004 (grid naming), STR-002, STR-001 (trapped title) |
