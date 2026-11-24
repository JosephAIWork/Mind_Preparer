# Where a loop or instances would fit: a structure check (prototype, 2026-10-22)

Status: **the finder is in the app since 1.12.5** (`app/instance_finder.py` with `app/finder_shapes.py`,
`app/finder_patterns.py` and `app/finder_grouping.py`; `docs/INSTANCE_FINDER.md`); the scripts below
stay as the prototype's record. The builder half of plan item 3 is in the app since 1.12.5
(`app/instance_builder.py`, `docs/MIND_INSTANCES.md`). Scripts: `tools/structure_check/`
(`extract.py`, `analyze.py`, `report.py`, `build_control.py`; how to run them in
its README). Nothing in `app/` imports them; they import the app's R1C1 parser
(`app.autofix.parse_r1c1`) read-only. Their `data/` folder holds client formulas
and is git-ignored. The app's finder was ported from these scripts in pure Python:
no Excel, it reads the app's own scan and turns every formula into R1C1 itself.

## Plan (order agreed 2026-10-26)

1. **Next: the other VBA loop keywords in the loop builder.** Today it recognises only
   `For i = ... To ... Next` (`app/model_map.py`, `app/loop_builder.py`); a macro that loops
   any other way gets no Mind loop and no reason why. To add: `Do While ... Loop` /
   `Do Until ... Loop` / `Do ... Loop While|Until` and `While ... Wend` with a counter
   raised by hand (`i = i + 1`) that run until a cell is empty -- an `MM_LOOP` over the
   filled rows of the table, like `For`; `For Each c In Range(...)` -- an `MM_LOOP` over
   that range. A loop that stops on a computed condition is listed as not built, with the
   reason. Every new form with a test, as the `For` loops have. About 2-3 h.
2. **Grid naming and flags.**
3. **Where instances and loops would help.** Both halves are done. The **builder** (1.12.5)
   turns the macro's run loop into Mind instances -- one copy of the model per key of the
   model-point table, a loop inside each copy over that copy's rows, a main workbook with the
   keys and the results (`app/instance_builder.py`, Advanced > Mind loop > Mind instances,
   `docs/MIND_INSTANCES.md`). The **finder** (1.12.5) points to where a loop or instances would
   help -- the grouping column of the looped table, repeated blocks, repeated sheets, near
   copies -- blind and deterministic, merged into dimensions by the input the copies step
   through and ranked, with the cells that read the macro's paste area sorted by what each
   needs in Mind (`app/instance_finder.py`, the **Where instances could help** section of the
   same panel, `docs/INSTANCE_FINDER.md`). What remains: the other VBA loop keywords (item 1),
   the grid flags (item 2), a run-time prediction (formula counts are not run time), and copies
   across separate workbooks.

## Purpose

Check a workbook for structure that a Mind loop or Mind instances would express,
to remove repeated formulas or to compute every choice in one run. It **checks;
it does not choose.** Every finding lists three options (loop, instances, leave
as is), each with "fits / does not fit as is" and the reasons. "Leave as is" is
always listed. The cost note is stated as the team's own observation from its
Mind projects, not a sentence of Mind's documentation: instances use more
memory and processor than loops.

## How it reads a workbook

Excel reads each sheet's `FormulaR1C1` (read-only, macros off, calculation
manual). In R1C1 a formula filled down or across, or copied, is the same text,
so repetition becomes visible.

| Check | What it finds |
|---|---|
| Sheet copies | Sheets with the same formulas at the same addresses |
| Repeated blocks | A block of columns or rows repeated side by side in one sheet |
| Parallel or chain | Does a copy read an earlier copy? A chain (periods) is a sequence: neither method expresses it, set aside |
| Time axes | Filled columns or rows that read the previous one (a projection): never flagged |
| Switch cells | A typed-in cell used as the *key* of many lookups (VLOOKUP/MATCH value, CHOOSE index, INDEX/OFFSET row or column, SUMIFS criterion), also through a defined name: Excel computes one choice per run |

Where copies differ, every cell is sorted into: identical / reads a different
input (becomes the loop index) / differs only in a typed-in number (an
assumption to move into a table) / different logic (an instance keeps it as an
override; a loop would need IF). Also counted: assumption numbers that differ,
cells present in some copies only (different sizes), formulas that add the
copies up (MM_RESULT for a loop, a plain link for instances).

Noise removed on the way: a column of row labels each read by its own row is
not a switch (neighbouring cells read about equally often are dropped); a cell
read only as a lookup *table* or *column number* is a parameter, not a switch.

## Results

A control workbook with known answers: all five cases right (copied product
sheets with one different formula, the Total that adds them, a chain of year
sheets rejected, independent scenario blocks found, time projections ignored).

| Workbook | Found | Against what is known |
|---|---|---|
| UC-1 | Switch: the loan-number cell (4,674 readers); the rest a time projection | The cell the macro steps through |
| UC-2 | Switch: the policy-run cell (+ the valuation date, a run setting); 1 chain; 1 small block | The loop the app built from the macro |
| UC-6 | Already 4 MM_LOOP; one possible selector cell | Plausible |
| UC-5 | Nothing (347,000 formulas, all filled areas) | Plausible |
| UC-3 | Switches: the active year, the projection year, the active cluster (a region name), the active scenario; Category A-D blocks in many sheets; two output sheets that are copies (100 cells of different logic, 8,680 differing numbers: loop does not fit as is, instances fit); two more sheets with identical formulas and different data; 33 chains set aside | Checked against the cells: labels "Active Cluster", "Category A..D", years 2026-2030 |

Extraction time through Excel: UC-2 76 s, UC-3 194 s.

## Not good enough yet

1. Too many findings on a big model (UC-3: 200) -- they need merging into
   dimensions (next step below).
2. Items or periods: the formulas cannot tell "2026, 2027" from "Category A, B".
3. Formula counts are not run time; Mind's speed and memory are not predicted.
4. Switch strength counts direct readers only, not everything downstream.
5. The list of choices (a dropdown, a table column) is not in the formulas;
   the model map reads dropdowns and macros.
6. Copies across separate workbooks are not covered.

## Merge findings by the input they step through (done in 1.12.5)

Header labels are only what the author typed: they can be missing, wrong or
inconsistent ("Cat A" / "Category A"), or formulas. So do not merge on labels.

Merge on what the copies **read**: if block 1 reads row 1 of an input table,
block 2 row 2, and so on, the dimension is that input table. Blocks in
different sheets that step through the same table are the same dimension, even
when their labels say nothing or disagree. This uses the same evidence as the
rest of the check (the formulas). Labels are shown only as names, to help the
reader -- never as proof.

Expected result on UC-3: a few named dimensions (e.g. the category table, the
year table, the cluster switch) instead of 200 lines.

Done in 1.12.5: `app/instance_finder.py` merges every finding by the anchors its copies
read -- a table column, a list, an area cut into slots, a band, a fill, a group of sheets
-- including copies that step by a computed index (`ROW()`, `COLUMN()`, `COUNTA`, `MATCH`),
which the prototype's relative-reference test could not see, and lists that hold the same
values. The macro's own chain (the lookup column, the list it matches into, the slots it
pastes into) is one dimension with them. A header stays a label, never evidence. The blind
check of the app's finder is on UC-2 and is recorded in `docs/INSTANCE_FINDER.md`; UC-3 has
not been through the app's finder yet.

## In the app

It went into the instances panel under Advanced > Mind loop, not next to the model map: the
**Where instances could help** section (1.12.5), below the instances builder, so that **Use as
the key** can fill the builder with the column it found. It reports, never writes.
