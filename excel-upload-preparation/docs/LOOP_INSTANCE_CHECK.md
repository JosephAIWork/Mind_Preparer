# Where a loop or instances would fit: a structure check (prototype, 2026-10-22)

Status: **prototype only, not in the app.** Scripts: `tools/structure_check/`
(`extract.py`, `analyze.py`, `report.py`, `build_control.py`; how to run them in
its README). Nothing in `app/` imports them; they import the app's R1C1 parser
(`app.autofix.parse_r1c1`) read-only. Their `data/` folder holds client formulas
and is git-ignored. Start the implementation from these scripts.

## Purpose

Check a workbook for structure that a Mind loop or Mind instances would express,
to remove repeated formulas or to compute every choice in one run. It **checks;
it does not choose.** Every finding lists three options (loop, instances, leave
as is), each with "fits / does not fit as is" and the reasons. "Leave as is" is
always listed. Mind's own cost note is stated: instances use more RAM and CPU
than loops.

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

## Next step (not done): merge findings by the input they step through

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

## If it goes into the app later

A read-only "Structure check" next to the model map: it reports, never writes.
