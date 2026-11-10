# Runbook — prep a workbook and run it in Milliman Mind

End-to-end: take a raw model workbook, fix it with the local **Mind Ready** app,
then upload it to **Milliman Mind**, convert it, and run it. This reproduces the
result achieved on a client IFRS risk model (converts + runs cleanly).

Two rules to remember up front:
- **Mind only accepts `.xlsx` / `.xlsm`** (not `.xlsb`). If your model is `.xlsb`,
  open it in Excel → *Save As* → **Excel Macro-Enabled Workbook (`.xlsm`)** first.
- The fix that makes this particular model convert (rewriting broken `#REF!`
  references to `NA()`) is now **opt-in** — you must turn it on in step A4.

---

## Part A — Fix the workbook in the local Mind Ready app

1. **Start the app** (if not already running): run `run_mind_ready_web.bat` in
   `excel-upload-preparation`, then open **http://localhost:8600**.
   (Header check: it's v1.6.6 / 96 rules.)
2. **Workbook screen → upload** your `.xlsm` (drag it onto the drop zone or
   *Load workbook*). The app analyses it and lands on **Findings**.
3. **Findings** — review what it caught. On that model you'll see, among
   others: **REF-001** (broken `#REF!` references — the conversion blocker),
   **REP-001** (value reconciliation), plus WARNINGs. No `FILTER` error anymore.
4. **Prep screen → turn on the fixes and Apply.** In the prep action list, make
   sure **"Replace broken (#REF!) references with NA()"** is **ticked** (it's
   off by default because it edits formulas). Leave the other default fixes on.
   Click **Apply** — the app edits a *copy* via Excel (your original is never
   touched) and creates a new version.
5. **Download the fixed workbook** — use the **Download** area in the left
   sidebar (it offers the latest version, e.g. `…_v2.xlsm`). Save it. This is the
   file you upload to Mind.

*(Optional: click **Recalculate** to confirm the fixed copy has no formula errors,
and **Reports** for a written findings report.)*

---

## Part B — Convert and run it in Milliman Mind

6. **Open Milliman Mind** and go to your workspace's **Project Manager**
   (`shared.milliman-mind.com` → *Access the Project manager*, or your normal
   Mind entry). Sign in if prompted.
7. **Create a project for the model:** click **New project**, give it a name,
   **Create project**. Then **open it** (click its tile — it opens the project).
8. The empty project shows **"Drag your Excel files here to convert them."**
   Click **Browse files** (or drag) and pick the **app-fixed `.xlsm`** from A5.
9. Mind runs a 3-step wizard: **Upload → Convert → Test.** Wait for all three to
   complete → **"Upload successful"**.
   - If it stops at **Convert** with a red step, click **Display Error Logs** —
     that message tells you exactly what to fix back in the app (this is how the
     app's rules were built). Fix, re-download, re-upload.
10. Click **Add template to Milliman Mind** (the blue button; **Add templates …**
    when you uploaded several workbooks). You'll see **"Successfully added
    model"** and the project opens with a **Run** button. If the model has
    outside input files, upload them now (see *Outside input files* below).
11. Click **Run** (top-left). Wait for **"Run in Progress"** → the button turns
    **green** and you get **"Calculations completed — Run completed in N seconds
    · Consistent with the audit trail."** The model has run.
12. **Get the results:** use **Exports** / the **download icon** (top-right) to
    export outputs, the **chart icon** for dashboards, or the **Grid validation
    manager** (export/right side) to review grid results. Audit trail of runs is
    under the history/clock icon (top-left).

---

### What "good" looks like
- Convert wizard: **Upload ✓ Convert ✓ Test ✓** ("Upload successful").
- After **Run**: green Run button, **"Calculations completed … Consistent with
  the audit trail."**

### If Convert fails
The error log names the cause. Common ones this app now handles automatically:
broken `#REF!` references (REF-001 → NA() rewrite, opt-in in A4) and unsupported
functions (FRM-002; note Mind *does* support `FILTER`). Fix in the app, re-download,
re-upload.

---

## Outside input files (Input Manager)

Some models get data from other files through a macro. The app turns each of those files
into a **Mind input**: a grid `#IN_<name> /Input` in the model, plus one CSV named
`IN_<name>.csv` (or a TAB `.txt`) in **Mind inputs.zip** (see `OUTSIDE_INPUTS.md`).

**By hand** (and each quarter, with the new CSVs — no new model upload):

1. Open the project. A model with `/Input` grids has an **Inputs** button left of **Run**.
2. Click **Inputs**. The window lists model › sheet › grid, with **Status** and **File**.
3. Leave **Culture** (English - United Kingdom) and **CSV Separator** (Default) as they are.
4. Click **Browse** and pick the **folder** that holds the CSVs (or drop the files).
   Mind puts each file on the grid of the same name: `IN_BBA_01.csv` → `IN_BBA_01`.
5. Every row should show **SUCCESS** and its file. If Mind opens a window saying **Some of
   your imported files have not been automatically matched** (older Mind: **Match unmatched
   inputs**), a file's name matches no grid: click **Close** (Cancel) and fix the file name.
   Never pick a grid from the list: a file in the wrong grid changes the wrong numbers.
6. Close the window (nothing to save) and **Run**. The run uses the uploaded values,
   also in a later session.

**A model with copies (Mind instances).** The window lists an input grid of a copied workbook
once, as **Rates (3 instances)**, collapsed; the small caret opens one row per copy (K1, K2 ...)
with its own Status and File. A file named `<key>-<grid>.csv` (`K2-Rates.csv`) changes that one
copy. A plain `Rates.csv` is **not** matched (Mind cannot tell which copy it is for) unless the
**Split between instances** switch is on, which gives every copy the rows of its own key.

**Run in Mind does this by itself** once the app has made the Mind inputs: in every
attempt, after the model is added and before the run, it uploads the folder of CSVs
(only the `.csv` / `.txt` files) and checks that each one shows SUCCESS. It opens the
collapsed rows of a copied grid first, so a copy's file is read on its own row. If a file is
not matched or not accepted, the attempt stops and names the file, in plain words ("Mind
could not tell which grid X.csv belongs to. Name it <grid>.csv or <key>-<grid>.csv ..."); the
window is closed with its own Close button and the app never picks a grid for a file itself.
It never changes Culture or CSV Separator: if Mind shows a culture that may read a decimal
point differently (German, French…) or another separator, it uploads nothing and says why.
The details are in the loop report, under `iterations[n].inputs_upload` (status per grid --
a copy's row as `Rates [K2]` -- unmatched files, `advice`).

**Open the project again after a copy's file.** Seen on 10 November 2026 in three invented
3-copy projects: the run right after a copy's first input file still used the copies as they
were loaded before the import (every copy held every row: K2 should have been 915 and showed
120, KeyRead K1 everywhere); a reload of the project before the run gave the right numbers
(K1 242 / K2 915 / K3 160, only K2 changed). By hand: close the Inputs window, reload the
project, then Run. `upload_inputs(..., project_url=<the project's link>)` does that itself when a
file went into a grid with copies (`reloaded` in its result).

## A second run in the same browser session

Mind's project page keeps the line **Consistent with the audit trail.** of the run before
on the page the whole time, also while the next run is in progress ("Run in Progress"); the
line **Last run in N seconds.** is taken down during the run and comes back with the new
run's time, with the notes "Calculations completed" / "Run completed in N seconds" for a
few seconds. `mind_client.run_model` therefore notes the page's run marker before it clicks
Run (`run_marker`: the Last-run line and the number of notes) and waits until a new run was
seen (the marker changed, or "Run in Progress") and the page shows that run's own
Last-run line. Its result says `completed`, `audit_consistent` and `seconds` **of this run**,
`started` (a new run was seen) and, when it did not complete, `reason` ("Mind showed no new
run after the Run click ..." when the page only ever showed the run before). Proven on
10 November 2026 in `zzz_mindready_runtwice_1` (a 3-copy invented model): run, one copy's
input file, run again -- the second run waited for its own result and its numbers differed in
that copy only; a plain `Rates.csv` opened Mind's unmatched window, which was closed and
explained, and the next run went through.
