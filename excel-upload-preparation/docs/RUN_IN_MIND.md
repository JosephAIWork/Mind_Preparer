c# Runbook — prep a workbook and run it in Milliman Mind

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
10. Click **Add template to Milliman Mind** (the blue button). You'll see
    **"Successfully added model"** and the project opens with a **Run** button.
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
