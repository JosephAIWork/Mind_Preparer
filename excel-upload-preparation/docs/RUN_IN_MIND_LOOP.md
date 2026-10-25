# Runbook — let the app run itself in Milliman Mind

Third of three runbooks. [RUN_IN_MIND.md](RUN_IN_MIND.md) walks the manual path
(fix in the app, upload, convert, run); [RUN_IN_MIND_NAMING.md](RUN_IN_MIND_NAMING.md)
audits grid names. This one covers the **autonomous loop** (Mind Ready 1.6.8):
you point the app at a raw model and walk away; it comes back with a converged
workbook and a report, or with the exact reason it could not get there.

It reproduces the run that took a raw client model (UC-6) to *converged* on
2026-08-30 with nobody in the loop.

---

## What the loop does

```
raw model ─► analyse ─► name grids (assistant, once) ─► plan ─► apply via Excel
                                                                     │
      ┌──────────────────────── local gates ────────────────────────┘
      │  numbers · structure · names
      ▼
   upload to Mind ─► Convert / Test ─► add template ─► Run ─► read template list
      │
      ▼
   decide:  converged  |  retry (one change, from the ORIGINAL source)  |  stuck (reason)
```

Every decision is a rule in `app/mind_loop.py`. No person and no assistant is
consulted while it runs; the assistant is asked **once**, before the first
iteration, to name grids — and only if it is reachable.

Two rules you should know before reading any result:

- **A gate that cannot be checked is a failure, never a pass.** If Excel could
  not recalculate, the loop stops rather than skipping the check.
- **Every retry starts from the original source** with one thing changed. The
  loop never re-preps its own output (that grows a workbook for ever fewer new
  titles).

---

## Before you start

| Need | Check |
|---|---|
| App at 1.6.8 or later | <http://localhost:8600/api/health> → `"version":"1.6.8"` |
| Excel installed | same endpoint → `"excel":true` (prep, and both recalculations of the numbers gate, go through it) |
| A Mind session | `python -m app.mind_client check` → `"authenticated": true`. If not: `python -m app.mind_client login <your /init/… link>` opens Edge; sign in once; the session persists for about a week. Nobody around to sign in? `python scripts/run_after_login.py <workbook>` opens the login window, waits for the sign-in, then runs the loop by itself (Part B2) |
| Assistant (optional) | health → `"assistant":true` (the AI key saved in the web app). Without it grid names are deterministic |
| The source | `.xlsx` / `.xlsm` / `.xlsb` (`.xlsb` is converted through Excel first) |
| Disk | about 70 MB per iteration for a 13 MB model (analysis copy, prepared copy, re-analysis copy, two recalculated copies). Put `--work-dir` outside OneDrive |
| Leave it alone | while it runs it owns the Excel COM session and the Mind browser profile — do not use the Prep / Recalculate screens or `run_in_mind` twice at once |

The source workbook is never modified. Everything is done on immutable copies.

---

## Part A — Run it from the command line

```bash
cd excel-upload-preparation
python scripts/run_in_mind.py "C:\path\Model.xlsm" --work-dir "C:\loops\model" --max-iterations 4
```

Options:

| Option | Meaning | Default |
|---|---|---|
| `--work-dir DIR` | where iterations, copies, screenshots and `loop_report.json` go | `runs/<stem>_<stamp>` |
| `--max-iterations N` | cap on retries; the loop usually converges or stops well before | 4 |
| `--enable ACTION` | force an opt-in prep action on from iteration 1 (repeatable) | none — the loop enables one itself when its rule is an ERROR finding |
| `--disable ACTION` | never use this prep action (repeatable) | none |
| `--no-assistant` | deterministic grid names only | assistant used when reachable |
| `--no-run` | convert in Mind but do not run the model | run |
| `--no-numbers` | skip the value comparison (only for experiments — you lose the guarantee) | compared |
| `--skip-mind` | local gates only; never opens a browser (dry run) | Mind used |
| `--keep-projects` | do not try to delete sandbox projects afterwards | delete what can be deleted |
| `--headed` | show the Edge window | headless |
| `--project-prefix P` | sandbox project name prefix; must start with `mind.project_prefix` of settings.yaml | that prefix (`zzz_mindready`) |

Exit code: **0** converged · **1** stuck / exhausted · **2** bad input.

What you see, line by line (times are real, from the reference run):

```
[13:37:44] iteration 1: analysing the source
[13:37:49] asking the assistant to name 73 grid(s)
[13:38:14] assistant named 73 grid(s)
[13:38:14] enabling opt-in action 'fix_broken_refs': its rule is an ERROR finding
[13:38:14] iteration 1: applying 136 operation(s) from fix_broken_refs, create_grid_titles, explicit_colors
[13:39:49] iteration 1: grids 155 (baseline 181), unnamed 75, blocked 55
[13:39:49] iteration 1: recalculating source and prepared copies in Excel
[13:40:02] iteration 1: numbers match (63174 cells compared, 0 differ, 1 volatile skipped)
[13:40:03] iteration 1: creating sandbox project zzz_mindready_20260830_133744_i1
[13:40:40] iteration 1: uploading Model.xlsm and waiting for Upload / Convert / Test
[13:41:12] iteration 1: convert succeeded {'Upload': 'Complete', 'Convert': 'Complete', 'Test': 'Complete'}
[13:41:24] iteration 1: Mind lists 160 template entries, 51 Untitled; app predicted 75, unexplained 8, app-only 32
[13:41:24] iteration 1: running the model
[13:41:31] iteration 1: run completed, consistent with the audit trail
[13:41:38] iteration 1: converged -- every gate passed; ...
```

About four minutes per iteration on this model: two locally (prep + two
recalculations), two in Mind.

---

## Part B — Run it from the app

1. Upload the workbook on the **Workbook** screen (an `.xlsb` is converted on
   upload).
2. Open **Mind** in the left sidebar. Set *Max iterations*, leave *Run the
   model after converting* ticked, click **Run in Mind**.
3. Watch the **live log** and the iteration cards fill in — each card shows the
   seven gates as chips (✓ passed · ✕ failed · – skipped) and the decision.
4. When it ends, the banner shows the verdict; the final workbook path and any
   leftover Mind projects are listed under it.

The screen polls `GET /api/sessions/{id}/mind-loop?after=N`; you can close and
reopen it while the loop runs. Only one loop runs at a time — a second start
returns *409 a Mind loop is already running*.

Programmatically: `POST /api/sessions/{id}/mind-loop` with any of
`maxIterations`, `useAssistant`, `runModel`, `checkNumbers`, `enable`,
`disable`, `deleteProjects`, `skipMind`.

### Part B2 — Sign in later, let it run by itself (1.7.1)

When the Mind session has expired and you will not be there to click:

```bash
cd excel-upload-preparation
python scripts/run_after_login.py "C:\path\Model.xlsm"
```

with the app already running on port 8600. An Edge window opens on the Mind
login page and the script waits (24 h by default, `--login-timeout-h`). Sign
in there whenever you are back; the script then uploads the workbook to the
app, starts the loop and follows it to the end. Read
`runs/after_login_<stamp>/summary.txt` afterwards (the log next to it has every
event). Defaults: the numbers gate is **off** (`--check-numbers` turns it on)
and the sandbox projects are **kept** so the converted model stays visible in
Mind (`--delete-projects` removes the deletable ones).

---

## Part C — Read the result

The CLI prints a summary; the same text is in the app banner and the full
detail is `<work-dir>/loop_report.json` (rewritten after every iteration, so a
killed run still leaves a record).

```
verdict: converged -- every gate passed; Mind lists 51 Untitled grid(s), the app predicted 75: 8 Mind-only, 32 app-only (see detection_gaps)
source: ...\Model.xlsm
  iter 1: actions=fix_broken_refs,create_grid_titles,explicit_colors | apply=APPLIED | grids 181->155 unnamed=75 blocked=55 | numbers=match | convert=ok run=ok audit=ok | converged: ...
untitled grids left: app 75 (of which 55 the app cannot title safely), Mind shows 51
detection gaps vs Mind: 8 untitled grid(s) only Mind sees (AB2, D158, ...), 32 only the app sees
leftover Mind projects (delete by hand): zzz_mindready_20260830_133744_i1
final workbook: ...\iter_1\prep\Model.xlsm
```

### The three verdicts

| Verdict | Meaning | What to do |
|---|---|---|
| `converged` | every gate passed in the last iteration | use `final workbook`; read the residuals below |
| `stuck` | a gate failed and no rule can change anything | read the reason (Part D) — it names the cause |
| `exhausted` | `--max-iterations` retries without converging | look at the last iteration's decision; raise the cap only if each retry was making progress |

### The seven gates, per iteration

| Gate | Passes when | In the report |
|---|---|---|
| Prepared | the plan applied through Excel and the copy opens | `iterations[].apply` |
| Numbers | full Excel recalculation of the prepared copy equals one of the source, cell for cell (Part E) | `iterations[].numbers` |
| Structure | the grid count did not balloon (a shattered workbook is the signature of a bad row insert) | `names.fragmented`, `names.grids` vs `baseline_grids` |
| Names | every grid the plan could title safely is titled; what is left is listed with the reason | `names.unnamed`, `names.blocked[]` |
| Convert | Mind's Upload / Convert / Test all *Complete* | `mind.convert.steps`, `mind.convert.errors[]` (Mind's error log, verbatim) |
| Run | Mind ran the model and reports *consistent with the audit trail* | `mind.run` |
| Counts | *reported, never enforced*: Mind's `Untitled(r,c)` entries matched cell by cell against the app's untitled anchors | `mind.counts`, top-level `detection_gaps` |

### Residuals a converged run still reports

- **`residual_untitled`** — grids left without a title: `blocked` lists each
  one with the formula that made a title unsafe (Part E). Nothing in that list
  can be fixed by the loop; it needs a modelling decision.
- **`detection_gaps`** — `mind_only`: grids Mind lists as `Untitled` that the
  app did not know about; `app_only`: grids the app split that Mind reads as
  part of a larger block. Mind's grid detection differs from the app's in some
  layouts (on one client model, Mind starts several blocks one column later than the app
  does). This is a limitation of the app's detection, not of your workbook.
- **`leftover_projects`** — sandbox projects that were *run*; Mind's
  automation cannot delete those (Part F).

---

## Part D — When it stops: the `stuck` reasons

| Reason begins with | Cause | Fix |
|---|---|---|
| `numbers could not be verified` | Excel could not recalculate one of the copies (COM busy, file locked, workbook Excel refuses to open) | close other Excel automation, re-run; if the prepared copy is the one Excel refuses, that is an app bug — keep the work dir |
| `values changed in N cell(s) and no active action can be blamed` | a value moved and the precedent walk found no prep operation upstream of it | open `numbers.listed` in the report (Part E); this is a gap in the safety rules — keep the work dir |
| `Mind's Convert failed for a reason the app has no fix for` | Mind's error log is quoted verbatim | the app has no rule for this Mind message yet: the message is the developer's specification |
| `Mind automation stopped before the model was verified` | not logged in, project not created, page changed | `python -m app.mind_client check`; log in again if needed; re-run |
| `prep could not be applied` | Excel refused an operation (`failed` count > 0) | the failed operation is in the prepared copy's `.changelog.json` |
| `Mind ran the model but does not report it consistent with the audit trail` | Mind computed something other than the workbook's cached values | the prepared copy's cached values are stale or wrong — keep the work dir |
| `grid count ballooned with no row-inserting action active` | fragmentation with nothing to disable | keep the work dir; app bug |

A `retry` line tells you what the loop changed for the next iteration — for
example `values changed in 790 cell(s); disabling create_grid_titles` or
`Mind's Convert asked for fix_broken_refs`.

---

## Part E — The numbers gate, and why it is the proof

Mind's *consistent with the audit trail* only says that Mind computed what the
**prepared** workbook's cached values say. It says nothing about whether the
prepared workbook still computes what the **original** did. The first
unattended run showed the difference: titling had changed **790 values** —
a `#Title` written into column A of the policy table was counted by
`Calc!D5 =COUNTA(Policies!A:A)`, the `MATCH` over `$C:$C` shifted,
and the model priced a *different policy* — and Mind ran that workbook and
called it consistent.

What the gate does, every iteration:

1. copies the source and the prepared workbook, recalculates **both** in Excel
   (`CalculateFullRebuild`);
2. compares every cell, mapping the prepared coordinates back through the
   plan's row inserts (`row_map`);
3. skips cells the plan wrote as **content** (titles, moved labels) — that is
   the plan's intent;
4. compares cells whose **formula** the plan rewrote (`fix_broken_refs`) under
   the conservative rule: unchanged, or an error became another error — never a
   valid value into anything else;
5. treats an error becoming another error anywhere as equal (no valid value
   involved); skips `NOW()`/`TODAY()`/`RAND()` cells and their dependents; ignores
   chart sheets;
6. reports `compared`, `differences`, and up to 40 `listed` cells with
   `sheet`, `cell` (source), `prepared_cell`, `source`, `prepared`, and
   `rewritten` (true when the plan rewrote that cell's formula).

To look at a difference yourself: open the source and the prepared copy
(`iterations[].prepared_path`) side by side at `cell` and `prepared_cell`.

The rules that came out of that first run now live in the planner and fire
before anything is written — each refusal appears in `names.blocked` with the
reading formula:

- `... is read by Calc!D5 -- a title there would change that result`
- `... inserting a row would shift what Calc!M4 counts or indexes over whole columns of this sheet`
- `... label kept in place (its text became the title of ...) because Policies!A1 reads it`
- `... the caption is read by ... -- turning it into a title would change that result`

---

## Part F — Clean up in Mind

Sandbox projects are created in the **MindReady_Sandbox** folder of the
production Project Manager, named `zzz_mindready_<stamp>_i<N>`. The loop
deletes the ones it can. A project that has been **run** cannot be deleted by
automation (the confirmation does not take); it is listed under
*leftover projects*. Delete those by hand: Project Manager → MindReady_Sandbox
→ the tile's `…` menu → **Delete project** → type `delete` → **Delete**.

Use `--no-run` when you only need to know whether a workbook converts: then
every sandbox project is deletable and nothing is left behind.

---

## Reference run — a raw client model (UC-6), 2026-08-30

| | Run 1 | Run 2 (after the fixes run 1 forced) |
|---|---|---|
| Iteration 1 | numbers **DIFFER**: 790 of 63,169 cells → `retry`, `create_grid_titles` disabled | numbers **match**: 0 of 63,174 cells |
| Iteration 2 | numbers match; Convert ✓ Run ✓ audit ✓; `stuck` on counts (Mind 154 Untitled vs app 181) | — |
| Outcome | two app changes: reference-safety rules in the planner; counts gate reported, not enforced | **converged** in one iteration |
| Untitled in Mind | 154 | 51 (8 Mind-only, 32 app-only detection gaps) |
| Kept under | — | a local `runs/` folder (not in the repository) |

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `authenticated: false` | `python -m app.mind_client login <init link>`; sign in in the Edge window once |
| The log says `assistant unavailable` | no AI key saved (paste it in the web app); names are deterministic — the run is still valid |
| `409 a Mind loop is already running` | one loop at a time; wait for it or restart the server |
| Convert waits the full 7 minutes then fails | Mind's wizard never finished; check `iter_N/mind/convert.png` |
| `Mind lists 0 template entries (selector could not be read)` | the `<stem> ▾` selector did not open; the counts gate is skipped for that iteration, nothing else is affected |
| Work dir on OneDrive | copies of a 13 MB model per iteration churn the sync — use a local folder |

---

## Reference

| Piece | Where |
|---|---|
| The loop | `app/mind_loop.py` — `run_loop`, `LoopConfig`, `decide`, `numbers_gate`, `compare_values`, `names_gate`, `untitled_counts` |
| Mind operations | `app/mind_client.py` — `open_manager`, `ensure_folder`, `create_blank_project`, `open_project`, `upload_and_convert`, `add_template`, `run_model`, `template_names`, `untitled_entries`, `export_summary`, `delete_project` |
| Safety rules in the planner | `app/prep.py` — `reference_index`, `referenced_by`, `whole_reference_to` |
| CLI | `scripts/run_in_mind.py` |
| API | `POST` / `GET /api/sessions/{id}/mind-loop` |
| UI | **Advanced › Run in Mind** (`MindReadyUI/src/app/pages/advanced/mind-run-panel.ts`) |
| Tests | `tests/unit/test_mind_loop.py`, `test_mind_loop_review.py`, `test_prep_reference_safety.py` |
