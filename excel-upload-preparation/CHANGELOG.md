# Changelog

## 1.15.0 (13 November 2026): Basic, Custom and Advanced screens, Remove all, range-read and numbers-check fixes, a copy of the project that stands on its own

### Cancel on every piece of work, and Start that does not stop for a person (17 Nov 2026)

- **Cancel.** Wherever the app works there is a **Cancel** button: Start / Do everything and each of its steps, the first check, Prep's Apply, Recalculate, the numbers check, the Mind loop and the instances builders, bringing in the outside files and making the stand-ins, next quarter, Suggest names, the assistant's answer, Upload to Mind / Run in Mind. One click, no question; the button says "Stopping…" until the work has really stopped on the server.
  - The step that was running is thrown away and the workbook is put back as it was before it: "Stopped. Nothing was changed." In the middle of Start the finished steps are kept: "Stopped. Your workbook is as it was after <the last finished step>."
  - Only the Excel the app itself started for that work is ended (checked against the list of Excels this copy of the app opened). An Excel the user has open is never touched.
  - During Upload to Mind / Run in Mind the half-made Mind project is kept: the app says its name and offers **Delete it from Mind**. Nothing is deleted without that click. The names are kept in `mind_projects.json` in the data folder.
  - When the work goes on after you leave its screen, a strip at the top says "Still working: …" with its own Cancel.
  - For developers: `app/cancel.py` (the jobs of each workbook), `POST /api/sessions/{id}/cancel`, `GET /api/sessions/{id}/jobs`, `POST /api/sessions/{id}/mind-loop/delete-project`. A cancelled blocking request answers 409 `{"detail": "Stopped. Nothing was changed.", "cancelled": true, "kind": ...}`; the status routes say `cancelled`.
- **Start does not stop for a person when the app can fix it.**
  - The Prep step repeats while a round still changes something and stops only when a round changes nothing.
  - What is still blocking then goes to the assistant. Its fix is written, then the numbers check and a new scan judge it: if any value changed, or the problem did not get smaller, the fix is undone. Only then is the person asked, in two plain sentences, with **Show me** (it opens the place) and **Send feedback** (the problem is already written in the draft).
  - Every stop sentence on Basic has a **Show me**. With no AI key saved the app stops and says so.
  - The cost: at most 5 problems a run and 2 tries for each. A try is one turn of the assistant (up to 5 model calls) plus an Apply and a recalculation of two versions in Excel, so a run is at most 50 model calls, and on a big model a try takes minutes. A problem the assistant could not fix is remembered for the workbook while the app runs and is not paid for again.
  - For developers: `app/handoff.py`, `POST` / `GET /api/sessions/{id}/prep-run`; `scripts/night_chain.py` uses the same job (`--no-assistant` runs the rounds only).
- **Review fixes (17 Nov).** A Cancel during the loop step's first moments (while its plan is built) now stops it. While a step that changes the workbook runs, Recalculate, adding input files, grid names, reports and the pivot choice answer "Wait a moment: ..." instead of being taken back by that step's Cancel. The Cancel beside Start stops only Start's own steps (`POST .../cancel` takes `{"kinds": [...]}`): a Run in Mind or an assistant answer started elsewhere keeps running. The assistant's fix is undone when a cell that showed an error now shows a value. An assistant that could not be reached says so and is asked again next time. After the loop, any blocking problem left goes to the assistant, and Start never says "every check passed" while the workbook is blocked.
- **Formula errors that show only for the Mind loop's first value no longer stop Start** (Palermo: 667 #N/A in Decrements once the loop made Excel compute policy 1, whose projection runs past the mortality table). When a loop is built and the recalculation finds new errors, the version before the loop is recalculated once too (or its recalculation reused) and the two are compared cell by cell, its cells followed through Prep's inserted rows. When every new error shows only for the loop's first value, the verdict is acceptable: Start finishes to the banner with one plain line ("667 cells show an error only for the Mind loop's first value, for example Decrements!M704: ... Check the results in Mind."). Errors that also show without the loop, or in a cell a change after the loop wrote, still stop it, counted alone ("Recalculated, 1 formula errors (the other 10 only for the Mind loop's first value)"); when the version before the loop cannot be recalculated it stops as before and says it could not be checked. Cost: one extra recalculation, only in that case (31 s on Palermo, inside a Recalculate of about 2 min), and Cancel stops it like the rest. For developers: `_loop_check` / `_before_loop` / `_written_since` in `app/web/server.py`, `recalc_state` in `app/readiness.py` (`loop_errors`, `loop_note`), `recalc_ok` in `scripts/night_chain.py`.
- **Phoenix's lone text cells are fixed by the app (18 Nov).** A cell holding only "#" with a header to its right is a column header, not a grid title (Phoenix CFs!B29: the table stays whole). Prep writes "Row N" labels into a short gap in a table's first column so the table takes in a lone cell (Phoenix B101:B103, refused when any formula could read those cells), and, as a last resort, writes the text itself into a formula in place of the reference (the same value). The lone-text check now counts only formulas that need the text (not SUM/MAX/COUNT over a range, nor a lookup that neither searches nor returns that column). Phoenix: Start reaches "ready", no value changed (347,223 formula cells compared); Palermo unchanged. For developers: `is_bare_hash` in `app/grids.py`, `_labels_closing_gap_above` / `_text_into_readers` / `readers_needing` in `app/prep.py`, `lone_text_cells_read` in `app/validators/structure.py`.
- **The numbers check after Prep no longer counts file-path texts as changed (18 Nov).** Prep writes a neutral folder or file name in place of a formula that asks Excel for its own file (FRM-012); a text built on it (Palermo 'Saved down'!C3, C4, C6, C9) now shows as expected, not as changed. It is let off only when everything it reads can be followed back, at least one cell read is a frozen one, and nothing else it reads was written by Prep or computes another value. Numbers are never let off. Palermo: no computed value changed (56,647 formula cells).
- **The header links stay on narrow windows (18 Nov).** Below about 1,050 px wide the Basic / Custom / Advanced links used to disappear. They now stay on the same line, smaller; "Milliman" drops from the name, and below about 770 px the Beta tag and the version too.
- **For developers: the exe build can run fewer compilers at once (19 Nov).** `tools/package/build.py --jobs N` hands `--jobs=N` to Nuitka, for a machine short of memory. Without it nothing changes.
- **A column heading that INDEX cannot return no longer stops Start (19 Nov).** CNHI stopped on "a formula reads a lone text cell" for 337 cells. 332 of them were one column heading, "Per Capita Cost", inside ranges that eight INDEX formulas cover. Their row number is always 1 + a whole number x 117, so INDEX never returns a heading, and Mind leaving the heading empty changes no result. Such a cell is no longer listed, and Prep no longer writes row labels for it. This applies only when the row number is plain whole-number arithmetic (plus, minus, times, brackets) over whole constants and typed-in whole numbers, or a Mind loop over such numbers. Anything else is listed as before: a division or a function in the row number, a year that is a formula, OFFSET, LET or a macro's function anywhere in the formula, a range built from the result. For developers: `prep._row_steps`, `prep._index_result_may_move`, `tests/unit/test_lone_text_stepped_index.py`; one refuter on Haiku 5.5 found the LET case.
- **A lone text read as the criterion of SUMIFS is fixed by the app (19 Nov).** CNHI's last lone text cell, 'Cost data'!AF10, is read by 20 formulas as the criterion of SUMIFS, and no title could go above it. Prep now writes the text itself into the criterion in place of the cell, as it already does in other formulas: `SUMIFS($H:$H,$C:$C,$AF10)` becomes `SUMIFS($H:$H,$C:$C,"Deliveries and inpatient services")`, the same value in Excel. This holds for the criteria of SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS and MINIFS; their ranges stay references. A text is never written beside the intersection operator (a space) or as one end of a range, where Excel would refuse it. For developers: `prep._CRITERION_FIRST`, `prep._beside_a_reference`, `tests/unit/test_lone_text_ifs_criterion.py`; one refuter on Haiku 5.5 found the two refused shapes.
- **The Mind sign-in says what to do in plain sentences (19 Nov).** While the app's own Edge window waits, the screen says where it is: "Sign in to Mind in the Edge window the app opened (also when this page is in Chrome). It closes by itself. If you do not see it, look for Edge on the taskbar." Not signed in: "Not signed in to Mind. Press Sign in to Mind and sign in once in the Edge window that opens. It lasts about a week." A window that cannot open gets one of three sentences (Edge is not installed; the app's Edge window is already open; anything else) and never the program's own error text: that goes with **Send feedback**, which now shows beside a sign-in that could not start. Advanced > Run in Mind also says where the sign-in is kept: "The sign-in is kept on this PC only, in <folder>." The app asks Windows to bring the Edge window to the front. For developers: `mind_client.sign_in_failure`; `GET /api/mind/sign-in` also returns `detail`, `feedback_ref` and `kept_in`; `core/sign-in-text.ts`.
- **For developers: the Mind sign-in after its refuter (19 Nov).** A second click while the window waits answers with the same "already open" sentence the screens use; "already open" is said for a failed start only when Edge really holds the profile (Edge keeps a `lockfile` in the profile folder while it runs: `mind_client.profile_in_use`), else the general sentence; Basic's finish banner has Send feedback beside a sign-in that could not start, as Advanced has. One refuter on Haiku 5.5.
- **"Show me" after a Cancel (19 Nov).** When you stop Start (or Do everything) with Cancel, the sentence "Stopped. Your workbook is as it was after ..." now has a **Show me** that leads to the step that was stopped: the place where that step runs by itself (the Outside files card, Fix all issues on Custom, Advanced > Prep actions, Advanced > Mind loop, Recalculate). When the click came too late to stop the running step, it leads to the next step that would have run (a step that is left out, such as a loop the model does not have, is passed over). For developers: `PLACE_OF_STEP` and `endCancelled` in `home.page.ts`, the new place `fix-all` (`main-fix-all`) in `core/show-me.ts`, `docs/SHOW_ME.md`.
- **The assistant's fixes: wider with care, and 10 rounds (19 Nov).** Two of the four limits on what Start lets the assistant write by itself were changed at Joseph's word. (1) A fix may now also rewrite cells that are not the problem's own, with care: only a formula cell that reads one of the problem's cells (for example the formula that reads a lone text cell), at most 5 such cells in one fix, never a cell whose result changes by itself (NOW, TODAY, RAND), and never a typed-in value, a range or anything that moves cells. Whether the fix is kept is decided as before: only when no computed value changed, the problem got smaller and no other problem got worse; else it is taken back. (2) The Prep rounds stop after 10 rounds at most, down from 20. Unchanged: a fix that does not make its problem smaller is undone, and the assistant gets a second try. For developers: `handoff.MAX_OTHER_CELLS`, `handoff.MAX_PASSES`, `handoff.narrow_problem`.
- **Checked in real Mind: which formulas need a lone text cell (20 Nov).** A made-up workbook (invented numbers, 30 formulas) went to Mind twice, as written and after Prep. Mind differs on every formula the app lists (10 of 10) and agrees on every one it treats as harmless (15 of 15): MAX, MIN, SUM, AVERAGE, COUNT, PRODUCT and SUMPRODUCT over a range with a lone text inside it, VLOOKUP and HLOOKUP past it, N(), the stepped INDEX of CNHI, a table whose first header is a bare #. After Prep the only differences left are on the cells the app itself says are left for a person. Two things learned about Mind's own list: a cell where Excel shows a text and Mind returns nothing is not named in it, and a sheet that holds nothing but lone text cells stops the conversion ("Spreadsheet not found"), as rule FRM-006 says. No app code changed, so the exe of 20 Nov is the same program. For developers: `scripts/build_lone_text_probe.py`, `docs/MIND_VS_EXCEL.md`.
- **CNHI: INDIRECT over a numbered list of names becomes INDEX (20 Nov).** `=INDIRECT("Cluster"&n)`, where the names Cluster1 to Cluster20 are 20 neighbouring cells, is now written `=INDEX('Validation list'!$B$7:$B$26, n)`: the same cell, and Mind does not follow INDIRECT. This works in any workbook with that shape (names Something1, Something2 ... side by side, no gap). Only when the number can only be whole and never below 1 (built from row and column numbers, counts, rounding to 0 digits); otherwise the formula is left as before. On CNHI: 2,320 of 2,320. For developers: `validators/references.py` `_named_list_index`, `_bounds`.
- **Start no longer stops after Prep (20 Nov).** Joseph: "App tries, never stops". When Prep leaves something, or no AI key is saved, Start goes on with the Mind loop and the recalculation, and the question comes at the end if the problem is still there. If a later step stops Start, the question's first sentence follows the stop sentence.
- **A label beside a lone text inside a range (20 Nov).** Where no title can go above such a text and it cannot move, Prep writes a label in the empty cell beside it, so Mind imports it. Checked in real Mind on a made-up workbook: Mind then agrees with Excel. Never where a formula could read the label's cell. For developers: `prep._label_beside`, `docs/MIND_VS_EXCEL.md`, `scripts/build_label_probe.py`.
- **Send feedback hides whole paths (20 Nov).** A folder name with spaces ("OneDrive - Milliman Inc") was cut at the first space, and the rest stayed in the message. Now the whole path is hidden.
- **Lone text cells are fixed in one round.** The Prep planner now sees the sheet as its own earlier changes leave it, so a diagonal "staircase" of lone text cells (each one became lone only after the one before it was fixed) no longer takes one round per cell.
- **Basic shows progress.** A bar under the step line fills with a percentage (a moving bar while there is no estimate yet), the button shows pulsing dots while Start runs, and a step that runs late says one thing: "This step is taking longer than usual; about 7 min left (around 22:09), maybe more."

### Also

- `scripts/night_chain.py` now skips the outside-files step the way the screen does when no cell names the files (or no macro brings them in): their data stays inside the model, and the chain goes on instead of stopping (found on Phoenix).
- **Do everything fixes again after the loop.** Building the Mind loop changes what the rules see (which policy row the model computes, the cells it clears), so a blocking action can appear that was not there before (Palermo: "an empty cell is compared with a number" went from pass to error once the loop was in). The loop step is now followed by one more Prep pass when something blocking is left; the same in `scripts/night_chain.py` (`prepare_after_loop`).
- A lone text cell read only through a formula that is wholly `N(...)` no longer blocks (Palermo `=N(VLOOKUP(...))`: `N()` turns the text into 0, as Mind's empty cell does), and the Mind upload script says "could not be read" instead of "Mind agrees" when it cannot open the list of differences.

### Plain words on screen (16 Nov 2026)

- No technical codes on the screens: a rule is said in its plain words ("an empty cell is compared with a number", not FRM-009); an Excel error keeps its value in brackets ("the lookup found nothing (#N/A)"), since that is what the workbook shows. Status tags read Problem / Fine / Warning / Not checked / Needs an answer; apply summaries say "was a problem, now fine". The codes stay in the bug reports (Send feedback) and in what the assistant is given.
- The "Works until" date is no longer in the header (the "ends soon" and "stopped" messages stay). Basic: the button is **Start** / **Start again**, the progress line says in plain words what each step does, with an ETA and "an estimate only". A **Recalculate** button appears wherever the recalculation is the only thing left; errors found after a Mind loop come with a plain note that they may come from the loop's first value.
- A thinking motion (three pulsing dots; still for people who ask for less motion) wherever the assistant works: the chat, naming the grids, Upload to Mind while it names grids, the stand-in files.
- Run in Mind's end says "Mind converted the workbook and ran the model, and the run agrees with its audit trail" (was "every gate passed"); the Mind browser keeps Edge's sandbox on, so Edge no longer warns "unsupported command-line flag: --no-sandbox".

### Three screens

- **Basic, Custom, Advanced** in the header. **Basic** is the new first screen: the upload box, one tick
  box "Build the Mind loop", one **Go** button that runs the steps of Do everything one after the other
  (each with its check, stopping at the first check that fails), one progress line and the Recent workbooks
  list, closed until you click it. **Custom** is the former main screen, unchanged. **Advanced** is as before,
  with a link back to Basic and one to Custom.
- **The end of the check says what to do next** (Basic). "The check is finished. The workbook can be uploaded to
  Mind." with **Upload to Mind** and **Download**; or "N blocking problems are left." with Download only and
  Upload greyed with the reason; or "The check is not finished" (nothing recalculated yet), Upload greyed:
  "Press Go first". **Upload to Mind** checks the Mind sign-in, opens the sign-in window when it is missing and
  then goes on by itself, runs the workbook in Mind and shows the stage and the result.
- **Show me** knows the three screens: the places that were on the main screen now point to Custom, and to the
  Go button, the list and Download on Basic.
- **Remove all** in the Recent workbooks list removes every saved workbook after one question ("Remove all N saved
  workbooks?"). A workbook another running copy of the app holds, or one with a step running, is kept, and a note
  says so. (`DELETE /api/saved-workbooks`.)

### Fixes

- **A title is never written into an empty cell that a formula reads.** Prep wrote a "#name value" title into empty
  cells of a range an INDEX/MATCH read, which changed ten computed values of a risk model. Such a table is now
  reported ("could not name it -- the empty cell above lies inside a range <the reader> reads") and left as it
  is. A whole-column or whole-row reference counts only when its function looks at text cells (lookups, MATCH, INDEX,
  COUNTA); SUM, AVERAGE, MIN and MAX skip text, so a title in such a range is still allowed.
- **The numbers check counts a replaced caption as expected.** A formula caption that Prep replaced by its title on
  purpose is no longer reported as a lost formula; the check says it was expected, and Do everything and Fix all
  issues stop on it no more. Other lost formulas are still reported.
- **CELL("filename") is frozen to a neutral text.** The folder and file name written in its place are neutral
  (same length, same separators), not the model owner's own path; every number stays the same.
- **Documents and package files no longer carry old versions and ports** (manifest, SKILL.md, README); a test keeps
  VERSION, manifest.json and SKILL.md equal. The guide lists Run in Mind from the exe as done (it needs
  Microsoft Edge on the computer) and its steps name Basic and Custom. "Main screen" in the app's texts now says
  Basic or Custom.

### The project copy

- The new copy of the project (its own folder) leaves out the retired Figma front end, the original-skill
  leftovers and the internal notes; its tests find their local files through `MIND_READY_DATA`, and the app is
  pointed at its data by `MIND_READY_SETTINGS`.

## 1.14.1 (12 November 2026): the assistant's token count, grid titles for formula captions, comment spacing

The night of 11/12 November: the items of Joseph's note for tonight, each built generally (the
anonymized IFRS 17 tool is one test case; synthetic workbooks in the tests are the proof), then a
review of the night's changes with every confirmed finding fixed (the last part of this section).

### The app shows how many tokens its assistant used

- **"AI used: 12,340 tokens in 7 calls"**, counted per workbook. Every completed call of the
  assistant -- the chat (plain or streamed), a suggestion on a finding, the grid namer, the instance
  finder's explanation, Fix all issues' advisor, the stand-in input files, next quarter's files, Run in
  Mind's grid names, the feedback's opening lines -- adds its tokens (input, output, cache read, cache
  write) and one call to the open session. One mechanism for all of them: `app/llm.py` keeps a usage
  sink in a context variable (`usage_scope`); the server sets it around every step that may ask the
  assistant (`_ai_scope` in the request's thread, `_ai_thread` for the jobs' threads, which start with
  no context of their own). The accounting never breaks a call: a sink that fails is printed and the
  reply returned as it was; a failed call counts nothing. The count is saved with the workbook at once
  (`ai_usage` in the state file; a file from before opens with an empty count) and comes back after a
  restart. Shown under Advanced > Assistant, by the AI key box, with one plain sentence on what a token
  and a call are, read again after each chat turn; and in the Send feedback message's log part; "AI
  not used yet on this workbook" before the first call. Input and output are shown together; the parts stay in the data
  (`aiUsage` in `GET /api/sessions/{id}` and after Reopen, `GET /api/sessions/{id}/ai-usage` with the
  line, `aiUsage` in the `POST /api/feedback` answer). (`app/llm.py`, `app/web/server.py`,
  `app/session_store.py`, `app/feedback.py`, `tests/unit/test_ai_usage.py` -- a fake gateway answering
  with a `usage` object, plain and streamed, so the real `chat_completion` is what is proven --
  `test_web_api.py`, `test_session_store.py`, `test_feedback.py`; `MindReadyUI` `core/models.ts`
  `aiUsageLine`, `core/api.service.ts`, `core/session.store.ts`, `pages/advanced/assistant-panel.*`,
  `core/feedback-dialog.*`, `core/ai-usage.spec.ts`, `core/feedback-dialog.spec.ts`.)

### Grid titles: a formula caption is replaced by the title (noted as an open issue)

- **A table whose caption is a formula gets its title after all** (Joseph, 11 November 2026). A key
  column with a caption on top and the block beside it are one table, titled into the caption cell;
  when that cell holds a formula (`="Running scenario ["&E1&"]"`, say) 1.14.0 left the table under
  "could not name". Now the cell loses its formula and gets '#' + the text
  the formula last showed (the value Excel stored for it, from the inventory's cached values) + the
  flags, for the automatic namer, the Grid Namer's own name and the assistant's name alike. The
  action's note says plainly: "this cell held a formula (…); it is replaced by the title. Open issue:
  if the caption must stay a formula, leave an empty row above the table instead." When the formula's
  last value is empty, an error or not text, the table is reported instead (see the review fixes below).
  Nothing else in the title rules changes: no row is ever inserted, and a caption cell something reads
  is still refused with the reader named. (`app/prep.py` `_caption_shown`, `_title_site`,
  `plan_create_grid_titles`; `app/grids.py` docstring; `docs/RUN_IN_MIND_NAMING.md`, `README.md`;
  `tests/unit/test_grid_titles_never_move_cells.py`: the formula caption replaced with the cached text,
  an empty, error or number cached value -> the table reported, the Grid Namer's name replaces it too, a read
  caption still refused.) Proven on the anonymized IFRS 17 model: 408 title operations, 0 row inserts,
  one operation on its formula-caption table (the caption cell -> '#' + the text it showed), nothing
  reads that cell (the app's index and a brute-force pass over 110,194 formulas and names: 0), both copies recalculated in Excel: 506,746 cells
  compared, 0 differences with the plan's writes allowed; the app's numbers check over 110,174 formula
  cells reports exactly one change, that cell itself, no longer a formula (`%LOCALAPPDATA%\MindReady\iteration_1_14_1\namer\`).

### Comment spacing

- Comment spacing kept consistent by a pre-commit hook (`tools/comment_spacing.py`,
  `tools/install_hooks.py`, `.githooks/pre-commit`).

### Review fixes (12 November)

- **Fix all issues' assistant calls are counted.** The advisor asks the gateway from a thread pool, and
  a pool's threads start with no context: its calls reached no usage sink, so the heaviest spender in
  the app counted nothing. Each batch now runs in a copy of the caller's context
  (`app/autofix_advisor.py`; `tests/unit/test_ai_usage.py`).
- **No made-up name over a formula caption.** When a formula caption's last value is empty, an error
  or not text (a workbook never calculated, a year), the table is reported with what to do --
  recalculate the workbook and name it again, or leave an empty row above the table by hand --
  instead of losing its formula to a name made from the surroundings; a literal year in that cell was
  protected already (`app/prep.py`, `app/grids.py`; `tests/unit/test_grid_titles_never_move_cells.py`).
- **A caption cell an INDIRECT or a 3-D reference may read is refused**, text or formula, naming the
  reader: the index pins such readers to a sheet, not to a cell, and a title there could change what
  they read -- the rule the fixer's moves followed already (`app/prep.py`).
- **The Assistant tab says what a token and a call are** in one plain sentence beside the count; the
  Send feedback dialog shows the token line once, inside the message, not again under it
  (`MindReadyUI` `pages/advanced/assistant-panel.html`, `core/feedback-dialog.*`, the two specs).
- The formula-caption rule is stated with a synthetic example, before the list of "could not name"
  reasons (`docs/RUN_IN_MIND_NAMING.md`, `README.md`).
- The comment-spacing hook works on the staged copies and leaves a partial commit alone; the tool's
  rule starts again at every top-level block, so a diff stays within what was edited.

## 1.14.0 (11 November 2026): saved workbooks, Do everything asks for the input files, skipped sheets, grid titles, Send feedback

The night of 10/11 November: the five items of Joseph's note, in his order, each built generally and
proven on synthetic workbooks (and on the anonymized IFRS 17 tool where the report came from), then a
review of the night's changes with every confirmed finding fixed. No Mind run, no exe.

### Saved workbooks

- **Workbooks are kept between sessions.** A session's files now live in the data folder
  (`<data folder>\sessions\<id>\`, the same layout the %TEMP% folder had) next to a state file the app
  writes after every step that changes the session: the upload, each scan, each Apply, the end of Fix
  all issues, of Bring in the outside files and of next quarter's files, a recalculation, the pivot
  setting, the grid names, the model map, the instance finder, the instances build, the reports, added
  input files. A restart of the app or a new browser window loses nothing. (`app/session_store.py`,
  `docs/SAVED_WORKBOOKS.md`, `tests/unit/test_session_store.py`.)
- **Recent workbooks, on the main screen.** Under the upload box, only when at least one workbook is
  saved: name, when it was last opened, how many versions, **Reopen** and **Remove**. Reopen loads the
  workbook with everything it had -- every version, the last check result, the plan, the last
  recalculation (Advanced > Recalculate shows its list without recalculating), the Mind loop and
  instances plans, the outside files and their results, the downloads -- and shows the same screens as
  after its last step; nothing runs by itself. Remove asks once, in place ("Remove <name> from this
  computer?" Yes / No), then deletes the folder. On start the app shows the list and waits. A job that
  was running when the app was closed shows as stopped: "The app was closed while this ran." A browser
  tab still on an old `?session=` link is told "Open it again from Recent workbooks on the main screen",
  with a Show me to the list. (`GET /api/saved-workbooks`, `POST /api/saved-workbooks/{id}/open`,
  `DELETE /api/saved-workbooks/{id}`; `MindReadyUI` `pages/recent-workbooks.ts`,
  `core/recent-workbooks.spec.ts`, `core/show-me.ts`.)
- **Review fixes (saved workbooks).** (1) The check result -- the whole analysis, 87 MB of JSON on the
  anonymized IFRS 17 tool -- is no longer written into `state.json` after every step: it has a file of
  its own, `result.json`, written only when a scan replaced it; the state file stays small, is written in
  one pass (no more dumps-loads-dumps), and the Recent workbooks list reads the small file alone. The
  plan -- every formula fix as an operation, 61 MB on that tool -- has its own file too, `plan.json`,
  written only when a scan or a re-plan (the pivot setting, the grid names) replaced it: the state file
  is under 1 MB on that tool instead of 63 MB, and a step that keeps the plan (an Apply without a
  re-scan, a reopen) no longer waits seconds for it. A state file of the first shape (the result or
  the plan inside) still opens. (2) Two copies of the app on the same
  data folder (two ports, say 8600 and 8601) no longer list, reopen or remove each other's open
  workbooks: the copy that holds a workbook keeps its `open.lock` locked (the OS lets go of it when the
  process ends, however it ends); the other copy's list leaves it out and its Reopen / Remove answer 409
  "This workbook is open in another copy of the app (port 8600)". (3) Remove is refused while Run in Mind
  runs on the workbook (it was not in the busy check). (4) A Fix all issues the restart cut shows on the
  main screen after the reopen: "Stopped: the app was closed while it ran", with the version it had made
  when it got that far. (5) The sentence "This workbook is not open any more (the app was restarted).
  Open it again from Recent workbooks on the main screen." is now shown where it can be acted on: any
  call about such a workbook makes the page let go of it and the main screen shows the sentence above
  the Recent workbooks list, with its Show me (before, it sat in a tile of the old page where the list
  is not drawn and the Show me found nothing). (6) A saved workbook that was uploaded but never checked
  is scanned after Reopen without starting Do everything, whatever the "Do everything right after the
  upload" box says. (`app/session_store.py`, `app/web/server.py`, `tests/unit/test_session_store.py`;
  `MindReadyUI` `core/restart.interceptor.ts`, `core/restart.spec.ts`, `core/session.store.ts`,
  `pages/home.page.*`, `core/recent-workbooks.spec.ts`.)

### Do everything and the input files

- **Do everything stops and asks for the input files, then goes on from where it stopped.** When its
  first step finds that files the model's macro opens were not picked, it no longer has the assistant
  make stand-ins quietly: the server's bring-in job waits (`waitForInputs`, state `waiting_for_inputs`;
  nothing runs, the workbook is unchanged) and the main screen shows one box under Do everything --
  "Do everything is waiting for 2 input files: BBA_05, VFA_01. Pick them, or let the assistant make
  stand-ins." with the list (picked / missing, what each fills), **Pick the files**, **Go on** (greyed
  out until every file is there), **Make them with AI** (only with an AI key saved; the stand-in path
  with its warning) and **Stop**. Go on (`POST .../outside-inputs/continue`) starts the job at the
  import itself -- the stages after the wait and only those (`stages` in `GET .../outside-inputs`);
  the chain was held at this step, so the other steps follow. Stop (`POST .../outside-inputs/stop`)
  ends the job in plain words and Do everything stops at its first step. A page reloaded while it
  waited gets the box back; a workbook reopened after a restart shows the job as stopped ("The app
  was closed while Do everything waited for the input files."). The card's own buttons and the
  plain run are as before. (`app/web/server.py`, `app/session_store.py` `_stopped`,
  `docs/OUTSIDE_INPUTS.md`, `tests/unit/test_outside_inputs_api.py` section 8; `MindReadyUI`
  `pages/home.page.*`, `core/models.ts` `outsideWaitingLine`, `core/api.service.ts`, `core/show-me.ts`
  `outside-waiting`, `core/outside-inputs.spec.ts`, `core/show-me.spec.ts`.)
- **Review fixes (Do everything and the input files).** While Do everything waits for the input files the
  countdown pauses ("4 min so far · paused: waiting for the input files"): the wait is no longer counted
  as the step's time, so it neither says "this step is taking longer than usual" nor inflates the
  estimate of the steps after it. Go on is disabled while its request is on its way (a double click sent
  a second continue that the server refused, and the refusal stayed on the Outside files card for the
  rest of the run), and the box goes as soon as the server took the Go on. Make them with AI is offered
  only while a file is missing (with every file there it did nothing of what it said). (`pages/home.page.*`,
  `core/outside-inputs.spec.ts`.)

### INDIRECT with skipped sheets

- **A sheet left out of the scan still exists for the formulas that point at it.** Joseph ticked the
  six calculation sheets of the Tokyo IFRS 17 tool as "do not check"; the 480 INDIRECT lookups on
  Result_CF that build addresses on them (`INDIRECT($A3&"!A:B")`) were then written as
  `INDEX($A$1,NA())` instead of `BBA_01!A:B`, as if the sheets were not in the workbook. The FRM-007
  fix looked the target sheet up among the scanned sheets only. Now a skipped sheet counts as existing
  wherever a target sheet is looked up (its name and tab index are known); what was never read about it
  (its cells, its grids) is not guessed: an INDIRECT whose target is on a skipped sheet becomes the
  plain reference exactly as without the skip and the note says the sheet was left out of the scan; an
  INDIRECT whose text needs a cell of a skipped sheet is left for a person, with that reason; a sheet
  that truly does not exist still becomes NA(). A 3-D range (`'First:Last'!`) runs over a skipped sheet
  at its own tab position (it used to be put at the end of the tab order). The Mind-value rules
  (FRM-008, 009, 010, 013) no longer read a skipped sheet's cells as empty: the formulas that read one
  are left as they are and the finding says how many and why. Proven on the anonymized IFRS 17 tool:
  with the six sheets skipped, Result_CF's formulas after Fix the formulas are identical to the
  no-skip run's (523 formulas, 0 differences; 480 before the fix), and the rest of the plan differs only
  in the actions on the skipped sheets' own cells (`%LOCALAPPDATA%\MindReady\iteration_1_14_0\indirect_skips`).
  Synthetic proofs in `tests/unit/test_skipped_sheet_references.py`. (`app/inventory.py` `sheet_order`;
  `app/validators/references.py` FRM-006 and FRM-007; `app/validators/mind_values.py` `_Book`;
  `app/prep.py` planner notes; `rules/formula-rules.yaml` FRM-007.)
- **Review fix (skipped sheets).** With two or more sheets skipped, the Mind-value rules (FRM-008, 009,
  010, 013) answered "not checked" with a TypeError (their cache key sorted the skipped sheets as dicts)
  and the four Prep actions planned nothing -- the very case the skipped-sheet handling was built for.
  One line; proven with two skipped sheets in `tests/unit/test_skipped_sheet_references.py`.
  (`app/validators/mind_values.py` `scan`.)

### Grid titles

- **Grid titles never move cells: one title per table, written into the header cell.** Joseph's
  report (the IFRS 17 tool's `Result_CF`): after the automatic grid namer the key column (column A,
  the sheet name each row's lookups use) sat one row below its rows. Reproduced on the anonymized
  model: the app read two untitled grids there, the column `A1:A8` (a caption formula on top of the
  keys) and the block `B2:CI8`; the title action inserted cells `A1:A1` over column A only to make
  room for its title, so column A moved down and the block stayed. Three changes, all general:
  (1) a key column with a caption on top and the block beside it that reads it (or that holds plain
  keys) are **one table** from the scan on (`grids._merge_key_columns`: one column wide, directly
  left of the block, starting one row above it and ending on its last row); two lists that only
  touch, differ in length, or where the column is already titled stay apart. (2) The title goes into
  the one cell Mind reads a title from, directly above the table's top-left cell, and **nothing is
  ever inserted to make room**: an empty cell above gets the made-up title; a caption on top of the
  key column becomes the title with its own text ('#' + the caption); a caption that is a formula, a
  grid on row 1, a cell above that belongs to another grid (a section heading row), and a title that
  would touch a cell Mind reads on through (the cell right above it, or the cell beside it on the
  left, when that cell is a lone text or a grid starting there; a lone label nothing reads is moved
  into the title instead) are reported under "could not name" with what to do by hand (leave an
  empty row, move the formula). The old namer wrote those touching titles: the app's own re-scan of
  its output on the anonymized model found 85 grids with a `#Title` trapped inside (122 cells),
  which the next scan would have "freed" with more row inserts. (3) The Grid Namer's user-chosen
  names follow the same rule (the user's name replaces the caption's text; the same refusals). The
  "a title there would change that result" checks stay; the apply path still carries out an older
  plan's `insert_row` (STR-001's separate-the-grids action still inserts rows, as before). On the
  anonymized model the new plan writes nothing on `Result_CF` and says why (`A1` holds the caption
  formula `="実行中のシナリオ　["&Ctrl_01!$C$27&"]"`, so the table `A2:CI8` cannot be titled
  without replacing it), no row is inserted anywhere (12 before), the titled copy's own re-scan
  holds no trapped title of the namer's making, and an Excel recalculation of the titled copy
  against the source shows 0 changed values (506,766 cells;
  `%LOCALAPPDATA%\MindReady\iteration_1_14_0\namer`). Synthetic proofs in
  `tests/unit/test_grid_titles_never_move_cells.py`; the older tests that expected a row insert now
  expect the report. (`app/grids.py` `_grid`, `_reads_column`, `_merge_key_columns`; `app/prep.py`
  `_title_site`, `plan_create_grid_titles`, `plan_named_areas`; `README.md`;
  `docs/RUN_IN_MIND_NAMING.md`.)
- **Review fixes (grid titles).** (1) The caption on top of a key column must be text: a year, a date or
  TRUE above a key column is a value something may read (a macro, a dropdown, a person) and is never
  merged as the caption, so no title is written over it (the column is reported as a grid on row 1, as
  before). (2) A lone label beside the title cell is moved into the title only when the title's name was
  made from its own text; the Grid Namer (the user's own name) and the assistant's name leave the label
  where it is and report the grid ("... beside it, which is a label of its own"), instead of clearing
  the label and losing its words. (3) The merge of a key column and its block looks at the shape alone
  (one column, directly left, one caption row, same last row): the keys may be formulas the block no
  longer reads -- on the IFRS 17 tool's `Result_CF`, once FRM-007 has rewritten `INDIRECT($A3&"!A:B")`
  as `BBA_01!A:B`, the next scan still reads one table, so the table stays nameable after the formula
  fix. (`app/grids.py`, `app/prep.py`, `tests/unit/test_grid_titles_never_move_cells.py`.)

### Send feedback

- **Send feedback, without a team name.** The box under the header, the dialog and the docs no longer name
  a team: the button is **Send feedback**, the hint "You can send the details: where it
  stopped and why, with nothing from your workbook.", the dialog is titled Send feedback and its sent step
  says "An email to <address> should be open." with the address the backend gives (`feedback.email` in the
  settings, unchanged: the reply still comes back to Joseph). The assistant's opening lines of the message
  are told to say what happened in their own words and never to copy the user's note; a sentence that still
  repeats the note word for word is dropped from the opening (the note has its own part, "The user's
  note"). The unit tests write their incidents and timings into a folder of their own, never into the
  user's `%LOCALAPPDATA%\MindReady` (`tests/unit/conftest.py` points `SETTINGS.data_folder` at tmp_path).
  (`app/feedback.py` `assistant_lines`, `without_note`, `compose`; `MindReadyUI` `app.html`,
  `core/feedback-dialog.*`; `core/feedback-dialog.spec.ts`.)
- **Review fix (feedback).** The remaining "Tel Aviv team" mentions are gone from what ships: the user
  guide (`tools/package/guide/guide.html`: the byline says "from Milliman", the button is "Send
  feedback"; `render_guide.py` falls back to "the Mind Ready feedback address"), the server's docstrings
  and the screens' code comments. The SharePoint paths in `internal-notes/` are folder names and stay.

## 1.13.0 (10 November 2026): Mind loop and Mind instances, with plainer screens

One version for the work of 9 and 10 November: the Mind loop and Mind instances builders as Joseph will
see them (one sentence and one button each, the choices folded away, a tab of their own for the
instances), copies of the model from the macro's loops (a scenario x a cohort), each loop with its own
choices and one macro at a time, next quarter's Mind input files, three Run in Mind fixes, and the
macro's text read in its own code page. Checked in Mind on small invented models and on the anonymized
IFRS 17 tool (24 copies, 12,528 cells, 0 differences from Excel). The 1.13.0 / 1.14.0 numbers the side
branch used in October are taken by this release; the work there is in this section.

### The screens: Advanced > Mind loop and Advanced > Instances

- **Advanced > Mind loop opens with one sentence and one button.** "The macro runs CohortRng 2021 to
  2026 and Scenario 0 to 3. Mind can do the same." and **Build it (N formulas)**. Everything else is
  folded: **Change the choices** (each loop's Name / Runs from / To / A cell for the number of passes,
  the Macro dropdown only with two or more macros, the two options, Preview) and **What gets written
  (N)** (the output cards, the cleared areas, how it runs, the repairs, Show the formulas). What MM_LOOP
  and MM_RESULT are sits behind **What is this?**, on the tab and on the main screen's loop tile. A loop
  that still needs one answer shows only its missing box above the folds (a cell for the number of
  passes, else the first value, else the last), with the plan's reason in plain words as the label, a
  one-click **Use <cell>** and a Check button. After a build the result line adds "Done. Next: Run in
  Mind" with a Show me to that tab. Words: "needs one answer" (was "needs a decision"), "left as is"
  (was "not built"), "the cell the macro changes" (was "driver"). (`MindReadyUI`, `loop-panel.*`,
  `core/loop-tab.spec.ts`.)
- **The instances have a tab of their own, "Instances"**, right after Mind loop; one line at the top
  of the Mind loop tab links to it (one model that loops, or one copy of the model per scenario). The
  tab asks one thing at a time. Step 1: **Make one copy per...** -- tick the macro's loops, or pick a
  column of a table from a dropdown; nothing else is drawn until there is a choice, except two closed
  folds: **Not sure? Let the app look** (the finder: the top suggestion with **Use it**, the rest
  behind **Show all N**, the assistant tick and **Look again** inside) and **What is this?**. Step 2:
  one line "N copies - M result cells", the status chip and **Build it (N changes)**; the cards (Copies
  from, One copy per <key>, Loop, Output block, Main workbook), the "Also keys with no rows" tick, the
  emptied areas, how it runs and Show the changes sit under a closed **Details** fold. Step 3, after
  the build: "Done. Download both files (zip). Next: Run in Mind - upload both files as one project",
  with the download button and a Show me to Run in Mind. MM_INSTANCE is named only behind What is
  this?. (`instances-panel.*`, `core/instances-tab.spec.ts`.)
- **The main screen's loop tile** says "Mind runs no macro. This writes the macro's run loop into the
  model, so Mind runs every pass itself." with a What is this? fold for the two functions; its summary
  reads "1 loop needs one answer (Advanced > Mind loop)" or "2 loops need one answer each".
- **Show me** knows the new places: on the Mind loop tab the answer box (else the choices fold), Build
  it and the result line; on the Instances tab the question, Build it and the result line. "Advanced >
  Instances" and "Advanced > Mind loop" in a message are told apart (`core/show-me.ts`).
- **After the review of the two tabs** (the same evening; `core/loop-tab.spec.ts`,
  `core/instances-tab.spec.ts`, `core/show-me.spec.ts`):
  - The finder's fold **Not sure? Let the app look** is there with or without a run loop, and when the
    plan could not be read: it reads the formulas, not the macro (a workbook with repeated year sheets
    built by hand had lost it).
  - The Mind loop tab starts from the choices the Instances tab followed, so coming back from Instances
    keeps the answer. Before, the tab was made anew with empty boxes, the answer was wiped and the
    Instances plan dropped the loop that needed it without a word (24 copies became 6).
  - The main screen's loop tile counts only the loops that wait for an answer. Another macro's loops
    and the ones left as is need none: on the IFRS 17 tool the tile said "2 loops need one answer
    each" and the tab had no answer to give.
  - A loop whose two values are known and still cannot run (the last before the first) gets both
    boxes with the plan's reason above them; the sentence never promises an answer box that is not
    there.
  - **Try again** under "Could not read the plan", on both tabs (a scan still running, a session opened
    again).
  - While a loop still waits for its answer, Build it stays on and says under it "Builds CohortRng
    only; Scenario waits for the answer above."
  - **What gets written (N)** counts each cleared cell once (they are operations already).
  - Both tabs show the plan's own notes: Prep's blocking items that are not formulas as one line above
    Build it, the rest behind **Notes (N)** -- the two notes added in this version reached no screen
    before.
  - The Instances tab's question and notes go through the same plain words as the loop tab
    (`core/plain-words.ts`: "the cell the macro changes", "a cell for the number of passes", "needs one
    answer", "Mind loop tab"); the backend's own notes now say "tick the loops under Make one copy
    per..." and "the Mind loop tab" (`app/instance_builder.py`).
  - "click Build it" in a sentence that names Advanced > Instances points at the Instances tab's
    button, not the loop's.

### Mind instances: copies from loops (a scenario x a cohort)

- **Copies from loops** (`app/instance_builder.py`, the Instances tab's question "Make one copy per...";
  `GET` / `POST .../mind-instances-plan` take `loops`, `results` and `loop_choices`). For a model whose
  macro runs the whole model once per scenario and cohort (or any counter or list loop), not one row
  of a table at a time: one copy per combination of the chosen loops' values. In the model: a split
  grid `#Instance values /InstanceSplit` two rows below the used range of the first loop's sheet (the
  key, then each loop's value, one row per copy), each chosen driver `=INDEX(<its column>,1)` (the one
  row Mind keeps in the copy; in Excel the first copy), the driver's dropdown removed, a cell with
  `MM_INSTANCEKEY`; the macro's other loops are written as `MM_LOOP`s, as the Mind loop plan has them,
  so they run inside every copy. The main workbook lists the keys and shows one block of the model for
  every copy -- the formula block of the output sheet by default, or the range asked for -- one row per
  copy and block row, `MM_INSTANCE(<cell>, <key>)`; no sum over the copies (it would add scenarios up).
  The notes say what Mind does with the model's `/Input` grids (a file named `<key>-<grid>.csv` changes
  one copy), that every copy computes (also the passes the macro skips), and name any cell that still
  uses a chosen loop in an `MM_` function. Every plan lists `loop_options` (each loop, and why copies
  cannot be made from it); a plan by the table that cannot be built points to them. The Mind loop
  plan's operations now say which loop they belong to (`loop`).
- **In Mind, the anonymized IFRS 17 tool as 24 copies** (scenario 0-3 x cohort 2021-2026; the group and
  roll-forward-year loops stay loops), built through the app's API from the prepped model:
  `zzz_mindready_tokyo_instances_1` converted in 129 s, its upload test shows the same 63 cells as the
  model without instances (rounding residues below 4e-6 and 8 "TRUE" text cells), the run took 3 min
  06 s, consistent with the audit trail. The main workbook's export after the run against Excel with
  each copy's cohort and scenario set in the driver cells: 24 x 6 x 87 = 12,528 cells, 0 differences.
- **Checked again in Mind on 10 November 2026** with three small invented projects
  (`docs/MIND_INSTANCES.md`): each copy runs MM_LOOPINSTANCE to its own count; `MM_INSTANCE` reads one
  copy by key cell, literal key or index, and on a looped cell gives only the first pass -- the builder
  points it at the `MM_RESULT` cells, so it is right; a plain link is the sum over the copies. Nothing
  in the builder had to change.
- **Written with the formula repairs Prep still has**, as the Mind loop is now: a version built before
  Prep goes to Mind with CELL / INFO frozen, OFFSET ranges written out and the rest of the "Fix the
  formulas" set, so Mind computes what Excel does. A cell the builder writes keeps the builder's
  formula; a plan that is not ready writes nothing. The plan's notes list the repairs.
- **Copies from loops** keep the clears of the macro's paste areas; copies' keys are always different
  (by position when two values would clash).
- **Prep's repairs with the instances** are compared with the plan's cells at their addresses before
  the plan's row and column inserts: no repair is dropped or overwritten because of another cell that
  only shares an address.
- **Screens**: the Instances tab sends the Mind loop tab's choices (`loop_choices`, `macro`) and
  previews again when they change; a new preview drops the earlier write's download link.
- **`scripts/mind_upload.py`**: the log reader clicks "Hide Logs" only. It took "Close", which closes the
  upload window without adding the templates, so `--run` found no Run button.

### Mind instances: one copy per key of a table (the builder)

Built in October on a side branch as 1.13.0 and 1.14.0 and never released; brought onto this branch in
this version and fitted to what it has since (outside inputs, stand-ins, Show me, plain codes, per-loop
choices).

- **Build the Mind instances** (Advanced > Instances; until the screen pass of 10 November a section
  below the loop panel). The macro's run loop turned into Mind instances: one copy of the model per
  value of a key column of the model-point table (a tariff, a product, a group), a loop inside each
  copy over that copy's own rows,
  and a second small workbook -- the main one, which Mind does not copy -- with the keys and the results.
  Preview, write it into a new version, then download both files as one zip and drop them together on
  an empty Mind project (`app/instance_builder.py`; `GET` / `POST /api/sessions/{id}/mind-instances-plan`
  `[/apply]`, `GET .../mind-instances/download`; `docs/MIND_INSTANCES.md`).
- **What changes in the model.** The model-point table becomes a split grid: a key column as its first
  column, the key of every row under the header `Instance key`, the title `#<Name> /InstanceSplit`
  directly above and an empty row above the title (the column and the rows inserted where needed). The
  loop's count cell gets `=COUNTA(<key column>)`, so each copy counts its own rows; a free cell holds
  the biggest count -- the loop size -- with a label; the driver cell gets
  `=MM_LOOPINSTANCE("<loop name>", <size cell>, <count cell>)` -- plus 1 when the range the loop
  indexes starts on the header row (the macro ran `2 To n`), so every pass reads a data row. Every
  formula of the block the macro copies out on each pass becomes `=MM_RESULT(<its source cell>)`, the
  sum over the copy's loop; the macro's paste area is emptied; a free cell shows the copy's key,
  `=MM_INSTANCEKEY("<first key>")`, with a label. Filled columns touching the table on its left are part
  of the grid in Mind's eyes, so the key column goes before them (and a headerless one gets a header).
  Cells that only read the paste area are left as they are and listed (they show 0 in a copy).
- **The main workbook** `<stem>_Main.xlsx`, written through Excel next to the model so its references
  resolve: sheet Instances with `#<key header> instances /InstanceKeys` (the model's file name as header,
  one row per key); sheet Summary with `#Results by <key header>` (one row per key, one column per row of
  the output block, each `=MM_INSTANCE(<model cell>, <key>)`) and `#All copies` (the whole output block
  as plain links to the model, its header row kept, with a Total row: in Mind a plain link to a copied
  workbook is the sum over the copies, checked in a real project). Keys stay what they are: a text
  key like `001` is written as text in the model's key column and in the main workbook alike. In Excel
  every grid shows the one copy Excel has.
- **The choices.** The key: a dropdown of the table's text columns with 2 to 200 distinct values, each
  with its row count. The default is the column the macro itself looks the loop row up in, else a header
  that says tariff, product, group, portfolio, segment, class, category, type, country, fund, treaty or
  cohort, else the column with the fewest values. "Every key of the list, even without rows" makes a copy
  per value of the list the macro reads, when the app finds it. The model is uploaded as
  `<stem>_instances` with its own extension; Mind finds the workbook to copy by that name, so the files
  keep their names.
- **A loop without a count cell** is not built: the plan says so in one sentence, and the loop panel's
  choice of a count cell comes first. The builder takes the first counter loop of the loop plan that has
  one.
- **A new change, `set_range_values`** (`app/prep.py`): one range written in one go, through Excel or
  openpyxl, counted in the change log like the other changes. The key column of a long table is one
  change, not one per row. After the write through Excel the last cell planned to hold something is read
  back (not the range's last cell, which is empty by design when the table's last rows have no key).
  `excel_text` now keeps every string a string: text Excel would turn into a number, a date or a
  boolean (`001`, `1E3`, `2024-01`, `TRUE`) gets the text prefix.
- **A refused change** (the key column, the driver, a result cell) leaves a model the plan does not
  describe: the version is kept, but the main workbook and the zip are not made, and the answer says
  which changes were refused. A paste-area cell that stayed does not stop the zip.
- **Two fixes on the way.** RZS-006 no longer counts an `/InstanceSplit` grid as resizable: a copy keeps a subset of the template's
  rows, never more, and Mind reads a fixed range over the whole key column as the copy's column -- so the
  readiness verdict no longer turns red right after the instances are built. The outcome card names the
  Mind loop and Mind instances steps.
- **In Mind:** the two files of the hand-built prototype this builder mirrors, dropped together on an
  empty project, converted, tested and ran as one (28 October 2026; the results-by-key grid matched the
  macro's sums by key). This builder's own two files, made through the API from the fixed model cut to
  the macro's 19 test policies (29 October 2026): converted and tested as one project, run in 3 s, the
  results-by-key grid (3 x 128) and the all-copies grid (128 x 10) equal to Excel's sums by tariff, 0
  differences, Total row equal.
- The feedback message names the step "Build the Mind instances". The guide for the Mind team has a
  section on it.

### Where instances could help (the finder)

- **Find where instances could help** (Advanced > Instances, the fold **Not sure? Let the app look**;
  until the screen pass of 10 November a button below the builder). The app reads every formula of
  the model and its macro and points to what repeats: a column of the looped table that splits its rows into groups (one copy of the model per
  value), the same formulas side by side in one sheet (bands of columns or of rows), sheets with the
  same layout and formulas, and near copies that differ only in typed-in numbers or in a few cells of
  logic -- with the edits that would unify them. Each place is a card: one plain sentence, whether
  instances, a loop or leaving it as it is fit (yes, partly, no) and why, and **Use as the key**,
  which fills the builder's key column and previews the instances. The finder is blind -- no hint of
  what the key is -- and deterministic: pure Python over the scan the app already made, no Excel, no
  numpy, no model call; the same workbook gives the same answer. It writes nothing. With an AI key
  saved, the tick **Explain and rank with the assistant** adds a plain explanation and the assistant's
  order on top; the findings themselves never change (`app/instance_finder.py`,
  `app/finder_shapes.py`, `app/finder_patterns.py`, `app/finder_grouping.py`;
  `GET /api/sessions/{id}/instance-finder` `[?assistant=1][&refresh=1]`;
  `python -m app.instance_finder <workbook> [--json] [--assistant]`; `docs/INSTANCE_FINDER.md`).
- **Merged into dimensions** by the input the copies step through, never by labels. The column the
  macro looks the loop row up in, the list it matches the value into and the slots it pastes each
  pass into are one dimension; so are blocks on other sheets whose copies read consecutive rows or
  column slots of that area, including blocks that step by a computed index -- `OFFSET` with
  `ROW()`, `COLUMN()`, an expanding `COUNTA` or a `MATCH` on the key -- and two lists that hold the
  same values. Lookups keyed by the current row's value, also through a mapping table, and formulas
  that branch on it with `IF(<key>="...")` count as evidence; a header word is shown as a hint and
  weighs nothing. Chains (each copy reads the one before it) are set aside with the reason; time axes
  and fills that join nothing are counted, not reported.
- **Cells reading the macro's paste area, by class**, each with what it needs in Mind. Cells that add
  every slot up move to the main workbook as plain links to the copied model -- there a link is the
  sum over the copies; inside a copy they would see only that copy. Rows that each read one slot by a
  computed index become one row per copy in the main workbook. The block that reads the current key's
  slot and adds the pass to it stays as it is: in a copy it is the copy's own block, and `MM_RESULT`
  does the adding. Cells reading one fixed slot need `MM_INSTANCE(cell, key)`. The builder
  listed all of these as "left as they are"; the finder says what to do with each.
- **Checked on the use case** against the builder's own choice (`docs/INSTANCE_FINDER.md`, "The
  blind result on the 2c use case"): the finder runs blind on the same session as the builder's plan,
  and the two JSON answers are set side by side, fact by fact -- the table, the key column and its
  header, the evidence, the distinct values and rows, the other candidates, the cells reading the
  paste area by class, the time taken. `builder.matches_top` says in the JSON whether both pick the
  same column. On the use case it does: the finder's first place is the builder's key column on the
  19-row cut, on the full fixed model (26 groups over 94,932 rows) and on the raw file, with the
  macro's grouping as its first evidence; the finder adds what the builder leaves as it is -- which
  cells read the paste area and what each needs in Mind -- and the lookups and branches keyed by the
  column.
- **Limits.** Chains are set aside, not solved. Formula counts are not run time: Mind's speed and
  memory are not predicted (the biggest copy's row count is the nearest hint). The request is one
  long GET -- up to about two minutes on 350,000 formulas, no background job yet; a step the budget
  cut is written in `skipped` in plain words, nothing is skipped silently. The app keeps two workbooks
  open: after a third upload or an Apply the finder has to read the workbook again before it starts
  (minutes on a 28 MB model) and says so, with the seconds, in its answer. Sheet copies are matched at
  the same addresses, so a copy shifted by a row is not seen. Copies across separate workbooks are not
  covered. A block or sheet dimension is reported with its edits; the builder builds grouping columns
  only.

### Each Mind loop with its own choices, one macro at a time

- **One macro's loops at a time** (review of 10 November 2026; `docs/MIND_LOOP.md`). Excel runs each
  run macro on its own, with the other macro's drivers fixed. The Mind loop plan wrote the loops of
  every macro into one model, so Mind ran them inside each other (the IFRS 17 tool: 6 x 4 x 6 x 6
  passes) and the values could differ from Excel. Now the plan builds one macro (`macro=` on
  `GET` / `POST .../mind-loop-plan`; a **Macro** dropdown on the Mind loop panel), lists the others'
  loops as not built, and names a driver of another macro that already holds a loop. Copies from loops
  take the loops of one macro (two macros: refused), pick that macro's loop plan, and no longer write
  another macro's loops into every copy. The Tokyo instances built before this fix
  (`zzz_mindready_tokyo_instances_1`) had the group and roll-forward loops inside every copy. Built
  again through the API (91 changes, no other macro's loops) as `zzz_mindready_tokyo_instances_2`:
  Mind's upload test 57 cells (a subset of the 63: rounding residues below 4e-6 and the 8 "TRUE"
  cells), run consistent with the audit trail, export against Excel 12,528 cells, 0 differences.
- **Loops told apart by procedure and loop variable**: two macros both using `i` no longer overwrite
  each other, and an output block is written by its own macro's loop.
- **Count cells**: a count cell that already holds the count this plan writes is accepted (the loop
  can be planned again after a write); the count's label never goes on another loop's count or label,
  a cell the plan writes, or a cell a formula reads.
- **A first value typed as a cell** counts as chosen, like a number: the header pass is kept.
- **Advanced > Mind loop: every loop has its own boxes** (under **Change the choices** since the screen
  pass: Name / Runs from / To), and a box for a cell for the number of passes when the loop asks for
  one (with "Use <free cell>"). First and last take a number or a cell
  (Sheet!A1). Sent as `loops` by driver cell (`GET .../mind-loop-plan?loops=<json>`, POST
  `{"loops": {...}}`); the old shorthands (`name`, `first`, `last`, `count_cell`) still work.
- **More loops are built from the workbook alone** (seen on the Tokyo tool):
  - a driver set to item i of a named list (`GroupGrid(i + 1, 1)`): MM_LOOP over the list, its end
    counted as the macro counts it (to the last filled cell);
  - a counter whose own dropdown lists its values (0, 1, 2, 3): MM_LOOP over that list (fixed range);
  - a bound reached through plain references (`=C10`, `=CalcYear`) reads the cell holding the number.
- **A workbook-level name wins over a sheet's own name of the same spelling** (the macros'
  `ThisWorkbook.Names("GroupName")`): the loop goes in the cell the macro writes, not a copy.
- **Count cells are safe**: never a filled cell, a driver, or another loop's count cell; a label
  ("<loop> count") is written to its left; a loop starting below 1 counts 1..n and shifts
  (`=MM_LOOP("Scenario",X1)-1`).
- **The Mind loop is written with the formula repairs Prep still has** for the version (CELL / INFO
  frozen, INDIRECT written out, ...: the "Fix the formulas" set). Found on Tokyo: the four loops were
  written before Prep, and `CELL("filename")` in B3 of the six calculation sheets -- the key of every
  lookup there -- is NaN in Mind (Mind answers with the workbook's name alone, no "]"), so Mind's test
  showed NaN on every formula of those sheets (14,737 cells). Not the loops: a probe gives the same NaN
  without any loop. The loop panel lists the repairs; `repairs: false` leaves them out.
- **New rule FRM-014 "math on a range that OFFSET returns"** and the Prep action **Write a fixed OFFSET
  range out where arithmetic uses it**: `SUMPRODUCT(1-OFFSET($H$2,2,0,3,1),OFFSET($O$2,2,0,3,1))` is
  -16,220 in Excel and -5,513,051 in Mind (probe workbook in real Mind); written as `1-$H$4:$H$6` Mind
  gives Excel's number. Tokyo: 84 formulas (14 per calculation sheet), 396 differences in Mind's test
  of the Do-everything version without loops.
- Proven in real Mind on Tokyo with its four loops, built through the app: Mind's test went from 14,737
  differences to 63 (Do-everything version + loops, zzz_mindready_tokyo_loops_2) and 75 (loop written
  on the unrepaired version), none a value: check cells that are 0 in Excel and a rounding residue below
  0.00002 in Mind, and the text "TRUE" read as 1 on Excel's side.

### Next quarter's Mind input files

- **Make next quarter's Mind input files** (Outside files card, once the files are brought in;
  `app/next_quarter_inputs.py`; `POST` / `GET /api/sessions/{id}/next-quarter-inputs`). Joseph put a
  raw quarterly file (`2026_Q1 - BBA_01.csv`, 1,340 columns) in Mind's Inputs window and Mind refused
  it: its grid `IN_BBA_01` has 93 columns. Raw input files cannot go into Mind directly: Mind takes one
  file per input grid, shaped like the grid. Now the user picks the new quarter's raw files; the app
  runs the model's own import macro on a copy of the version the files were first brought into, reads
  each grid's block and writes it under the grid's name with the grid's own header. One zip:
  `README.md` (which file goes into which grid, and why raw files cannot) and the folder `Mind inputs`.
  No new model upload, no new version. A grid without a new file is left out (Mind keeps its data); a
  file already shaped like a grid goes in as it is; a raw file named like a grid is recognised; a file
  that fills more than its grid is refused with why; the macro's own message boxes are shown.
- **The macro gets the raw files exactly as picked** (review of 10 November 2026). A first version
  gave the macro a `.xlsx` copy of each raw `.csv` with every digit (Excel keeps 15 significant
  digits when it opens a CSV). That skipped the macro's own way of opening the file (OpenText
  settings, `Local:=True`, reading it line by line -- the Tokyo tool's `CountCSVRow` does
  `Open ... For Input` -- or checks of the extension), and the bring-in from the same CSVs gave
  other values. Now the macro opens each file itself, as at the bring-in and in the client's tool;
  a CSV with numbers past 15 digits is said ("Excel keeps 15 when it opens a CSV, as in the
  client's tool"). "Values changed" no longer counts a number that differs only past its 15th digit:
  Joseph's raw CSVs against the bring-in's `.xlsx` files give 0 changed (15,011 such cells in
  IN_BBA_01 before). Also gone with the copy: a raw `X.csv` and `X.xlsx` picked together no longer
  reach Excel under the same name.
- **Grid files checked before they go in**: a file shaped like a grid but not UTF-8 (a CSV saved on a
  Japanese PC as "CSV (comma delimited)", cp932) is refused with what to do, and no longer stops
  the whole job. A `.txt` is read as TAB-separated, as Mind reads it: a comma-separated `.txt`
  shaped like a grid goes into the zip as `<grid>.csv` (refused, with why, when it holds quotes or
  TABs).
- **One Excel job at a time**: while next quarter's files are being made, the bring-in, Run in Mind,
  the fixer, a recalculation, Apply and the other workbook changes wait (409, in words), as the
  job already waited for them. A second bring-in could have changed the grids the job writes for.
- **Screens**: the card picks a running job back up after a reload or a visit to Advanced (and shows
  the last result and its zip); a new workbook clears the last job's line, files and zip, and an
  old job's polling stops, so its result never lands on another workbook's card. The zip link is
  only offered in the session that made it.
- **Checked again after the review** (dev server, the anonymized IFRS 17 tool, Joseph's six raw
  CSVs): the macro opened all six as picked; 6 files made in 84 s (239 s with the copy); 0 values
  changed against the bring-in; set B (BBA_04 x 1.1) changes 26,456 values of `IN_BBA_04.csv` only.
  In real Mind (`zzz_mindready_tokyo_inputs_1`): all six SUCCESS, run consistent with the audit
  trail, Result_CF the same 120 numbers as the first proof (Mind = Excel). The project holds the
  original quarter (set A) again.
- **Mind's refusal in words**: `mind_client.upload_inputs` reads Mind's reason for a refused file from
  the row's LOG ("CSV Logs") into `errors`, and for a raw file says what to do instead (`advice`).
- **Fix: Inputs window read too early.** `upload_inputs` read the rows while Mind's "Inputs import" box
  was still up; they show the last upload's status and file (the same names when files replace earlier
  ones), so a refused file could pass as SUCCESS. It now waits for the box to come and go.
- **Proven in real Mind** (`zzz_mindready_tokyo_inputs_1`, `docs/OUTSIDE_INPUTS.md`): the raw file
  refused, with the app's advice; set A (Joseph's raw CSVs) all six files SUCCESS (incl.
  `IN_VFA_01.txt`), Mind = Excel 120/120; set B (BBA_04 x 1.1) all SUCCESS, Mind = Excel 120/120, and
  16 numbers changed against set A, all in the BBA_04 row.

### Run in Mind: a second run, and the Inputs window of a copied workbook

- **Fix: a second run in the same browser session returned at once.** `mind_client.run_model` took
  the "Consistent with the audit trail" line of the run before -- Mind keeps it on the page the whole
  time, also while the next run is in progress -- as this run's end. Now it notes the page's run
  marker before the click (`run_marker`: the "Last run in N seconds" line, taken down during a run and
  put back with the new time, and the number of "Run completed" notes) and waits until a new run was
  seen ("Run in Progress" or a changed marker) and that run's own line is up. The result says
  `completed`, `audit_consistent` and `seconds` of this run, `started`, and `reason` when the page only
  ever showed the run before. Mind's unmatched-inputs window, when left open, is closed first (it
  covered the Run button). Fake-page tests on a fake clock (`test_mind_inputs_upload.py`).
- **Fix: Mind's Inputs window for a copied workbook (Mind instances).** `upload_inputs` could not read
  the rows: a copied grid is listed once, collapsed, as "Rates (3 instances)", so it waited its full
  300 s and said "still waiting". Now the collapsed rows are opened first (their carets), each copy's
  row is read with its key (`statuses` "Rates [K2]", `rows[].key` / `copies`), and `<key>-<grid>.csv`
  is found on its copy's row. Mind's new window for a file it could not match ("Some of your imported
  files have not been automatically matched ... Match selection"; the old one said "Match unmatched
  inputs") is recognised, closed with its own Close button (never "Match selection") and explained in
  plain words: "Mind could not tell which copy of Rates Rates.csv belongs to: the grid has 3 copies
  (K1, K2, K3). Name the file <key>-Rates.csv (K1-Rates.csv) to change one copy, or turn on 'Split
  between instances' ..." / "Mind could not tell which grid X.csv belongs to. Name it <grid>.csv or
  <key>-<grid>.csv for one copy." (`unmatched_advice`, also in `advice`).
- **The project is opened again after a copy's file** (`upload_inputs(..., project_url=)`, `reloaded`):
  the run right after a copy's first input file still used the copies as loaded before the import
  (every copy with every row: K2 should have been 915 and showed 120); after a reload the run was
  right. `app/mind_loop.py` passes the project's link, so Run in Mind on an instances model with a
  copy's own input file runs on the reloaded project.
- **Fix: a run of a minute or more was never seen as finished.** Mind writes such a run as a clock
  time -- "Last run in 00:03:01." / "Run completed in 00:03:01" (the anonymized IFRS 17 tool, 3 min
  06 s) -- and the patterns above only knew "N seconds", so `run_model` waited its whole timeout and
  the Run in Mind loop said the run did not complete, on every iteration. Both forms are read now
  (`run_seconds`: "00:03:01" -> 181 s); fake-page test with a clock-time run.
- **Proven in real Mind on 10 November 2026** with a 3-copy invented model (probe 1's InstModel +
  InstMain, grid `Rates /Input`), `docs/RUN_IN_MIND.md`, evidence in
  `%LOCALAPPDATA%\MindReady\iteration_1_13_0\mind_client\`: `zzz_mindready_runtwice_1` -- run 1
  (0.22 s), `K2-Rates.csv` read on K2's row in 12 s (300 s and "still waiting" before), run 2 waited
  for its own result (marker `(None, 0)` -> `('Last run in 0.68 seconds', 1)`), a plain `Rates.csv`
  opened the new window, which was closed and explained, and run 3 went through (the click timed out
  on that window the night before). `zzz_mindready_runtwice_2`: opening the collapsed rows changes no
  number; `zzz_mindready_runtwice_3`: reload before the run after `K2-Rates.csv` -> K1 242 / K2 915 /
  K3 160, only K2 changed (Mind = the hand-worked values).

### The macro's text in its own code page, and two notes on the plans

- **The macro's text in its own code page** (`model_map.read_vba`, `decode_vba_source`;
  `docs/MIND_LOOP.md`). The VBE stores module text in the ANSI code page of the PC that saved the
  workbook, and Excel 365 records 0 as the project's code page; olevba answered 0 with UTF-8 and
  replacement characters, so the Japanese comments and string literals of the IFRS 17 tool
  (`Find("年月")`, `MsgBox "完了しました"`; 363 replacement characters in M1_Main) arrived as mojibake and
  the stand-in assistant and the message-box matching could not read them. The module bytes are now
  decoded with the stored code page, else UTF-8, cp932 (when it gives Japanese), cp949 (Korean), the PC's
  own code page, cp936 / cp950 (ideographs), cp1252 -- each only when it decodes every byte. The
  anonymized IFRS 17 tool reads clean (0 replacement characters, 290 Japanese characters in M1_Main);
  the loops are found as before. Fixture `tests/unit/fixtures/vba_japanese_cp932.xlsm` (a small
  row-per-record model, its module cp932 with code page 0 as a Japanese PC writes it; made once through
  Excel, the module stream's source replaced at the same length) and `tests/unit/test_model_map_vba_text.py`.
- **Kanji-only Japanese is Japanese** (review of the same day). The first check asked cp932 for a
  quarter of kana, so a module whose non-ASCII text was kanji only -- `Find("年月")`,
  `Worksheets("計算")`, a comment `' 全契約計算実行` -- came back as Korean (cp949 gives clean-looking
  Hangul for it) or as Chinese, and `MsgBox "ｴﾗｰ"` (half-width katakana) as Hebrew even with 932
  stored. Now the stored code page is trusted as it is; cp932 is taken when at least half of what it
  gives comes from the kana, the full-width symbols and the common (JIS level 1) kanji -- Hebrew,
  Western, Korean or Chinese bytes read as cp932 give half-width katakana, rare kanji and private-use
  characters, never those; cp949 when it gives the common Hangul syllables (Japanese bytes read as
  cp949 give the rare ones of the UHC extension); the PC's own code page is tried with its check first
  among the East Asian ones and trusted after them (a Korean file on a Japanese PC reads as Korean).
  The limit that stays: half-width katakana alone, with no code page recorded, on a Hebrew or Western
  PC read in that PC's code page. A test case for each.
- **The loop and instances plans name Prep's blocking items that are not formulas.** Both steps
  write Prep's formula repairs with the loop, but said nothing about a lone text cell to title
  (STR-008), a pivot table (STR-009) or a trapped grid title (STR-001), which Mind's converter refuses
  or computes wrong without. One line now: "Prep still has N blocking items that are not formulas, e.g.
  a formula reads a lone text cell (STR-008). Run Fix all issues before uploading." (`prep_blockers`
  on the plan; `loop_builder.prep_blockers`).
- **The two harmless kinds of cell in Mind's upload test are said once** on every ready loop or
  instances plan (`loop_builder.UPLOAD_TEST_NOTE`; `docs/MIND_VS_EXCEL.md`): check cells that are 0 in
  Excel and a rounding residue below 0.00001 in Mind, and cells holding the text TRUE / FALSE, which the
  test reads as 1 / 0 on Excel's side. The values agree; nobody has to chase them.

### Housekeeping

- **The stash of 9 November** ("release-cleanup uncommitted work", 22 files) was checked hunk by hunk
  against the branch and dropped; it is saved in full in
  `%LOCALAPPDATA%\MindReady\iteration_1_13_0\stash0_release_cleanup_2026-11-09.patch`. Three pieces were
  still live and came across: `tests/unit/test_feedback.py` dates its stage-timing row at run time
  (`feedback._timings()` keeps the last day only, so the fixed date had aged out and the crash test
  failed); a test that a workbook-level defined name wins over a sheet-local one of the same spelling;
  and an optional token meter in `app/llm.py` (a try-import of the `TokenUsage` `token_meter`, a no-op
  where it is not installed). Everything else was already on the branch or superseded (the side branch's
  1.13.0 / 1.14.0 numbering, an older built dist).
- **Version 1.13.0**: `VERSION`, this changelog, the README's status, and the docs name the tabs as they
  are now (`docs/MIND_LOOP.md`, `docs/MIND_INSTANCES.md`, `docs/INSTANCE_FINDER.md`).

## 1.12.4 (9 November 2026): outside input files as Mind inputs (the Tokyo IFRS 17 tool)

- **A model and its input files, picked together.** The upload takes several files at once (File
  Explorer or drag and drop); the app finds the model (the workbook whose macros open other files)
  and lists the files its macro reads: matched, missing, and files it does not need. Missing files
  can be added later; a file that is not picked keeps the data the model already holds.
- **Bring in the outside files** (the first step of "Do everything", and a button). The app runs the
  model's OWN import macro on a copy, in an Excel of its own, with each path cell pointing at the
  picked file, and watches it through Excel's events: which file it opens and which cells it writes
  while that file is open. The values arrive exactly as the client's tool puts them. Message boxes
  are answered (No / Cancel on a question) and reported; a macro that stops early is reported, not
  hidden. The numbers check then compares later steps with this version.
- **Each outside file becomes a Mind input.** Its area moves to a sheet `IN_<sheet>` holding the grid
  `#IN_<sheet> /Input`; every formula that read the area reads the same cells there (a SUMIF whose
  ranges run past the side of the area is split in two: the same total); one CSV per input
  (`Mind inputs.zip`). Each quarter the Mind team imports new CSVs in Mind's **Inputs** window and
  runs -- no new upload of the model. An area some formula cannot follow (a 3-D reference over it,
  OFFSET from inside it, a range that covers only part of it) stays in the model and is reported.
- **Run in Mind uploads the input files** through the Inputs window after the template step.
- **New rule INP-010 "Outside input files"**: blocking until the outside files are brought in.
- **No input files? "Make them with AI".** When files are missing, the assistant reads the import
  macro and the model's sheets and describes each file; the app builds stand-ins from the numbers
  the model already holds, proves them with the model's own macro (every written cell equals the
  model's, then every computed value; the assistant tries again with what went wrong, 4 tries) and
  brings them in as usual. "Do everything" does it behind the scenes when files are missing.
  INP-010 then warns instead of blocking -- best results come with the actual input files -- and an
  actual file added later replaces its stand-in (`POST .../outside-inputs/run {"standIns": true}`,
  `app/stand_in_inputs.py`; `chat_completion(effort=)` for this one call).
- Proven on the Tokyo tool (anonymized, invented input files): 6 of 6 files matched, import 11 s,
  6 Mind inputs, 46,800 formulas moved, no computed value changed; in Mind 6 of 6 inputs SUCCESS and
  Mind = Excel on the result grid (120 of 120 numbers); a "next quarter" (two inputs changed, model
  not re-uploaded) gives Excel's new answer too (120 of 120).
- **Hardened after an independent review** (4 reviewers, each finding challenged by a refuter on
  another model; 11 findings, none refuted, 10 fixed): a plain OK message box no longer counts as an
  early stop (Windows gives a lone OK button control id 2); a macro that clears every area first is
  paired correctly; whole-column clears of a file not picked are put back; the model's own
  Workbook_Open does not run and a box there cannot hang the job; when making the Mind inputs would
  change a value the import version stays current and nothing is offered for Mind; input files changed
  after a bring-in are noticed (fingerprint) and Run in Mind refuses stale CSVs; Do everything retries
  a failed bring-in; numbers that could not be checked stop the chain; the macro's message box reads
  correctly on a Windows of another language ("BBA_01のinput領域には93項目まで…" instead of mojibake).
- **RZS-006** leaves the Mind inputs made from outside files alone: their size is the model's import
  area, as the model's own macro keeps it, and their CSVs are written at exactly that size.
- Final check on Tokyo through the app: Acceptable by Mind (0 blocking), then the app's own Run in
  Mind: numbers match (1,703,906 cells), the 6 input files SUCCESS in Mind's Inputs window, run
  consistent with the audit trail, converged; 56 untitled grids in Mind (76 in 1.12.3).
- `scripts/night_chain.py` writes UTF-8 (a Japanese sheet name stopped it on a Hebrew console) and takes
  `--inputs <folder>`.
- **A shared copy stops on a date** (`build.py --expires 2026-12-15 --contact "..."`, 15 December
  2026 by default, `--no-expiry` for none): "Works until 15 Dec 2026" in the header, a line in the
  last 7 days, then a plain message at start and a calm page instead of the app (410 on every API
  call); a clock moved back counts too (each build remembers its own latest date; a date that has
  not come yet, from a clock once set ahead, is put right by the internet's date at start); the
  developer version has no date (`app/expiry.py`, docs/SHARING.md). Build the screens first
  (`npx ng build`): build.py does not package a copy with a date from screens that cannot show it.
- **Plain English next to every code**: "FRM-009 · an empty cell is compared with a number" on the screens, in the
  texts and in the Excel report (`rules/plain_names.yaml`, required by a test); Excel's error values get their meaning
  once, Excel / VBA errors a plain sentence with Excel's words in brackets (`app/plain_codes.py`, docs/PLAIN_CODES.md).
- **"Show me" pointers.** Next to every instruction that sends the user somewhere (the banner's next
  step, "Apply the blocking repairs in Prep", "Check them under Advanced › Deliver", "on the Outside
  files card", "paste an AI key under Advanced › Assistant", a finding's "Use Bring in the outside
  files" or "The Prep action '…'", "the list at the top of this page") a quiet "Show me" link: one
  arrow with a short bubble at the next click ("Click Advanced", "Click Prep actions") until the place
  itself, outlined, with "Got it"; Esc or a click elsewhere ends it, and it never takes a click from
  the page. "Fix that under Advanced" now names the tab. `docs/SHOW_ME.md`.
- **"Show me", after an independent review** (8 findings, 7 fixed): in a window narrower than 66rem,
  where Carbon hides the header's Main and Advanced, the arrow uses the page's own links (it showed
  nothing for 7.5 s, then ended); a click on the way that cannot be shown ends the pointer at once; a
  double click goes on instead of ending it; a stopped "Fix all issues" names the right place (Deliver
  when its comparison did not run, none when it did not run at all: "Press Do everything again"); a
  clean recalculation names no list; the assistant's own answers get no link (only the app's words);
  the findings' messages and Prep's "Add to the list" note have theirs; the readiness banner, now one
  component for both screens, keeps its name for screen readers (the title carries Carbon's id).

## 1.12.3 (8 November 2026): sheet names in any language (the Tokyo IFRS 17 tool)

- **Sheet and name text in any language.** A reference such as `=ロールフォワード!B10` was read as
  `=B10`: the reference reader only knew English letters, so on a workbook with Japanese sheet
  names the scan reported dozens of false circular references (a cell reading "itself") and the
  banner stayed on "blocked" while the recalculation was clean. Every reader of sheet names and
  defined names now takes letters of any alphabet (Japanese, Hebrew, accented, Cyrillic ...).
- **`CELL("filename", C6)` in C6 is not a circular reference.** CELL asks where the cell is, not
  what it holds (Excel computes it); it is read like ROW and COLUMN.
- **Text typed as `'=...` is text.** The scan took any text starting with "=" for a formula.
- **A CELL formula that shows "" is frozen as `=""`.** A formula whose saved value is empty text
  was read as "no value", so the fix left it "by hand"; and a CELL formula that was an error in
  the original and shows "" after Fix all issues is frozen to "" as well.
- **INDIRECT used as a range is rewritten too (FRM-007).** Checked in Mind: `VLOOKUP(x, INDIRECT($A5&"!A:B"), 2, FALSE)`
  is not computed there and the IFERROR around it showed "-" in 480 cells of the result sheet; the same
  lookup on `BBA_04!$A:$B` gives Excel's values. The Prep action ("Replace INDIRECT by the reference its
  text builds") now writes the direct reference wherever INDIRECT stands. An empty cell in the text is
  read as "" (a sheet chooser left blank: Excel's INDIRECT errs already); a missing target becomes NA(),
  or `INDEX($A$1,NA())` where a reference must stand (Excel refuses `OFFSET(NA(), ...)`).
- **A one-cell array formula over 255 characters is written with Formula2** (FormulaArray refuses longer
  text; the same one-cell array, the same value).
- **No title row is inserted inside a range a formula reads.** A row inserted inside `D29:H32` stretched it
  and `INDEX($D$29:$H$32, Scenario+1, 4)` read another row ("New" became "Old"); such an insert is now
  skipped (`prep.stretched_reader`).
- **A lone text cell read only through a whole column is not flagged (STR-008)**: the formula reads the
  sheet's table there, not the cell (Mind computed the lookup with it).
- **Mind's new "type of project" step.** Creating a project now first asks Milliman Mind or Milliman OpS;
  the automation picks Milliman Mind.
- `scripts/night_chain.py --no-loop`: the chain without a Mind loop.
- Checked on the Tokyo tool (anonymized): no computed value changed (110,175 formula cells), converted
  and ran in Mind at the first attempt, and Mind's results equal Excel's on the result sheet and a
  cash-flow grid.

## 1.12.2 (7 November 2026): the macro's loop through named ranges

- **The loop is found when the macro reaches its cells through `Names("X").RefersToRange`.** A macro
  that writes `Set cohortRng = ThisWorkbook.Names("CohortRng").RefersToRange` and then
  `cohortRng.Value = cohortYear` or `scenarioRng(1, 1).Value = i` had its loops read with no input
  cell, so no Mind loop was offered (the Tokyo IFRS 17 tool). The macro reader now writes these
  shortcuts out in full (`app/model_map.py`), including the loop limits read from them
  (`startcohortyear = startRng.Value`) and `Range("X").Cells(r, c)`.

## 1.12.1 (5 November 2026): what the Mind team saw when they tried it

- **The program no longer stops with "No module named win32timezone".** Excel hands back dates and
  times through a part of pywin32 that is loaded by name, which the compiled program left out; it is
  now built in (`tools/package/build.py`).
- **Run in Mind says which version it uploads: always the latest.** The screen names the version and
  its file before the click. While Fix all issues, a recalculation, a Prep round or an analysis is
  still changing the workbook, Run in Mind waits ("Wait a moment: ...") instead of uploading the version
  before that change (`GET /api/sessions/{id}/mind-loop/options`, 409 on `POST .../mind-loop`).
- **Choose the changes made before uploading.** Every change Run in Mind can make first is a tick box
  (ticked as before by default); an unticked change is never made, not even when Mind asks for it, and
  nothing ticked uploads the version as it is (`actions` in the `POST` body).
- **A clearer status while it runs, with the time.** The stages of a round (reading, changes, checks,
  opening Mind, upload and convert, run, results), the one going on now, the time spent, the time left
  and the clock time it should be done by. The estimate is kept on the long side on purpose (1.5 times
  the measured figures); past it, the screen says it is taking longer and is still working. The
  detailed log is one click away (`app/estimate.py` `estimate_mind_run`).
- **A row inserted for one grid no longer moves the grids beside it.** The empty row Prep inserts above
  a '#Title' trapped inside a grid (STR-001) or above a grid to make room for its title (STR-004) now
  shifts only that grid's columns down when another grid sits beside it at or below that row, so the
  neighbour stays where it is (before, the trapped-title fix was skipped there, and a title row moved
  the neighbour down, away from its own title). The columns that shift widen to take in anything that
  must move as one: a grid below, a merged range, a pivot table, a range a formula reads across them.
  When nothing sits beside the grid, or the sheet has an Excel table or a formula reading whole rows,
  a whole row is inserted as before. The numbers check and Run in Mind follow the moved cells
  (`app/prep.py` `plan_row_insert`). The grid names chosen by hand in Advanced (Grid Namer) insert their
  title rows the same way.
- **Sign in to Mind from the app.** Run in Mind has a "Mind sign-in" box: "Sign in to Mind" opens
  Edge with the app's own profile on Mind; once you are in, it saves the sign-in and closes by itself
  (the sign-in lasts about a week). "Check the sign-in" loads Mind without a window and says whether it
  still works. Before, signing in needed a command line, which the program does not have
  (`GET`/`POST /api/mind/sign-in`, `app/mind_client.py` `sign_in`). Run in Mind and the sign-in window
  wait for each other: they share the profile.

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
- **No hidden Excel stays behind.** The app writes down every Excel it starts until it quits it. Closing
  the app's window ends its Excels first, and the next start ends any left by an app that was stopped
  another way (a crash, Task Manager). Only an Excel the app itself started is ever ended
  (`app/excel_com.py`).
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
