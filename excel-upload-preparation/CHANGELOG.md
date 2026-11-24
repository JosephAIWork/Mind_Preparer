# Mind Ready changelog

> Release notes for Mind Ready, including new features, improvements, and bug fixes by version.

The full notes of each version, with the reasons, the proofs and the file names for developers, are in [docs/CHANGELOG_DETAILED.md](docs/CHANGELOG_DETAILED.md). The version you run is shown in the app's header.

**1.15.0**

November 13, 2026

**Added**

* Added three screens, Basic, Custom and Advanced: Basic has the upload box, one **Start** button that runs every step with its check, a progress line and the Recent workbooks list
* Added a **Cancel** button wherever the app works; it stops the work on the server, puts the workbook back as it was before the step and ends only the Excel the app itself started
* Added a "Still working" strip with its own Cancel when work goes on after you leave its screen
* Added **Delete it from Mind** for a half-made Mind project left by a cancelled Upload to Mind or Run in Mind; nothing is deleted without that click
* Added an end-of-check banner on Basic that says what to do next, with **Upload to Mind** and **Download**
* Added **Remove all** to the Recent workbooks list
* Added a **Recalculate** button wherever the recalculation is the only thing left
* Added a progress bar, an ETA marked "an estimate only", and pulsing dots while Start runs
* Added a **Show me** to every stop sentence on Basic, including the sentence after a Cancel
* Added a thinking motion wherever the assistant works
* Added the assistant's last try at what Prep left: its fix is kept only when no computed value changed and the problem got smaller
* Added INDIRECT over a numbered list of names (`INDIRECT("Cluster"&n)`) rewritten as INDEX over those cells, when the number can only be whole and never below 1
* Added a label written beside a lone text cell inside a range, so Mind imports it (checked in real Mind)
* Added a lone text read as the criterion of SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS or MINIFS written into the formula as text
* Added "Row N" labels that close a short gap in a table's first column
* exe: Added `--jobs N` to the build, for a machine short of memory

**Changed**

* Changed Start so it never stops at Prep: it goes on with the Mind loop and the recalculation and asks at the end if a problem is left
* Changed the Prep step to repeat while a round still changes something, up to 10 rounds
* Changed the screens to plain words: rule codes appear only in Send feedback and in what the assistant is given
* Changed the Mind sign-in messages to plain sentences that say where the Edge window is and how long the sign-in lasts
* Changed the "Works until" date so it no longer shows in the header (the "ends soon" and "stopped" messages stay)
* Changed Run in Mind's last line to "Mind converted the workbook and ran the model, and the run agrees with its audit trail"
* Changed the lone-text check to count only formulas that need the text (not SUM, MAX or COUNT over a range, nor a lookup that never returns that column)
* Changed the project into a copy that stands on its own, without the retired screens and the internal notes

**17 bug fixes**

* Fixed formula errors that show only for the Mind loop's first value stopping Start; they are now said once in a plain line
* Fixed the numbers check counting file-path texts that Prep froze on purpose as changed values
* Fixed the numbers check reporting a caption Prep replaced by its title as a lost formula
* Fixed a column heading that INDEX can never return being listed as a lone text cell
* Fixed a cell holding only "#" above a column being read as a grid title, which cut the table in two
* Fixed lone text cells in a diagonal "staircase" taking one Prep round per cell
* Fixed a title written into an empty cell that an INDEX/MATCH reads, which changed computed values
* Fixed `CELL("filename")` being frozen to the model owner's own folder and file name
* Fixed Do everything leaving a blocking problem that appeared only after the Mind loop was built
* Fixed `scripts/night_chain.py` stopping when no cell names the outside files
* Fixed the Mind upload script saying "Mind agrees" when it could not read the list of differences
* Fixed Send feedback hiding only part of a path when a folder name has spaces
* Fixed the Basic, Custom and Advanced links disappearing on windows narrower than about 1,050 px
* Fixed a Cancel in the loop step's first seconds not stopping it
* Fixed the Cancel beside Start stopping work started elsewhere
* Fixed Edge warning "unsupported command-line flag: --no-sandbox" in the Mind browser
* Fixed documents and package files that still carried old versions and ports

**1.14.1**

November 12, 2026

**Added**

* Added the assistant's token count per workbook ("AI used: 12,340 tokens in 7 calls") under Advanced > Assistant and in Send feedback
* Added grid titles for tables whose caption is a formula: the caption becomes the title, and the note says so
* Added a pre-commit hook that keeps comment spacing consistent

**4 bug fixes**

* Fixed Fix all issues' assistant calls not being counted
* Fixed a made-up title replacing a formula caption whose last value was empty or an error
* Fixed a title being written into a caption cell that INDIRECT or a 3-D reference may read
* Fixed the comment-spacing hook touching lines outside a partial commit

**1.14.0**

November 11, 2026

**Added**

* Added saved workbooks: every workbook is kept in the data folder after each step and survives a restart
* Added the Recent workbooks list with **Reopen** and **Remove**
* Added a wait in Do everything for missing input files, with **Pick the files**, **Go on**, **Make them with AI** and **Stop**

**Changed**

* Changed grid titles so they never move cells: one title per table, written into the cell above it, and a table with no room is reported with what to do
* Changed a key column and the block beside it so they are read as one table
* Changed Send feedback so it names no team, and the assistant's opening lines never repeat the user's note

**6 bug fixes**

* Fixed INDIRECT that points at a sheet left out of the scan being rewritten as an error
* Fixed the Mind-value rules reading the cells of a skipped sheet as empty
* Fixed the Mind-value rules failing when two or more sheets are skipped
* Fixed two copies of the app on one data folder listing, reopening or removing each other's open workbooks
* Fixed the saved state file growing to tens of MB by keeping the check result and the plan in files of their own
* Fixed the countdown counting the wait for input files as the step's time

**1.13.0**

November 10, 2026

**Added**

* Added the Mind instances builder (Advanced > Instances): one copy of the model per key of a table, with a main workbook of results by key
* Added copies from the macro's loops, one copy per combination such as a scenario and a cohort
* Added the instances finder ("Not sure? Let the app look"), which points to what repeats in the model and says whether instances fit
* Added each Mind loop's own choices, and one macro's loops at a time
* Added more loops found from the workbook alone: a named list, a counter with its own dropdown, a bound read through references
* Added next quarter's Mind input files: the model's own import macro runs on the new raw files and the app writes one file per input grid
* Added rule FRM-014, math on a range that OFFSET returns, and the Prep action that writes the range out
* Added a note on every loop and instances plan naming Prep's blocking items that are not formulas
* Mind: Added reading of Mind's reason when it refuses an input file, with advice on what to do instead
* Mind: Added support for the Inputs window of a copied workbook, where each copy has its own row

**Changed**

* Changed Advanced > Mind loop to open with one sentence and one **Build it** button, with the choices folded away
* Changed the Mind loop and instances to be written with the formula repairs Prep still has
* Changed the macro's text to be read in its own code page, so Japanese and Korean comments and strings read clean

**5 bug fixes**

* Fixed a second Mind run in the same session returning at once with the run before's result
* Fixed a Mind run of a minute or more never being seen as finished
* Fixed Mind's Inputs window being read before its import box closed
* Fixed the run after a copy's first input file using the copies as loaded before the import
* Fixed the Mind upload script closing the upload window without adding the templates

**1.12.4**

November 9, 2026

**Added**

* Added picking a model and its outside input files together; the app finds the model and lists the files its macro reads
* Added **Bring in the outside files**: the model's own import macro runs on a copy and each file becomes a Mind input grid with its CSV
* Added rule INP-010, outside input files
* Added **Make them with AI**: stand-in input files built from the numbers the model already holds and proven with its own macro
* Added an end date for shared copies, with an "ends soon" line and a calm stop page
* Added plain English next to every rule code
* Added **Show me** pointers that lead to the next click

**Changed**

* Changed Run in Mind to upload the input files through Mind's Inputs window
* Changed RZS-006 to leave Mind inputs made from outside files alone

**1 bug fix**

* Fixed `scripts/night_chain.py` stopping on a sheet name in Japanese on a Hebrew console

**1.12.3**

November 8, 2026

**Changed**

* Changed every reader of sheet names and defined names to take letters of any alphabet
* Changed INDIRECT used as a range to be rewritten too (FRM-007)
* Changed the automation to pick Milliman Mind at Mind's new "type of project" step

**6 bug fixes**

* Fixed `CELL("filename", C6)` in C6 being read as a circular reference
* Fixed text typed as `'=...` being read as a formula
* Fixed a CELL formula that shows "" being left for a person
* Fixed one-cell array formulas over 255 characters being refused
* Fixed a title row inserted inside a range a formula reads, which changed what INDEX read
* Fixed a lone text cell read only through a whole column being flagged

**1.12.2**

November 7, 2026

**1 bug fix**

* Fixed the macro's loop not being found when the macro reaches its cells through `Names("X").RefersToRange`

**1.12.1**

November 5, 2026

**Added**

* Added tick boxes for the changes Run in Mind makes before uploading
* Added the stages, the time spent, the time left and the finish time while Run in Mind runs
* Added **Sign in to Mind** and **Check the sign-in** in the app

**3 bug fixes**

* Fixed the program stopping with "No module named win32timezone"
* Fixed Run in Mind uploading an older version while another step was still changing the workbook
* Fixed a row inserted for one grid moving the grids beside it

**1.12.0**

October 27, 2026

**Added**

* Added **Send feedback**: a plain message in codes, never a file name, sheet name, cell address or value
* Added the AI key box, with the key checked and saved encrypted for the Windows user
* Added one settings file, `settings.yaml`
* Added "Do everything right after the upload"
* Added the Beta tag in the header and the browser tab

**Changed**

* Changed every pivot table to block Mind's converter, with values as the default setting
* Changed a formula filled down to be written as one fill
* Changed lone text cells with no room for a title to get a label beside them or move to a column of their own

**5 bug fixes**

* Fixed hidden Excels staying behind after the app closed or crashed
* Fixed GETPIVOTDATA readers keeping their pivots, which Mind refused
* Fixed a broken defined name with its sheet becoming `Sheet!NA()`
* Fixed one-cell dynamic-array formulas being left for a person
* Fixed labels flagged only because an unused name covered the whole sheet

**Up to 1.11.10**

* Added the readiness check against the rules mined from Mind's knowledge base
* Added **Fix all issues**, **Fix the formulas** (Prep), **Build the Mind loop**, **Recalculate** with the numbers check, and **Do everything**
* Added Advanced: the findings rule by rule, the Model map, the Grid Namer, the assistant, Run in Mind and every version
