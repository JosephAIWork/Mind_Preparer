# Mind Ready

Gets an Excel model ready for **Milliman Mind**. It fixes the formula errors, makes the formulas and the
structure acceptable to Mind, and turns the model's macro run loop into a Mind loop (`MM_LOOP` /
`MM_RESULT`). After each step it recalculates in Excel to show that the model's numbers did not change.

Your original file is never changed. Every step writes a new version that you download.

This is a first version shared for feedback. Please tell us what works, what does not, and what is
missing (see [Feedback](#feedback)).

## What you need

- **Windows** with **desktop Excel**. The app writes workbooks through Excel itself, because saving a
  real model any other way can damage it.
- The **MM for Excel add-in**. The app loads it into its own Excel session. Without it, every `MM_`
  formula shows `#NAME?` and the app cannot confirm the workbook is ready.
- **Python 3.11 or later** (tested on 3.13). When installing it, tick *Add python.exe to PATH*.
- **Microsoft Edge**, only for running in Mind from the app.
- **A key for the Milliman AI gateway**, only for the assistant (see [The AI key](#the-ai-key)).

## Install (once)

Open a command prompt and run:

```bat
git clone https://github.com/Milliman-Tel-Aviv-Life/Mind_Ready.git MindReady
cd MindReady\excel-upload-preparation
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Start

Double-click `excel-upload-preparation\run_mind_ready_web.bat`. A server window opens and then the app
opens in your browser at <http://localhost:8600>. The browser tab shows the version, for example
*Mind Ready 1.11.10*.

To stop the app, close the server window. If port 8600 is taken, change `port` in the settings (below).

The app runs only on your machine (`127.0.0.1`). Nothing is sent anywhere except the assistant's
questions to the AI gateway and, when you choose it, your workbook to Mind.

## Settings

Everything you might need to change is in `excel-upload-preparation\settings.yaml`, each line explained
in the file: the port, the folder where the app keeps its own files, the AI gateway, the Mind addresses,
the Mind folder and project names it uses, and where the MM for Excel add-in is. The values in the file
work as they are. After a change, restart the app. If a line is wrong, the app does not start and the
server window names the line.

## Use it

1. **Open a workbook** (`.xlsx`, `.xlsm` or `.xlsb`). For a large file the app first asks whether to
   skip some sheets. Skipped sheets are left exactly as they are.
2. Click **Do everything**. It runs four steps, each followed by its check, and stops at the first check
   that fails:
   - **Fix all issues**: finds the cells where errors start, gives each a value to fall back on, and
     undoes any fix that changed a value that was correct. Errors it cannot fix are listed with the
     reason.
   - **Fix the formulas** and the structure: every change Mind needs before it accepts the file, such
     as grid titles starting with `#`, frozen external links and pivot tables turned into values.
   - **Build the Mind loop**: the loop the workbook's macro runs, written as `MM_LOOP` / `MM_RESULT`.
   - **Recalculate** in Excel with the MM for Excel add-in, and compare the numbers with the original.

   You can also run the steps one at a time with their own buttons.
3. **Download this version**. You can also download the report on its own, or the workbook with the
   report added as sheets.

Everything else is under **Advanced**: the findings rule by rule, the individual fixes to tick or untick,
the Mind loop choices, the Model map (where the inputs, the loop and the outputs are), the Grid Namer, the
assistant, Run in Mind, and every version made so far.

## The AI key

The assistant (grid names from context, the chat, explaining and suggesting fixes) uses the Milliman AI
gateway with your own key. The app does **not** need it for the main steps.

When no key is saved, the main screen and **Advanced > Assistant** show a box to paste it in. The app
first checks the key with the gateway and does not save a key the gateway refuses. It saves the key
encrypted for your Windows account in the app's data folder (`%LOCALAPPDATA%\MindReady` unless the
settings say otherwise), so only you on this machine can read
it. After that the app shows only the key's last 4 characters, with a button to replace it.

The assistant only advises. Nothing it suggests is applied until you click Apply.

## Run in Mind (optional)

The app can upload a prepared workbook to Mind, convert it, run it and report back. First sign in once
from the `excel-upload-preparation` folder:

```bat
python -m app.mind_client login <the /init/ link of any Mind project you can open>
python -m app.mind_client check
```

`check` should say `"authenticated": true`. The sign-in lasts about a week. When it expires, run `login`
again.

The app works only inside a folder called `MindReady_Sandbox`, which it creates if needed. It names its
projects `zzz_mindready_…` and never touches any other project. Test projects that never got as far as a
run are deleted afterwards. Both names can be changed in the settings.

## Known limits

- **Windows and desktop Excel only.**
- **The Mind loop is built only from a loop the macro already runs.** The app does not invent loops or
  decide where a loop or instances would help.
- **`MM_INSTANCE(cell, k)` with k above 1** shows 0 in Mind's upload check before the first run. That is
  a false alarm: after a run, Mind shows the right values.
- **Pivot tables:** Mind does not accept them, so the app turns them into values. The choice is under
  Advanced, on the Prep screen.
- **External links** are frozen to the values Excel last kept for them.
- **Large models** take several minutes per step. The largest ones tested needed about 7 GB of free
  memory on top of Excel.

## Feedback

Please contact the **Milliman Tel Aviv team** with your feedback.

The most useful feedback includes:

- the version shown on the browser tab;
- what you did, what you expected, and what happened, with a screenshot;
- the report (**Download the report**), and the workbook if you are allowed to share it.

## What is in this repository

| Folder | What it is |
|---|---|
| `excel-upload-preparation/` | The app: server, rules, checks, fixes, Excel and Mind automation. |
| `MindReadyUI/` | The screens (Angular, Carbon). The built version is included, so you do not need Node. |
| `0*_*.md` | The Mind knowledge-base documents the rules were drawn from. The app does not read them while it runs. |

### For developers

- **Tests:** `pip install -r requirements-dev.txt`, then `python -m pytest tests -q` from
  `excel-upload-preparation`. With Excel the full run takes about 50 minutes. Without Excel, the tests
  that need it skip themselves.
- **Rebuilding the screens:** `npm install` and then `npx ng build --configuration production` in
  `MindReadyUI`.
