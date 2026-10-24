# Plan: share Mind Ready with the Mind team (deliver Monday 2026-10-26, build Sunday 2026-10-25)

**Goal.** The team that built Mind runs the app on their own machines, with their own Excel, MM for Excel
add-in, APIM and Mind accounts. They get a compiled package, not the repo: no git history, no branches, no
way to push back, and the code hard to read (Nuitka compiles the Python to C). Only hosting would hide the
code completely; this is the agreed balance.

Start from the latest branch: `mind-loop-palermo` (pushed to origin + milliman). Work on a new branch, do not
push unless asked.

## Decisions to ask Joseph at the start (then work unattended)

1. Which features they get: Mind upload / run in Mind (needs Playwright, +~105 MB), the assistant (their
   APIM), the Model map, the Mind loop builder, MCP.
2. Their APIM: endpoint and how the key is given (settings file next to the app, or an environment variable).
3. Package form: a folder with a start shortcut (recommended) or a single .exe (same size, slower start).

## Steps

1. **Scrub** (2-3 h)
   - Client names (Palermo, Phoenix, CNHI, Horizon ...) out of the 15 files in `app/` that mention them
     (comments, examples); tests, CHANGELOG and docs are not shipped at all.
   - One settings file (paths, APIM endpoint + key, port, MM add-in path, Mind sign-in folder) instead of
     what is set up for Joseph; check first how the app finds the APIM key today.
   - Feature switches per decision 1.
   - FastAPI `/docs` and `/openapi.json` off in the release (`docs_url=None, openapi_url=None`).
   - Production front end without source maps (today's build has none; keep it so).
   - A release file list: ship the app, the built front end (MindReadyUI/dist), rules, schemas, settings;
     leave out `scripts/`, `tests/`, `runs/`, internal docs, `CLAUDE_CODE_PROMPT.md`, `streamlit_run.log`,
     CHANGELOG; drop Streamlit from the requirements (old front end), and pandas / numpy if nothing in the
     app imports them (check).
   - Full test suite still green.
2. **Nuitka setup** (~15 min): install Nuitka, let it fetch its C compiler.
3. **First build** (20-45 min on this laptop; rebuilds 10-20 min), standalone folder.
4. **Fix build issues** (1-3 h, the least predictable part): Excel COM (pywin32), Playwright's driver
   files, rule / schema / data files Nuitka does not pick up by itself, dynamic imports.
5. **Test install** (~30 min): copy the package to a clean folder, start it, run a full Do everything on a
   small workbook (Horizon), check the front end, `/docs` gone, no client names in the shipped files.
6. **Hand-over note** for the Mind team: how to start it, where the settings file is, what each switch does.

## Expected size

| Part | Approx. |
|---|---|
| Python runtime + pywin32, lxml, pydantic, cryptography | 60-80 MB |
| App compiled (~24,000 lines) | 30-60 MB |
| Playwright (Mind upload; uses the installed Edge) | ~105 MB |
| Front end | 2 MB |
| Rules, schemas, settings | < 5 MB |

With the Mind upload: ~200-300 MB as a folder, 80-120 MB zipped. Without: ~100-150 MB, 40-60 MB zipped.

## Expected time

About half a day to a day, mostly unattended. Estimates, not measured: the first build gives the real
figures.

## Before handing it over

- Nobody on the Mind team has access to the GitHub repos (check the Milliman-Tel-Aviv-Life organisation:
  Mind_AI_Excel_converter, which holds every branch).
- No key, token, sign-in or personal path inside the package.
