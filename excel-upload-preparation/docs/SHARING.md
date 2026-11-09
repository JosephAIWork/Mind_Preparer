# Sharing a copy that stops on a date

A copy of Mind Ready shared with other people (the single-file `MindReady.exe`) can stop working
after an end date. It works all day on that date, on the computer's own clock, and stops from the
next day. The code is in `app/expiry.py`.

## Build a copy that stops

In the build venv (see `tools/package/build.py`):

```
python tools/package/build.py <out dir> --onefile --expires 2026-12-15 --contact "Joseph Gestetner"
```

- `--expires` is the last day the copy works. Without it the build uses **15 December 2026**.
- `--contact` is the person the message tells users to ask for a new copy. Without it the message
  says "the person who shared it".
- A copy that never stops needs `--no-expiry`, written out. A date already past stops the build.

The build writes a small file `EXPIRY` (the date and the contact) into the package, next to
`VERSION`. It never writes it next to the source's `VERSION`.

## What the user sees

- In the header, quietly: **Works until 15 Dec 2026**.
- In the last 7 days, one line under the header: **This copy stops working on 15 Dec 2026.**
- From the next day, when they start the program, the window says
  **This copy of Mind Ready stopped working on 15 December 2026. Ask Joseph Gestetner for a new copy.**
  Nothing else starts. Pressing Enter closes the window.
- If the program was already running, every screen gives way to the same message on a calm page.
  The server answers every call with 410 and that message; `/api/health` still answers and says the
  date (`"expires": "2026-12-15"`).

Their own workbooks are never changed by any of this.

## A clock moved back

The copy remembers the latest date it has seen, in `%LOCALAPPDATA%\MindReady\latest_date_seen.json`
(written at most once a day). If the computer's clock is more than 2 days behind that date, the copy
stops, and the message says to set the right date. At start the copy also asks the internet for the
date, in the background: one quick request (2 seconds at most) to the AI gateway in
`settings.yaml`, and only if that one does not answer, one to `https://www.microsoft.com`. That
answer can only move the remembered date forward. Start-up never waits for it, and without a
network nothing changes.

## The developer version

The source has no `EXPIRY` file, so the developer servers never stop. To try the feature, start a
server with a date (here on its own port, beside the usual one):

```
set MIND_READY_EXPIRES=2026-12-15
set MIND_READY_PORT=8623
python -m app.web.server
```

A date in the past shows the stopped page; a date a few days ahead shows the line for the last 7
days. A server with a date writes the remembered date (above) on this computer too: try it with
today's real clock, never with a clock set ahead.

An `EXPIRY` file placed next to `VERSION` works too (it is in `.gitignore`: never commit one). In a
built copy, `MIND_READY_EXPIRES` can make the date earlier, never later.

## What this is not

This is a deterrent, not copy protection. Someone who finds the program's unpacked files and edits
`EXPIRY`, or who deletes the remembered date and moves the clock back, can get round it. It stops an
honest user from running an old copy by mistake, which is what it is for.
