"""Build Mind Ready as a compiled package (Nuitka): a folder with MindReady.exe that needs no Python.

    <build venv>\\Scripts\\python.exe tools/package/build.py <out dir> [--onefile]
        [--expires YYYY-MM-DD | --no-expiry] [--contact "Name"]

The build venv holds requirements.txt and nuitka (nothing else, so nothing else is pulled in). Nuitka
fetches its own C compiler the first time. The result is <out dir>/MindReady/, the folder to zip:
MindReady.exe plus what the app reads next to it (rules, references, config, settings.yaml, VERSION)
and the built screens in ui/ (run `npx ng build --configuration production` in MindReadyUI first).
A second build reuses Nuitka's compile cache and is much shorter.

--onefile makes one file instead, <out dir>/MindReady.exe, to share as a link: it unpacks itself the
first time it starts (into %LOCALAPPDATA%\\MindReady-program\\<version and build time>, so later
starts are quick) and reads a settings.yaml placed next to it (mind_ready_app.py).

A copy to share stops working after an end date (docs/SHARING.md, app/expiry.py): --expires is the
last day it works, local time (by default 15 December 2026), --contact the person its message tells
to ask for a new copy. Both go into a small file EXPIRY next to VERSION in the package. A copy that
never stops needs --no-expiry, said explicitly.

The package takes the screens as built in MindReadyUI/dist: build them first (npx ng build in
MindReadyUI). A copy with an end date is never built from screens made before the end date was added
(they could not show it).
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

APP = Path(__file__).resolve().parents[2]  # excel-upload-preparation
UI = APP.parent / "MindReadyUI" / "dist" / "mind-ready-ui" / "browser"
sys.path.insert(0, str(APP))
from app.expiry import end_date_text, long_date, parse_day  # noqa: E402  (one format: the app reads what this writes)

# The copies shared from 1.12.4 work through 15 December 2026 (Joseph, 9 November 2026).
DEFAULT_EXPIRES = "2026-12-15"
# The selector of the screens' notice for a copy past its end date (MindReadyUI/src/app/core/
# expired-notice.ts). Built screens without it were made before the end date was added: they would
# show neither "Works until" nor the stopped page, so a copy with a date is never packaged with them.
END_DATE_NOTICE = "mr-expired-notice"
START_HERE = """Mind Ready {version}

Start: double-click MindReady.exe. A window opens (close it to stop Mind Ready) and the app opens in
your browser at http://localhost:8600.

Settings (port, folders, the Mind and AI addresses): settings.yaml in this folder. Restart after a change.
The AI key: pasted in the app the first time (main screen, or Advanced > Assistant); it is saved for
your Windows account, not in this folder.

Needs: Windows, desktop Excel, the MM for Excel add-in; Microsoft Edge for Run in Mind.
"""


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="build.py", description="Build Mind Ready as a compiled package (Nuitka).")
    parser.add_argument("out", help="the folder the build goes to")
    parser.add_argument("--onefile", action="store_true", help="one MindReady.exe, to share as a link")
    when = parser.add_mutually_exclusive_group()
    when.add_argument("--expires", default=DEFAULT_EXPIRES, metavar="YYYY-MM-DD", help="the last day the copy works, local time (default: %(default)s)")
    when.add_argument("--no-expiry", action="store_true", help="a copy that never stops working")
    parser.add_argument("--contact", default="", help='who to ask for a new copy, named in the message (e.g. "Joseph Gestetner")')
    return parser.parse_args(argv)


def end_date(args: argparse.Namespace, today: dt.date | None = None) -> dt.date | None:
    """The copy's last day (None with --no-expiry). A date that is not one, or one already past (the
    copy would not even start), stops the build with a message."""
    if args.no_expiry:
        return None
    try:
        day = parse_day(args.expires)
    except ValueError:
        raise SystemExit(f"--expires should be a date such as 2026-12-15, not {args.expires!r}") from None
    if day < (today or dt.date.today()):
        raise SystemExit(f"--expires {day.isoformat()} is past: the copy would not start. Give a later date, or --no-expiry for a copy that never stops.")
    return day


def write_expiry_file(folder: Path, last_day: dt.date, contact: str = "", built: str | None = None) -> Path:
    """The EXPIRY file the package carries next to VERSION (app/expiry.py reads it). It is written in
    the build folder, never next to the source's VERSION: there it would stop the developer servers.
    `built` (now, by default) makes this build its own copy for the dates it remembers: a date an
    older copy saw never stops this one."""
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "EXPIRY"
    path.write_text(end_date_text(last_day, contact, built or dt.datetime.now().isoformat(timespec="seconds")), encoding="utf-8")
    return path


def screens_show_end_date(ui: Path) -> bool:
    """Whether the built screens in `ui` can show a copy's end date: one of their scripts holds the
    notice (Angular cleans the folder at every build, so no older script is left beside them)."""
    for script in ui.glob("*.js"):
        try:
            if END_DATE_NOTICE in script.read_text(encoding="utf-8", errors="replace"):
                return True
        except OSError:
            continue
    return False


def nuitka_command(out: Path, version: str, onefile: bool, expiry_file: Path | None) -> list[str]:
    cmd = [
        sys.executable, "-m", "nuitka", str(APP / "mind_ready_app.py"),
        "--standalone", "--assume-yes-for-downloads", "--lto=no",
        f"--output-dir={out}", "--output-filename=MindReady.exe",
        "--include-package=app",  # the rules engine imports app.validators.<category> by name
        "--include-package=uvicorn",  # uvicorn picks its protocol and loop modules by name
        # pywintypes imports win32timezone by name when Excel hands back a date or time (a cell
        # value, a file's save time): Nuitka cannot see that import (Mind team, 2026-11: "No module
        # named win32timezone")
        "--include-module=win32timezone",
        "--enable-plugin=playwright",
        "--nofollow-import-to=*.tests",
        f"--include-data-dir={APP / 'rules'}=rules",
        f"--include-data-dir={APP / 'references'}=references",
        f"--include-data-dir={APP / 'config'}=config",
        f"--include-data-files={APP / 'settings.yaml'}=settings.yaml",
        f"--include-data-files={APP / 'VERSION'}=VERSION",
        f"--include-data-dir={UI}=ui",
    ]
    if expiry_file is not None:
        cmd.append(f"--include-data-files={expiry_file}=EXPIRY")
    if onefile:
        # unpacked once per build (a rebuild of the same version must not reuse an older unpacking)
        cmd += ["--onefile", f"--onefile-tempdir-spec={{CACHE_DIR}}/MindReady-program/{version}-{time.strftime('%Y%m%d-%H%M')}"]
    return cmd


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    out = Path(args.out).resolve()
    version = (APP / "VERSION").read_text(encoding="utf-8").strip()
    if not (UI / "index.html").is_file():
        print(f"the built screens are missing: {UI}")
        return 2
    last_day = end_date(args)
    if last_day and not screens_show_end_date(UI):
        print(f"the built screens cannot show the end date (they were built before it was added): build them first (npx ng build in {UI.parents[2]}), then run build.py again")
        return 2
    data = out / "package-data"
    if last_day:
        expiry_file: Path | None = write_expiry_file(data, last_day, args.contact)
    else:
        expiry_file = None
        (data / "EXPIRY").unlink(missing_ok=True)  # an earlier build's date, so nobody mistakes it for this one's
    until = f"This copy works until {long_date(last_day)}." if last_day else "This copy never stops working (--no-expiry)."
    print(until, flush=True)
    cmd = nuitka_command(out, version, args.onefile, expiry_file)
    print(" ".join(cmd), flush=True)
    # A fresh compiler cache for every build. Zig (Nuitka's C compiler here) caches an object by its C
    # source, and the files that embed the program's constants and the onefile payload (`#embed
    # "blobs/..."`) read the same every time: on 2026-11-05 it handed back the previous build's objects,
    # so the program carried old constants with new code (a segmentation fault at start) and the
    # onefile exe unpacked the previous version. Costs the compiler's own runtime once per build.
    env = dict(os.environ)
    zig = out / "zig-cache"
    if zig.exists():
        shutil.rmtree(zig, ignore_errors=True)
    env["ZIG_LOCAL_CACHE_DIR"] = str(zig / "local")
    env["ZIG_GLOBAL_CACHE_DIR"] = str(zig / "global")
    rc = subprocess.call(cmd, cwd=str(APP), env=env)
    if rc:
        return rc
    if args.onefile:
        print(f"program: {out / 'MindReady.exe'}\n{until}", flush=True)
        return 0
    built, final = out / "mind_ready_app.dist", out / "MindReady"
    if final.exists():
        shutil.rmtree(final)
    built.rename(final)
    (final / "START HERE.txt").write_text(START_HERE.format(version=version) + (f"\n{until}\n" if last_day else ""), encoding="utf-8")
    print(f"package: {final}\n{until}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
