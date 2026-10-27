"""Build Mind Ready as a compiled package (Nuitka): a folder with MindReady.exe that needs no Python.

    <build venv>\\Scripts\\python.exe tools/package/build.py <out dir> [--onefile]

The build venv holds requirements.txt and nuitka (nothing else, so nothing else is pulled in). Nuitka
fetches its own C compiler the first time. The result is <out dir>/MindReady/, the folder to zip:
MindReady.exe plus what the app reads next to it (rules, references, config, settings.yaml, VERSION)
and the built screens in ui/ (run `npx ng build --configuration production` in MindReadyUI first).
A second build reuses Nuitka's compile cache and is much shorter.

--onefile makes one file instead, <out dir>/MindReady.exe, to share as a link: it unpacks itself the
first time it starts (into %LOCALAPPDATA%\\MindReady-program\\<version and build time>, so later
starts are quick) and reads a settings.yaml placed next to it (mind_ready_app.py).
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

APP = Path(__file__).resolve().parents[2]  # excel-upload-preparation
UI = APP.parent / "MindReadyUI" / "dist" / "mind-ready-ui" / "browser"
START_HERE = """Mind Ready {version}

Start: double-click MindReady.exe. A window opens (close it to stop Mind Ready) and the app opens in
your browser at http://localhost:8600.

Settings (port, folders, the Mind and AI addresses): settings.yaml in this folder. Restart after a change.
The AI key: pasted in the app the first time (main screen, or Advanced > Assistant); it is saved for
your Windows account, not in this folder.

Needs: Windows, desktop Excel, the MM for Excel add-in; Microsoft Edge for Run in Mind.
"""


def main() -> int:
    args = [a for a in sys.argv[1:] if a != "--onefile"]
    onefile = "--onefile" in sys.argv[1:]
    if len(args) != 1:
        print(__doc__)
        return 2
    out = Path(args[0]).resolve()
    version = (APP / "VERSION").read_text(encoding="utf-8").strip()
    if not (UI / "index.html").is_file():
        print(f"the built screens are missing: {UI}")
        return 2
    cmd = [
        sys.executable, "-m", "nuitka", str(APP / "mind_ready_app.py"),
        "--standalone", "--assume-yes-for-downloads", "--lto=no",
        f"--output-dir={out}", "--output-filename=MindReady.exe",
        "--include-package=app",  # the rules engine imports app.validators.<category> by name
        "--include-package=uvicorn",  # uvicorn picks its protocol and loop modules by name
        "--enable-plugin=playwright",
        "--nofollow-import-to=*.tests",
        f"--include-data-dir={APP / 'rules'}=rules",
        f"--include-data-dir={APP / 'references'}=references",
        f"--include-data-dir={APP / 'config'}=config",
        f"--include-data-files={APP / 'settings.yaml'}=settings.yaml",
        f"--include-data-files={APP / 'VERSION'}=VERSION",
        f"--include-data-dir={UI}=ui",
    ]
    if onefile:
        # unpacked once per build (a rebuild of the same version must not reuse an older unpacking)
        cmd += ["--onefile", f"--onefile-tempdir-spec={{CACHE_DIR}}/MindReady-program/{version}-{time.strftime('%Y%m%d-%H%M')}"]
    print(" ".join(cmd), flush=True)
    rc = subprocess.call(cmd, cwd=str(APP))
    if rc:
        return rc
    if onefile:
        print(f"program: {out / 'MindReady.exe'}", flush=True)
        return 0
    built, final = out / "mind_ready_app.dist", out / "MindReady"
    if final.exists():
        shutil.rmtree(final)
    built.rename(final)
    (final / "START HERE.txt").write_text(START_HERE.format(version=version), encoding="utf-8")
    print(f"package: {final}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
