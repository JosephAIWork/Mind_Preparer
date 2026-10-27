"""Mind Ready as one program: the entry point of the compiled package (tools/package/build.py).

Starts the server on the port in settings.yaml and opens the browser once it answers. In the
package, the built screens sit next to this file in `ui/`; from the source it is the same as
run_mind_ready_web.bat. The single-file MindReady.exe unpacks itself out of sight, so a
settings.yaml placed next to the .exe is read instead of its own copy. When the program cannot
start, the window stays open on the message.
"""
from __future__ import annotations

import os
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
if (HERE / "ui").is_dir():
    os.environ.setdefault("MIND_READY_DIST", str(HERE / "ui"))
_beside = Path(sys.argv[0]).resolve().parent / "settings.yaml"
if _beside.is_file() and _beside.parent != HERE:
    os.environ.setdefault("MIND_READY_SETTINGS", str(_beside))


def _open_when_ready(port: int) -> None:
    import urllib.request

    for _ in range(90):
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{port}/api/health", timeout=2)
        except Exception:
            time.sleep(1)
            continue
        webbrowser.open(f"http://localhost:{port}")
        return


def _run() -> None:
    from app.settings import SETTINGS
    from app.web import server

    port = int(os.environ.get("MIND_READY_PORT") or SETTINGS.port)
    if not server._port_is_taken(port):  # already running: server.main() opens it itself
        threading.Thread(target=_open_when_ready, args=(port,), daemon=True).start()
    print("Mind Ready -- close this window to stop it.")
    server.main()


def main() -> None:
    try:
        _run()
    except KeyboardInterrupt:
        return
    except SystemExit as exc:
        if exc.code in (None, 0):
            return
        _wait()
    except Exception:
        traceback.print_exc()
        _wait()


def _wait() -> None:
    """Keep the window open on the message (a double-clicked program's window closes on exit)."""
    try:
        input("\nMind Ready could not start (see above). Press Enter to close this window.")
    except Exception:
        pass


if __name__ == "__main__":
    main()
