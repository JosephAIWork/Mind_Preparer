"""Stage timings: where the time goes on a big workbook.

Every slow step (starting Excel, opening, writing, recalculating, saving, the check that the file opens,
the scan) is timed with `stage(name)` and appended as one JSON line to
`perf\\timings.jsonl` in the data folder (settings.yaml) (`MINDREADY_PERF_LOG` overrides the file;
`MINDREADY_PERF_LOG=0` turns it off). Writing the line never fails the step it measures.
"""
from __future__ import annotations

import contextlib
import json
import os
import threading
import time
from pathlib import Path
from typing import Any, Iterator

from .settings import SETTINGS

_LOCK = threading.Lock()


def log_path() -> Path | None:
    env = os.environ.get("MINDREADY_PERF_LOG")
    if env is not None:
        if env.strip().lower() in ("", "0", "off", "no", "false"):
            return None
        return Path(env)
    return SETTINGS.data_folder / "perf" / "timings.jsonl"


def record(name: str, seconds: float, **facts: Any) -> None:
    path = log_path()
    if path is None:
        return
    line = {"at": time.strftime("%Y-%m-%d %H:%M:%S"), "stage": name, "s": round(seconds, 2), "thread": threading.current_thread().name, **facts}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with _LOCK, path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(line, default=str) + "\n")
    except Exception:
        pass


_STAGES = threading.local()


def current_stage() -> str | None:
    """The innermost stage this thread is in (for a problem report)."""
    stack = getattr(_STAGES, "stack", None)
    return stack[-1] if stack else None


@contextlib.contextmanager
def stage(name: str, **facts: Any) -> Iterator[dict[str, Any]]:
    """Times the block; the caller may add facts to the yielded dict (they are logged with it).
    An error leaving the block carries the innermost stage it left (`mind_ready_stage`), for the
    problem report (app/feedback.py)."""
    extra: dict[str, Any] = {}
    stack = _STAGES.__dict__.setdefault("stack", [])
    stack.append(name)
    t = time.perf_counter()
    try:
        yield extra
    except BaseException as exc:
        if getattr(exc, "mind_ready_stage", None) is None:
            try:
                exc.mind_ready_stage = name  # type: ignore[attr-defined]
            except Exception:
                pass
        raise
    finally:
        stack.pop()
        record(name, time.perf_counter() - t, **facts, **extra)
