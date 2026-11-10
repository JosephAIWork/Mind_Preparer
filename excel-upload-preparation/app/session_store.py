"""Saved workbooks: a session's files and state kept on disk between runs of the app.

Each session lives in `SETTINGS.data_folder / "sessions" / <session id>` (its work folder: the
upload, every version, the analysis copies, the recalculations, the Mind inputs, the instances zip,
with the same layout the %TEMP% folder had) next to one `state.json` that holds everything the
`Session` dataclass (app/web/server.py) can save: the versions and their files, the last check
result, the plan, the recalculation, the plans of the Mind loop and the instances, the bring-in and
next-quarter results, the grid names, the settings chosen. Locks, threads and live job state are not
saved: a job that was running when the app stopped shows as stopped when the session is opened
again, with the sentence `CLOSED_WHILE_RUNNING`.

`save(s)` is called by the server after every step that changes a session's state; it never raises
(a save must not break the step: the problem is printed and the step's answer goes out as before).
`load(session_id, Session)` reads the folder back into a `Session`; `list_saved()` reads the state
files alone (no session is loaded); `remove(session_id)` deletes the folder. The app loads nothing at
start: the screens show the list and wait (docs/SAVED_WORKBOOKS.md).
"""
from __future__ import annotations

import json
import os
import re
import shutil
import threading
import time
from pathlib import Path
from typing import Any

from app.settings import SETTINGS

STATE_FILE = "state.json"
CLOSED_WHILE_RUNNING = "The app was closed while this ran."
_ID = re.compile(r"^[A-Za-z0-9_-]{1,64}$")
_WRITE_LOCK = threading.Lock()  # one state file written at a time (jobs end in their own threads)

# The Session fields saved as they are (JSON-safe already) ...
_PLAIN_FIELDS = (
    "mode", "versions", "current_version_id", "result", "plan", "last_delta", "recalc", "recalc_version_id",
    "grid_names", "created_at", "size", "ignore_sheets", "pivot_mode", "pending", "scan", "plan_history",
    "apply_history", "apply_status", "model_map", "model_map_version_id", "input_seq", "outside_vba",
    "outside_inputs", "outside_run", "next_quarter", "brought_in_files", "instance_finder",
    "instance_finder_version_id", "instance_finder_assistant", "instance_plan_default",
    "instance_plan_default_version_id", "name", "last_opened",
)
# ... the fields that hold a path (kept relative to the session's folder when inside it; the Mind
# inputs folders are held as text in the Session, the others as Paths) ...
_PATH_FIELDS = ("recalc_path", "instances_zip", "mind_inputs_dir")
_PATH_DICT_FIELDS = ("paths", "files", "mind_inputs_dirs")
_TEXT_PATHS = ("mind_inputs_dir", "mind_inputs_dirs")
# ... and the live job states: "running" cannot survive a restart
_JOB_FIELDS = ("scan", "apply_status", "outside_run", "next_quarter")


def valid_id(session_id: str) -> bool:
    return bool(_ID.match(session_id or ""))


def sessions_root() -> Path:
    """`SETTINGS.data_folder` is read at call time: the tests point it at a folder of their own."""
    return Path(SETTINGS.data_folder) / "sessions"


def folder_of(session_id: str) -> Path:
    if not valid_id(session_id):
        raise ValueError(f"not a session id: {session_id!r}")
    return sessions_root() / session_id


def new_folder(session_id: str) -> Path:
    """The work folder of a new session (made, empty)."""
    folder = folder_of(session_id)
    folder.mkdir(parents=True, exist_ok=False)
    return folder


# --- paths and odd shapes ---------------------------------------------------------------------
def _pack_path(p: Any, folder: Path) -> str | None:
    if p is None:
        return None
    path = Path(p)
    try:
        return path.relative_to(folder).as_posix()
    except ValueError:
        return str(path)


def _unpack_path(text: Any, folder: Path) -> Path | None:
    if text is None:
        return None
    path = Path(str(text))
    return path if path.is_absolute() else folder / path


def _pack_pairs(pairs: set[tuple[str, str]] | None) -> list[list[str]] | None:
    return None if pairs is None else sorted([list(p) for p in pairs])


def _unpack_pairs(rows: Any) -> set[tuple[str, str]] | None:
    return None if rows is None else {(str(r[0]), str(r[1])) for r in rows}


def _stopped(job: dict[str, Any] | None) -> dict[str, Any]:
    """A job state as it is, unless it was running: then stopped, with the sentence, as an error."""
    job = dict(job or {})
    if job.get("state") == "running":
        now = time.time()
        job.update(state="error", error=CLOSED_WHILE_RUNNING, message=CLOSED_WHILE_RUNNING, title="Stopped", finished=job.get("finished") or now, finished_at=job.get("finished_at") or now, updated=now, updated_at=now)
    return job


# --- the state file -----------------------------------------------------------------------------
def to_state(s: Any, autofix: dict[str, Any] | None = None) -> dict[str, Any]:
    """Everything of `s` that can be saved, JSON-safe, paths relative to the session's folder."""
    folder = Path(s.work_dir)
    state: dict[str, Any] = {"id": s.id, "saved_at": time.time()}
    for name in _PLAIN_FIELDS:
        state[name] = getattr(s, name, None)
    for name in _PATH_FIELDS:
        state[name] = _pack_path(getattr(s, name, None), folder)
    for name in _PATH_DICT_FIELDS:
        state[name] = {k: _pack_path(v, folder) for k, v in (getattr(s, name, None) or {}).items()}
    state["inputs"] = [{**i, "path": _pack_path(i.get("path"), folder)} for i in (s.inputs or [])]
    state["original_ref_formulas"] = _pack_pairs(s.original_ref_formulas)
    state["ref_formulas"] = {k: _pack_pairs(v) for k, v in (s.ref_formulas or {}).items()}
    state["original_environment"] = None if s.original_environment is None else [[k[0], k[1], v] for k, v in s.original_environment.items()]
    state["baselines"] = {k: {**b, "ref_formulas": _pack_pairs(b.get("ref_formulas"))} for k, b in (s.baselines or {}).items()}
    state["original_baseline"] = next((k for k, b in (s.baselines or {}).items() if b is s.original_baseline), None) if s.original_baseline is not None else None
    # the automatic fixer's state lives next to the sessions in the server (AUTOFIX): its public part
    state["autofix"] = None if autofix is None else {k: v for k, v in autofix.items() if k != "stop"}
    return json.loads(json.dumps(state, default=str))  # dates, Paths and openpyxl values as text


def apply_state(s: Any, state: dict[str, Any]) -> dict[str, Any] | None:
    """Fill `s` (a Session made with its id, folder and mode) from a state dict; returns the saved
    automatic-fixer state (None when there was none), stopped when it was running."""
    folder = Path(s.work_dir)
    for name in _PLAIN_FIELDS:
        if name in state and name != "mode":
            setattr(s, name, state[name])
    for name in _PATH_FIELDS:
        path = _unpack_path(state.get(name), folder)
        setattr(s, name, str(path) if path is not None and name in _TEXT_PATHS else path)
    for name in _PATH_DICT_FIELDS:
        unpacked = {k: _unpack_path(v, folder) for k, v in (state.get(name) or {}).items()}
        setattr(s, name, {k: str(v) for k, v in unpacked.items()} if name in _TEXT_PATHS else unpacked)
    s.inputs = [{**i, "path": str(_unpack_path(i.get("path"), folder))} for i in (state.get("inputs") or [])]
    s.original_ref_formulas = _unpack_pairs(state.get("original_ref_formulas"))
    s.ref_formulas = {k: _unpack_pairs(v) or set() for k, v in (state.get("ref_formulas") or {}).items()}
    env = state.get("original_environment")
    s.original_environment = None if env is None else {(str(r[0]), str(r[1])): r[2] for r in env}
    s.baselines = {k: {**b, "ref_formulas": _unpack_pairs(b.get("ref_formulas")) or set()} for k, b in (state.get("baselines") or {}).items()}
    key = state.get("original_baseline")
    s.original_baseline = s.baselines.get(key) if key else None
    for name in _JOB_FIELDS:
        setattr(s, name, _stopped(getattr(s, name, None)))
    if not s.scan:
        s.scan = {}
    autofix = state.get("autofix")
    return _stopped(autofix) if autofix else None


def save(s: Any, autofix: dict[str, Any] | None = None) -> Path | None:
    """Write the session's state file; never raises (the problem is printed)."""
    try:
        folder = Path(s.work_dir)
        folder.mkdir(parents=True, exist_ok=True)
        text = json.dumps(to_state(s, autofix), ensure_ascii=False)
        target = folder / STATE_FILE
        with _WRITE_LOCK:
            tmp = folder / f"{STATE_FILE}.tmp"
            tmp.write_text(text, encoding="utf-8")
            os.replace(tmp, target)
        return target
    except Exception as exc:  # a save must never break the step it follows
        print(f"Could not save the session {getattr(s, 'id', '?')}: {type(exc).__name__}: {str(exc)[:200]}", flush=True)
        return None


def read_state(session_id: str) -> dict[str, Any] | None:
    """The state file of a saved session, or None when there is none."""
    if not valid_id(session_id):
        return None
    path = folder_of(session_id) / STATE_FILE
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def load(session_id: str, session_cls: Any) -> tuple[Any, dict[str, Any] | None] | None:
    """A Session built from its folder, with the saved automatic-fixer state; None when there is no
    saved session of that id."""
    state = read_state(session_id)
    if state is None:
        return None
    s = session_cls(id=session_id, work_dir=folder_of(session_id), mode=state.get("mode") or "plan")
    autofix = apply_state(s, state)
    return s, autofix


def _display_name(state: dict[str, Any]) -> str:
    if state.get("name"):
        return str(state["name"])
    versions = state.get("versions") or []
    return str(versions[0].get("file_name") or "") if versions else ""


def _file_bytes(folder: Path, state: dict[str, Any]) -> int:
    paths = state.get("paths") or {}
    path = _unpack_path(paths.get(state.get("current_version_id")) or next(iter(paths.values()), None), folder)
    try:
        return path.stat().st_size if path is not None else 0
    except OSError:
        return 0


def summary_of(state: dict[str, Any]) -> dict[str, Any]:
    """One line of the Recent workbooks list."""
    sid = str(state.get("id") or "")
    folder = folder_of(sid) if valid_id(sid) else Path()
    versions = state.get("versions") or []
    return {
        "id": sid,
        "name": _display_name(state),
        "uploaded_at": state.get("created_at"),
        "last_opened": state.get("last_opened") or state.get("saved_at") or state.get("created_at"),
        "versions": len(versions),
        "current_version": state.get("current_version_id"),
        "bytes": _file_bytes(folder, state),
        "has_result": bool(state.get("result")),
        "pending": bool(state.get("pending")),
    }


def list_saved() -> list[dict[str, Any]]:
    """What is on disk, the last opened first; a folder without a readable state file is skipped."""
    root = sessions_root()
    if not root.is_dir():
        return []
    rows = []
    for child in root.iterdir():
        if not child.is_dir() or not valid_id(child.name):
            continue
        state = read_state(child.name)
        if state is None:
            continue
        state.setdefault("id", child.name)
        rows.append(summary_of(state))
    rows.sort(key=lambda r: (r.get("last_opened") or 0, r.get("uploaded_at") or 0), reverse=True)
    return rows


def remove(session_id: str) -> bool:
    """Delete the session's folder; True when there was one. A file still open (an Excel that was not
    released) leaves the folder behind for the next remove: the state file goes first, so the list
    never shows it again."""
    if not valid_id(session_id):
        return False
    folder = folder_of(session_id)
    if not folder.exists():
        return False
    with _WRITE_LOCK:
        for name in (STATE_FILE, f"{STATE_FILE}.tmp"):
            with_suppress_unlink(folder / name)
    shutil.rmtree(folder, ignore_errors=True)
    return True


def with_suppress_unlink(path: Path) -> None:
    try:
        path.unlink()
    except OSError:
        pass
