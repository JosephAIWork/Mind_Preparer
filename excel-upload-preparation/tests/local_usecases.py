"""Client models' local copies, by anonymous code (UC-1, UC-2 ...), for the tests that check a real
model when its copy is on this machine.

Nothing about the models is in the repo: the files and what the tests expect of them (their sheet,
cell and name values) live next to the copies, in `%LOCALAPPDATA%\\MindReady\\autofix\\usecases`:
  index.json     {"UC-1": "<file name in that folder, or an absolute path>", ...}
  expected.json  {"UC-1": {"driver": "...", ...}, ...}
Without them (any other machine) those tests skip.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

USECASES = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "MindReady" / "autofix" / "usecases"


def _local_json(name: str) -> dict[str, Any]:
    try:
        return json.loads((USECASES / name).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def usecase(code: str) -> Path | None:
    """The local copy of the model coded `code`, or None when this machine has none."""
    entry = _local_json("index.json").get(code)
    if not entry:
        return None
    path = Path(entry)
    path = path if path.is_absolute() else USECASES / path
    return path if path.is_file() else None


def usecase_or_skip(code: str) -> tuple[Path, dict[str, Any]]:
    """The local copy and what the tests expect of it; the test skips when either is missing."""
    path, expected = usecase(code), _local_json("expected.json").get(code)
    if path is None or not expected:
        pytest.skip(f"{code}: no local copy here")
    return path, expected
