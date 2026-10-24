"""PREP_MIND_LOOPS: focuses on loop/result/resize rules (modes/prep-mind-loops.md).
This mode analyzes, validates and asks only: it applies nothing, so `output` stays null and the
status is capped by the NOT_SUPPORTED / REQUIRES_USER_INPUT findings that would block a
correction. Changes are made through Prep, the automatic fixer and the Mind loop builder."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from ..rules_engine import RulesEngine
from ..models import Mode
from .common import run_mode

CATEGORIES = ["file", "recalculation", "loop"]


def run(source_path: Path, work_dir: Path, config: dict[str, Any], engine: RulesEngine | None = None, ignore_sheets: list[str] | None = None, progress=None) -> dict[str, Any]:
    return run_mode(Mode.PREP_MIND_LOOPS, source_path, work_dir, config, categories=CATEGORIES, engine=engine, ignore_sheets=ignore_sheets, progress=progress)
