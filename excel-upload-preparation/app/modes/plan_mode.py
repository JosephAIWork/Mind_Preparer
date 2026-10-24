"""PLAN_MODE: analyze and report only, across every active rule (modes/plan-mode.md). It never
applies changes, so it needs no TRANSFORM step."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from ..rules_engine import RulesEngine
from ..models import Mode
from .common import run_mode


def run(source_path: Path, work_dir: Path, config: dict[str, Any], engine: RulesEngine | None = None, ignore_sheets: list[str] | None = None, progress=None) -> dict[str, Any]:
    return run_mode(Mode.PLAN_MODE, source_path, work_dir, config, categories=None, engine=engine, ignore_sheets=ignore_sheets, progress=progress)
