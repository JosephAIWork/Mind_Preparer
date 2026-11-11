"""STRUCTURE_FIX:  focuses on  structure/format/risk rules (modes/structure-fix.md).
ANALYZE+VALIDATE+ASK only -- moving/splitting/merging grids needs change-set
application (phase 8), not built yet."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from ..rules_engine import RulesEngine
from ..models import Mode
from .common import run_mode

CATEGORIES = ["file", "recalculation", "structure", "format", "risk"]


def run(source_path: Path, work_dir: Path, config: dict[str, Any], engine: RulesEngine | None = None, ignore_sheets: list[str] | None = None, progress=None) -> dict[str, Any]:
    return run_mode(Mode.STRUCTURE_FIX, source_path, work_dir, config, categories=CATEGORIES, engine=engine, ignore_sheets=ignore_sheets, progress=progress)
