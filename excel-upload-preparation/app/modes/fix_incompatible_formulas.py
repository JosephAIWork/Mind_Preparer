"""FIX_INCOMPATIBLE_FORMULAS:  focuses on  formula rules (modes/fix-incompatible-formulas.md).
ANALYZE+VALIDATE+ASK only -- rewriting formulas needs the LLM abstraction
(phase 7) and change-set application (phase 8), neither built yet."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from ..rules_engine import RulesEngine
from ..models import Mode
from .common import run_mode

CATEGORIES = ["file", "recalculation", "formula"]


def run(source_path: Path, work_dir: Path, config: dict[str, Any], engine: RulesEngine | None = None, ignore_sheets: list[str] | None = None, progress=None) -> dict[str, Any]:
    return run_mode(Mode.FIX_INCOMPATIBLE_FORMULAS, source_path, work_dir, config, categories=CATEGORIES, engine=engine, ignore_sheets=ignore_sheets, progress=progress)
