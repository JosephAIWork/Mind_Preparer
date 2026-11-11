"""READY-001.  A real  recalculation adapter exists (app/recalc.py, Excel COM
automation), but per modes/plan-mode.md ("Recalculation is optional...") and
instructions/workflow.md (recalculation is a VERIFY-step action, not part of
plain ANALYZE/VALIDATE), it is NOT run automatically as part of every  rule
pass -- that would launch Excel and take minutes on every  analysis. It's a
deliberate, separate action: scripts/recalculate_workbook.py,  or the UI's
Recalculate button. This finding reports  NOT_SUPPORTED because recalculation
was not run *as part of this pass* --  not because no adapter exists."""
from __future__ import annotations

from typing import Any


def successful_recalculation(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    return {
        "status": "NOT_SUPPORTED",
        "evidence": "DOCUMENTED_RULE",
        "message": (
            "Recalculation was not run as part of this analysis pass (it's a separate, "
            "explicit step -- see scripts/recalculate_workbook.py or the UI's Recalculate "
            "button), so PASS cannot be claimed for this workbook yet."
        ),
    }
