"""Scan progress (1.6.6): the stages a scan goes through, their weights and
the overall figure -- shared by the web server's status endpoint
(GET /api/sessions/{id}/status) and the Streamlit status bar.

The engine reports `progress(stage, message, fraction, **facts)` from
app/inventory.py (copy, load, inventory, names) and app/modes/common.py
(rules, report); the callers add convert (an .xlsb goes through Excel
first) and plan. Weights were measured on real models: reading the
workbook dominates, then the cell pass, then the rules.
"""
from __future__ import annotations

# Measured: an 8-sheet 52 MB model = load 23 s / inventory 12 s / rules 21 s;
# a 27-sheet 50 MB .xlsb with 3 sheets skipped = convert 47 s / load 13 s /
# inventory 13 s / rules 117 s (103k formulas) / plan 3 s. The weights are a
# compromise between the two profiles.
SCAN_STAGES: list[tuple[str, float]] = [
    ("convert", 0.15),
    ("copy", 0.02),
    ("load", 0.28),
    ("inventory", 0.20),
    ("names", 0.01),
    ("rules", 0.30),
    ("report", 0.01),
    ("plan", 0.03),
]

STAGE_TITLES = {
    "convert": "Converting .xlsb through Excel",
    "copy": "Copying the workbook",
    "load": "Reading the workbook",
    "inventory": "Scanning cells and grids",
    "names": "Reading defined names",
    "rules": "Checking the rules",
    "report": "Building the report",
    "plan": "Planning the fixes",
    "done": "Done",
}

# 1.7.2 -- the stages of an Apply (Prep, Fix panel, Grid Namer), reported by
# app/prep.py `apply_operations(progress=...)`. Only "apply_write" carries
# facts `done` / `total` (operations written so far / to write): the others
# are single Excel calls that cannot say how far they are.
APPLY_STAGES: list[str] = ["apply_copy", "apply_write", "apply_save", "apply_verify", "apply_log"]

APPLY_STAGE_TITLES = {
    "apply_copy": "Copying the workbook",
    "apply_write": "Writing the changes",
    "apply_save": "Saving the copy",
    "apply_verify": "Checking the copy opens in Excel",
    "apply_log": "Writing the change log",
    "done": "Done",
}


def overall_progress(stage: str | None, fraction: float | None, needs_convert: bool = False) -> float:
    """0..1 across the whole scan: the completed stages' weights plus the
    current stage's weight scaled by its in-stage fraction (half when the
    stage reports none)."""
    stages = [(st, w) for st, w in SCAN_STAGES if st != "convert" or needs_convert]
    names = [st for st, _ in stages]
    if stage == "done":
        return 1.0
    if stage not in names:
        return 0.0
    idx = names.index(stage)
    total = sum(w for _, w in stages) or 1.0
    done = sum(w for _, w in stages[:idx])
    current = stages[idx][1] * (float(fraction) if isinstance(fraction, (int, float)) else 0.5)
    return round(min(1.0, (done + current) / total), 3)
