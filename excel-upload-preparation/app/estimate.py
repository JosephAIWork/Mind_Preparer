"""A rough time estimate for "Do everything" (1.11.6), per step, before it runs.

Built from what was measured on Palermo (27 MB), CNHI (37 MB) and Phoenix TVOG (179 MB): Excel opens a
workbook at about 1.2 s per MB and saves it at about 0.9 s per MB, each Excel session starts in about
8 s, and the app's own scan of the workbook (measured on this very file, the last scan) is the unit for
the rest -- the fixer's first full read of every cell takes about 0.6 of a scan, each fix pass about one
scan, and every change is followed by a re-scan. How many fix passes a workbook needs cannot be known
before Excel recalculates it, so every step has a low / likely / high figure; the front end shows the
range before the click and a countdown that corrects itself with the real duration of each step.
"""
from __future__ import annotations

from typing import Any

EXCEL_START_S = 8.0
OPEN_S_PER_MB = 1.2
SAVE_S_PER_MB = 0.9
FIRST_READ_PER_SCAN = 0.6  # the fixer reads every cell's value and formula once, after Excel recalculates
PASS_PER_SCAN = 1.0  # one fix pass: write, recalculate, read back, judge
WRITE_S_PER_OP = 0.02  # a Prep operation written through Excel
FALLBACK_SCAN_S_PER_FORMULA = 0.0035  # when the file was never scanned in this session (Palermo, CNHI)

CAVEAT = (
    "A rough guess: it depends on how many formula errors Excel finds when it recalculates, how many "
    "fix passes they need, and how busy this computer is (the same file has run 2-3 times slower at night)."
)


def _excel_session(mb: float) -> tuple[float, float]:
    """(start + open, save) seconds for a workbook of `mb` megabytes."""
    return EXCEL_START_S + max(5.0, OPEN_S_PER_MB * mb), max(3.0, SAVE_S_PER_MB * mb)


def estimate_chain(file_mb: float, scan_s: float | None, formula_count: int = 0, stored_errors: int | None = None, blocking_ops: int = 0, has_loop: bool = True) -> dict[str, Any]:
    """-> {"steps": [{id, low_s, mid_s, high_s, why}], "total": {low_s, mid_s, high_s}, "caveat", "basis"}.
    `scan_s`: the last scan of this file (without the .xlsb conversion); `stored_errors`: formula cells
    whose value Excel last stored is an error (None = not known)."""
    scan = float(scan_s) if scan_s and scan_s > 0 else max(30.0, FALLBACK_SCAN_S_PER_FORMULA * formula_count)
    open_s, save_s = _excel_session(file_mb)
    verify_s = open_s  # a second Excel opens the saved file
    steps: list[dict[str, Any]] = []

    # Fix all issues: open, recalculate and read every cell; then per pass write/recalculate/read; then
    # save, check it opens, re-scan. A clean workbook stops after the first read.
    first = open_s + FIRST_READ_PER_SCAN * scan
    finish = save_s + verify_s + scan

    def with_passes(n: float) -> float:
        return first + n * PASS_PER_SCAN * scan + finish

    if stored_errors == 0:
        low, mid, high = first, first, with_passes(1)  # Excel may still find errors the stored values hide
        why = "no formula error in the values Excel last stored: probably one read and done"
    elif stored_errors:
        low, mid, high = with_passes(1), with_passes(2), with_passes(4)
        why = f"{stored_errors:,} formula cells hold an error now: likely 1-4 fix passes"
    else:
        low, mid, high = first, with_passes(2), with_passes(4)
        why = "the errors are known only after Excel recalculates: 0-4 fix passes"
    steps.append({"id": "fixAll", "low_s": low, "mid_s": mid, "high_s": high, "why": why})

    # Fix the formulas and the structure: one or two rounds of open, write, save, re-scan (the check
    # that the file opens runs alongside the re-scan)
    if blocking_ops:
        one = open_s + WRITE_S_PER_OP * blocking_ops + save_s + max(scan, verify_s)
        steps.append({"id": "prepare", "low_s": one, "mid_s": 1.5 * one, "high_s": 2.5 * one, "why": f"{blocking_ops:,} planned change(s), usually one or two rounds"})
    else:
        steps.append({"id": "prepare", "low_s": 5.0, "mid_s": 5.0, "high_s": 2.0 * (open_s + save_s + scan), "why": "nothing planned now (the fixer may leave something)"})

    # Build the Mind loop: one open, write, save, re-scan
    one = open_s + save_s + max(scan, verify_s)
    if has_loop:
        steps.append({"id": "loop", "low_s": one, "mid_s": one, "high_s": 1.5 * one, "why": "one round through Excel"})
    else:
        steps.append({"id": "loop", "low_s": 0.0, "mid_s": 0.0, "high_s": 0.0, "why": "skipped: no loop to build"})

    # Recalculate: open, full rebuild (seconds), save, find the error cells
    rc = open_s + 5.0 + save_s + 5.0
    steps.append({"id": "recalc", "low_s": rc, "mid_s": rc, "high_s": 1.5 * rc, "why": "one full recalculation in Excel"})

    for s in steps:
        for k in ("low_s", "mid_s", "high_s"):
            s[k] = round(s[k])
    total = {k: sum(s[k] for s in steps) for k in ("low_s", "mid_s", "high_s")}
    return {
        "steps": steps,
        "total": total,
        "caveat": CAVEAT,
        "basis": {"file_mb": round(file_mb, 1), "scan_s": round(scan), "scan_measured": bool(scan_s and scan_s > 0), "stored_errors": stored_errors, "blocking_ops": blocking_ops},
    }
