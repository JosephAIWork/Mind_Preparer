"""A rough time estimate for "Do everything", per step, before it runs.

Built from what was measured (four Do everything runs on two large workbooks, 27 MB and 170 MB, with
bulk writes and the per-session Excel worker): an Excel session starts in about 5 s, opens a
workbook in about 8 s + 0.3 s per MB and saves it in about 10 s + 0.18 s per MB, and the app's own scan
of the workbook (measured on this very file, the last scan) is the unit for the rest -- the fixer's first
full read of every cell takes about 0.3 of a scan, each fix pass about 0.1 of a scan (it writes in bulk),
and every change is followed by a re-scan of 0.7-1.5 scans (the worker's Excel opens the next copy
alongside it). How many fix passes a workbook needs cannot be known
before Excel recalculates it, so every step has a low / likely / high figure; the front end shows the
range before the click and a countdown that corrects itself with the real duration of each step.
"""
from __future__ import annotations

from typing import Any

EXCEL_START_S = 5.0
OPEN_S = 8.0
OPEN_S_PER_MB = 0.3
SAVE_S = 10.0
SAVE_S_PER_MB = 0.18
FIRST_READ_PER_SCAN = 0.3  # the fixer reads every cell's value and formula once, after Excel recalculates
PASS_PER_SCAN = 0.1  # one fix pass: write (in bulk), recalculate, read back, judge
RESCAN = {"low": 0.7, "mid": 1.2, "high": 1.5}  # the re-scan of a new version, in scans of the last one
RECALC_AHEAD_S = {"low": 5.0, "mid": 15.0}  # Recalculate worked out ahead by the worker: wait, read, answer
BASELINE_EXTRA_S = 25.0  # the original's recalculation besides its Excel and open: rebuild, error read, answer
WRITE_S_PER_OP = 0.02  # a Prep operation written through Excel
FALLBACK_SCAN_S_PER_FORMULA = 0.0035  # when the file was never scanned in this session (measured on two large models)
# a pivot table turned into cells: copy, paste and clear through a scratch sheet, then (to formulas)
# about ten Excel calls per value cell -- what the pivot adds up, its value, the formula, the check
PIVOT_S_EACH = 3.0
PIVOT_S_PER_CELL = 0.01

CAVEAT = (
    "A rough guess: it depends on how many formula errors Excel finds when it recalculates, how many "
    "fix passes they need, and how busy this computer is (the same file has run 2-3 times slower at night)."
)


def _excel_session(mb: float) -> tuple[float, float]:
    """(start + open, save) seconds for a workbook of `mb` megabytes."""
    return EXCEL_START_S + OPEN_S + OPEN_S_PER_MB * mb, SAVE_S + SAVE_S_PER_MB * mb


def pivot_seconds(pivots: list[tuple[str, int]]) -> float:
    """Seconds to turn the pivot tables of a Prep round into cells: `pivots` = (operation, cells of the
    pivot) per pivot; only "pivot_to_formulas" pays per cell (above MAX_FORMULA_CELLS it becomes values)."""
    from .pivots import MAX_FORMULA_CELLS

    return sum(PIVOT_S_EACH + (PIVOT_S_PER_CELL * cells if op == "pivot_to_formulas" and cells <= MAX_FORMULA_CELLS else 0.0) for op, cells in pivots)


def estimate_chain(file_mb: float, scan_s: float | None, formula_count: int = 0, stored_errors: int | None = None, blocking_ops: int = 0, has_loop: bool = True, worker: bool = False,
                   recalc_ahead: bool = True, baseline_mb: float = 0.0, pivots: list[tuple[str, int]] = ()) -> dict[str, Any]:
    """-> {"steps": [{id, low_s, mid_s, high_s, why}], "total": {low_s, mid_s, high_s}, "caveat", "basis"}.
    `scan_s`: the last scan of this file (without the .xlsb conversion); `stored_errors`: formula cells
    whose value Excel last stored is an error (None = not known). `worker`: the session's
    Excel worker opens each next copy while the version is analysed, so Prep rounds and the loop
    wait for no Excel start, open or second check, and (`recalc_ahead`) Recalculate after the loop
    is ready. `baseline_mb`: the size of the original when its own recalculation (the baseline the
    errors left are compared with) has not run yet in this session, 0 when it has or will not be
    needed. `pivots`: (operation, cells) of every pivot table Prep turns into cells (pivot_seconds)."""
    scan = float(scan_s) if scan_s and scan_s > 0 else max(30.0, FALLBACK_SCAN_S_PER_FORMULA * formula_count)
    open_s, save_s = _excel_session(file_mb)
    verify_s = open_s  # a second Excel opens the saved file
    steps: list[dict[str, Any]] = []

    # Fix all issues: open, recalculate and read every cell; then per pass write/recalculate/read; then
    # save, check it opens, re-scan. A clean workbook stops after the first read.
    first = open_s + FIRST_READ_PER_SCAN * scan

    def rescan(k: str) -> float:
        return RESCAN[k] * scan

    def with_passes(n: float, k: str) -> float:
        return first + n * PASS_PER_SCAN * scan + save_s + verify_s + rescan(k)

    if stored_errors == 0:
        low, mid, high = first, first, with_passes(1, "high")  # Excel may still find errors the stored values hide
        why = "no formula error in the values Excel last stored: probably one read and done"
    elif stored_errors:
        low, mid, high = with_passes(1, "low"), with_passes(2, "mid"), with_passes(4, "high")
        why = f"{stored_errors:,} formula cells hold an error now: likely 1-4 fix passes"
    else:
        low, mid, high = first, with_passes(2, "mid"), with_passes(4, "high")
        why = "the errors are known only after Excel recalculates: 0-4 fix passes"
    steps.append({"id": "fixAll", "low_s": low, "mid_s": mid, "high_s": high, "why": why})

    # Fix the formulas and the structure: one round of open, write, save, re-scan (the check that the
    # file opens runs alongside the re-scan); a second round, when there is one, is usually empty
    if blocking_ops:
        piv = pivot_seconds(list(pivots))  # once, in the first round (the pivot tables are cells after it)

        def one(k: str) -> float:
            return (0.0 if worker else open_s) + WRITE_S_PER_OP * max(0, blocking_ops - len(pivots)) + save_s + (rescan(k) if worker else max(rescan(k), verify_s))

        why = f"{blocking_ops:,} planned change(s), usually one round" + (f", {len(pivots):,} pivot table(s) turned into cells" if pivots else "")
        steps.append({"id": "prepare", "low_s": one("low") + piv, "mid_s": one("mid") + piv, "high_s": 2.0 * one("high") + 1.5 * piv, "why": why})
    else:
        steps.append({"id": "prepare", "low_s": 5.0, "mid_s": 5.0, "high_s": 2.0 * (open_s + save_s + rescan("high")), "why": "nothing planned now (the fixer may leave something)"})

    # Build the Mind loop: one open, write, save, re-scan
    def loop_round(k: str) -> float:
        return (save_s + rescan(k)) if worker else open_s + save_s + max(rescan(k), verify_s)

    if has_loop:
        steps.append({"id": "loop", "low_s": loop_round("low"), "mid_s": loop_round("mid"), "high_s": loop_round("high"), "why": "one round through Excel"})
    else:
        steps.append({"id": "loop", "low_s": 0.0, "mid_s": 0.0, "high_s": 0.0, "why": "skipped: no loop to build"})

    # Recalculate: open, full rebuild (seconds), save, find the error cells
    ahead = worker and recalc_ahead and has_loop
    rc_low = RECALC_AHEAD_S["low"] if ahead else open_s + 5.0 + save_s + 5.0
    rc = RECALC_AHEAD_S["mid"] if ahead else rc_low
    # with errors left, the original is recalculated once per session too (another Excel and a full
    # open); the fixer's step may already have paid it, so the low figure leaves it out
    base = _excel_session(baseline_mb)[0] + BASELINE_EXTRA_S if baseline_mb > 0 else 0.0
    steps.append({"id": "recalc", "low_s": rc_low, "mid_s": rc + base, "high_s": 1.5 * rc + base,
                  "why": "one full recalculation in Excel" + (", and one of the original to compare the errors left with" if base else "")})

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
