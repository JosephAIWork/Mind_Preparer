"""CAL-*  validators: KB  'Iterations: design sequential calculations in your
model' (/CalculationSteps grid: workbook name | integer step, equal steps
run in parallel; MM_ITERATIONS once per workbook; /iterationinput.(name)
paired with /iterationoutput.(name); /NoKeepResults) and 'MM_ITERATIONS
function'."""
from __future__ import annotations

from typing import Any

from ..formula_utils import as_number, unquote
from ..grids import all_grids, grid_flags
from ..inventory import mm_calls
from ._common import finding, flagged_grids, fmt_grids, fmt_sites, grid_location, grid_rows, is_blank


def _steps_grid(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return flagged_grids(analysis, "calculationsteps")


def _step_rows(analysis: dict[str, Any], g: dict[str, Any]) -> list[dict[str, Any]]:
    rows = grid_rows(analysis, g)
    body = rows[1:] if g["header_is_all_text"] and len(rows) > 1 and not isinstance(rows[0][1] if len(rows[0]) > 1 else None, (int, float)) else rows
    out = []
    for r, row in enumerate(body, start=g["first_row"] + (1 if body is not rows else 0)):
        if not row or all(is_blank(v) for v in row):
            continue
        out.append({"row": r, "workbook": row[0], "step": row[1] if len(row) > 1 else None})
    return out


def cal_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _steps_grid(analysis)
    if len(grids) > 1:
        return finding("ERROR", f"{len(grids)} /CalculationSteps grids -- one per project: {fmt_grids(grids)}", {"grids": [f"{g['sheet']}!{g['ref']}" for g in grids]}, location=grid_location(grids[0]))
    if not grids:
        return finding("PASS", "No /CalculationSteps grid (optional for a single-workbook model; workbooks are otherwise calculated in one pass).", {"grids": [], "steps": []})
    rows = _step_rows(analysis, grids[0])
    return finding("PASS", f"/CalculationSteps grid at {grids[0]['sheet']}!{grids[0]['ref']} orders {len(rows)} workbook(s).", {"grids": [f"{grids[0]['sheet']}!{grids[0]['ref']}"], "steps": [{"workbook": r["workbook"], "step": r["step"]} for r in rows]})


def cal_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _steps_grid(analysis)
    if not grids:
        return finding("PASS", "No /CalculationSteps grid to validate.", {"problems": []})
    g = grids[0]
    problems, warnings = [], []
    if g["n_cols"] < 2:
        problems.append({"issue": "grid needs 2 columns: workbook name | calculation step"})
    rows = _step_rows(analysis, g)
    steps = []
    for r in rows:
        if is_blank(r["workbook"]):
            problems.append({"row": r["row"], "issue": "empty workbook name"})
        step = r["step"]
        n = step if isinstance(step, (int, float)) and not isinstance(step, bool) else as_number(str(step)) if step is not None else None
        if n is None or n != int(n) or n < 1:
            problems.append({"row": r["row"], "workbook": r["workbook"], "issue": f"step '{step}' is not a positive integer"})
        else:
            steps.append(int(n))
    if steps:
        if min(steps) != 1:
            warnings.append(f"steps start at {min(steps)}, not 1")
        gaps = sorted(set(range(1, max(steps) + 1)) - set(steps))
        if gaps:
            warnings.append(f"step number(s) {gaps} are skipped")
    this_wb = (analysis["workbooks"][0].get("file_name") or "").strip().lower()
    names = {str(r["workbook"]).strip().lower() for r in rows if not is_blank(r["workbook"])}
    if this_wb and this_wb not in names:
        warnings.append(f"this workbook ('{analysis['workbooks'][0].get('file_name')}') is not listed -- names must match the Excel file names without extension")
    observed = {"problems": problems, "warnings": warnings, "steps": [{"workbook": r["workbook"], "step": r["step"]} for r in rows]}
    if problems:
        return finding("ERROR", f"/CalculationSteps problem(s): {problems[:6]}", observed, location=grid_location(g))
    if warnings:
        return finding("WARNING", f"/CalculationSteps ordering is valid but: {'; '.join(warnings)}", observed, location=grid_location(g))
    return finding("PASS", f"/CalculationSteps ordering is valid: {len(rows)} workbook(s), steps 1..{max(steps) if steps else 0}.", observed)


def cal_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _steps_grid(analysis)
    if not grids:
        return finding("PASS", "No /CalculationSteps grid; parallelisation applies only to multi-workbook projects.", None, evidence="RECOMMENDATION")
    rows = _step_rows(analysis, grids[0])
    steps = [int(r["step"]) for r in rows if isinstance(r["step"], (int, float)) and not isinstance(r["step"], bool)]
    parallel = len(steps) - len(set(steps))
    if len(steps) >= 3 and parallel == 0:
        return finding("WARNING", f"{len(steps)} workbooks each get a distinct step -- nothing runs in parallel. Give independent workbooks the same step number (KB: 'it is more optimized to parallelize as possible').", {"workbooks": len(steps), "parallel_pairs": 0}, location=grid_location(grids[0]), evidence="RECOMMENDATION")
    return finding("PASS", f"{len(steps)} workbook(s) over {len(set(steps))} step(s) ({parallel} share a step with another).", {"workbooks": len(steps), "parallel_pairs": parallel}, evidence="RECOMMENDATION")


def cal_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    calls = mm_calls(analysis, "MM_ITERATIONS")
    if not calls:
        return finding("PASS", "No MM_ITERATIONS (no sequential iterations on this workbook).", {"calls": []})
    sites = [{"sheet": c["sheet"], "cell": c["cell"], "args": c["args"]} for c in calls]
    bad = [s for s in sites if len(s["args"]) != 2 or unquote(s["args"][0]) is None]
    if len(calls) > 1:
        return finding("ERROR", f"MM_ITERATIONS appears {len(calls)} times; the KB allows it only once per workbook: {fmt_sites(sites)}", {"calls": sites}, location={"sheet": sites[0]["sheet"], "cell": sites[0]["cell"]})
    if bad:
        return finding("ERROR", f"MM_ITERATIONS must be MM_ITERATIONS(\"Name\", Size): {fmt_sites(bad)}", {"calls": sites}, location={"sheet": bad[0]["sheet"], "cell": bad[0]["cell"]})
    return finding("PASS", f"One MM_ITERATIONS call ({sites[0]['sheet']}!{sites[0]['cell']}, name {unquote(sites[0]['args'][0])!r}).", {"calls": sites})


def cal_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    inputs: dict[str, list[dict[str, Any]]] = {}
    outputs: dict[str, list[dict[str, Any]]] = {}
    unnamed = []
    for g in all_grids(analysis):
        for f in grid_flags(g, "iterationinput") + grid_flags(g, "iterationoutput"):
            if not f["args"]:
                unnamed.append({"sheet": g["sheet"], "cell": g["anchor"], "flag": f["raw"]})
                continue
            (inputs if f["name"] == "iterationinput" else outputs).setdefault(f["args"][0], []).append(g)
    if not inputs and not outputs and not unnamed:
        return finding("PASS", "No /iterationinput or /iterationoutput grids.", {"inputs": [], "outputs": []})
    problems = []
    for name in sorted(set(inputs) - set(outputs)):
        problems.append(f"/iterationinput.{name} has no /iterationoutput.{name}")
    for name in sorted(set(outputs) - set(inputs)):
        problems.append(f"/iterationoutput.{name} has no /iterationinput.{name}")
    for name, gs in inputs.items():
        if len(gs) > 1:
            problems.append(f"/iterationinput.{name} used on {len(gs)} grids")
    for name, gs in outputs.items():
        if len(gs) > 1:
            problems.append(f"/iterationoutput.{name} used on {len(gs)} grids")
    if unnamed:
        problems.append(f"flags without a name: {[u['flag'] for u in unnamed]}")
    observed = {"inputs": sorted(inputs), "outputs": sorted(outputs), "problems": problems}
    first = next(iter(inputs.values()), next(iter(outputs.values()), None))
    if problems:
        return finding("ERROR", "; ".join(problems), observed, location=grid_location(first[0]) if first else None)
    if not mm_calls(analysis, "MM_ITERATIONS"):
        return finding("WARNING", f"{len(inputs)} iteration input/output pair(s) but no MM_ITERATIONS call defines iterations on this workbook.", observed, location=grid_location(first[0]) if first else None)
    return finding("PASS", f"{len(inputs)} iteration input/output pair(s), all matched: {sorted(inputs)}.", observed)
