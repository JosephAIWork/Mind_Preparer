"""PAR-*  validators: the  /Parameters grid (KB 'Setup parameters editor':
columns Label | Type | PossibleValues | Values...; nine types; PossibleValues
syntax per type)."""
from __future__ import annotations

import re
from typing import Any

from ._common import finding, flagged_grids, fmt_grids, grid_location, grid_rows, header_names, is_blank

ALLOWED_TYPES = ["number", "text", "switch", "checkbox", "radio", "slider", "dropdown", "percentage", "date"]
RANGE_TYPES = {"number", "slider", "percentage"}
LIST_TYPES = {"checkbox", "radio", "dropdown"}
FREE_TYPES = {"text", "switch", "date"}
NUMBER_RE = re.compile(r"^[+-]?\d+(\.\d+)?$")


def _summary(g: dict[str, Any]) -> dict[str, Any]:
    return {"sheet": g["sheet"], "ref": g["ref"], "name": g["display_name"], "rows": g["n_rows"] - 1, "value_columns": max(g["n_cols"] - 3, 0)}


def par_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = flagged_grids(analysis, "parameters")
    return finding("PASS", f"{len(grids)} /Parameters grid(s).", {"parameters_grids": [_summary(g) for g in grids]})


def par_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = flagged_grids(analysis, "parameters")
    if not grids:
        return finding("PASS", "No /Parameters grids.", {"problems": []})
    problems = []
    for g in grids:
        headers = [h.lower() for h in header_names(g)]
        issues = []
        if len(headers) < 4:
            issues.append(f"{len(headers)} column(s); at least Label, Type, PossibleValues and one value column are needed")
        for i, expected in enumerate(("label", "type", "possiblevalues")):
            if i >= len(headers) or headers[i] != expected:
                issues.append(f"column {i + 1} header is '{header_names(g)[i] if i < len(headers) else ''}', expected '{expected}'")
        if issues:
            problems.append({"sheet": g["sheet"], "cell": g["anchor"], "grid": g["display_name"], "issues": issues})
    if problems:
        return finding("ERROR", f"{len(problems)} /Parameters grid(s) with a wrong structure: " + "; ".join(f"{p['grid']}: {', '.join(p['issues'])}" for p in problems[:4]), {"problems": problems}, location={"sheet": problems[0]["sheet"], "cell": problems[0]["cell"]})
    return finding("PASS", f"All {len(grids)} /Parameters grid(s) have Label | Type | PossibleValues | Values... columns.", {"problems": []})


def _typed_rows(analysis: dict[str, Any], g: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for r, row in enumerate(grid_rows(analysis, g)[1:], start=g["first_row"] + 1):
        if all(is_blank(v) for v in row[:3]):
            continue
        out.append({"row": r, "label": row[0] if row else None, "type": (str(row[1]).strip().lower() if len(row) > 1 and not is_blank(row[1]) else ""), "possible": row[2] if len(row) > 2 else None, "values": row[3:]})
    return out


def par_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = flagged_grids(analysis, "parameters")
    if not grids:
        return finding("PASS", "No /Parameters grids.", {"problems": []})
    problems = []
    checked = 0
    for g in grids:
        for row in _typed_rows(analysis, g):
            checked += 1
            if row["type"] not in ALLOWED_TYPES:
                problems.append({"sheet": g["sheet"], "row": row["row"], "label": row["label"], "type": row["type"]})
    if problems:
        return finding("ERROR", f"{len(problems)} parameter row(s) with a Type outside {ALLOWED_TYPES}: {problems[:6]}", {"problems": problems, "checked": checked}, location={"sheet": problems[0]["sheet"], "cell": f"B{problems[0]['row']}"})
    return finding("PASS", f"All {checked} parameter row(s) use an allowed Type.", {"problems": [], "checked": checked}, expected=ALLOWED_TYPES)


def _check_possible(row: dict[str, Any]) -> str | None:
    t, p = row["type"], row["possible"]
    text = "" if is_blank(p) else str(p).strip()
    if t in RANGE_TYPES:
        parts = text.split("|")
        if len(parts) not in (2, 3) or not all(NUMBER_RE.match(x.strip()) for x in parts):
            return f"'{text}' is not 'min|max' or 'min|max|step'"
        lo, hi = float(parts[0]), float(parts[1])
        if lo >= hi:
            return f"min {lo} is not below max {hi}"
        if len(parts) == 3 and float(parts[2]) <= 0:
            return "step must be positive"
    elif t in LIST_TYPES:
        items = [x.strip() for x in text.split("|")] if text else []
        if not items or any(not x for x in items):
            return f"'{text}' is not a non-empty pipe-separated list"
        if t == "checkbox":
            for v in row["values"]:
                if is_blank(v) or (isinstance(v, str) and v.startswith("=")):
                    continue
                chosen = [x.strip() for x in str(v).split("|")]
                bad = [x for x in chosen if x and x not in items]
                if bad:
                    return f"checkbox value(s) {bad} are not among the possible values"
    elif t == "switch":
        for v in row["values"]:
            if is_blank(v) or (isinstance(v, str) and v.startswith("=")):
                continue
            if str(v).strip().lower() not in ("true", "false"):
                return f"switch value '{v}' must be true or false"
    return None


def par_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = flagged_grids(analysis, "parameters")
    if not grids:
        return finding("PASS", "No /Parameters grids.", {"problems": []})
    problems = []
    checked = 0
    for g in grids:
        for row in _typed_rows(analysis, g):
            if row["type"] not in ALLOWED_TYPES:
                continue
            checked += 1
            issue = _check_possible(row)
            if issue:
                problems.append({"sheet": g["sheet"], "row": row["row"], "label": row["label"], "type": row["type"], "issue": issue})
    if problems:
        return finding("ERROR", f"{len(problems)} PossibleValues/Values problem(s): " + "; ".join(f"{p['label']} ({p['type']}): {p['issue']}" for p in problems[:6]), {"problems": problems, "checked": checked}, location={"sheet": problems[0]["sheet"], "cell": f"C{problems[0]['row']}"})
    return finding("PASS", f"PossibleValues syntax valid for all {checked} typed parameter row(s).", {"problems": [], "checked": checked})
