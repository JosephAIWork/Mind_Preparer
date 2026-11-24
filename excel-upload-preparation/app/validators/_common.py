"""Small  helpers shared  by the validator modules."""
from __future__ import annotations

import re
from typing import Any

from ..grids import all_grids, grid_label
from ..inventory import cell_value

TRUE_WORDS = {"true", "1", "yes", "vrai", "oui"}
FALSE_WORDS = {"false", "0", "no", "faux", "non", ""}


def finding(status: str, message: str, observed: Any = None, location: dict | None = None, evidence: str | None = None, expected: Any = None, severity: str | None = None) -> dict:
    out: dict[str, Any] = {"status": status, "message": message}
    if observed is not None:
        out["observed"] = observed
    if location:
        out["location"] = location
    if evidence:
        out["evidence"] = evidence
    if expected is not None:
        out["expected"] = expected
    if severity:
        out["severity"] = severity
    return out


def site(sheet: str, cell: str) -> dict:
    return {"sheet": sheet, "cell": cell}


def fmt_sites(items: list[dict], limit: int = 8) -> str:
    """'Sheet!A1,  Sheet!B2, ...  (+N more)' from dicts with sheet/cell."""
    texts = [f"{i['sheet']}!{i['cell']}" for i in items[:limit]]
    extra = len(items) - limit
    return ", ".join(texts) + (f", ... (+{extra} more)" if extra > 0 else "")


def fmt_grids(grids: list[dict], limit: int = 8) -> str:
    texts = [grid_label(g) for g in grids[:limit]]
    extra = len(grids) - limit
    return ", ".join(texts) + (f", ... (+{extra} more)" if extra > 0 else "")


def grid_location(grid: dict) -> dict:
    return {"sheet": grid["sheet"], "cell": grid.get("title_cell") or grid["anchor"]}


def grid_rows(analysis: dict, grid: dict) -> list[list[Any]]:
    """Body  rows of  a grid (header row first): the rows stored at inventory
    time, or read from the cached workbook for grids that were too big to store."""
    rows = grid.get("rows")
    if rows is not None:
        return rows
    out = []
    for r in range(grid["first_row"], grid["last_row"] + 1):
        out.append([cell_value(analysis, grid["sheet"], r, c) for c in range(grid["first_col"], grid["last_col"] + 1)])
    return out


def header_index(grid: dict, header: str) -> int | None:
    """Column  index (0-based)  of a header, case-insensitive, whitespace-trimmed."""
    target = header.strip().lower()
    for i, h in enumerate(grid.get("header_values", [])):
        if isinstance(h, str) and h.strip().lower() == target:
            return i
    return None


def header_names(grid: dict) -> list[str]:
    return [h.strip() if isinstance(h, str) else ("" if h is None else str(h)) for h in grid.get("header_values", [])]


def as_bool(value: Any) -> bool | None:
    """TRUE/FALSE/1/0  (Excel or  text) -> bool; None when not a boolean-ish value."""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        s = value.strip().lower()
        if s in TRUE_WORDS:
            return True
        if s in FALSE_WORDS:
            return False
    return None


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def flagged_grids(analysis: dict, *flag_names: str) -> list[dict]:
    wanted = {f.lower() for f in flag_names}
    return [g for g in all_grids(analysis) if wanted & set(g.get("flag_names", []))]


def grids_by_name(analysis: dict) -> dict[str, list[dict]]:
    out: dict[str, list[dict]] = {}
    for g in all_grids(analysis):
        if g.get("name"):
            out.setdefault(g["name"].strip().lower(), []).append(g)
    return out


DEFAULT_SHEET_NAME_RE = re.compile(r"^(sheet|feuil|hoja|tabelle|blad|foglio|planilha|arkusz|ark|лист|工作表)\s?\d+$", re.IGNORECASE)
DEFAULT_WORKBOOK_NAME_RE = re.compile(r"^(book|classeur|mappe|libro|cartel|pasta)\s?\d+$", re.IGNORECASE)
