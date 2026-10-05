"""Mind-style grid detection and grid-title / flag parsing.

Source: KB "General structure and guidelines" (basic-model-design): sheets
are read left to right, top to bottom; on a non-empty cell not already in a
grid, Mind looks right and down until it finds an empty cell and creates a
grid of that size; an alone text cell is ignored; a first row made only of
strings is the header row; the grid name is a cell starting with '#' above
the grid ('#GridName /Flag1 /Flag2'), otherwise the grid is 'untitled' plus
its first-cell coordinates. Flags are '/Name' tokens in that title cell (or,
for a few documented ones, in a header cell: /HideRows, /InstanceSelect,
&&HideColumn, &hide).

This is a faithful re-implementation of the documented algorithm, not a
guess -- but it is still this app's reading of the doc, so findings built on
it are labeled INFERENCE where they depend on grid *boundaries* and
DETERMINISTIC_FINDING where they only depend on a title cell's own text.

1.7.3: a merged cell occupies every cell of its range. Mind reads a header
merged over B4:E4 as a header row four columns wide (a conversion of a real
model failed with "the array formula ... exceeds the grid size. Ensure that
the column headers cover the entire array formula" at the first column past
the merge); reading the merge as one cell hid that.
"""
from __future__ import annotations

import datetime as _dt
import re
from typing import Any

from .formula_utils import num_to_col, ref_text

TITLE_PREFIX = "#"
# Excel error *values* are stored as text that also begins with '#'. They are
# results, not grid titles: reading '#N/A' in a data table as a title makes the
# app believe a title is trapped inside the grid (STR-001) and, worse, offer to
# insert a row through a real data table to "free" it.
ERROR_LITERALS = frozenset(
    {
        "#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!",
        "#SPILL!", "#CALC!", "#GETTING_DATA", "#FIELD!", "#BLOCKED!",
        "#CONNECT!", "#UNKNOWN!", "#BUSY!", "#EXTERNAL!",
    }
)
FLAG_RE = re.compile(r"/([A-Za-z][A-Za-z0-9_]*)((?:\.(?:\([^)]*\)|[^\s./]+))*)")


class MergedInto:
    """A cell of a merged range other than its first cell: it takes its place
    in a grid like any occupied cell and shows the first cell's value."""

    __slots__ = ("anchor", "value")

    def __init__(self, anchor: tuple[int, int], value: Any) -> None:
        self.anchor, self.value = anchor, value

    def __repr__(self) -> str:
        return f"MergedInto({ref_text(self.anchor[1], self.anchor[0])})"


def looks_like_title(value: Any) -> bool:
    """A '#...' text cell that is a grid title -- not an Excel error value."""
    if not isinstance(value, str):
        return False
    text = value.strip()
    return text.startswith(TITLE_PREFIX) and text.upper() not in ERROR_LITERALS
MARKER_HIDE_COLUMN = "&&hidecolumn"
MARKER_HIDE_HEADER = "&hide"
STEP_LABELS_MARKER = "steplabels"
MAX_STORED_CELLS = 20000


def json_safe(value: Any) -> Any:
    """openpyxl cell values -> something json.dumps can serialize (formulas as text)."""
    if isinstance(value, MergedInto):
        value = value.value
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return value.isoformat()
    text = getattr(value, "text", None)  # ArrayFormula / DataTableFormula
    if isinstance(text, str):
        return text if text.startswith("=") else "=" + text
    return str(value)


def is_formula(value: Any) -> bool:
    if isinstance(value, str):
        return value.startswith("=")
    return getattr(value, "text", None) is not None


def is_occupied(value: Any) -> bool:
    return value is not None and not (isinstance(value, str) and value == "")


def parse_flags(text: str) -> list[dict[str, Any]]:
    """'/Group.(My group).0.1 /Input' -> [{'name': 'group', 'raw': ..., 'args': ['My group','0','1']}, {'name': 'input', ...}]"""
    flags = []
    for m in FLAG_RE.finditer(text or ""):
        args: list[str] = []
        rest = m.group(2)
        # split on '.' outside parentheses
        buf = ""
        depth = 0
        for ch in rest:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            if ch == "." and depth == 0:
                if buf:
                    args.append(buf)
                buf = ""
            else:
                buf += ch
        if buf:
            args.append(buf)
        args = [a[1:-1].strip() if a.startswith("(") and a.endswith(")") else a.strip() for a in args]
        flags.append({"name": m.group(1).lower(), "raw_name": m.group(1), "raw": m.group(0), "args": args})
    return flags


def parse_title(text: str) -> tuple[str, list[dict[str, Any]]]:
    """'#Grid name /Input /Resize.A' -> ('Grid name', [flags...])."""
    body = text[1:] if text.startswith(TITLE_PREFIX) else text
    flags = parse_flags(body)
    first_flag = FLAG_RE.search(body)
    name = body[: first_flag.start()] if first_flag else body
    return name.strip(), flags


def header_markers(header_values: list[Any]) -> list[dict[str, Any]]:
    """Documented header-cell flags/markers: /HideRows, /InstanceSelect, &&HideColumn, &hide."""
    found = []
    for idx, v in enumerate(header_values):
        if not isinstance(v, str):
            continue
        low = v.lower()
        for f in parse_flags(v):
            found.append({"column_index": idx, "name": f["name"], "raw": f["raw"], "header": v})
        if MARKER_HIDE_COLUMN in low:
            found.append({"column_index": idx, "name": "&&hidecolumn", "raw": "&&HideColumn", "header": v})
        elif MARKER_HIDE_HEADER in low:
            found.append({"column_index": idx, "name": "&hide", "raw": "&hide", "header": v})
    return found


def detect_grids(sheet_name: str, occupied: dict[tuple[int, int], Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Run the documented detection over an occupancy map {(row, col): value}.
    Returns (grids, standalone_text_cells)."""
    assigned: set[tuple[int, int]] = set()
    grids: list[dict[str, Any]] = []
    standalone: list[dict[str, Any]] = []

    for (r, c) in sorted(occupied):
        if (r, c) in assigned:
            continue
        value = occupied[(r, c)]
        if isinstance(value, MergedInto):
            # its first cell came earlier in reading order: a title cell, or a block this
            # cell was not reached from (merged downwards past the block) -- never a new grid
            assigned.add((r, c))
            continue
        title_cell: tuple[int, int] | None = None
        ar, ac = r, c
        if looks_like_title(value) and (r + 1, c) in occupied and (r + 1, c) not in assigned:
            title_cell = (r, c)
            ar = r + 1

        width = 0
        while (ar, ac + width) in occupied and (ar, ac + width) not in assigned:
            width += 1
        height = 0
        while (ar + height, ac) in occupied and (ar + height, ac) not in assigned:
            height += 1
        width = max(width, 1)
        height = max(height, 1)

        if title_cell is None and width == 1 and height == 1 and isinstance(value, str) and not value.startswith("="):
            assigned.add((r, c))
            standalone.append(
                {
                    "sheet": sheet_name,
                    "cell": ref_text(c, r),
                    "text": value,
                    "looks_like_title": looks_like_title(value),
                }
            )
            continue

        for rr in range(ar, ar + height):
            for cc in range(ac, ac + width):
                assigned.add((rr, cc))
        if title_cell:
            assigned.add(title_cell)

        header_values = [json_safe(occupied.get((ar, ac + i))) for i in range(width)]
        header_is_all_text = all(isinstance(v, str) and not v.startswith("=") for v in header_values)
        inner_titles = [
            ref_text(cc, rr)
            for rr in range(ar, ar + height)
            for cc in range(ac, ac + width)
            if looks_like_title(occupied.get((rr, cc))) and (rr, cc) != (ar, ac)
        ]
        formula_count = sum(
            1 for rr in range(ar, ar + height) for cc in range(ac, ac + width) if is_formula(occupied.get((rr, cc)))
        )
        empty_inside = sum(
            1 for rr in range(ar, ar + height) for cc in range(ac, ac + width) if not is_occupied(occupied.get((rr, cc)))
        )

        name: str | None = None
        flags: list[dict[str, Any]] = []
        title_text: str | None = None
        if title_cell:
            title_text = occupied[title_cell]
            name, flags = parse_title(title_text)
        title_marker_text = (title_text or "").replace(" ", "").lower()
        grid: dict[str, Any] = {
            "sheet": sheet_name,
            "title_cell": ref_text(title_cell[1], title_cell[0]) if title_cell else None,
            "title": title_text,
            "name": name,
            "display_name": name or f"untitled {ref_text(ac, ar)}",
            "flags": flags,
            "flag_names": sorted({f["name"] for f in flags}),
            "anchor": ref_text(ac, ar),
            "ref": ref_text(ac, ar, ac + width - 1, ar + height - 1),
            "first_row": ar,
            "last_row": ar + height - 1,
            "first_col": ac,
            "last_col": ac + width - 1,
            "n_rows": height,
            "n_cols": width,
            "header_values": header_values,
            "header_is_all_text": header_is_all_text,
            "header_has_formula": any(isinstance(v, str) and v.startswith("=") for v in header_values),
            "header_has_number": any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in header_values),
            "header_markers": header_markers(header_values),
            "inner_title_cells": inner_titles,
            "formula_count": formula_count,
            "empty_cells_inside": empty_inside,
            "looks_like_step_labels": STEP_LABELS_MARKER in title_marker_text
            or STEP_LABELS_MARKER in " ".join(str(v) for v in header_values if isinstance(v, str)).replace(" ", "").lower(),
        }
        if (flags or name) and width * height <= MAX_STORED_CELLS:
            grid["rows"] = [[json_safe(occupied.get((ar + i, ac + j))) for j in range(width)] for i in range(height)]
        grids.append(grid)
    return grids, standalone


def grid_containing(grids: list[dict[str, Any]], row: int, col: int) -> dict[str, Any] | None:
    for g in grids:
        if g["first_row"] <= row <= g["last_row"] and g["first_col"] <= col <= g["last_col"]:
            return g
    return None


def all_grids(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return [g for wb in analysis["workbooks"] for s in wb["sheets"] for g in s.get("grids", [])]


def grids_with_flag(analysis: dict[str, Any], flag_name: str) -> list[dict[str, Any]]:
    flag_name = flag_name.lower()
    return [g for g in all_grids(analysis) if flag_name in g.get("flag_names", [])]


def grid_flags(grid: dict[str, Any], flag_name: str) -> list[dict[str, Any]]:
    return [f for f in grid.get("flags", []) if f["name"] == flag_name.lower()]


def grid_label(grid: dict[str, Any]) -> str:
    return f"{grid['sheet']}!{grid['anchor']} ({grid['display_name']})"


def column_letter(col: int) -> str:
    return num_to_col(col)
