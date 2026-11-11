"""Shapes  of a  workbook's formulas, for the instance finder (1.12.5).

A formula cell "has the same shape" as another when their R1C1 texts are equal: cells filled down
or across, or copied to another sheet, share  one text. The finder reads repetition on these shapes.

  * `a1_to_r1c1(formula, row, col)`:  the A1 text of a cell  turned into R1C1 (the inverse of
    `autofix.r1c1_to_a1`); string  literals and structured references are copied verbatim, the
    `_xlfn.` / `_xll.` storage  prefixes are dropped so copies compare equal.
  * `loose_keys(r1c1)`:  three looser keys of one text -- references masked (sheets kept), references
    and sheets masked, typed-in numbers masked too. They sort how two copies differ.
  * `ShapeBook` / `build_book(analysis, budget)`: one shape id per distinct R1C1 text, and per sheet
    the id of every formula cell (by cell, by column, by row). Constants are not stored.
  * `Budget`: the time budget the finder checks between its steps;  a skipped step is written down.

Pure Python: no Excel, no numpy. Nothing here names a client workbook.
"""
from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from typing import Any

from .autofix import _REF_RE as R1C1_REF_RE, _SHEET_PREFIX, _STRING_RE, _TABLE_RE, _WORD, Parsed, parse_r1c1
from .formula_utils import STORAGE_PREFIX_RE, col_to_num, num_to_col, parse_ref
from .inventory import open_cached

# ---  A1 ->  R1C1 ------------------------------------------------------------------------------
# The A1 reference grammar with the FULL sheet-prefix grammar of autofix's R1C1 parser (quoted,
# plain, external `[1]Sheet!`, 3-D `S1:S3!`). The lookbehind holds `[` and `]` so a column name that
# looks like  a cell inside a structured reference (`Table1[Q1]`) is never converted, and  `$` so the
# engine cannot restart inside  `$B$3`. A leading `@` (implicit intersection) is  dropped.
_W = _WORD
_PFX = _SHEET_PREFIX
_CELL = r"\$?[A-Za-z]{1,3}\$?\d{1,7}"
A1_REF_RE = re.compile(
    rf"(?<![{_W}\[\]\\$])(?P<at>@)?(?P<sheet>{_PFX})?"
    rf"(?P<ref>{_CELL}(?::{_CELL})?|\$?[A-Za-z]{{1,3}}:\$?[A-Za-z]{{1,3}}|\$?\d{{1,7}}:\$?\d{{1,7}})"
    rf"(?![{_W}(\[!'\\])"
)
_CELL_PARTS_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})")
_COL_PART_RE = re.compile(r"(\$?)([A-Za-z]{1,3})")
_ROW_PART_RE = re.compile(r"(\$?)(\d{1,7})")
# what  the normaliser  copies verbatim: a string literal, or a structured reference `Name[` ... `]`
_SKIP_RE = re.compile(_STRING_RE.pattern + "|" + _TABLE_RE.pattern)
_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")


def _axis(kind: str, absolute: str, n: int, cur: int) -> str:
    if absolute:
        return f"{kind}{n}"
    return kind if n == cur else f"{kind}[{n - cur}]"


MAX_ROW = 1_048_576  # Excel's  last row
MAX_COL = 16_384  # Excel's  last column  (XFD)


class _PastExcel(Exception):
    """A  text that  looks like a cell but lies past Excel's last row or column (ABC1234567, XFE1): Excel
    allows such a text as a defined name, so the normaliser copies it as it is."""


def _col_axis(absolute: str, letters: str, cur: int) -> str:
    n = col_to_num(letters)
    if n > MAX_COL:
        raise _PastExcel
    return _axis("C", absolute, n, cur)


def _row_axis(absolute: str, digits: str, cur: int) -> str:
    n = int(digits)
    if not 1 <= n <= MAX_ROW:
        raise _PastExcel
    return _axis("R", absolute, n, cur)


def _cell_endpoint(text: str, row: int, col: int) -> str:
    cabs, letters, rabs, digits = _CELL_PARTS_RE.fullmatch(text).groups()
    return _row_axis(rabs, digits, row) + _col_axis(cabs, letters, col)


def _convert_ref(m: re.Match, row: int, col: int) -> str:
    try:
        return _convert_in_bounds(m, row, col)
    except _PastExcel:
        return m.group(0)


def _convert_in_bounds(m: re.Match, row: int, col: int) -> str:
    sheet = m.group("sheet") or ""
    ref = m.group("ref")
    if ":" not in ref:
        return sheet + _cell_endpoint(ref, row, col)
    a, b = ref.split(":", 1)
    if _CELL_PARTS_RE.fullmatch(a):
        return f"{sheet}{_cell_endpoint(a, row, col)}:{_cell_endpoint(b, row, col)}"
    mc = _COL_PART_RE.fullmatch(a)
    if mc:  # whole  columns A:A  / $F:$F / A:C
        mb = _COL_PART_RE.fullmatch(b)
        ea = _col_axis(mc.group(1), mc.group(2), col)
        eb = _col_axis(mb.group(1), mb.group(2), col)
    else:  # whole rows 5:5 / $5:$5 / 1:3
        mr, mb = _ROW_PART_RE.fullmatch(a), _ROW_PART_RE.fullmatch(b)
        ea = _row_axis(mr.group(1), mr.group(2), row)
        eb = _row_axis(mb.group(1), mb.group(2), row)
    return sheet + (ea if ea == eb else f"{ea}:{eb}")


def a1_to_r1c1(formula: str, row: int, col: int) -> str:
    """The  R1C1 text  of an A1 formula sitting at (row, col): the inverse of `autofix.r1c1_to_a1`.
    Storage prefixes (`_xlfn.`, `_xll.`) are dropped everywhere; string literals, structured
    references, defined names (also one that looks like a cell past  Excel's last row or column),
    numbers and function names stay as they  are."""
    text = STORAGE_PREFIX_RE.sub("", formula) if "_xl" in formula else formula

    def conv(m: re.Match) -> str:
        return _convert_ref(m, row, col)

    if '"' not in text and "[" not in text:
        return A1_REF_RE.sub(conv, text)
    out: list[str] = []
    pos = 0
    n = len(text)
    while True:
        m = _SKIP_RE.search(text, pos)
        if not m:
            break
        out.append(A1_REF_RE.sub(conv, text[pos:m.start()]))
        if text[m.start()] == '"':
            out.append(m.group(0))
            pos = m.end()
            continue
        depth = 0  # a structured reference: copy up to  the matching `]` (they nest: [[#This Row],[Q1]])
        i = m.end() - 1
        while i < n:
            if text[i] == "[":
                depth += 1
            elif text[i] == "]":
                depth -= 1
                if depth == 0:
                    break
            i += 1
        end = i + 1 if i < n else n
        out.append(text[m.start():end])
        pos = end
    out.append(A1_REF_RE.sub(conv, text[pos:]))
    return "".join(out)


# ---  loose keys --------------------------------------------------------------------------------
_NUM_RE = re.compile(r"(?<![A-Za-z_\]\[\d.])\d+(?:\.\d+)?(?:E[+-]?\d+)?(?![\d\]])", re.IGNORECASE)


def _outside_strings(text: str, fn) -> str:
    out: list[str] = []
    pos = 0
    for m in _STRING_RE.finditer(text):
        out.append(fn(text[pos:m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(fn(text[pos:]))
    return "".join(out)


def loose_keys(r1c1: str) -> tuple[str, str, str]:
    """(references  masked but  sheets kept, references and sheets masked, + typed-in numbers masked).
    String literals are left alone."""
    k1 = _outside_strings(r1c1, lambda s: R1C1_REF_RE.sub(lambda m: (m.group("sheet") or "") + "@", s))
    k2 = _outside_strings(r1c1, lambda s: R1C1_REF_RE.sub("@", s))
    k3 = _outside_strings(k2, lambda s: _NUM_RE.sub("#", s))
    return k1, k2, k3


def number_literals(r1c1: str) -> list[str]:
    """The  typed-in numbers  of a formula, in order, outside its string literals and references."""
    masked = _outside_strings(r1c1, lambda s: R1C1_REF_RE.sub("@", s))
    out: list[str] = []
    pos = 0
    for m in _STRING_RE.finditer(masked):
        out.extend(_NUM_RE.findall(masked[pos:m.start()]))
        pos = m.end()
    out.extend(_NUM_RE.findall(masked[pos:]))
    return out


# --- the time budget ---------------------------------------------------------------------------
class Budget:
    """Seconds  the finder  may spend; checked between steps. A step that is skipped says so in
    `skipped` (plain words), nothing is skipped silently."""

    def __init__(self, seconds: float = 120.0) -> None:
        self.seconds = float(seconds)
        self.started = time.monotonic()
        self.skipped: list[str] = []

    def restart(self) -> None:
        """Start the clock now (after the workbook is open: a reload is reported,  not charged)."""
        self.started = time.monotonic()

    def elapsed(self) -> float:
        return time.monotonic() - self.started

    def left(self) -> float:
        return self.seconds - self.elapsed()

    def out(self) -> bool:
        return self.left() <= 0

    def skip(self, text: str) -> None:
        if text and text not in self.skipped:
            self.skipped.append(text)


# --- the shape book ----------------------------------------------------------------------------
EMPTY_ID, NUMBER_ID, TEXT_ID = 0, 1, 2  # ids  below FIRST_FORMULA_ID  are reserved for constants
FIRST_FORMULA_ID = 10


@dataclass
class SheetShape:
    """The  formula cells  of one sheet as shape ids. Bounds are those of the FORMULA cells, not of
    the used range. `indices` are positions in wb0["formulas"], in the order they were read."""

    name: str
    r_lo: int = 0
    r_hi: int = 0
    c_lo: int = 0
    c_hi: int = 0
    cells: dict[tuple[int, int], int] = field(default_factory=dict)  # (row, col) -> id
    cols: dict[int, dict[int, int]] = field(default_factory=dict)  # col  -> {row: id}
    rows: dict[int, dict[int, int]] = field(default_factory=dict)  # row -> {col: id}
    n: int = 0
    indices: list[int] = field(default_factory=list)

    def put(self, row: int, col: int, sid: int) -> None:
        self.cells[(row, col)] = sid
        col_d = self.cols.get(col)
        if col_d is None:
            col_d = self.cols[col] = {}
        col_d[row] = sid
        row_d = self.rows.get(row)
        if row_d is None:
            row_d = self.rows[row] = {}
        row_d[col] = sid
        self.n += 1
        if self.n == 1:
            self.r_lo = self.r_hi = row
            self.c_lo = self.c_hi = col
        else:
            if row < self.r_lo:
                self.r_lo = row
            elif row > self.r_hi:
                self.r_hi = row
            if col < self.c_lo:
                self.c_lo = col
            elif col > self.c_hi:
                self.c_hi = col


class ShapeBook:
    """Every  distinct R1C1  text of the workbook once (`texts[id]`, ids from 10), the formula cells
    of every sheet as ids, and the three loose-key tables. The A1 text of a cell is re-read from the
    analysis  (`a1_text`) when a detector needs it; the workbook itself is opened  only on demand
    (`workbook()`, through the app's  cache) for constants."""

    def __init__(self, analysis: dict[str, Any]) -> None:
        self.analysis = analysis
        wb0 = analysis["workbooks"][0]
        self.formulas: list[dict[str, Any]] = wb0.get("formulas") or []
        self.sheet_order: list[str] = [s["name"] for s in wb0.get("sheets") or [] if s.get("name")]
        self.texts: list[str | None] = [None] * FIRST_FORMULA_ID
        self.id_of: dict[str, int] = {}
        self.loose: list[list[int]] = [[], [], []]
        self.sheets: dict[str, SheetShape] = {}
        self.n = 0
        self._index: dict[str, dict[tuple[int, int], int]] = {}
        self._wb: Any = None
        self._wb_tried = False

    # -- texts
    def parsed(self, sid: int) -> Parsed:
        return parse_r1c1(self.texts[sid] or "")

    def loose_id(self, sid: int, level: int = 1) -> int:
        return self.loose[level][sid]

    def sheets_read(self, sid: int, own: str) -> set[str]:
        """The other sheets a  shape reads (3-D spans expanded along the workbook order)."""
        out: set[str] = set()
        for spec, _a, _b in self.parsed(sid).refs:
            if spec is None:
                continue
            if spec[0] == "sheet":
                if spec[1] != own:
                    out.add(spec[1])
            elif spec[0] == "3d" and spec[1] in self.sheet_order and spec[2] in self.sheet_order:
                i, j = sorted((self.sheet_order.index(spec[1]), self.sheet_order.index(spec[2])))
                out.update(self.sheet_order[i:j + 1])
        return out

    # -- A1 texts,  by index into the analysis
    def index_of(self, sheet: str, row: int, col: int) -> int | None:
        m = self._index.get(sheet)
        if m is None:
            m = {}
            sh = self.sheets.get(sheet)
            for i in (sh.indices if sh else ()):
                f = self.formulas[i]
                rc = coord_of(f.get("cell") or "")
                if rc:
                    m[rc] = i
                ref = f.get("array_ref")
                if ref and ref != f.get("cell"):
                    for pos in range_cells(ref):
                        m.setdefault(pos, i)
            self._index[sheet] = m
        return m.get((row, col))

    def a1_text(self, sheet: str, row: int, col: int) -> str | None:
        i = self.index_of(sheet, row, col)
        return None if i is None else self.formulas[i].get("formula")

    # -- the workbook,  for constants (opened through the app's cache; None when it cannot be)
    def workbook(self) -> Any:
        if not self._wb_tried:
            self._wb_tried = True
            try:
                self._wb = open_cached(self.analysis)
            except Exception:
                self._wb = None
        return self._wb

    def worksheet(self, name: str) -> Any:
        wb = self.workbook()
        try:
            return wb[name] if wb is not None and name in wb.sheetnames else None
        except Exception:
            return None


def coord_of(cell: str) -> tuple[int, int] | None:
    """'B2'  -> (2,  2); None for anything else."""
    m = _COORD_RE.match(cell)
    if not m:
        return None
    return int(m.group(2)), col_to_num(m.group(1))


def range_cells(ref: str) -> list[tuple[int, int]]:
    """(row,  col) of  every cell of a literal range (no whole rows/columns), row-major."""
    r = parse_ref(ref)
    if not r or r["whole_column"] or r["whole_row"]:
        return []
    return [(row, col) for row in range(r["r1"], r["r2"] + 1) for col in range(r["c1"], r["c2"] + 1)]


def cell_name(sheet: str, row: int, col: int) -> str:
    """'Sheet'!A1  as a  formula writes it."""
    q = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else "'" + sheet.replace("'", "''") + "'"
    return f"{q}!{num_to_col(col)}{row}"


def build_book(analysis: dict[str, Any], budget: Budget | None = None) -> ShapeBook:
    """One  pass over  wb0["formulas"]: every formula cell gets the id of its R1C1 text (the same text
    on any sheet = the same id); an array formula gives its anchor's id to every member cell.  The
    budget is accepted for the finder's bookkeeping; the pass itself always  completes (a partial
    book would mislead every  detector)."""
    book = ShapeBook(analysis)
    texts, id_of, sheets = book.texts, book.id_of, book.sheets
    for i, f in enumerate(book.formulas):
        cell, sheet, text = f.get("cell"), f.get("sheet"), f.get("formula")
        if not cell or sheet is None or not text:
            continue
        rc = coord_of(cell)
        if rc is None:
            continue
        row, col = rc
        r1c1 = a1_to_r1c1(text, row, col)
        sid = id_of.get(r1c1)
        if sid is None:
            sid = len(texts)
            id_of[r1c1] = sid
            texts.append(r1c1)
        sh = sheets.get(sheet)
        if sh is None:
            sh = sheets[sheet] = SheetShape(sheet)
        sh.put(row, col, sid)
        sh.indices.append(i)
        ref = f.get("array_ref")
        if ref and ref != cell:
            for pos in range_cells(ref):
                if pos != rc and pos not in sh.cells:
                    sh.put(pos[0], pos[1], sid)
    book.n = sum(sh.n for sh in sheets.values())
    ids: list[dict[str, int]] = [{}, {}, {}]
    for lvl in range(3):
        book.loose[lvl] = list(range(FIRST_FORMULA_ID)) + [0] * (len(texts) - FIRST_FORMULA_ID)
    for sid in range(FIRST_FORMULA_ID, len(texts)):
        for lvl, key in enumerate(loose_keys(texts[sid] or "")):
            book.loose[lvl][sid] = ids[lvl].setdefault(key, len(ids[lvl]) + FIRST_FORMULA_ID)
    return book
