"""1.8.0 -- the automatic formula-error fixer ("Fix all automatically" on the
Recalculate step).

The manual way (Fix panel: one error or one group at a time, a proposal, an
Apply, a recalculation) stays. This module does the same job on its own, for
every error cell of a workbook, in ONE Excel session:

    open a copy -> full recalculation -> read every formula cell
    loop:
        find the ROOT error cells (an error whose own inputs are fine --
        everything downstream only repeats it)
        choose a fix for each root (the strategy ladder below, optionally
        ordered by the assistant)
        write the fixes -> recalculate -> read every formula cell again
        THE NUMBERS GATE: a formula cell that had a valid value in the
        workbook as it was opened must still have exactly that value.
        A fix that changed one is rolled back, and its next strategy is
        tried on the following pass.
    until no error is left, or nothing more can be fixed safely
    full rebuild (the confirmation) -> save

Two rules decided by the tool's owner (2026-10-05):
  * clean means ZERO error cells, including the ones the original already had;
  * a good number never changes: a fix that would change one is undone, another
    way is tried, and when none works the error stays and is listed for a person.

What a fix is. An error cell keeps its formula and gets a value to return when
the formula fails: =IFERROR(<the formula>, 0) -- or "" where the cells around
it hold text. A formula that can never compute again (its reference was
deleted: only NA() / #REF! is left of it) is replaced by that value. A constant
that is an error value is cleared. Nothing else is rewritten, so a cell that
works today computes exactly as before.

Formulas are handled in R1C1 notation: cells filled down or across share one
R1C1 text, so a block of thousands of cells is one group, one decision and one
write.
"""
from __future__ import annotations

import bisect
import re
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

from .change_apply import append_change_log
from .excel_com import close_quietly, com_available, excel_session, open_for_write, save_in_place, verify_opens_in_excel
from .formula_utils import num_to_col
from .inventory import make_immutable_copy, sha256_of
from .recalc import XL_ERROR_CODES

Progress = Callable[..., None]
Key = tuple[int, int, int]  # (sheet index, row, column)

CHUNK_ROWS = 2000
CHUNK_CELLS = 600_000
XL_MANUAL = -4135
MAX_ROW, MAX_COL = 1_048_576, 16_384
NUMERIC_TOL = 1e-9
TRANSIENT_COM_HRESULTS = {-2147418111, -2147417846}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
VOLATILE_FUNCS = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY"}
DYNAMIC_FUNCS = {"INDIRECT", "OFFSET", "CELL", "INFO", "GETPIVOTDATA", "EVALUATE"}
LOOKUP_FUNCS = {"VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "MATCH", "XMATCH", "INDEX"}
ERROR_CODE_OF = {v: k for k, v in XL_ERROR_CODES.items()}
MAX_REPORTED = 400  # groups / left-over cells listed in the result (the counts are always complete)

ERROR_MEANING = {
    "#DIV/0!": "a division by zero",
    "#N/A": "a value that is not available (a lookup that finds nothing)",
    "#VALUE!": "text where a number is expected",
    "#REF!": "a reference that no longer exists",
    "#NAME?": "a name or function Excel does not know",
    "#NUM!": "a number out of range",
    "#NULL!": "two ranges that do not intersect",
}


# --- R1C1 formulas: references, names, functions -------------------------------------------
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_ERROR_LITERAL_RE = re.compile(r"#(?:N/A|REF!|DIV/0!|NAME\?|VALUE!|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)", re.IGNORECASE)
_WORD = r"A-Za-z0-9_.À-￿"
_EP = r"R(?:\[-?\d+\]|\d*)C(?:\[-?\d+\]|\d*)|R(?:\[-?\d+\]|\d*)|C(?:\[-?\d+\]|\d*)"
_SHEET_PREFIX = rf"(?:'(?:[^']|'')+'|\[[^\]]*\][^!'+\-*/^&=<>(),;:]*|[{_WORD}]+(?::[{_WORD}]+)?)!"
_REF_RE = re.compile(rf"(?<![{_WORD}\]\\])(?P<sheet>{_SHEET_PREFIX})?(?P<a>{_EP})(?::(?P<b>{_EP}))?(?![{_WORD}(\[!'\\])")
_FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:_xl[a-z]+\.)?[A-Za-z_][A-Za-z0-9_.]*)\s*\(")
_NAME_RE = re.compile(rf"(?<![{_WORD}\\!'\]])([A-Za-z_\\À-￿][{_WORD}\\?]*)(?![{_WORD}\\?(\[!])")
_TABLE_RE = re.compile(rf"[A-Za-z_À-￿][{_WORD}]*\[")
_REF_LITERAL_RE = re.compile(rf"(?<![:{_WORD}'])(?:{_SHEET_PREFIX})?#REF!(?![:\w])")
_PURE_NA_RE = re.compile(r"^=\s*\+?\s*NA\(\s*\)\s*$", re.IGNORECASE)
_STORAGE_PREFIX_RE = re.compile(r"^_xl[a-z]+\.", re.IGNORECASE)


@dataclass(frozen=True)
class Parsed:
    """What an R1C1 formula text reads, independent of the cell it sits in."""
    refs: tuple[tuple[Any, Any, Any], ...]  # (sheet spec, endpoint a, endpoint b | None)
    names: tuple[str, ...]                  # candidate defined names, upper-cased
    funcs: frozenset[str]
    dynamic: bool                           # INDIRECT / OFFSET / a table reference: reads cells the text does not show
    external: bool                          # reads another workbook
    broken: bool                            # contains a literal #REF!


def _endpoint(text: str) -> tuple[Any, Any]:
    """'R[-1]C3' -> ((False, -1), (True, 3)); 'C[2]' -> (None, (False, 2)).
    An axis is None (the whole row/column), (True, n) absolute, (False, n) relative."""
    def axis(s: str) -> tuple[bool, int]:
        if not s:
            return (False, 0)
        if s[0] == "[":
            return (False, int(s[1:-1]))
        return (True, int(s))

    row = col = None
    rest = text
    if rest[:1] == "R":
        i = rest.find("C", 1)
        row = axis(rest[1:] if i < 0 else rest[1:i])
        rest = "" if i < 0 else rest[i:]
    if rest[:1] == "C":
        col = axis(rest[1:])
    return row, col


def _sheet_spec(prefix: str | None) -> Any:
    if not prefix:
        return None
    s = prefix[:-1]
    if s.startswith("'") and s.endswith("'"):
        s = s[1:-1].replace("''", "'")
    if "[" in s:
        return ("ext",)
    if ":" in s:
        a, b = s.split(":", 1)
        return ("3d", a, b)
    return ("sheet", s)


_PARSE_CACHE: dict[str, Parsed] = {}


def parse_r1c1(formula: str) -> Parsed:
    hit = _PARSE_CACHE.get(formula)
    if hit is not None:
        return hit
    text = _STRING_RE.sub('""', formula)
    broken = "#REF!" in text.upper()
    text = _ERROR_LITERAL_RE.sub(" ", text)
    refs: list[tuple[Any, Any, Any]] = []
    external = False

    def take(m: re.Match) -> str:
        nonlocal external
        spec = _sheet_spec(m.group("sheet"))
        if spec == ("ext",):
            external = True
        refs.append((spec, _endpoint(m.group("a")), _endpoint(m.group("b")) if m.group("b") else None))
        return " "

    rest = _REF_RE.sub(take, text)
    if "[" in rest and re.search(r"\[[^\]]*\]", rest) and not _TABLE_RE.search(rest):
        external = True  # an external workbook prefix that no reference followed (e.g. a name in another book)
    funcs = frozenset(_STORAGE_PREFIX_RE.sub("", f).upper() for f in _FUNC_RE.findall(rest))
    names = tuple(sorted({n.upper() for n in _NAME_RE.findall(rest)} - {"TRUE", "FALSE"}))
    dynamic = bool(funcs & DYNAMIC_FUNCS) or bool(_TABLE_RE.search(rest))
    parsed = Parsed(tuple(refs), names, funcs, dynamic, external, broken)
    if len(_PARSE_CACHE) < 200_000:
        _PARSE_CACHE[formula] = parsed
    return parsed


def _span(a: Any, b: Any, cur: int, limit: int) -> tuple[int, int]:
    """Rows (or columns) an axis pair covers for a formula sitting at `cur`."""
    lo = a[1] if a[0] else cur + a[1]
    hi = lo
    if b is not None:
        hi = b[1] if b[0] else cur + b[1]
    lo, hi = min(lo, hi), max(lo, hi)
    return max(1, lo), min(limit, hi)


def ref_rect(ref: tuple[Any, Any, Any], row: int, col: int) -> tuple[int, int, int, int]:
    """(r1, r2, c1, c2) a parsed reference covers for a formula at (row, col)."""
    _, a, b = ref
    ar, ac = a
    if b is None:
        r1, r2 = (1, MAX_ROW) if ar is None else _span(ar, None, row, MAX_ROW)
        c1, c2 = (1, MAX_COL) if ac is None else _span(ac, None, col, MAX_COL)
    else:
        br, bc = b
        r1, r2 = (1, MAX_ROW) if ar is None or br is None else _span(ar, br, row, MAX_ROW)
        c1, c2 = (1, MAX_COL) if ac is None or bc is None else _span(ac, bc, col, MAX_COL)
    return r1, r2, c1, c2


def _a1_endpoint(ep: tuple[Any, Any], row: int, col: int) -> tuple[str, str]:
    r, c = ep
    rs = "" if r is None else (f"${r[1]}" if r[0] else str(max(1, min(MAX_ROW, row + r[1]))))
    cs = "" if c is None else (f"${num_to_col(c[1])}" if c[0] else num_to_col(max(1, min(MAX_COL, col + c[1]))))
    return cs, rs


def r1c1_to_a1(formula: str, row: int, col: int) -> str:
    """The A1 text of an R1C1 formula sitting at (row, col) -- for reports and the
    change log (people read A1). Strings are left alone."""
    out: list[str] = []
    pos = 0
    for s in _STRING_RE.finditer(formula):
        out.append(_convert_refs(formula[pos:s.start()], row, col))
        out.append(s.group(0))
        pos = s.end()
    out.append(_convert_refs(formula[pos:], row, col))
    return "".join(out)


def _convert_refs(text: str, row: int, col: int) -> str:
    def conv(m: re.Match) -> str:
        a = _endpoint(m.group("a"))
        b = _endpoint(m.group("b")) if m.group("b") else None
        ac, ar = _a1_endpoint(a, row, col)
        if b is None:
            if a[0] is None or a[1] is None:  # a whole column / row needs both ends in A1
                one = ac or ar
                return f"{m.group('sheet') or ''}{one}:{one}"
            return f"{m.group('sheet') or ''}{ac}{ar}"
        bc, br = _a1_endpoint(b, row, col)
        return f"{m.group('sheet') or ''}{ac}{ar}:{bc}{br}"

    return _REF_RE.sub(conv, text)


def a1(row: int, col: int) -> str:
    return f"{num_to_col(col)}{row}"


def rect_a1(r1: int, c1: int, r2: int, c2: int) -> str:
    return a1(r1, c1) if (r1, c1) == (r2, c2) else f"{a1(r1, c1)}:{a1(r2, c2)}"


# --- cells by sheet, with range queries -------------------------------------------------------
class CellIndex:
    """A set of (sheet, row, col) that answers "is any of them in this rectangle?"
    in time proportional to the smaller side of the rectangle."""

    def __init__(self, cells: Iterable[Key] = ()):
        self.by_sheet: dict[int, set[tuple[int, int]]] = defaultdict(set)
        for si, r, c in cells:
            self.by_sheet[si].add((r, c))
        self._rows: dict[int, dict[int, list[int]]] = {}
        self._cols: dict[int, dict[int, list[int]]] = {}
        self._row_keys: dict[int, list[int]] = {}
        self._col_keys: dict[int, list[int]] = {}

    def __len__(self) -> int:
        return sum(len(v) for v in self.by_sheet.values())

    def _build(self, si: int) -> None:
        rows: dict[int, list[int]] = defaultdict(list)
        cols: dict[int, list[int]] = defaultdict(list)
        for r, c in self.by_sheet.get(si, ()):
            rows[r].append(c)
            cols[c].append(r)
        for v in rows.values():
            v.sort()
        for v in cols.values():
            v.sort()
        self._rows[si], self._cols[si] = rows, cols
        self._row_keys[si], self._col_keys[si] = sorted(rows), sorted(cols)

    def cells_in(self, si: int, r1: int, r2: int, c1: int, c2: int, limit: int = 1 << 30, exclude: tuple[int, int] | None = None) -> list[tuple[int, int]]:
        cells = self.by_sheet.get(si)
        if not cells:
            return []
        if r1 == r2 and c1 == c2:
            return [(r1, c1)] if (r1, c1) in cells and (r1, c1) != exclude else []
        if si not in self._rows:
            self._build(si)
        out: list[tuple[int, int]] = []
        row_keys, col_keys = self._row_keys[si], self._col_keys[si]
        i1, i2 = bisect.bisect_left(row_keys, r1), bisect.bisect_right(row_keys, r2)
        j1, j2 = bisect.bisect_left(col_keys, c1), bisect.bisect_right(col_keys, c2)
        if i2 - i1 <= j2 - j1:
            rows = self._rows[si]
            for r in row_keys[i1:i2]:
                cs = rows[r]
                k = bisect.bisect_left(cs, c1)
                while k < len(cs) and cs[k] <= c2:
                    if (r, cs[k]) != exclude:
                        out.append((r, cs[k]))
                        if len(out) >= limit:
                            return out
                    k += 1
        else:
            cols = self._cols[si]
            for c in col_keys[j1:j2]:
                rs = cols[c]
                k = bisect.bisect_left(rs, r1)
                while k < len(rs) and rs[k] <= r2:
                    if (rs[k], c) != exclude:
                        out.append((rs[k], c))
                        if len(out) >= limit:
                            return out
                    k += 1
        return out

    def any_in(self, si: int, r1: int, r2: int, c1: int, c2: int, exclude: tuple[int, int] | None = None) -> bool:
        return bool(self.cells_in(si, r1, r2, c1, c2, limit=1, exclude=exclude))


def rectangles(cells: Iterable[tuple[int, int]]) -> list[tuple[int, int, int, int]]:
    """Cover (row, col) cells with rectangles (r1, c1, r2, c2): runs along each
    row, then identical runs on consecutive rows stacked."""
    by_row: dict[int, list[int]] = defaultdict(list)
    for r, c in cells:
        by_row[r].append(c)
    runs: dict[tuple[int, int], list[int]] = defaultdict(list)
    for r, cols in by_row.items():
        cols.sort()
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
                continue
            if c == prev:
                continue
            runs[(start, prev)].append(r)
            start = prev = c
        runs[(start, prev)].append(r)
    out: list[tuple[int, int, int, int]] = []
    for (c1, c2), rows in runs.items():
        rows.sort()
        s = p = rows[0]
        for r in rows[1:]:
            if r == p + 1:
                p = r
                continue
            out.append((s, c1, p, c2))
            s = p = r
        out.append((s, c1, p, c2))
    return sorted(out)


# --- the strategy ladder --------------------------------------------------------------------
FALLBACK_LABEL = {"0": "0", '""': "empty text", "FALSE": "FALSE"}


def candidate_fixes(formula: str, parsed: Parsed, kind_of_value: str, preferred: str | None = None) -> list[dict[str, Any]]:
    """Every way this module may fix a root error formula, best first. Each is
    {"kind": "formula"|"value", "content", "strategy", "what"}.

    `kind_of_value` ("number" | "text" | "bool") is what the healthy cells with
    the same formula return; `preferred` (a fallback literal chosen by the
    assistant) moves to the front when it is one of the allowed ones."""
    order = {"number": ["0", '""'], "text": ['""', "0"], "bool": ["FALSE", "0", '""']}.get(kind_of_value, ["0", '""'])
    if preferred in FALLBACK_LABEL and preferred in order:
        order = [preferred] + [f for f in order if f != preferred]
    elif preferred in FALLBACK_LABEL:
        order = [preferred] + order
    body: str | None = formula[1:]
    dead = False
    if parsed.broken:
        substituted = _REF_LITERAL_RE.sub("NA()", formula)
        body = None if "#REF!" in substituted.upper() else substituted[1:]
        dead = body is None
    if body is not None and _PURE_NA_RE.match("=" + body):
        body, dead = None, True
    out: list[dict[str, Any]] = []
    if body is not None:
        for fb in order:
            out.append({
                "kind": "formula", "content": f"=IFERROR({body},{fb})", "strategy": f"IFERROR(…, {fb})",
                "what": f"the formula is kept and returns {FALLBACK_LABEL[fb]} when it fails",
            })
    if dead:
        for fb in order:
            value: Any = 0 if fb == "0" else (False if fb == "FALSE" else "")
            out.append({
                "kind": "value", "content": value, "strategy": f"replaced by {FALLBACK_LABEL[fb]}",
                "what": f"the formula could never compute again (its reference was deleted): replaced by {FALLBACK_LABEL[fb]}",
            })
    return out


CONSTANT_FIXES = [
    {"kind": "value", "content": "", "strategy": "error value cleared", "what": "the cell held an error value typed or pasted as data: cleared"},
    {"kind": "value", "content": 0, "strategy": "error value replaced by 0", "what": "the cell held an error value typed or pasted as data: replaced by 0"},
]


# --- configuration, units -------------------------------------------------------------------
@dataclass
class AutoFixConfig:
    max_passes: int = 40
    time_budget_s: float = 3 * 3600
    numeric_tol: float = NUMERIC_TOL
    # advisor(groups) -> {group key: {"fallback": '0'|'""'|'FALSE', "reason": str}} -- the assistant's say on
    # what an error group should return (app/autofix_advisor.py); None = the built-in order only
    advisor: Callable[[list[dict[str, Any]]], dict[str, dict[str, Any]]] | None = None
    progress: Progress | None = None
    should_stop: Callable[[], bool] | None = None


@dataclass
class Unit:
    """One thing a fix writes as a whole: a plain cell, a multi-cell array
    formula, the anchor of a spilled array, or a constant cell."""
    key: Key
    r2: int
    c2: int
    kind: str            # "cell" | "array" | "spill" | "const"
    old: Any             # R1C1 text, or the error text of a constant
    error: str           # the error that made it a root
    group: str           # group key: sheet + error + R1C1 text
    fixes: list[dict[str, Any]] = field(default_factory=list)
    tries: int = 0       # index of the next fix to try
    accepted: dict[str, Any] | None = None
    notes: list[str] = field(default_factory=list)

    @property
    def cells(self) -> list[Key]:
        si, r1, c1 = self.key
        return [(si, r, c) for r in range(r1, self.r2 + 1) for c in range(c1, self.c2 + 1)]


class Stop(Exception):
    """The time budget is used up, or the caller asked to stop: keep what is accepted."""


def _same(a: Any, b: Any, tol: float) -> bool:
    if a is b or a == b:
        return type(a) is type(b) or not (isinstance(a, bool) or isinstance(b, bool))
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= tol * max(1.0, abs(a), abs(b))
    return False


def _is_error(v: Any) -> bool:
    return v.__class__ is int and v in XL_ERROR_CODES


def _retry(fn: Callable[[], Any], attempts: int = 6) -> Any:
    """Excel answers RPC_E_CALL_REJECTED while it is busy for an instant."""
    for n in range(attempts):
        try:
            return fn()
        except Exception as exc:
            code = getattr(exc, "hresult", None) or (exc.args[0] if exc.args else None)
            if code in TRANSIENT_COM_HRESULTS and n < attempts - 1:
                time.sleep(0.4 * (n + 1))
                continue
            raise


# --- the engine -----------------------------------------------------------------------------
class _Sheet:
    __slots__ = ("name", "index", "r0", "c0", "nr", "nc", "per", "vals", "fml", "ws")

    def __init__(self, name: str, index: int, r0: int, c0: int, nr: int, nc: int):
        self.name, self.index, self.r0, self.c0, self.nr, self.nc = name, index, r0, c0, nr, nc
        self.per = max(1, min(CHUNK_ROWS, CHUNK_CELLS // max(1, nc)))
        self.vals: dict[int, tuple] = {}             # chunk start -> rows of values (formula-bearing chunks only)
        self.fml: dict[int, list[list | None]] = {}  # chunk start -> rows of R1C1 texts (None where no formula)
        self.ws: Any = None


class _Engine:
    def __init__(self, excel: Any, wb: Any, cfg: AutoFixConfig):
        self.excel, self.wb, self.cfg = excel, wb, cfg
        self.t0 = time.time()
        self.sheets: list[_Sheet] = []
        self.sheet_index: dict[str, int] = {}
        self.names: dict[str, tuple] = {}
        self.err: dict[Key, int] = {}            # formula cells that are errors now
        self.err0: set[Key] = set()              # ... that were errors when the workbook was opened
        self.const_err: dict[Key, int] = {}      # constants that are error values
        self.unstable: set[Key] = set()          # cells whose value changes from one recalculation to the next
        self.volatile_texts: set[str] = set()
        self.votes: dict[tuple[int, str], list[int]] = {}  # (sheet, R1C1 text) -> [numbers, texts, booleans] among healthy cells
        self.units: dict[Key, Unit] = {}
        self.cell_unit: dict[Key, Key] = {}      # every cell of an array / spill unit -> the unit's key
        self.advice: dict[str, dict[str, Any]] = {}
        self.group_sample: dict[str, dict[str, Any]] = {}
        self.timings: dict[str, float] = defaultdict(float)
        self.counts: dict[str, int] = defaultdict(int)
        self.formula_cells = 0
        self.passes = 0
        self.errors_before = 0
        self.confirm_moved = 0
        self.last_resort_used = False
        self.log: list[str] = []

    # -- plumbing --
    def note(self, text: str) -> None:
        self.log.append(f"{time.time() - self.t0:7.1f}s  {text}")

    def tell(self, stage: str, message: str, fraction: float | None = None, **facts: Any) -> None:
        if self.cfg.progress is not None:
            try:
                self.cfg.progress(stage, message, fraction, **facts)
            except Exception:
                pass

    def check_budget(self) -> None:
        if time.time() - self.t0 > self.cfg.time_budget_s:
            raise Stop(f"the time budget of {int(self.cfg.time_budget_s // 60)} min is used up")
        if self.cfg.should_stop is not None and self.cfg.should_stop():
            raise Stop("stopped on request")

    def _range(self, sh: _Sheet, r1: int, c1: int, r2: int, c2: int) -> Any:
        ws = sh.ws
        return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))

    def _read(self, sh: _Sheet, start: int, what: str) -> tuple:
        rows = min(sh.per, sh.nr - start)

        def go() -> Any:
            rng = self._range(sh, sh.r0 + start, sh.c0, sh.r0 + start + rows - 1, sh.c0 + sh.nc - 1)
            try:
                return getattr(rng, what)
            finally:
                rng = None

        v = _retry(go)
        if not isinstance(v, tuple):
            v = ((v,),)
        return v

    def calculate(self, full: bool = False) -> None:
        t = time.time()
        if full:
            _retry(lambda: self.excel.CalculateFullRebuild())
        else:
            _retry(lambda: self.excel.Calculate())
        self.timings["calculate"] += time.time() - t
        self.counts["calculations"] += 1

    def formula_at(self, key: Key) -> str | None:
        si, row, col = key
        sh = self.sheets[si]
        i, j = row - sh.r0, col - sh.c0
        if i < 0 or j < 0 or i >= sh.nr or j >= sh.nc:
            return None
        rows = sh.fml.get((i // sh.per) * sh.per)
        if rows is None:
            return None
        frow = rows[i % sh.per]
        return None if frow is None else frow[j]

    def _set_formula(self, key: Key, text: str | None) -> None:
        si, row, col = key
        sh = self.sheets[si]
        i, j = row - sh.r0, col - sh.c0
        rows = sh.fml.get((i // sh.per) * sh.per)
        if rows is None:
            return
        k = i % sh.per
        if rows[k] is None:
            if text is None:
                return
            rows[k] = [None] * sh.nc
        rows[k][j] = sys.intern(text) if text is not None else None

    def value_at(self, key: Key) -> Any:
        si, row, col = key
        sh = self.sheets[si]
        i, j = row - sh.r0, col - sh.c0
        rows = sh.vals.get((i // sh.per) * sh.per)
        return None if rows is None else rows[i % sh.per][j]

    # -- first read: every cell --
    def load(self) -> None:
        t = time.time()
        wb = self.wb
        count = int(wb.Worksheets.Count)
        for n in range(1, count + 1):
            ws = wb.Worksheets(n)
            ur = ws.UsedRange
            sh = _Sheet(str(ws.Name), n - 1, int(ur.Row), int(ur.Column), int(ur.Rows.Count), int(ur.Columns.Count))
            ur = None
            sh.ws = ws
            self.sheets.append(sh)
            self.sheet_index[sh.name.upper()] = sh.index
        total = sum(s.nr for s in self.sheets) or 1
        done = 0
        for sh in self.sheets:
            for start in range(0, sh.nr, sh.per):
                self.tell("read", f"Reading {sh.name}", done / total, sheet=sh.name)
                vals = self._read(sh, start, "Value2")
                done += len(vals)
                if not any(any(v is not None for v in row) for row in vals):
                    continue
                texts = self._read(sh, start, "FormulaR1C1")
                rows: list[list | None] = []
                any_formula = False
                for i, trow in enumerate(texts):
                    vrow = vals[i]
                    frow: list | None = None
                    for j, f in enumerate(trow):
                        v = vrow[j]
                        if f.__class__ is str and f[:1] == "=":
                            if frow is None:
                                frow = [None] * sh.nc
                            f = sys.intern(f)
                            frow[j] = f
                            self.formula_cells += 1
                            if v.__class__ is int and v in XL_ERROR_CODES:
                                self.err[(sh.index, sh.r0 + start + i, sh.c0 + j)] = v
                            else:
                                vote = self.votes.get((sh.index, f))
                                if vote is None:
                                    vote = self.votes[(sh.index, f)] = [0, 0, 0]
                                vote[2 if v.__class__ is bool else 1 if v.__class__ is str else 0] += 1
                        elif v.__class__ is int and v in XL_ERROR_CODES:
                            self.const_err[(sh.index, sh.r0 + start + i, sh.c0 + j)] = v
                    rows.append(frow)
                    any_formula = any_formula or frow is not None
                if any_formula:
                    sh.vals[start] = vals
                    sh.fml[start] = rows
        for (_, text) in self.votes:
            if text not in self.volatile_texts and parse_r1c1(text).funcs & VOLATILE_FUNCS:
                self.volatile_texts.add(text)
        self.err0 = set(self.err)
        self._load_names()
        self.timings["first_read"] = time.time() - t
        self.note(f"read {self.formula_cells} formula cells on {len(self.sheets)} sheets: {len(self.err)} error formulas, {len(self.const_err)} error constants")

    def _load_names(self) -> None:
        try:
            n = int(self.wb.Names.Count)
        except Exception:
            return
        for i in range(1, min(n, 5000) + 1):
            try:
                nm = self.wb.Names.Item(i)
                name, refers = str(nm.Name), str(nm.RefersToR1C1)
                nm = None
            except Exception:
                continue
            target = self._name_target(refers)
            short = name.split("!")[-1].upper()
            self.names.setdefault(short, target)
            self.names[name.upper()] = target

    def _name_target(self, refers: str) -> tuple:
        p = parse_r1c1(refers)
        if p.broken:
            return ("broken",)
        if not p.refs and not p.names and not p.funcs:
            return ("value",)
        if len(p.refs) == 1 and not p.funcs and not p.names and not p.dynamic:
            spec, a, b = p.refs[0]
            absolute = all(ax is None or ax[0] for ep in (a, b) if ep is not None for ax in ep)
            if spec is not None and spec[0] == "sheet" and absolute and spec[1].upper() in self.sheet_index:
                r1, r2, c1, c2 = ref_rect(p.refs[0], 1, 1)
                return ("range", self.sheet_index[spec[1].upper()], r1, r2, c1, c2)
        return ("dynamic",)

    # -- later reads: formula cells only --
    def read_changes(self) -> tuple[dict[tuple[int, int], tuple], list[tuple[Key, Any, Any]]]:
        """Re-read every chunk that holds formulas. -> (the chunks that differ,
        [(cell, old value, new value)] for every formula cell whose value moved)."""
        t = time.time()
        fresh: dict[tuple[int, int], tuple] = {}
        changed: list[tuple[Key, Any, Any]] = []
        tol = self.cfg.numeric_tol
        for sh in self.sheets:
            for start, old in sh.vals.items():
                new = self._read(sh, start, "Value2")
                if new == old:
                    continue
                fresh[(sh.index, start)] = new
                rows = sh.fml[start]
                for i, nrow in enumerate(new):
                    orow = old[i]
                    if nrow == orow:
                        continue
                    frow = rows[i]
                    if frow is None:
                        continue
                    for j, v in enumerate(nrow):
                        if frow[j] is not None and not _same(orow[j], v, tol):
                            changed.append(((sh.index, sh.r0 + start + i, sh.c0 + j), orow[j], v))
        self.timings["read"] += time.time() - t
        self.counts["reads"] += 1
        return fresh, changed

    def commit(self, fresh: dict[tuple[int, int], tuple], changed: list[tuple[Key, Any, Any]]) -> None:
        for (si, start), vals in fresh.items():
            self.sheets[si].vals[start] = vals
        for key, _, new in changed:
            if _is_error(new):
                self.err[key] = new
            else:
                self.err.pop(key, None)

    # -- roots --
    def _sheets_of(self, spec: Any, si: int) -> list[int]:
        if spec is None:
            return [si]
        if spec[0] == "sheet":
            i = self.sheet_index.get(spec[1].upper())
            return [] if i is None else [i]
        if spec[0] == "3d":
            a, b = self.sheet_index.get(spec[1].upper()), self.sheet_index.get(spec[2].upper())
            if a is None or b is None:
                return []
            return list(range(min(a, b), max(a, b) + 1))
        return []

    def precedent_rects(self, key: Key, parsed: Parsed) -> tuple[list[tuple[int, int, int, int, int]], bool]:
        """Every rectangle the formula at `key` reads that this module can see,
        and whether it also reads something it cannot (dynamic)."""
        si, row, col = key
        out: list[tuple[int, int, int, int, int]] = []
        dynamic = parsed.dynamic
        for ref in parsed.refs:
            for sj in self._sheets_of(ref[0], si):
                r1, r2, c1, c2 = ref_rect(ref, row, col)
                out.append((sj, r1, r2, c1, c2))
        for name in parsed.names:
            target = self.names.get(name)
            if target is None:
                continue
            if target[0] == "range":
                out.append(target[1:])
            elif target[0] == "dynamic":
                dynamic = True
        return out, dynamic

    def classify(self) -> tuple[list[Key], list[Key], list[Key]]:
        """-> (roots, uncertain, propagated) among the error formula cells.
        An error is propagated when a cell it reads holds the same error."""
        by_code: dict[int, CellIndex] = {}
        for key, code in self.err.items():
            by_code.setdefault(code, CellIndex()).by_sheet[key[0]].add((key[1], key[2]))
        for key, code in self.const_err.items():
            by_code.setdefault(code, CellIndex()).by_sheet[key[0]].add((key[1], key[2]))
        roots: list[Key] = []
        uncertain: list[Key] = []
        propagated: list[Key] = []
        for key, code in self.err.items():
            text = self.formula_at(key)
            if text is None:
                continue
            parsed = parse_r1c1(text)
            rects, dynamic = self.precedent_rects(key, parsed)
            idx = by_code[code]
            me = (key[1], key[2])
            hit = False
            for sj, r1, r2, c1, c2 in rects:
                if idx.any_in(sj, r1, r2, c1, c2, exclude=me if sj == key[0] else None):
                    hit = True
                    break
            if hit:
                propagated.append(key)
            elif dynamic:
                uncertain.append(key)
            else:
                roots.append(key)
        return roots, uncertain, propagated

    # -- units --
    def _kind_of_value(self, si: int, text: str, parsed: Parsed) -> str:
        n, t, b = self.votes.get((si, text), (0, 0, 0))
        if b > n and b > t:
            return "bool"
        if t > n:
            return "text"
        if n == 0 and t == 0:
            masked = _STRING_RE.sub('""', text)
            if "&" in masked:
                return "text"
        return "number"

    def group_key(self, si: int, code: int, text: str) -> str:
        return f"{self.sheets[si].name}|{XL_ERROR_CODES.get(code, '?')}|{text}"

    def make_units(self, cells: list[Key]) -> list[Unit]:
        """Turn root cells into units (resolving arrays and spills through Excel),
        each with its ladder of fixes. Cells already known keep their state."""
        fresh: list[Unit] = []
        by_text: dict[tuple[int, str], list[Key]] = defaultdict(list)
        for key in cells:
            if key in self.units or key in self.cell_unit:
                continue
            text = self.formula_at(key)
            if text is not None:
                by_text[(key[0], text)].append(key)
        for (si, text), keys in by_text.items():
            sh = self.sheets[si]
            parsed = parse_r1c1(text)
            for r1, c1, r2, c2 in rectangles((k[1], k[2]) for k in keys):
                plain = True
                try:
                    rng = self._range(sh, r1, c1, r2, c2)
                    try:
                        has_array = rng.HasArray
                        first = rng.Cells(1, 1)
                        try:
                            has_spill = bool(first.HasSpill)
                        except Exception:
                            has_spill = False
                        first = None
                    finally:
                        rng = None
                    plain = has_array is False and not has_spill
                except Exception:
                    plain = False
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        key = (si, r, c)
                        if key in self.units or key in self.cell_unit:
                            continue
                        code = self.err.get(key)
                        if code is None:
                            continue
                        unit = self._plain_unit(key, text, parsed, code) if plain else self._resolve_unit(key, text, parsed, code)
                        if unit is not None:
                            fresh.append(unit)
        return fresh

    def _ladder(self, si: int, code: int, text: str, parsed: Parsed) -> list[dict[str, Any]]:
        advice = self.advice.get(self.group_key(si, code, text)) or {}
        fixes = candidate_fixes(text, parsed, self._kind_of_value(si, text, parsed), advice.get("fallback"))
        if advice.get("reason"):
            fixes = [{**f, "why": advice["reason"]} if n == 0 else f for n, f in enumerate(fixes)]
        return fixes

    def _plain_unit(self, key: Key, text: str, parsed: Parsed, code: int) -> Unit:
        unit = Unit(key, key[1], key[2], "cell", text, XL_ERROR_CODES[code], self.group_key(key[0], code, text), self._ladder(key[0], code, text, parsed))
        self.units[key] = unit
        return unit

    def _resolve_unit(self, key: Key, text: str, parsed: Parsed, code: int) -> Unit | None:
        """A root cell that is part of an array formula or of a spilled range:
        the unit is the whole array / the anchor cell."""
        si, row, col = key
        sh = self.sheets[si]
        cell = sh.ws.Cells(row, col)
        try:
            kind, r1, c1, r2, c2 = "cell", row, col, row, col
            if cell.HasArray:
                area = cell.CurrentArray
                try:
                    r1, c1 = int(area.Row), int(area.Column)
                    r2, c2 = r1 + int(area.Rows.Count) - 1, c1 + int(area.Columns.Count) - 1
                finally:
                    area = None
                kind = "array"
            else:
                try:
                    spilled = bool(cell.HasSpill)
                except Exception:
                    spilled = False
                if spilled:
                    parent = cell.SpillParent
                    try:
                        r1, c1 = int(parent.Row), int(parent.Column)
                    finally:
                        parent = None
                    r2, c2 = r1, c1
                    kind = "spill"
        finally:
            cell = None
        anchor = (si, r1, c1)
        if anchor in self.units:
            self.cell_unit[key] = anchor
            return None
        anchor_text = self.formula_at(anchor) or text
        unit = Unit(anchor, r2, c2, kind, anchor_text, XL_ERROR_CODES[code], self.group_key(si, code, anchor_text), self._ladder(si, code, anchor_text, parse_r1c1(anchor_text)))
        self.units[anchor] = unit
        if kind == "array":
            for k in unit.cells:
                self.cell_unit[k] = anchor
        self.cell_unit[key] = anchor
        return unit

    def make_const_units(self) -> list[Unit]:
        fresh = []
        for key, code in self.const_err.items():
            if key in self.units:
                continue
            unit = Unit(key, key[1], key[2], "const", XL_ERROR_CODES[code], XL_ERROR_CODES[code], f"{self.sheets[key[0]].name}|{XL_ERROR_CODES[code]}|(constant)", [dict(f) for f in CONSTANT_FIXES])
            self.units[key] = unit
            fresh.append(unit)
        return fresh

    # -- writing --
    def _write_plain(self, si: int, cells: list[tuple[int, int]], content: Any, formula: bool) -> list[tuple[int, int, str]]:
        """Write one content into plain cells, a rectangle at a time. -> failures (row, col, message)."""
        sh = self.sheets[si]
        failures: list[tuple[int, int, str]] = []
        for r1, c1, r2, c2 in rectangles(cells):
            try:
                def go() -> None:
                    rng = self._range(sh, r1, c1, r2, c2)
                    try:
                        if formula:
                            rng.FormulaR1C1 = content
                        elif content == "":
                            rng.ClearContents()
                        else:
                            rng.Value2 = content
                    finally:
                        rng = None

                _retry(go)
                self.counts["writes"] += 1
            except Exception as exc:
                if (r1, c1) == (r2, c2):
                    failures.append((r1, c1, _com_text(exc)))
                    continue
                for r in range(r1, r2 + 1):  # one cell refused: find which
                    for c in range(c1, c2 + 1):
                        failures.extend(self._write_plain(si, [(r, c)], content, formula))
        return failures

    def _write_unit(self, unit: Unit, content: Any, formula: bool) -> str | None:
        """Write an array / spill / constant unit. -> the failure message, or None."""
        si, r1, c1 = unit.key
        sh = self.sheets[si]
        try:
            def go() -> None:
                rng = self._range(sh, r1, c1, unit.r2, unit.c2)
                try:
                    if unit.kind == "array" and formula:
                        rng.FormulaArray = content
                    elif unit.kind == "array":
                        rng.ClearContents()
                        if content != "":
                            rng.Value2 = content
                    elif unit.kind == "spill" and formula:
                        rng.Formula2R1C1 = content
                    elif formula:
                        rng.FormulaR1C1 = content
                    elif unit.kind == "const" and isinstance(content, str) and content.startswith("#"):
                        rng.Formula = content  # an error value back into a constant cell
                    elif content == "":
                        rng.ClearContents()
                    else:
                        rng.Value2 = content
                finally:
                    rng = None

            _retry(go)
            self.counts["writes"] += 1
            return None
        except Exception as exc:
            return _com_text(exc)

    def apply(self, units: list[Unit]) -> list[Unit]:
        """Write the current fix of every unit. -> the units Excel refused."""
        t = time.time()
        refused: list[Unit] = []
        plain: dict[tuple[int, bool, str], list[Unit]] = defaultdict(list)
        contents: dict[tuple[int, bool, str], Any] = {}
        for u in units:
            fix = u.fixes[u.tries]
            if u.kind == "cell":
                slot = (u.key[0], fix["kind"] == "formula", repr(fix["content"]))  # repr: 0 and False are different contents
                plain[slot].append(u)
                contents[slot] = fix["content"]
            else:
                msg = self._write_unit(u, fix["content"], fix["kind"] == "formula")
                if msg:
                    u.notes.append(f"{fix['strategy']}: Excel refused it ({msg})")
                    refused.append(u)
        for slot, group in plain.items():
            si, is_formula, _ = slot
            failed = self._write_plain(si, [(u.key[1], u.key[2]) for u in group], contents[slot], is_formula)
            if failed:
                bad = {(r, c): m for r, c, m in failed}
                for u in group:
                    m = bad.get((u.key[1], u.key[2]))
                    if m:
                        u.notes.append(f"{u.fixes[u.tries]['strategy']}: Excel refused it ({m})")
                        refused.append(u)
        self.timings["write"] += time.time() - t
        return refused

    def revert(self, units: list[Unit], count: bool = True) -> None:
        t = time.time()
        plain: dict[tuple[int, str], list[tuple[int, int]]] = defaultdict(list)
        for u in units:
            if u.kind == "cell":
                plain[(u.key[0], u.old)].append((u.key[1], u.key[2]))
            else:
                msg = self._write_unit(u, u.old, u.kind != "const")
                if msg:
                    u.notes.append(f"could not restore the original content ({msg})")
                    self.note(f"REVERT FAILED {self.sheets[u.key[0]].name}!{a1(u.key[1], u.key[2])}: {msg}")
        for (si, text), cells in plain.items():
            for r, c, m in self._write_plain(si, cells, text, True):
                self.note(f"REVERT FAILED {self.sheets[si].name}!{a1(r, c)}: {m}")
        if count:
            self.counts["rolled_back"] += len(units)
        self.timings["write"] += time.time() - t

    # -- the gate --
    def judge(self, changed: list[tuple[Key, Any, Any]], written: set[Key]) -> list[tuple[Key, Any, Any]]:
        """The changes that are violations: a formula cell that was fine when the
        workbook was opened and now holds another value."""
        out = []
        for key, old, new in changed:
            if key in written or key in self.err0 or key in self.unstable:
                continue
            text = self.formula_at(key)
            if text is not None and text in self.volatile_texts:
                continue
            out.append((key, old, new))
        return out

    def blame(self, violations: list[tuple[Key, Any, Any]], changed: list[tuple[Key, Any, Any]], written: dict[Key, Unit]) -> tuple[set[Key], bool]:
        """Walk back from each violation through the cells whose value moved, to
        the units written in this pass. -> (unit keys, whether the walk is complete)."""
        moved = CellIndex([k for k, _, _ in changed] + list(written))
        blamed: set[Key] = set()
        complete = True
        seen: set[Key] = set()
        stack = [k for k, _, _ in violations[:400]]
        start = set(stack)
        if len(violations) > 400:
            complete = False
        while stack:
            if len(seen) > 60_000:
                complete = False
                break
            key = stack.pop()
            if key in seen:
                continue
            seen.add(key)
            unit = written.get(key)
            if unit is not None:
                blamed.add(unit.key)
                continue
            text = self.formula_at(key)
            if text is None:
                continue
            rects, dynamic = self.precedent_rects(key, parse_r1c1(text))
            if dynamic:
                complete = False
            found = False
            for sj, r1, r2, c1, c2 in rects:
                for r, c in moved.cells_in(sj, r1, r2, c1, c2, limit=5000):
                    if (sj, r, c) != key:
                        found = True
                        stack.append((sj, r, c))
            if not found and key in start:
                complete = False
        return blamed, complete

    def settle(self, units: list[Unit], label: str, depth: int = 0) -> list[Unit]:
        """Apply the units' current fixes and keep the ones that pass the gate.
        A unit that fails moves on to its next fix (tried on a later pass).
        -> the accepted units."""
        self.check_budget()
        units = [u for u in units if u.tries < len(u.fixes)]
        if not units:
            return []
        self.tell("fix", f"{label}: writing {len(units)} fix{'es' if len(units) != 1 else ''}", None, pass_no=self.passes, writing=len(units))
        refused = self.apply(units)
        for u in refused:
            u.tries += 1
        if refused:
            self.revert([u for u in refused if u.kind != "cell"], count=False)  # a half-written array
            bad = {u.key for u in refused}
            units = [u for u in units if u.key not in bad]
            if not units:
                self.calculate()
                return []
        self.calculate()
        fresh, changed = self.read_changes()
        written: dict[Key, Unit] = {}
        for u in units:
            for k in u.cells:
                written[k] = u
        # a fix that leaves its own cell in error cured nothing
        new_value = {k: new for k, _, new in changed}
        useless = [u for u in units if u.kind != "const" and _is_error(new_value.get(u.key, self.value_at(u.key) if u.key not in new_value else None))]
        violations = self.judge(changed, set(written))
        if not violations and not useless:
            self._accept(units, fresh, changed)
            return units
        guilty: set[Key] = {u.key for u in useless}
        useless_keys = set(guilty)
        complete = True
        if violations:
            blamed, complete = self.blame(violations, changed, written)
            guilty |= blamed
        if complete and guilty:
            out = [u for u in units if u.key in guilty]
            keep = [u for u in units if u.key not in guilty]
            self.revert(out)
            for u in out:
                self._failed(u, violations, u.key in useless_keys)
            self.calculate()
            fresh, changed = self.read_changes()
            violations = self.judge(changed, {k for u in keep for k in u.cells})
            if not violations:
                self._accept(keep, fresh, changed)
                return keep
            if not keep:
                for key, _, _ in violations:  # nothing of ours is written any more: these move by themselves
                    self.unstable.add(key)
                self.commit(fresh, changed)
                return []
            units = keep  # the walk missed something: fall through to the split
        # everything back, then find the guilty ones by halves
        self.revert(units)
        self.calculate()
        fresh, changed = self.read_changes()
        leftover = self.judge(changed, set())
        if leftover:
            # these moved although nothing is written any more: they are not ours (volatile / iterative)
            for key, _, _ in leftover:
                self.unstable.add(key)
            self.note(f"{len(leftover)} cell(s) change by themselves between recalculations: left out of the gate")
        self.commit(fresh, [c for c in changed])
        if len(units) == 1:
            self._failed(units[0], violations, units[0].key in useless_keys)
            return []
        if depth > 24:
            for u in units:
                self._failed(u, violations, False)
            return []
        half = self._split(units)
        return self.settle(units[:half], label, depth + 1) + self.settle(units[half:], label, depth + 1)

    @staticmethod
    def _split(units: list[Unit]) -> int:
        """Cut at a group boundary near the middle when there is one."""
        mid = len(units) // 2
        for off in range(0, mid):
            for i in (mid - off, mid + off):
                if 0 < i < len(units) and units[i - 1].group != units[i].group:
                    return i
        return mid

    def _failed(self, unit: Unit, violations: list[tuple[Key, Any, Any]], useless: bool) -> None:
        fix = unit.fixes[unit.tries]
        if useless:
            unit.notes.append(f"{fix['strategy']}: the cell was still an error")
        else:
            sample = "; ".join(f"{self.sheets[k[0]].name}!{a1(k[1], k[2])} {_show(o)} → {_show(n)}" for k, o, n in violations[:3])
            unit.notes.append(f"{fix['strategy']}: it changed {len(violations)} good value(s), e.g. {sample}")
        unit.tries += 1

    def _accept(self, units: list[Unit], fresh: dict, changed: list) -> None:
        self.commit(fresh, changed)
        for u in units:
            fix = u.fixes[u.tries]
            u.accepted = fix
            if u.kind == "const":
                self.const_err.pop(u.key, None)
                continue
            new_text = fix["content"] if fix["kind"] == "formula" else None
            for k in (u.cells if u.kind == "array" else [u.key]):
                self._set_formula(k, new_text)
                if new_text is None:
                    self.err.pop(k, None)
        self.counts["accepted"] += len(units)

    # -- the loop --
    def ask_advisor(self, units: list[Unit]) -> None:
        if self.cfg.advisor is None:
            return
        groups: dict[str, dict[str, Any]] = {}
        for u in units:
            if u.kind == "const" or u.group in self.advice or u.tries > 0:
                continue
            g = groups.get(u.group)
            if g is None:
                si, row, col = u.key
                g = groups[u.group] = {
                    "key": u.group, "sheet": self.sheets[si].name, "error": u.error, "cell": a1(row, col),
                    "formula": r1c1_to_a1(u.old, row, col), "formula_r1c1": u.old, "count": 0,
                    "kind_of_value": self._kind_of_value(si, u.old, parse_r1c1(u.old)),
                    "context": self._context(u),
                }
            g["count"] += 1
        if not groups:
            return
        t = time.time()
        self.tell("advise", f"Asking the assistant about {len(groups)} kind(s) of error", None)
        try:
            answers = self.cfg.advisor(sorted(groups.values(), key=lambda g: -g["count"])) or {}
        except Exception as exc:
            self.note(f"assistant not available for this pass: {str(exc)[:160]}")
            answers = {}
        self.timings["assistant"] += time.time() - t
        for key in groups:
            self.advice[key] = answers.get(key) or {}
        for u in units:
            advice = self.advice.get(u.group)
            if advice and u.kind != "const" and u.tries == 0:
                si = u.key[0]
                code = ERROR_CODE_OF.get(u.error, 0)
                u.fixes = self._ladder(si, code, u.old, parse_r1c1(u.old))

    def _context(self, unit: Unit) -> dict[str, Any]:
        """What the assistant may look at: the labels left of and above the cell,
        and the values of the cells the formula reads (first few)."""
        si, row, col = unit.key
        sh = self.sheets[si]
        out: dict[str, Any] = {}
        try:
            ws = sh.ws
            if col > 1:
                left = self._retry_value(ws, row, max(1, col - 6), row, col - 1)
                out["left_of_cell"] = [v for v in left if isinstance(v, str) and v.strip()][-2:]
            if row > 1:
                up = self._retry_value(ws, max(1, row - 6), col, row - 1, col)
                out["above_cell"] = [v for v in up if isinstance(v, str) and v.strip()][-2:]
        except Exception:
            pass
        reads = []
        rects, _ = self.precedent_rects(unit.key, parse_r1c1(unit.old))
        for sj, r1, r2, c1, c2 in rects[:6]:
            if (r2 - r1 + 1) * (c2 - c1 + 1) <= 4:
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        v = self.value_at((sj, r, c))
                        if v is None and self.formula_at((sj, r, c)) is None:
                            try:
                                v = self._retry_value(self.sheets[sj].ws, r, c, r, c)[0]
                            except Exception:
                                v = None
                        reads.append(f"{self.sheets[sj].name}!{a1(r, c)} = {_show(v)}")
            else:
                reads.append(f"{self.sheets[sj].name}!{rect_a1(r1, c1, min(r2, MAX_ROW), min(c2, MAX_COL))} (a range)")
        out["reads"] = reads[:10]
        return out

    def _retry_value(self, ws: Any, r1: int, c1: int, r2: int, c2: int) -> list[Any]:
        def go() -> Any:
            rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
            try:
                return rng.Value2
            finally:
                rng = None

        v = _retry(go)
        if not isinstance(v, tuple):
            return [v]
        return [x for row in v for x in row]

    def genuine_errors(self) -> dict[Key, int]:
        """Error formula cells, minus MM_ calls that only fail because the
        MMForExcel add-in is not installed here (not a workbook defect)."""
        out = {}
        for key, code in self.err.items():
            if XL_ERROR_CODES.get(code) == "#NAME?":
                text = self.formula_at(key) or ""
                if "MM_" in text.upper():
                    continue
            out[key] = code
        return out

    def run(self) -> str:
        """-> the status: clean | partial | stuck | unchanged."""
        self.tell("read", "Recalculating the workbook in Excel", 0.0)
        self.calculate(full=True)
        self.load()
        # what moves by itself between two recalculations is not ours to guard
        self.calculate()
        fresh, changed = self.read_changes()
        for key, _, _ in changed:
            self.unstable.add(key)
        self.commit(fresh, changed)
        self.err0 |= set(self.err)
        if changed:
            self.note(f"{len(changed)} formula cell(s) change by themselves between two recalculations (volatile / iterative): left out of the gate")
        self.errors_before = len(self.genuine_errors()) + len(self.const_err)
        if not self.err and not self.const_err:
            return "unchanged"
        stalled = 0
        last_resort = False
        while self.passes < self.cfg.max_passes:
            self.check_budget()
            remaining = self.genuine_errors()
            if not remaining and not self.const_err:
                break
            self.passes += 1
            self.tell("find", f"Pass {self.passes}: looking for the cells where the {len(remaining)} error(s) start", None, pass_no=self.passes, errors=len(remaining))
            t = time.time()
            roots, uncertain, propagated = self.classify()
            self.timings["classify"] += time.time() - t
            name_gap = {k for k in roots + uncertain + propagated if k not in remaining}
            roots = [k for k in roots if k not in name_gap]
            uncertain = [k for k in uncertain if k not in name_gap]
            fresh_units = self.make_units(roots) + self.make_const_units()
            todo = self._open_units(roots) + [u for u in self.units.values() if u.kind == "const" and u.accepted is None and u.tries < len(u.fixes)]
            tier = "the cells where errors start"
            if not todo:
                self.make_units(uncertain)
                todo = self._open_units(uncertain)
                tier = "cells that read other cells indirectly (INDIRECT / OFFSET / a named formula)"
            if not todo and not any(u.accepted is None and u.tries >= len(u.fixes) for u in self.units.values()):
                # nothing is a clear starting point (errors that feed each other): take them all
                rest = [k for k in propagated if k not in name_gap]
                self.make_units(rest)
                todo = self._open_units(rest)
                tier = "the remaining error cells"
                last_resort = True
            if not todo:
                break
            self.note(f"pass {self.passes}: {len(remaining)} errors, {len(roots)} roots, {len(uncertain)} uncertain, {len(propagated)} propagated -> {len(todo)} unit(s) to fix ({tier})")
            self.ask_advisor(todo)
            todo.sort(key=lambda u: (u.group, u.key))
            before = len(remaining) + len(self.const_err)
            accepted = self.settle(todo, f"Pass {self.passes}")
            after = len(self.genuine_errors()) + len(self.const_err)
            self.note(f"pass {self.passes}: {len(accepted)} of {len(todo)} fix(es) kept, errors {before} -> {after}")
            self.tell("pass", f"Pass {self.passes}: {len(accepted)} fix(es) kept, {after} error(s) left", None, pass_no=self.passes, errors=after, kept=len(accepted))
            stalled = stalled + 1 if after >= before and not accepted else 0
            if stalled >= 6:
                break
        self.last_resort_used = last_resort
        left = len(self.genuine_errors()) + len(self.const_err)
        if left == 0:
            return "clean"
        return "partial" if self.counts["accepted"] else "stuck"

    def _open_units(self, cells: list[Key]) -> list[Unit]:
        seen: set[Key] = set()
        out = []
        for key in cells:
            ukey = key if key in self.units else self.cell_unit.get(key)
            if ukey is None or ukey in seen:
                continue
            seen.add(ukey)
            u = self.units[ukey]
            if u.accepted is None and u.tries < len(u.fixes):
                out.append(u)
        return out

    # -- results --
    def confirm(self) -> None:
        """The full rebuild the app's own Recalculate does, and one more read."""
        self.tell("confirm", "Full recalculation to confirm the result", None)
        self.calculate(full=True)
        fresh, changed = self.read_changes()
        moved = self.judge(changed, set())
        if moved:
            self.note(f"the confirming full rebuild moved {len(moved)} guarded cell(s): {', '.join(self.sheets[k[0]].name + '!' + a1(k[1], k[2]) for k, _, _ in moved[:5])}")
        self.confirm_moved = len(moved)
        self.commit(fresh, changed)

    def report(self) -> dict[str, Any]:
        groups: dict[tuple[str, str], dict[str, Any]] = {}
        applied: list[dict[str, Any]] = []
        by_strategy: dict[str, int] = defaultdict(int)
        cells_rewritten = 0
        accepted = [u for u in self.units.values() if u.accepted is not None]
        buckets: dict[tuple[int, str, str], list[Unit]] = defaultdict(list)
        for u in accepted:
            buckets[(u.key[0], u.group, u.accepted["strategy"])].append(u)
        for (si, group, strategy), units in buckets.items():
            sh = self.sheets[si]
            first = min(units, key=lambda u: u.key)
            fix = first.accepted
            n_cells = sum(len(u.cells) if u.kind == "array" else 1 for u in units)
            cells_rewritten += n_cells
            by_strategy[strategy] += n_cells
            before = r1c1_to_a1(first.old, first.key[1], first.key[2]) if first.kind != "const" else first.old
            after = r1c1_to_a1(fix["content"], first.key[1], first.key[2]) if fix["kind"] == "formula" else fix["content"]
            if first.kind == "cell":
                rects = rectangles((u.key[1], u.key[2]) for u in units)
            else:
                rects = [(u.key[1], u.key[2], u.r2, u.c2) for u in units]
            reason = f"{ERROR_MEANING.get(first.error, first.error)}: {fix['what']}" + (f" ({fix['why']})" if fix.get("why") else "")
            groups[(group, strategy)] = {
                "sheet": sh.name, "error": first.error, "cells": n_cells, "ranges": [rect_a1(*r) for r in rects[:12]], "more_ranges": max(0, len(rects) - 12),
                "cell": a1(first.key[1], first.key[2]), "before": before, "after": after, "strategy": strategy, "reason": reason,
                "kind": first.kind,
            }
            for r1, c1, r2, c2 in rects:
                op_before = r1c1_to_a1(first.old, r1, c1) if first.kind != "const" else first.old
                op_after = r1c1_to_a1(fix["content"], r1, c1) if fix["kind"] == "formula" else fix["content"]
                n = (r2 - r1 + 1) * (c2 - c1 + 1)
                applied.append({
                    "op": "set_array_formula" if first.kind == "array" and fix["kind"] == "formula" else "set_formula" if fix["kind"] == "formula" else ("clear_cell" if fix["content"] == "" else "set_value"),
                    "action_id": "autofix", "rule_id": "READY-001", "sheet": sh.name, "cell": rect_a1(r1, c1, r2, c2),
                    **({"range": rect_a1(r1, c1, r2, c2)} if first.kind == "array" else {}),
                    "before": op_before, "after": op_after,
                    "note": f"auto-fix ({first.error}): {fix['what']}" + (f"; same formula pattern in all {n} cells" if n > 1 else ""),
                })
        remaining = self.genuine_errors()
        left: list[dict[str, Any]] = []
        left_groups: dict[str, dict[str, Any]] = {}
        for key, code in sorted(remaining.items()):
            text = self.formula_at(key) or ""
            ukey = key if key in self.units else self.cell_unit.get(key)
            unit = self.units.get(ukey) if ukey else None
            gk = self.group_key(key[0], code, text)
            g = left_groups.get(gk)
            if g is None:
                why = "its error comes from another cell that could not be fixed"
                if unit is not None and unit.notes:
                    why = "; ".join(unit.notes[-3:])
                elif unit is not None and not unit.fixes:
                    why = "no safe fix is known for this formula"
                g = left_groups[gk] = {
                    "sheet": self.sheets[key[0]].name, "cell": a1(key[1], key[2]), "error": XL_ERROR_CODES.get(code, "?"),
                    "formula": r1c1_to_a1(text, key[1], key[2]) if text else "", "count": 0, "why": why, "tried": bool(unit is not None and unit.notes),
                }
            g["count"] += 1
            if len(left) < 5000:
                left.append({"sheet": self.sheets[key[0]].name, "cell": a1(key[1], key[2]), "error": XL_ERROR_CODES.get(code, "?"), "formula": r1c1_to_a1(text, key[1], key[2]) if text else ""})
        for key, code in sorted(self.const_err.items()):
            unit = self.units.get(key)
            left_groups[f"const|{key}"] = {
                "sheet": self.sheets[key[0]].name, "cell": a1(key[1], key[2]), "error": XL_ERROR_CODES.get(code, "?"), "formula": "(an error value typed or pasted as data)",
                "count": 1, "why": "; ".join(unit.notes[-3:]) if unit and unit.notes else "not fixed", "tried": bool(unit and unit.notes),
            }
        addin_gap = [k for k in self.err if k not in remaining]
        fix_groups = sorted(groups.values(), key=lambda g: -g["cells"])
        left_list = sorted(left_groups.values(), key=lambda g: (not g["tried"], -g["count"]))
        return {
            "errors_before": self.errors_before,
            "errors_after": len(remaining) + len(self.const_err),
            "cells_fixed": max(0, self.errors_before - len(remaining) - len(self.const_err)),
            "cells_rewritten": cells_rewritten,
            "passes": self.passes,
            "rolled_back": self.counts["rolled_back"],
            "by_strategy": dict(by_strategy),
            "fixes": fix_groups[:MAX_REPORTED], "fix_groups": len(fix_groups),
            "left": left_list[:MAX_REPORTED], "left_groups": len(left_list), "left_count": len(remaining) + len(self.const_err),
            "left_cells": left,
            "addin_gap_cells": [{"sheet": self.sheets[k[0]].name, "cell": a1(k[1], k[2]), "formula": r1c1_to_a1(self.formula_at(k) or "", k[1], k[2])} for k in sorted(addin_gap)[:2000]],
            "unstable_cells": len(self.unstable),
            "formula_cells": self.formula_cells,
            "applied": applied,
            "timings": {k: round(v, 1) for k, v in self.timings.items()},
            "counts": dict(self.counts),
            "log": self.log,
        }

    def release(self) -> None:
        for sh in self.sheets:
            sh.ws = None


def _show(v: Any) -> str:
    if v.__class__ is int and v in XL_ERROR_CODES:
        return XL_ERROR_CODES[v]
    if isinstance(v, float):
        return f"{v:.10g}"
    if v is None:
        return "(empty)"
    return repr(v) if isinstance(v, str) else str(v)


def _com_text(exc: BaseException) -> str:
    info = getattr(exc, "excepinfo", None)
    if info and len(info) > 2 and info[2]:
        return str(info[2])[:200]
    if info and len(info) > 5 and info[5] == -2146827284:
        return "Excel does not accept this formula"
    return str(exc)[:200]


def summary_sentence(res: dict[str, Any]) -> str:
    before, after = res.get("errors_before", 0), res.get("errors_after", 0)
    if res.get("status") == "unchanged":
        return "No formula error to fix: the workbook already recalculates clean."
    fixed = f"{res.get('cells_rewritten', 0)} cell(s) rewritten in {res.get('passes', 0)} pass(es); errors {before} → {after}"
    if res.get("status") == "clean":
        return f"Clean: every formula recalculates without an error. {fixed}. No value that was good before has changed."
    if res.get("status") == "partial":
        return f"{fixed}. {after} error(s) are left for a person: fixing them automatically would change a good value, or no safe fix is known."
    return f"Nothing could be fixed safely: {after} error(s) are left for a person."


def run_autofix(source_path: Path, work_dir: Path, cfg: AutoFixConfig | None = None) -> dict[str, Any]:
    """Fix the formula errors of `source_path` on a fresh copy in `work_dir`.

    -> {status: clean | partial | stuck | unchanged | error, ran, message, summary,
        errors_before, errors_after, cells_rewritten, passes, fixes, left, applied,
        output_path, verified_opens_in_excel, recalc (the shape of app.recalc.recalculate), ...}
    The source is never modified. With status unchanged / stuck / error no new file is kept."""
    cfg = cfg or AutoFixConfig()
    if not com_available():
        return {"status": "error", "ran": False, "message": "Excel is not available here (pywin32): the automatic fixer needs a real Excel."}
    source_path = Path(source_path).resolve()
    t0 = time.time()
    copy_path, source_sha = make_immutable_copy(source_path, Path(work_dir))
    out: dict[str, Any] = {"status": "error", "ran": False}
    stopped: str | None = None
    try:
        with excel_session() as excel:
            def session() -> None:
                nonlocal stopped
                wb = None
                engine = None
                try:
                    wb = open_for_write(excel, copy_path)
                    calc_mode = None
                    try:
                        calc_mode = excel.Calculation
                        excel.Calculation = XL_MANUAL
                    except Exception:
                        calc_mode = None
                    for attr, value in (("ScreenUpdating", False), ("CalculationInterruptKey", 0)):
                        try:
                            setattr(excel, attr, value)
                        except Exception:
                            pass
                    engine = _Engine(excel, wb, cfg)
                    try:
                        status = engine.run()
                    except Stop as exc:
                        stopped = str(exc)
                        engine.note(f"stopped: {stopped}")
                        left = len(engine.genuine_errors()) + len(engine.const_err)
                        status = "clean" if left == 0 else ("partial" if engine.counts["accepted"] else "stuck")
                    if engine.counts["accepted"]:
                        engine.confirm()
                        left = len(engine.genuine_errors()) + len(engine.const_err)
                        status = "clean" if left == 0 else "partial"
                    out.update(engine.report())
                    out["status"] = status
                    out["ran"] = True
                    if engine.counts["accepted"]:
                        engine.tell("save", "Saving the fixed copy", None)
                        if calc_mode is not None:
                            try:
                                excel.Calculation = calc_mode
                            except Exception:
                                pass
                        save_in_place(wb, copy_path)
                        out["saved"] = True
                finally:
                    if engine is not None:
                        engine.release()
                    close_quietly(wb)
                    wb = None

            session()
    except Exception as exc:
        out.update({"status": "error", "ran": False, "message": f"The automatic fixer failed: {type(exc).__name__}: {str(exc)[:300]}"})
        return out
    out["seconds"] = round(time.time() - t0, 1)
    out["stopped"] = stopped
    out["summary"] = summary_sentence(out) + (f" (Stopped early: {stopped}.)" if stopped else "")
    out["message"] = out["summary"]
    remaining = out.pop("left_cells", [])
    gaps = out.pop("addin_gap_cells", [])
    out["recalc"] = {
        "status": "ERROR" if remaining else ("NOT_SUPPORTED" if gaps else "PASS"),
        "message": f"Recalculated; {len(remaining)} genuine formula error cell(s) found." if remaining else "Recalculated successfully; no formula errors found.",
        "formula_errors": remaining, "addin_gap_errors": gaps, "ran": True,
    }
    if not out.get("saved"):
        try:
            copy_path.unlink()
        except OSError:
            pass
        out["output_path"] = None
        return out
    verification = verify_opens_in_excel(copy_path)
    out["output_path"] = str(copy_path)
    out["verified_opens_in_excel"] = verification.get("opens")
    entry = {
        "rule_id": ["READY-001"], "action": "autofix", "actions": ["autofix"], "method": "excel_com",
        "source_path": str(source_path), "source_sha256": source_sha, "output_path": str(copy_path), "output_sha256": sha256_of(copy_path),
        "applied": out["applied"], "failed": [], "verified_opens_in_excel": verification.get("opens"), "warnings": [],
    }
    append_change_log(copy_path, entry)
    out["change_log_entry"] = entry
    return out


if __name__ == "__main__":  # python -m app.autofix <workbook> <work dir>
    import json

    def _print(stage: str, message: str, fraction: float | None = None, **facts: Any) -> None:
        if stage != "read":
            print(f"  [{stage}] {message}", flush=True)

    result = run_autofix(Path(sys.argv[1]), Path(sys.argv[2]), AutoFixConfig(progress=_print))
    brief = {k: v for k, v in result.items() if k not in ("applied", "fixes", "left", "recalc", "change_log_entry", "log")}
    print(json.dumps(brief, indent=1, ensure_ascii=False, default=str))
    print("--- log")
    print("\n".join(result.get("log", [])))
    print("--- fixes")
    for g in result.get("fixes", [])[:40]:
        print(f"{g['cells']:6d} {g['sheet']}!{g['cell']} {g['error']} [{g['strategy']}] {g['before'][:90]}  ->  {str(g['after'])[:110]}")
    print("--- left")
    for g in result.get("left", [])[:40]:
        print(f"{g['count']:6d} {g['sheet']}!{g['cell']} {g['error']} {g['formula'][:90]} :: {g['why'][:200]}")
