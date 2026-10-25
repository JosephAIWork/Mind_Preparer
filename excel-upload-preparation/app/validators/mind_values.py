"""Places where Mind computes another value than Excel.

The first two were found by uploading a client model that recalculates clean in
Excel and confirmed with a 49-cell test workbook in Mind:

FRM-008  arithmetic on a cell that holds "" (empty text), inside IFERROR.
         Excel: ""-1 is #VALUE!, IFERROR returns its fallback.
         Mind:  "" counts as 0, the arithmetic gives a number, IFERROR never fires.
         =IFERROR(1+INT((A10-1)/12),"") is blank in Excel and a value in Mind
         where A10 holds "". =IF(A10="","",IFERROR(...)) gives both the same.

FRM-009  a number compared with a cell that is empty in Excel.
         Excel: the empty cell is 0.  Mind: it is not (0=empty is false, 1<=empty is true).
         The empty cell is read directly (=IF(1<=B8,0,1)) or through a formula
         that lands on it and shows 0 in Excel: a plain reference (=+Sheet2!P4)
         or an exact VLOOKUP / HLOOKUP. A lookup that lands on an empty cell,
         then compared (IF(A35<=$H$9,0,1)), returns another value in Mind for
         every live row. N() around the reading makes it a 0 in both.

What the same test showed to be the SAME in both, and is therefore not flagged:
SUMIFS with ">="&DATE() criteria, N(""), MIN(90,""), ""<=number, cell="" on an
empty cell, and arithmetic on an empty cell.

Both rules read the values Excel stored in the file: a formula cell with no
stored value returned "" (openpyxl gives None for an empty stored string), so a
workbook that was never calculated and saved by Excel cannot be checked.
"""
from __future__ import annotations

import datetime
import os
import re
from typing import Any

from ..formula_utils import REF_RE, _arg_spans, _closing_paren, as_number, col_to_num, mask_strings, num_to_col, parse_ref, unquote
from ._common import finding

MAX_SITES = 80
ARITHMETIC = set("+-*/^")
COMPARISON = set("=<>")
TWO_CHAR = ("<=", ">=", "<>")
# functions that hand an error in an argument straight on: arithmetic on "" inside them is still an error in Excel
TRANSPARENT = {
    "INT", "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "ABS", "SQRT", "EXP", "LN", "LOG", "LOG10", "POWER", "MOD", "SIGN", "CEILING", "FLOOR", "MROUND",
    "MAX", "MIN", "SUM", "PRODUCT", "AVERAGE", "YEAR", "MONTH", "DAY", "DATE", "EDATE", "EOMONTH", "VALUE", "TEXT",
    "SUMPRODUCT",  # an error in its array arithmetic (SUMPRODUCT(range*range) with a "" cell) is the call's error
}
NUMERIC_CALLS = (TRANSPARENT - {"TEXT"}) | {"N", "LEN", "ROW", "COLUMN", "ROWS", "COLUMNS", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "SUMPRODUCT", "TODAY", "NOW"}
_CELL_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})")
_COORD_RE = re.compile(r"^([A-Za-z]{1,3})(\d{1,7})")
_NAME_BEFORE_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*$")
_IFERROR_RE = re.compile(r"(?<![A-Za-z0-9_.])IFERROR\s*\(", re.IGNORECASE)
_IF_RE = re.compile(r"(?<![A-Za-z0-9_.])IF\s*\(", re.IGNORECASE)
_IS_TEST_RE = re.compile(r"^(ISNUMBER|ISERROR|ISERR)\s*\(", re.IGNORECASE)
_LOOKUP_RE = re.compile(r"(VLOOKUP|HLOOKUP)\s*\(", re.IGNORECASE)
_EXACT_LOOKUP_RE = re.compile(r"(?<![A-Za-z0-9_.])(VLOOKUP|HLOOKUP|MATCH)\s*\(", re.IGNORECASE)
_CALL_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")

_SCANS: dict[str, dict[str, Any]] = {}  # one scan serves both rules and both Prep actions


def _is_formula(v: Any) -> bool:
    return isinstance(v, str) and len(v) > 1 and v.startswith("=")


def _is_error_value(v: Any) -> bool:
    return isinstance(v, str) and v.startswith("#") and v.rstrip("!?") in ("#REF", "#VALUE", "#DIV/0", "#N/A", "#NAME", "#NUM", "#NULL", "#SPILL", "#CALC", "#GETTING_DATA")


def _is_number(v: Any) -> bool:
    return (isinstance(v, (int, float)) and not isinstance(v, bool)) or isinstance(v, (datetime.date, datetime.time, datetime.timedelta))


def _same_key(a: Any, b: Any) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a == b
    if isinstance(a, str) and isinstance(b, str):
        return a.casefold() == b.casefold()
    return False


def _bare(body: str) -> str:
    """=+(X) -> X"""
    while True:
        body = body.strip()
        if body.startswith("+"):
            body = body[1:]
        elif body.startswith("(") and _closing_paren(mask_strings(body), 0) == len(body) - 1:
            body = body[1:-1]
        else:
            return body


def _one_cell(text: str) -> dict | None:
    r = parse_ref(text)
    if r and not r["whole_column"] and not r["whole_row"] and r["c1"] == r["c2"] and r["r1"] == r["r2"]:
        return r
    return None


class _Book:
    """What each cell holds (formula text or constant) and the value Excel stored for it."""

    def __init__(self, analysis: dict[str, Any]):
        from ..inventory import _values_workbook, ignored_sheet_names, open_cached

        wb0 = analysis["workbooks"][0]
        self.f = open_cached(analysis)
        self.v = _values_workbook(wb0["copy_path"], ignored_sheet_names(analysis))  # the same view REP-001 loads
        self.sheets = set(self.f.sheetnames) & set(self.v.sheetnames)
        self._fc: dict[str, dict] = {}
        self._vc: dict[str, dict] = {}
        self._land: dict[tuple[str, int, int], tuple[str, int, int] | None] = {}
        self._vectors: dict[tuple, tuple[bool, bool]] = {}
        # defined names that stand for a range (e.g. VLOOKUP(Key_name, Table_name, MATCH(B67, Headings_name, 0), FALSE))
        self.names: dict[str, dict] = {}
        for d in wb0.get("defined_names", []) or []:
            val = str(d.get("value") or "").strip()
            if d.get("name") and val and "#REF!" not in val and d.get("local_sheet_id") is None:
                r = parse_ref(val[1:] if val.startswith("=") else val)
                if r and r.get("sheet"):
                    self.names[str(d["name"]).upper()] = r

    def _named(self, text: str) -> dict | None:
        """The range a bare defined name stands for, else None."""
        t = text.strip()
        if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.]*", t):
            return None
        return self.names.get(t.upper())

    def _range_of(self, sheet: str, text: str) -> dict | None:
        r = parse_ref(text.strip())
        return r if r else self._named(text)

    def _match_index(self, sheet: str, text: str) -> int | None:
        """MATCH(key, one-line range, 0) worked out from the workbook's values: the position, else None."""
        m = re.match(r"\s*MATCH\s*\(", text, re.IGNORECASE)
        if not m:
            return None
        masked = mask_strings(text)
        open_at = m.end() - 1
        close = _closing_paren(masked, open_at)
        if close != len(text.rstrip()) - 1:
            return None
        spans = _arg_spans(masked, open_at + 1, close)
        if len(spans) != 3 or text[spans[2][0]:spans[2][1]].strip() != "0":
            return None
        key = self._scalar(sheet, text[spans[0][0]:spans[0][1]].strip())
        r = self._range_of(sheet, text[spans[1][0]:spans[1][1]])
        if key is None or r is None or r["whole_column"] or r["whole_row"]:
            return None
        tsheet = r["sheet"] or sheet
        if tsheet not in self.sheets:
            return None
        if r["r1"] == r["r2"]:
            cells = [(r["r1"], c) for c in range(r["c1"], r["c2"] + 1)]
        elif r["c1"] == r["c2"]:
            cells = [(rr, r["c1"]) for rr in range(r["r1"], r["r2"] + 1)]
        else:
            return None
        for i, (rr, cc) in enumerate(cells, start=1):
            if _same_key(self.value(tsheet, rr, cc), key):
                return i
        return None

    def _cells(self, wb, cache: dict[str, dict], sheet: str) -> dict:
        cells = cache.get(sheet)
        if cells is None:
            cells = cache[sheet] = wb[sheet]._cells if sheet in self.sheets else {}  # read without creating cells
        return cells

    def content(self, sheet: str, row: int, col: int) -> Any:
        c = self._cells(self.f, self._fc, sheet).get((row, col))
        v = c.value if c is not None else None
        if v is None or isinstance(v, (str, int, float, bool, datetime.date, datetime.time, datetime.timedelta)):
            return v
        text = getattr(v, "text", None)  # an array formula
        if isinstance(text, str):
            return text if text.startswith("=") else "=" + text
        return str(v)

    def stored(self, sheet: str, row: int, col: int) -> Any:
        c = self._cells(self.v, self._vc, sheet).get((row, col))
        return c.value if c is not None else None

    def value(self, sheet: str, row: int, col: int) -> Any:
        v = self.content(sheet, row, col)
        return self.stored(sheet, row, col) if _is_formula(v) else v

    def size(self, sheet: str) -> tuple[int, int]:
        ws = self.f[sheet]
        return ws.max_row, ws.max_column

    def lookup_vector(self, sheet: str, r: dict, name: str) -> tuple[bool, bool] | None:
        """For the column (VLOOKUP) or row (HLOOKUP) a lookup searches, or a one-line MATCH range:
        (holds an empty cell, holds a formula that returned "") within the sheet's used area."""
        tsheet = r["sheet"] or sheet
        if tsheet not in self.sheets:
            return None
        if name == "VLOOKUP" or (name == "MATCH" and r["c1"] == r["c2"]):
            cells = [(row, r["c1"]) for row in range(r["r1"], min(r["r2"], self.size(tsheet)[0]) + 1)]
        elif name == "HLOOKUP" or (name == "MATCH" and r["r1"] == r["r2"]):
            cells = [(r["r1"], col) for col in range(r["c1"], min(r["c2"], self.size(tsheet)[1]) + 1)]
        else:
            return None
        key = (tsheet, r["ref"], name == "HLOOKUP" or (name == "MATCH" and r["r1"] == r["r2"]))
        hit = self._vectors.get(key)
        if hit is None:
            empty = blank = False
            for row, col in cells:
                v = self.content(tsheet, row, col)
                if v is None:
                    empty = True
                elif _is_formula(v) and self.stored(tsheet, row, col) is None:
                    blank = True
            hit = self._vectors[key] = (empty, blank)
        return hit

    # --- a formula that lands on an empty cell --------------------------------------------
    def landing(self, sheet: str, row: int, col: int) -> tuple[str, int, int] | None:
        """The empty cell this formula cell hands on (Excel shows 0), else None."""
        key = (sheet, row, col)
        if key in self._land:
            return self._land[key]
        self._land[key] = None  # a cycle ends here
        v = self.content(sheet, row, col)
        out = None
        s = self.stored(sheet, row, col)
        if _is_formula(v) and _is_number(s) and s == 0:
            out = self._follow(sheet, v, 0)
        self._land[key] = out
        return out

    def _follow(self, sheet: str, formula: str, depth: int) -> tuple[str, int, int] | None:
        if depth > 8:
            return None
        body = _bare(formula[1:])
        r = _one_cell(body)
        if r is None:
            n = self._named(body)
            if n and n["r1"] == n["r2"] and n["c1"] == n["c2"] and not n["whole_column"] and not n["whole_row"]:
                r = n
        target = (r["sheet"] or sheet, r["r1"], r["c1"]) if r else self._lookup(sheet, body)
        if target is None or target[0] not in self.sheets:
            return None
        v = self.content(*target)
        if v is None:
            return target
        return self._follow(target[0], v, depth + 1) if _is_formula(v) else None

    def _scalar(self, sheet: str, text: str) -> Any:
        n = as_number(text)
        if n is not None:
            return n
        s = unquote(text)
        if s is not None:
            return s
        r = _one_cell(text)
        if r is None:
            n = self._named(text)
            if n and n["r1"] == n["r2"] and n["c1"] == n["c2"] and not n["whole_column"] and not n["whole_row"]:
                r = n
        return self.value(r["sheet"] or sheet, r["r1"], r["c1"]) if r else None

    def _lookup(self, sheet: str, body: str) -> tuple[str, int, int] | None:
        """The cell an exact VLOOKUP / HLOOKUP returns, worked out from the workbook's values."""
        m = _LOOKUP_RE.match(body)
        if not m:
            return None
        masked = mask_strings(body)
        open_at = m.end() - 1
        close = _closing_paren(masked, open_at)
        if close != len(body) - 1:
            return None
        spans = _arg_spans(masked, open_at + 1, close)
        if len(spans) != 4:
            return None  # no 4th argument: an approximate match
        args = [body[a:b].strip() for a, b in spans]
        idx = as_number(args[2])
        if idx is None:
            idx = self._match_index(sheet, args[2])  # MATCH(heading, headings, 0) as the column index
        table = self._range_of(sheet, args[1])  # a defined name as the table
        if args[3].upper() not in ("0", "FALSE") or idx is None or idx != int(idx) or idx < 1 or table is None:
            return None
        tsheet = table["sheet"] or sheet
        key = self._scalar(sheet, args[0])
        if tsheet not in self.sheets or key is None:
            return None
        idx = int(idx)
        max_row, max_col = self.size(tsheet)
        if m.group(1).upper() == "VLOOKUP":
            if idx > table["c2"] - table["c1"] + 1:
                return None
            for row in range(table["r1"], min(table["r2"], max_row) + 1):
                if _same_key(self.value(tsheet, row, table["c1"]), key):
                    return tsheet, row, table["c1"] + idx - 1
        else:
            if idx > table["r2"] - table["r1"] + 1:
                return None
            for col in range(table["c1"], min(table["c2"], max_col) + 1):
                if _same_key(self.value(tsheet, table["r1"], col), key):
                    return tsheet, table["r1"] + idx - 1, col
        return None


# --- reading one formula ----------------------------------------------------------------------
def _around(masked: str, s: int, e: int) -> tuple[int, int]:
    """Index of the first non-blank character before `s` (-1 if none) and after `e` (len if none)."""
    i = s - 1
    while i >= 0 and masked[i].isspace():
        i -= 1
    j, n = e, len(masked)
    while j < n and masked[j].isspace():
        j += 1
    return i, j


def _operand_end(masked: str, k: int) -> int:
    depth = 0
    for i in range(k, len(masked)):
        ch = masked[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0:
                return i
            depth -= 1
        elif depth == 0 and (ch in ",;" or ch in COMPARISON):
            return i
    return len(masked)


def _operand_start(masked: str, a: int) -> int:
    depth = 0
    for i in range(a - 1, 0, -1):
        ch = masked[i]
        if ch == ")":
            depth += 1
        elif ch == "(":
            if depth == 0:
                return i + 1
            depth -= 1
        elif depth == 0 and (ch in ",;" or ch in COMPARISON):
            return i + 1
    return 1


def _other_side(masked: str, s: int, e: int) -> tuple[int, int] | None:
    """If the reference at [s, e) is one whole side of a comparison, the span of the other side."""
    i, j = _around(masked, s, e)
    n = len(masked)
    prev = masked[i] if i >= 0 else ""
    nxt = masked[j] if j < n else ""
    if nxt and nxt in COMPARISON and (i == 0 or (prev and prev in "(,;") or (prev == "+" and i == 1)):
        k = j + (2 if masked[j:j + 2] in TWO_CHAR else 1)
        return k, _operand_end(masked, k)
    if prev and prev in COMPARISON and i > 0 and (j >= n or nxt in "),;"):
        a = i - 1 if (i - 1 > 0 and masked[i - 1:i + 1] in TWO_CHAR) else i
        return _operand_start(masked, a), a
    return None


def _numeric_side(book: _Book, text: str, masked: str, sheet: str) -> bool:
    """Is the other side of the comparison a number (as far as can be told without computing it)?"""
    t = text.strip()
    if not t or '"' in t or "&" in masked:
        return False
    if as_number(t) is not None:
        return True
    if t.upper() in ("TRUE", "FALSE"):
        return False
    r = _one_cell(t)
    if r:
        return _is_number(book.value(r["sheet"] or sheet, r["r1"], r["c1"]))
    if parse_ref(t):
        return False
    if any(ch in masked for ch in ARITHMETIC):
        return True
    m = _CALL_RE.match(masked.strip())
    return bool(m and m.group(1).upper() in NUMERIC_CALLS)


def _ref_sheet(formula: str, m: re.Match, home: str) -> str:
    sheet = m.group("sheet")
    if not sheet:
        return home
    if sheet.startswith("'"):
        return formula[m.start("sheet") + 1:m.end("sheet") - 1].replace("''", "'")
    return sheet


def _ref_cell(m: re.Match) -> tuple[int, int]:
    mm = _CELL_RE.fullmatch(m.group("ref"))
    return int(mm.group(4)), col_to_num(mm.group(2))


def _is_arithmetic_operand(masked: str, s: int, e: int, lo: int, hi: int) -> bool:
    i, j = _around(masked, s, e)
    return (i >= lo and masked[i] in ARITHMETIC) or (j < hi and (masked[j] in ARITHMETIC or masked[j] == "%"))


def _reaches_unconditionally(masked: str, pos: int, call_open: int) -> bool:
    """Only parentheses and error-transparent functions between `pos` and the IFERROR that encloses it."""
    depth = 0
    for i in range(pos - 1, call_open - 1, -1):
        ch = masked[i]
        if ch == ")":
            depth += 1
        elif ch == "(":
            if depth:
                depth -= 1
                continue
            if i == call_open:
                return True
            m = _NAME_BEFORE_RE.search(masked[:i])
            if m and m.group(1).upper() not in TRANSPARENT:
                return False
    return False


def _tests_for_empty_text(compact: str, ref_compact: str) -> bool:
    r = re.escape(ref_compact)
    return bool(re.search(r'(?<![A-Z0-9_.!$])' + r + r'(=|<>)""|""(=|<>)' + r + r'(?![0-9])|(?<![A-Z0-9_.])(ISNUMBER|ISTEXT|ISBLANK|LEN|N)\(' + r + r'\)', compact))


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text).upper().replace("$", "")


def _iferror_calls(masked: str) -> list[tuple[int, int, int, list[tuple[int, int]]]]:
    out = []
    for m in _IFERROR_RE.finditer(masked):
        open_at = m.end() - 1
        close = _closing_paren(masked, open_at)
        if close == -1:
            continue
        out.append((m.start(), open_at, close, _arg_spans(masked, open_at + 1, close)))
    return out


def _guard_calls(masked: str) -> list[tuple[int, int, int, list[tuple[int, int]]]]:
    """Every construct that hides an error of X behind a fallback: IFERROR(X, fb), and
    IF(ISNUMBER(X), ..., fb) / IF(ISERROR(X), fb, ...) / IF(ISERR(X), fb, ...). Each as
    (start, open paren of X's call, close of the whole construct, [X span, fallback span]):
    Excel takes the fallback when X errors on a "" cell, Mind computes X with "" as 0
    (e.g. IF(ISNUMBER(Sheet2!H171*NOT(Flag_name)),...,"") where H171 holds "")."""
    out = _iferror_calls(masked)
    for m in _IF_RE.finditer(masked):
        open_at = m.end() - 1
        close = _closing_paren(masked, open_at)
        if close == -1:
            continue
        spans = _arg_spans(masked, open_at + 1, close)
        if len(spans) != 3:
            continue
        a0, a1 = spans[0]
        test = masked[a0:a1]
        lead = len(test) - len(test.lstrip())
        test_s = test.strip()
        mt = _IS_TEST_RE.match(test_s)
        if not mt:
            continue
        t_open = a0 + lead + mt.end() - 1
        if _closing_paren(masked, t_open) != a0 + lead + len(test_s) - 1:
            continue  # the test is more than the one IS... call
        fallback = spans[2] if mt.group(1).upper() == "ISNUMBER" else spans[1]
        out.append((m.start(), t_open, close, [(t_open + 1, a0 + lead + len(test_s) - 1), fallback]))
    out.sort(key=lambda c: c[0])
    return out


def _has_guard_form(formula: str) -> bool:
    upper = formula.upper()
    return "IFERROR" in upper or (("ISNUMBER" in upper or "ISERR" in upper) and re.search(r"(?<![A-Z0-9_.])IF\s*\(", upper) is not None)


def _relative_form(formula: str, masked: str, refs: list[re.Match], row: int, col: int) -> str:
    """The formula with every reference written relative to its own cell: equal for a filled-down block."""
    def rel(mm: re.Match) -> str:
        c_abs, c, r_abs, r = mm.groups()
        return (f"C{col_to_num(c)}" if c_abs else f"C[{col_to_num(c) - col}]") + (f"R{r}" if r_abs else f"R[{int(r) - row}]")

    out, last = [], 0
    for m in refs:
        out.append(formula[last:m.start("ref")])
        out.append(_CELL_RE.sub(rel, m.group("ref")).upper())
        last = m.end()
    out.append(formula[last:])
    return "".join(out)


def _guarded(formula: str, refs: list[re.Match], calls: list, guards: dict[int, set[int]], extra: dict[int, list[str]] | None = None) -> str:
    """IFERROR(X, fallback) -> IF(cell="", fallback, IFERROR(X, fallback)) for each call in `guards`;
    `extra` adds tests written as text (the "" cells found inside a range operand)."""
    keys = set(guards) | set(extra or {})
    for ci in sorted(keys, key=lambda i: calls[i][0], reverse=True):
        start, _open, close, spans = calls[ci]
        tests = []
        for k in sorted(guards.get(ci, ())):
            t = formula[refs[k].start():refs[k].end()] + '=""'
            if t not in tests:
                tests.append(t)
        for t in (extra or {}).get(ci, []):
            if t not in tests:
                tests.append(t)
        test = tests[0] if len(tests) == 1 else "OR(" + ",".join(tests) + ")"
        fallback = formula[spans[1][0]:spans[1][1]].strip()
        formula = formula[:start] + f"IF({test},{fallback},{formula[start:close + 1]})" + formula[close + 1:]
    return formula


def _fallback_fired(book: _Book, formula: str, calls: list, ci: int, sheet: str, row: int, col: int) -> bool:
    """When the IFERROR is the whole formula, the cell must show the fallback in Excel."""
    start, _open, close, spans = calls[ci]
    if _bare(formula[1:]) != formula[start:close + 1]:
        return True  # inside a larger formula: nothing to read off the cell
    fallback = formula[spans[1][0]:spans[1][1]].strip()
    stored = book.stored(sheet, row, col)
    if fallback == '""':
        return stored is None
    n = as_number(fallback)
    if n is not None:
        return _is_number(stored) and abs(stored - n) < 1e-9
    return True


# --- the scan ---------------------------------------------------------------------------------
def scan(analysis: dict[str, Any]) -> dict[str, Any]:
    wb0 = analysis["workbooks"][0]
    path = str(wb0["copy_path"])
    try:
        stamp = os.stat(path).st_mtime_ns
    except OSError:
        stamp = 0
    key = f"{path}|{stamp}|{len(wb0.get('formulas', []))}|{sorted(wb0.get('ignored_sheets') or [])}"
    hit = _SCANS.get(key)
    if hit is None:
        hit = _scan(analysis)
        if len(_SCANS) >= 4:
            _SCANS.clear()
        _SCANS[key] = hit
    return hit


def _scan(analysis: dict[str, Any]) -> dict[str, Any]:
    book = _Book(analysis)
    wb0 = analysis["workbooks"][0]
    formulas = []
    array_cells: set[tuple[str, int, int]] = set()
    for f in wb0.get("formulas", []):
        m = _COORD_RE.match(str(f.get("cell") or ""))
        if m and isinstance(f.get("formula"), str) and not f.get("data_table") and f["sheet"] in book.sheets:
            formulas.append((f, int(m.group(2)), col_to_num(m.group(1))))
            if f.get("array_ref"):
                array_cells.add((f["sheet"], int(m.group(2)), col_to_num(m.group(1))))
    out: dict[str, Any] = {"checked": True, "formulas": len(formulas), "blank_arithmetic": [], "origins": [], "direct": [], "readers": 0, "blank_errors": []}
    if not any(book.stored(f["sheet"], row, col) is not None for f, row, col in formulas):
        out["checked"] = False
        return out

    blank_flagged: dict[tuple[str, int, int], dict[int, set[int]]] = {}
    range_extra: dict[tuple[str, int, int], dict[int, list[str]]] = {}
    arithmetic_marks: dict[tuple[str, int, int], set[int]] = {}
    lookup_marks: dict[tuple[str, int, int], set[int]] = {}
    iferror_by_column: dict[tuple[str, int], list[tuple[dict, int, int]]] = {}
    origins: dict[tuple[str, int, int], dict[str, Any]] = {}
    for f, row, col in formulas:
        formula, sheet = f["formula"], f["sheet"]
        rest = formula[1:]
        compares = "=" in rest or "<" in rest or ">" in rest
        has_iferror = _has_guard_form(formula)
        if not (compares or has_iferror or book.stored(sheet, row, col) == "#VALUE!"):
            continue
        masked = mask_strings(formula) if ('"' in formula or "'" in formula) else formula
        refs = [m for m in REF_RE.finditer(masked) if not (m.start() and masked[m.start() - 1] == "]")]  # not another workbook's cell

        compact = None
        if compares:
            spans = []
            for m in refs:
                if ":" in m.group("ref"):
                    continue
                other = _other_side(masked, m.start(), m.end())
                if other is None:
                    continue
                tsheet = _ref_sheet(formula, m, sheet)
                if tsheet not in book.sheets:
                    continue
                trow, tcol = _ref_cell(m)
                held = book.content(tsheet, trow, tcol)
                land = None
                if held is not None:
                    land = book.landing(tsheet, trow, tcol) if _is_formula(held) else None
                    if land is None:
                        continue
                else:
                    # =IF(B5="","",IF(B5>0,...)): the formula asks first whether the cell is empty -- both take the same branch
                    if compact is None:
                        compact = _compact(masked)
                    if _tests_for_empty_text(compact, _compact(masked[m.start():m.end()])):
                        continue
                if not _numeric_side(book, formula[other[0]:other[1]], masked[other[0]:other[1]], sheet):
                    continue
                out["readers"] += 1
                if land is None:
                    spans.append((m.start(), m.end(), f"{tsheet}!{num_to_col(tcol)}{trow}"))
                else:
                    o = origins.setdefault((tsheet, trow, tcol), {"sheet": tsheet, "cell": f"{num_to_col(tcol)}{trow}", "formula": held, "lands_on": f"{land[0]}!{num_to_col(land[2])}{land[1]}", "readers": 0, "first_reader": None})
                    o["readers"] += 1
                    if o["first_reader"] is None:
                        o["first_reader"] = {"sheet": sheet, "cell": f["cell"], "formula": formula[:200]}
            if spans:
                fixed = formula
                for s, e, _ in sorted(spans, reverse=True):
                    fixed = fixed[:s] + "N(" + fixed[s:e] + ")" + fixed[e:]
                array = bool(f.get("array_ref"))
                out["direct"].append({
                    "sheet": sheet, "cell": f["cell"], "formula": formula, "empty_cells": sorted({c for _, _, c in spans}),
                    "suggested_formula": None if array else fixed, "issue": "an array formula: rewrite it by hand" if array else None,
                })

        own = book.stored(sheet, row, col)
        if _is_error_value(own) and own in ("#VALUE!",) and any(ch in rest for ch in "+-*/^"):
            # FRM-013: the cell shows #VALUE! -- which of its arithmetic operands holds ""?
            blanks = []
            for k, m in enumerate(refs):
                if ":" in m.group("ref") or not _is_arithmetic_operand(masked, m.start(), m.end(), 1, len(masked)):
                    continue
                tsheet = _ref_sheet(formula, m, sheet)
                trow, tcol = _ref_cell(m)
                if _is_formula(book.content(tsheet, trow, tcol)) and book.stored(tsheet, trow, tcol) is None:
                    if _reaches_unconditionally(masked, m.start(), -1) or True:
                        blanks.append(k)
            if blanks and not _IFERROR_RE.search(masked):
                tests = []
                for k in blanks:
                    t = formula[refs[k].start():refs[k].end()] + '=""'
                    if t not in tests:
                        tests.append(t)
                test = tests[0] if len(tests) == 1 else "OR(" + ",".join(tests) + ")"
                array = bool(f.get("array_ref"))
                out["blank_errors"].append({"sheet": sheet, "cell": f["cell"], "formula": formula, "blank_cells": [formula[refs[k].start():refs[k].end()] for k in blanks],
                                            "suggested_formula": None if array else f"=IF({test},NA(),{formula[1:]})", "issue": "an array formula: rewrite it by hand" if array else None})
        if has_iferror:
            iferror_by_column.setdefault((sheet, col), []).append((f, row, col))
            calls = _guard_calls(masked)
            guards: dict[int, set[int]] = {}
            range_blanks: dict[int, list[str]] = {}
            for ci, (_start, open_at, _close, spans2) in enumerate(calls):
                if len(spans2) != 2:
                    continue
                lo, hi = spans2[0]
                for k, m in enumerate(refs):
                    if m.start() < lo or m.end() > hi or not _is_arithmetic_operand(masked, m.start(), m.end(), lo, hi):
                        continue
                    if ":" in m.group("ref"):
                        # a range in array arithmetic (SUMPRODUCT(I53:DL53*...)): its "" cells, named one by one
                        r = parse_ref(formula[m.start():m.end()])
                        if not r or r["whole_column"] or r["whole_row"] or (r["r2"] - r["r1"] + 1) * (r["c2"] - r["c1"] + 1) > 600:
                            continue
                        tsheet = r["sheet"] or sheet
                        if tsheet not in book.sheets or not _reaches_unconditionally(masked, m.start(), open_at):
                            continue
                        prefix = formula[m.start():m.start("ref")]  # the sheet part, quoted as written
                        if compact is None:
                            compact = _compact(masked)
                        for rr in range(r["r1"], r["r2"] + 1):
                            for cc in range(r["c1"], r["c2"] + 1):
                                if _is_formula(book.content(tsheet, rr, cc)) and book.stored(tsheet, rr, cc) is None:
                                    text = f"{prefix}{num_to_col(cc)}{rr}"
                                    if _tests_for_empty_text(compact, _compact(text)):
                                        continue  # already guarded (a second Prep round)
                                    range_blanks.setdefault(ci, []).append(text + '=""')
                        continue
                    tsheet = _ref_sheet(formula, m, sheet)
                    trow, tcol = _ref_cell(m)
                    if not _is_formula(book.content(tsheet, trow, tcol)) or book.stored(tsheet, trow, tcol) is not None:
                        continue  # only a formula cell that returned ""
                    if not _reaches_unconditionally(masked, m.start(), open_at):
                        continue
                    if compact is None:
                        compact = _compact(masked)
                    if _tests_for_empty_text(compact, _compact(masked[m.start():m.end()])):
                        continue
                    guards.setdefault(ci, set()).add(k)
                    arithmetic_marks.setdefault((sheet, row, col), set()).add(k)
                # FRM-010: an exact lookup whose key is a cell that returned "", over a column / row with an empty cell
                for m2 in _EXACT_LOOKUP_RE.finditer(masked, lo, hi):
                    open2 = m2.end() - 1
                    close2 = _closing_paren(masked, open2)
                    if close2 == -1 or close2 > hi:
                        continue
                    spans3 = _arg_spans(masked, open2 + 1, close2)
                    name = m2.group(1).upper()
                    exact = (name != "MATCH" and len(spans3) == 4 and formula[spans3[3][0]:spans3[3][1]].strip().upper() in ("0", "FALSE")) or (name == "MATCH" and len(spans3) == 3 and formula[spans3[2][0]:spans3[2][1]].strip() == "0")
                    if not exact:
                        continue
                    key_text = formula[spans3[0][0]:spans3[0][1]].strip()
                    k = next((i for i, m in enumerate(refs) if spans3[0][0] <= m.start() and m.end() <= spans3[0][1] and formula[m.start():m.end()] == key_text), None)
                    if k is None or ":" in refs[k].group("ref"):
                        continue
                    tsheet = _ref_sheet(formula, refs[k], sheet)
                    trow, tcol = _ref_cell(refs[k])
                    if not _is_formula(book.content(tsheet, trow, tcol)) or book.stored(tsheet, trow, tcol) is not None:
                        continue  # the key is a formula cell that returned ""
                    r = parse_ref(formula[spans3[1][0]:spans3[1][1]].strip())
                    vec = book.lookup_vector(sheet, r, name) if r else None
                    if vec is None or not vec[0] or vec[1]:
                        continue  # no empty cell to match -- or a "" that Excel would match too
                    if not _reaches_unconditionally(masked, m2.start(), open_at):
                        continue
                    if compact is None:
                        compact = _compact(masked)
                    if _tests_for_empty_text(compact, _compact(masked[refs[k].start():refs[k].end()])):
                        continue
                    guards.setdefault(ci, set()).add(k)
                    lookup_marks.setdefault((sheet, row, col), set()).add(k)
                if (ci in guards or ci in range_blanks) and not _fallback_fired(book, formula, calls, ci, sheet, row, col):
                    guards.pop(ci, None)
                    range_blanks.pop(ci, None)
            if range_blanks and not guards:
                # a site of its own (no filled-down-block treatment: each row names its own "" cells)
                array = bool(f.get("array_ref"))
                cells = sorted({t[:-3] for ts in range_blanks.values() for t in ts})
                out["blank_arithmetic"].append({
                    "sheet": sheet, "cell": f["cell"], "formula": formula, "blank_cells": cells, "blank_now": True, "arithmetic": True, "lookup_key": False,
                    "suggested_formula": None if array else _guarded(formula, refs, calls, {}, range_blanks), "issue": "an array formula: rewrite it by hand" if array else None,
                })
            elif guards:
                if range_blanks:
                    range_extra[(sheet, row, col)] = range_blanks
                blank_flagged[(sheet, row, col)] = guards

    # FRM-008: the whole filled-down block gets the same test, so a column keeps one formula
    patterns: dict[tuple[str, int, str], dict[int, set[int]]] = {}
    pattern_kinds: dict[tuple[str, int, str], set[str]] = {}
    parsed: dict[tuple[str, int, int], tuple] = {}
    for (sheet, col) in {(s, c) for s, _, c in blank_flagged}:
        for f, row, _ in iferror_by_column.get((sheet, col), []):
            formula = f["formula"]
            masked = mask_strings(formula) if ('"' in formula or "'" in formula) else formula
            refs = [m for m in REF_RE.finditer(masked) if not (m.start() and masked[m.start() - 1] == "]")]
            rel = _relative_form(formula, masked, refs, row, col)
            parsed[(sheet, row, col)] = (f, masked, refs, rel)
            g = blank_flagged.get((sheet, row, col))
            if g:
                union = patterns.setdefault((sheet, col, rel), {})
                for ci, ks in g.items():
                    union.setdefault(ci, set()).update(ks)
                kinds = pattern_kinds.setdefault((sheet, col, rel), set())
                kinds.update(["arithmetic"] if (sheet, row, col) in arithmetic_marks else [])
                kinds.update(["lookup"] if (sheet, row, col) in lookup_marks else [])
    for (sheet, row, col), (f, masked, refs, rel) in sorted(parsed.items(), key=lambda kv: (kv[0][0], kv[0][2], kv[0][1])):
        union = patterns.get((sheet, col, rel))
        if not union:
            continue
        formula = f["formula"]
        calls = _guard_calls(masked)
        array = bool(f.get("array_ref"))
        cells = sorted({formula[refs[k].start():refs[k].end()] for ks in union.values() for k in ks})
        kinds = pattern_kinds.get((sheet, col, rel), set())
        # A cell the new test reads must not hold an error: IF(#REF!="", ...) is #REF!, where the
        # IFERROR it replaces returned the fallback (seen on a client model: 19 values changed, 9 into errors).
        # Such a row keeps its formula -- Excel and Mind both stop on the error and fall back.
        errors = sorted({
            formula[refs[k].start():refs[k].end()]
            for ks in union.values() for k in ks
            if _is_error_value(book.value(_ref_sheet(formula, refs[k], sheet), *_ref_cell(refs[k])))
        })
        issue = "an array formula: rewrite it by hand" if array else (f"{', '.join(errors[:3])} holds an error: the IFERROR already returns its fallback in Excel and in Mind; left as it is" if errors else None)
        out["blank_arithmetic"].append({
            "sheet": sheet, "cell": f["cell"], "formula": formula, "blank_cells": cells, "blank_now": (sheet, row, col) in blank_flagged,
            "arithmetic": "arithmetic" in kinds, "lookup_key": "lookup" in kinds,
            "suggested_formula": None if issue else _guarded(formula, refs, calls, union, range_extra.get((sheet, row, col))), "issue": issue,
        })

    # every formula cell that lands on an empty cell and is read by some formula -- AVERAGE,
    # COUNT, MIN and MAX skip an empty where Excel saw a 0
    try:
        from ..prep import reference_index, referenced_by
        index = reference_index(analysis)
    except Exception:
        index = None
    if index is not None:
        for f, row, col in formulas:
            key = (f["sheet"], row, col)
            if key in origins:
                continue
            held = f["formula"]
            if not (_bare(held[1:])[:1] in "'[ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz$" or _LOOKUP_RE.match(_bare(held[1:]))):
                continue
            if book.landing(f["sheet"], row, col) is None:
                continue
            owner = referenced_by(index, f["sheet"], row, col)
            if not owner:
                continue
            land = book.landing(f["sheet"], row, col)
            origins[key] = {"sheet": f["sheet"], "cell": f["cell"], "formula": held, "lands_on": f"{land[0]}!{num_to_col(land[2])}{land[1]}", "readers": 1, "first_reader": {"sheet": owner.split("!")[0], "cell": owner.split("!")[-1], "formula": "(reads it: AVERAGE, COUNT, MIN, MAX skip an empty cell where Excel saw 0)"}}
    for key in sorted(origins):
        o = origins[key]
        array = key in array_cells
        body = _bare(o["formula"][1:])
        o["suggested_formula"] = None if array else f"=N({body})"
        o["issue"] = "an array formula: rewrite it by hand" if array else None
        out["origins"].append(o)
    return out


def _short(formula: str | None, n: int = 110) -> str:
    text = re.sub(r"\s+", " ", formula or "")
    return text if len(text) <= n else text[: n - 1] + "…"


def _not_checked(exc: Exception | None) -> dict:
    if exc is not None:
        return finding("NOT_SUPPORTED", f"The stored values of the workbook could not be read ({type(exc).__name__}: {str(exc)[:120]}): not checked.", {"sites": []})
    return finding("NOT_SUPPORTED", "The workbook holds no stored values (it was never calculated and saved by Excel), so it cannot be told which cells are empty text or land on an empty cell: not checked. Open it in Excel, save it and upload it again.", {"sites": []})


def blank_text_arithmetic(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = MAX_SITES) -> dict:
    """FRM-008: arithmetic on a cell that holds "" inside IFERROR -- a fallback in Excel, a number in Mind."""
    try:
        res = scan(analysis)
    except Exception as exc:
        return _not_checked(exc)
    if not res["checked"]:
        return _not_checked(None)
    sites = [s for s in res["blank_arithmetic"] if s["arithmetic"]]
    observed, first, now, with_fix = _blank_text_observed(sites, limit)
    if not sites:
        return finding("PASS", 'No IFERROR (or IF on ISNUMBER / ISERROR) hides arithmetic on a cell that holds "" (empty text).', observed)
    return finding(
        "ERROR",
        f'{now} formula(s) do arithmetic on a cell that holds "" (empty text) inside IFERROR, or inside an IF that tests ISNUMBER / ISERROR on it. Excel stops on the "" (#VALUE!) and returns the fallback; '
        'Mind counts "" as 0, computes a number and never reaches the fallback, so these cells get another value in Mind (confirmed in a real Mind conversion: '
        "1,242 cells blank in Excel held a value). "
        f"First: {first['sheet']}!{first['cell']} = {_short(first['formula'])}"
        + (f" -> proposed {_short(first['suggested_formula'], 150)}" if first["suggested_formula"] else "")
        + f". The Prep action '{BLANK_TEXT_ACTION}' rewrites {with_fix} formula(s) -- these and the rest of their filled-down blocks, "
        "so each column keeps one formula; Excel's results stay the same"
        + (f"; {len(sites) - with_fix} left as they are (an array formula, or a cell the test would read holds an error -- see the Prep action's skip list)." if len(sites) - with_fix else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )


BLANK_TEXT_ACTION = "Test for empty text before the arithmetic or lookup"


def _blank_text_observed(sites: list[dict], limit: int | None) -> tuple[dict, dict | None, int, int]:
    now = sum(1 for s in sites if s["blank_now"])
    with_fix = sum(1 for s in sites if s["suggested_formula"])
    shown = sorted(sites, key=lambda s: not s["blank_now"])
    observed = {
        "sites": [{**s, "formula": s["formula"][:300]} for s in (shown if limit is None else shown[:limit])],
        "sites_count": len(sites), "blank_now": now, "with_proposal": with_fix, "without_proposal": len(sites) - with_fix,
    }
    return observed, (shown[0] if shown else None), now, with_fix


def blank_text_lookup_key(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = MAX_SITES) -> dict:
    """FRM-010: an exact lookup whose key is a cell that holds "", inside IFERROR -- Excel finds nothing and
    returns the fallback; Mind matches an empty cell of the looked-up column (the header of a titled table)."""
    try:
        res = scan(analysis)
    except Exception as exc:
        return _not_checked(exc)
    if not res["checked"]:
        return _not_checked(None)
    sites = [s for s in res["blank_arithmetic"] if s["lookup_key"]]
    observed, first, now, with_fix = _blank_text_observed(sites, limit)
    if not sites:
        return finding("PASS", 'No IFERROR hides an exact lookup (VLOOKUP, HLOOKUP, MATCH) whose key is a cell that holds "" over a column or row with an empty cell.', observed)
    return finding(
        "ERROR",
        f'{now} formula(s) look up (VLOOKUP, HLOOKUP or MATCH, exact match) a key that holds "" (empty text) in a column or row that has an empty cell, inside IFERROR. '
        'Excel finds nothing and returns the IFERROR fallback; Mind matches the empty cell -- the header cell of a table Prep gave a title -- and returns that row '
        "(confirmed in a real Mind conversion: =1+(<a flag>=1)*IFERROR(VLOOKUP(<a cell holding \"\">,<a titled table>,2,0),...) came back NaN on 621 rows, the lookup returned the header text). "
        f"First: {first['sheet']}!{first['cell']} = {_short(first['formula'])}"
        + (f" -> proposed {_short(first['suggested_formula'], 150)}" if first["suggested_formula"] else "")
        + f". The Prep action '{BLANK_TEXT_ACTION}' rewrites {with_fix} formula(s) -- these and the rest of their filled-down blocks; Excel's results stay the same"
        + (f"; {len(sites) - with_fix} left as they are (an array formula, or a cell the test would read holds an error -- see the Prep action's skip list)." if len(sites) - with_fix else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )


def empty_cell_compared(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = MAX_SITES) -> dict:
    """FRM-009: a number compared with a cell that is empty in Excel -- 0 in Excel, not 0 in Mind."""
    try:
        res = scan(analysis)
    except Exception as exc:
        return _not_checked(exc)
    if not res["checked"]:
        return _not_checked(None)
    origins, direct = res["origins"], res["direct"]
    sites = [{"kind": "lands_on_empty", **o, "formula": o["formula"][:300]} for o in origins] + [{"kind": "reads_empty", **d, "formula": d["formula"][:300]} for d in direct]
    with_fix = sum(1 for s in sites if s["suggested_formula"])
    observed = {
        "sites": sites if limit is None else sites[:limit], "sites_count": len(sites), "landing_formulas": len(origins), "formulas_reading_an_empty_cell": len(direct),
        "comparisons": res["readers"], "with_proposal": with_fix, "without_proposal": len(sites) - with_fix,
    }
    if not sites:
        return finding("PASS", "No formula compares a number with a cell that is empty (directly, or through a reference or lookup that lands on an empty cell).", observed)
    parts = []
    if origins:
        o = origins[0]
        reader = o["first_reader"] or {}
        parts.append(
            f"{len(origins)} formula cell(s) land on an empty cell and are then compared with a number in {sum(x['readers'] for x in origins)} place(s). "
            f"First: {o['sheet']}!{o['cell']} = {_short(o['formula'], 80)} lands on {o['lands_on']} (empty), read by {reader.get('sheet')}!{reader.get('cell')} = {_short(reader.get('formula'), 80)}"
            + (f" -> proposed {o['sheet']}!{o['cell']} {_short(o['suggested_formula'], 90)}" if o["suggested_formula"] else "")
        )
    if direct:
        d = direct[0]
        parts.append(
            f"{len(direct)} formula(s) compare an empty cell with a number directly. First: {d['sheet']}!{d['cell']} = {_short(d['formula'], 80)} ({', '.join(d['empty_cells'][:3])} empty)"
            + (f" -> proposed {_short(d['suggested_formula'], 100)}" if d["suggested_formula"] else "")
        )
    first = origins[0] if origins else direct[0]
    return finding(
        "ERROR",
        "A number is compared with a cell that is empty. Excel reads the empty cell as 0; Mind does not (0 = empty is false there, 1 <= empty is true), so the comparison goes the "
        "other way in Mind (confirmed in a real Mind conversion: IF(<a number> <= <an empty cell>, 0, 1) returned 0 on 157 live rows where Excel returns 1, and the reserves came out wrong). "
        + " ".join(p + "." for p in parts)
        + f" The Prep action 'Read empty cells as 0 where they are compared' puts N() around {with_fix} of them (N of an empty cell is 0 in both; N of a number is the number); "
        "Excel's results stay the same"
        + (f"; {len(sites) - with_fix} left as they are (an array formula, or a cell the test would read holds an error -- see the Prep action's skip list)." if len(sites) - with_fix else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )


# --- FRM-011: a line break inside a string literal -------------------------------------------
# =B62&"<line break>/ text - "&D8 came back from Mind as another text, while =B62&CHAR(10)&"x"
# came back the same (observed in Mind, 2026-10-08). The break written as CHAR(10) is the same
# text in Excel.
_BREAK_RE = re.compile(r"\r\n|\r|\n")


def string_literal_spans(formula: str) -> list[tuple[int, int]]:
    """(start, end) of every "..." literal, quotes included, "" inside a literal being an escaped quote."""
    spans, i, n = [], 0, len(formula)
    while i < n:
        if formula[i] == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            spans.append((i, min(j, n - 1)))
            i = j + 1
        else:
            i += 1
    return spans


def line_breaks_as_char(formula: str) -> str | None:
    """The formula with every line break inside a string literal written as CHAR(10) / CHAR(13); None when there is none."""
    if not formula or not formula.startswith("=") or not _BREAK_RE.search(formula):
        return None
    out, last, changed = [], 0, False
    for a, b in string_literal_spans(formula):
        text = formula[a + 1:b]
        if not _BREAK_RE.search(text):
            continue
        parts = []
        pos = 0
        for m in _BREAK_RE.finditer(text):
            if m.start() > pos:
                parts.append('"' + text[pos:m.start()] + '"')
            parts.append("&".join("CHAR(13)" if ch == "\r" else "CHAR(10)" for ch in m.group(0)))
            pos = m.end()
        if pos < len(text):
            parts.append('"' + text[pos:] + '"')
        out.append(formula[last:a])
        out.append("&".join(parts))
        last = b + 1
        changed = True
    if not changed:
        return None
    out.append(formula[last:])
    return "".join(out)


def line_break_in_text(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = MAX_SITES) -> dict:
    """FRM-011: a line break typed inside a formula's string literal -- Mind computes another text."""
    wb0 = analysis["workbooks"][0]
    sites = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        fixed = line_breaks_as_char(formula)
        if fixed is None:
            continue
        array = bool(f.get("array_ref"))
        sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:300], "suggested_formula": None if array else fixed, "issue": "an array formula: rewrite it by hand" if array else None})
    observed = {"sites": sites if limit is None else sites[:limit], "sites_count": len(sites), "with_proposal": sum(1 for s in sites if s["suggested_formula"])}
    if not sites:
        return finding("PASS", "No formula holds a line break inside a string literal.", observed)
    first = sites[0]
    return finding(
        "ERROR",
        f"{len(sites)} formula(s) hold a line break typed inside a string literal (\"...\"): Mind computes another text for them (confirmed in a real Mind conversion; "
        "the same break written as CHAR(10) came back the same). "
        f"First: {first['sheet']}!{first['cell']} = {_short(first['formula'])}"
        + (f" -> proposed {_short(first['suggested_formula'], 150)}" if first["suggested_formula"] else "")
        + f". The Prep action 'Write line breaks inside text as CHAR(10)' rewrites {observed['with_proposal']} of them; the text Excel shows is the same.",
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )


# --- FRM-012: a formula that asks Excel about itself --------------------------------------------
# =MID(CELL("filename",Sheet2!A1),FIND("[",CELL("filename",...))+1,...): the file's path and
# name, used to build save paths for a macro. Mind has no file: its CELL("filename")
# returned a text without "[", FIND failed and the run stopped ("startIndex ('-1') must be a
# non-negative value", 2026-10-08). The value such a formula shows today is written in its place.
_ENV_RE = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?(CELL|INFO)\s*\(", re.IGNORECASE)
# The app's own session folders (tempfile.mkdtemp(prefix="mind_ready_"), app/web/server.py). Once the
# app's Excel has saved a version there, CELL("filename") shows that folder -- the app's working copy,
# with the Windows user's name in it, not the user's own path.
_APP_FOLDER_MARK = "\\mind_ready_"


def in_app_folder(value: Any) -> bool:
    return isinstance(value, str) and _APP_FOLDER_MARK in value.lower()


def environment_formulas(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = MAX_SITES) -> dict:
    """FRM-012: CELL(...) / INFO(...) -- the Excel session, the file, its path: nothing Mind can answer."""
    wb0 = analysis["workbooks"][0]
    try:
        book = _Book(analysis)
    except Exception as exc:
        return _not_checked(exc)
    sites = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        if not _ENV_RE.search(mask_strings(formula)):
            continue
        m = _COORD_RE.match(str(f.get("cell") or ""))
        stored = book.stored(f["sheet"], int(m.group(2)), col_to_num(m.group(1))) if m else None
        calls = sorted({x.upper() for x in _ENV_RE.findall(mask_strings(formula))})
        multi_cell = bool(f.get("array_ref")) and str(f["array_ref"]).replace("$", "") not in (f["cell"], f"{f['cell']}:{f['cell']}")
        if multi_cell:
            value, issue = None, "an array formula over several cells: rewrite it by hand"
        elif stored is None:
            value, issue = None, 'the cell shows "" or no value is stored: by hand'
        elif _is_error_value(stored):
            value, issue = None, f"the cell shows {stored} in Excel: by hand"
        else:
            value, issue = stored, None
        sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:300], "calls": calls, "value": value, "issue": issue, "array": bool(f.get("array_ref")), "multi_cell": multi_cell})
    observed = {"sites": sites if limit is None else sites[:limit], "sites_count": len(sites), "with_proposal": sum(1 for x in sites if x["issue"] is None)}
    if not sites:
        return finding("PASS", "No formula asks Excel about itself (CELL, INFO).", observed)
    first = sites[0]
    return finding(
        "ERROR",
        f"{len(sites)} formula(s) ask Excel about itself with {', '.join(sorted({c for x in sites for c in x['calls']}))} -- the file's name or path, the session. Mind has no file to answer with: "
        "in a real Mind conversion, MID(CELL(\"filename\",...),FIND(\"[\",...)+1,...) stopped the run (\"startIndex ('-1') must be a non-negative value\"). "
        f"First: {first['sheet']}!{first['cell']} = {_short(first['formula'])}"
        + (f" ({first['issue']})" if first["issue"] else " (shows the app's working folder in this version)" if in_app_folder(first["value"]) else f" (shows {str(first['value'])[:60]!r} today)")
        + f". The Prep action 'Write what CELL / INFO formulas show today in their place' freezes {observed['with_proposal']} of them to the value Excel shows"
        + (f"; {len(sites) - observed['with_proposal']} by hand." if len(sites) - observed["with_proposal"] else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )


def environment_values(analysis: dict[str, Any]) -> dict[tuple[str, str], Any]:
    """What each CELL / INFO formula shows in this workbook, by (sheet, cell); None where it shows
    no usable value. Taken from the original upload, these are the values the user's own Excel saved."""
    sites = (environment_formulas(None, analysis, {}, limit=None).get("observed") or {}).get("sites", [])
    return {(x["sheet"], x["cell"]): x["value"] for x in sites}


# --- FRM-013: an error cell kept on purpose that Mind computes ---------------------------------
# =AB77-AB84 with AB84 holding "": #VALUE! in Excel, kept on purpose (the fixer would have changed
# a good value by hiding it); 0 in Mind, which counts "" as 0 -- and the IFERROR cells fed by it
# (a text label in Excel) turned into numbers there. Written as IF(AB84="",NA(),AB77-AB84) it is an
# error in both (another error in Excel: #N/A for #VALUE!; what reads it sees an error as before).
def blank_text_error_kept(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = MAX_SITES) -> dict:
    """FRM-013: arithmetic on a cell that holds "", with no IFERROR around it, in a cell that shows an error."""
    try:
        res = scan(analysis)
    except Exception as exc:
        return _not_checked(exc)
    if not res["checked"]:
        return _not_checked(None)
    sites = res.get("blank_errors", [])
    with_fix = sum(1 for x in sites if x["suggested_formula"])
    observed = {"sites": sites if limit is None else sites[:limit], "sites_count": len(sites), "with_proposal": with_fix, "without_proposal": len(sites) - with_fix}
    if not sites:
        return finding("PASS", 'No error cell comes from arithmetic on a cell that holds "" (empty text).', observed)
    first = sites[0]
    return finding(
        "ERROR",
        f'{len(sites)} cell(s) show an error in Excel because their formula does arithmetic on a cell that holds "" (empty text), with no IFERROR to catch it -- errors kept on purpose. '
        'Mind counts "" as 0 and computes a number there, and everything that read the error (IFERROR fallbacks, text labels) changes with it '
        "(confirmed in a real Mind conversion: =AB77-AB84 with AB84 holding \"\"). "
        f"First: {first['sheet']}!{first['cell']} = {_short(first['formula'])}"
        + (f" -> proposed {_short(first['suggested_formula'], 150)}" if first["suggested_formula"] else "")
        + f". The Prep action 'Keep an error cell an error in Mind' rewrites {with_fix} of them as IF(cell=\"\",NA(),...): an error in both (#N/A in place of #VALUE!), what reads it sees an error as before"
        + (f"; {len(sites) - with_fix} left as they are." if len(sites) - with_fix else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )
