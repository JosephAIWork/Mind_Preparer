"""1.8.1 -- two places where Mind computes another value than Excel.

Both were found by uploading a model that recalculates clean in Excel (Horizon,
2026-10-06) and confirmed with a 49-cell test workbook in Mind:

FRM-008  arithmetic on a cell that holds "" (empty text), inside IFERROR.
         Excel: ""-1 is #VALUE!, IFERROR returns its fallback.
         Mind:  "" counts as 0, the arithmetic gives a number, IFERROR never fires.
         =IFERROR(1+INT((D192-1)/12),"") was blank in Excel and a value in Mind
         on 1,242 cells. =IF(D192="","",IFERROR(...)) gives both the same.

FRM-009  a number compared with a cell that is empty in Excel.
         Excel: the empty cell is 0.  Mind: it is not (0=empty is false, 1<=empty is true).
         The empty cell is read directly (=IF(1<=B8,0,1)) or through a formula
         that lands on it and shows 0 in Excel: a plain reference (=+OUTPUT!P4)
         or an exact VLOOKUP / HLOOKUP. Horizon: H9 = VLOOKUP(...) landed on an
         empty cell, IF(D35<=$H$9,0,1) returned 0 on 157 live rows and the
         reserves came out wrong. N() around the reading makes it a 0 in both.

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
}
NUMERIC_CALLS = (TRANSPARENT - {"TEXT"}) | {"N", "LEN", "ROW", "COLUMN", "ROWS", "COLUMNS", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "SUMPRODUCT", "TODAY", "NOW"}
_CELL_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})")
_COORD_RE = re.compile(r"^([A-Za-z]{1,3})(\d{1,7})")
_NAME_BEFORE_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*$")
_IFERROR_RE = re.compile(r"(?<![A-Za-z0-9_.])IFERROR\s*\(", re.IGNORECASE)
_LOOKUP_RE = re.compile(r"(VLOOKUP|HLOOKUP)\s*\(", re.IGNORECASE)
_CALL_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")

_SCANS: dict[str, dict[str, Any]] = {}  # one scan serves both rules and both Prep actions


def _is_formula(v: Any) -> bool:
    return isinstance(v, str) and len(v) > 1 and v.startswith("=")


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
        table = parse_ref(args[1])
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


def _guarded(formula: str, refs: list[re.Match], calls: list, guards: dict[int, set[int]]) -> str:
    """IFERROR(X, fallback) -> IF(cell="", fallback, IFERROR(X, fallback)) for each call in `guards`."""
    for ci in sorted(guards, key=lambda i: calls[i][0], reverse=True):
        start, _open, close, spans = calls[ci]
        tests = []
        for k in sorted(guards[ci]):
            t = formula[refs[k].start():refs[k].end()] + '=""'
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
    out: dict[str, Any] = {"checked": True, "formulas": len(formulas), "blank_arithmetic": [], "origins": [], "direct": [], "readers": 0}
    if not any(book.stored(f["sheet"], row, col) is not None for f, row, col in formulas):
        out["checked"] = False
        return out

    blank_flagged: dict[tuple[str, int, int], dict[int, set[int]]] = {}
    iferror_by_column: dict[tuple[str, int], list[tuple[dict, int, int]]] = {}
    origins: dict[tuple[str, int, int], dict[str, Any]] = {}
    for f, row, col in formulas:
        formula, sheet = f["formula"], f["sheet"]
        rest = formula[1:]
        compares = "=" in rest or "<" in rest or ">" in rest
        has_iferror = "IFERROR" in formula.upper()
        if not (compares or has_iferror):
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

        if has_iferror:
            iferror_by_column.setdefault((sheet, col), []).append((f, row, col))
            calls = _iferror_calls(masked)
            guards: dict[int, set[int]] = {}
            for ci, (_start, open_at, _close, spans2) in enumerate(calls):
                if len(spans2) != 2:
                    continue
                lo, hi = spans2[0]
                for k, m in enumerate(refs):
                    if m.start() < lo or m.end() > hi or ":" in m.group("ref") or not _is_arithmetic_operand(masked, m.start(), m.end(), lo, hi):
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
                if ci in guards and not _fallback_fired(book, formula, calls, ci, sheet, row, col):
                    del guards[ci]
            if guards:
                blank_flagged[(sheet, row, col)] = guards

    # FRM-008: the whole filled-down block gets the same test, so a column keeps one formula
    patterns: dict[tuple[str, int, str], dict[int, set[int]]] = {}
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
    for (sheet, row, col), (f, masked, refs, rel) in sorted(parsed.items(), key=lambda kv: (kv[0][0], kv[0][2], kv[0][1])):
        union = patterns.get((sheet, col, rel))
        if not union:
            continue
        formula = f["formula"]
        calls = _iferror_calls(masked)
        array = bool(f.get("array_ref"))
        cells = sorted({formula[refs[k].start():refs[k].end()] for ks in union.values() for k in ks})
        out["blank_arithmetic"].append({
            "sheet": sheet, "cell": f["cell"], "formula": formula, "blank_cells": cells, "blank_now": (sheet, row, col) in blank_flagged,
            "suggested_formula": None if array else _guarded(formula, refs, calls, union), "issue": "an array formula: rewrite it by hand" if array else None,
        })

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
    sites = res["blank_arithmetic"]
    now = sum(1 for s in sites if s["blank_now"])
    with_fix = sum(1 for s in sites if s["suggested_formula"])
    shown = sorted(sites, key=lambda s: not s["blank_now"])
    observed = {
        "sites": [{**s, "formula": s["formula"][:300]} for s in (shown if limit is None else shown[:limit])],
        "sites_count": len(sites), "blank_now": now, "with_proposal": with_fix, "without_proposal": len(sites) - with_fix,
    }
    if not sites:
        return finding("PASS", 'No IFERROR hides arithmetic on a cell that holds "" (empty text).', observed)
    first = shown[0]
    return finding(
        "ERROR",
        f'{now} formula(s) do arithmetic on a cell that holds "" (empty text) inside IFERROR. Excel stops on the "" (#VALUE!) and returns the IFERROR fallback; '
        'Mind counts "" as 0, computes a number and never reaches the fallback, so these cells get another value in Mind (confirmed in Mind on the model Horizon: '
        "1,242 cells blank in Excel held a value). "
        f"First: {first['sheet']}!{first['cell']} = {_short(first['formula'])}"
        + (f" -> proposed {_short(first['suggested_formula'], 150)}" if first["suggested_formula"] else "")
        + f". The Prep action 'Test for empty text before the arithmetic' rewrites {with_fix} formula(s) -- these and the rest of their filled-down blocks, "
        "so each column keeps one formula; Excel's results stay the same"
        + (f"; {len(sites) - with_fix} array formula(s) are left to rewrite by hand." if len(sites) - with_fix else "."),
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
        "other way in Mind (confirmed in Mind on the model Horizon: IF(D35<=$H$9,0,1) returned 0 on 157 live rows and the reserves came out wrong). "
        + " ".join(p + "." for p in parts)
        + f" The Prep action 'Read empty cells as 0 where they are compared' puts N() around {with_fix} of them (N of an empty cell is 0 in both; N of a number is the number); "
        "Excel's results stay the same"
        + (f"; {len(sites) - with_fix} array formula(s) are left to rewrite by hand." if len(sites) - with_fix else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )
