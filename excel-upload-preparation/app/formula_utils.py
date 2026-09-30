"""Deterministic formula inspection helpers: tokenization, argument splitting,
literal extraction and A1-reference parsing -- not a full Excel formula
parser. Used by inventory.py, grids.py and the validators. Per
instructions/decision-rules.md, function extraction and formula tokenization
belong in code, not the LLM.

1.3.0: string literals and quoted sheet names are masked before scanning
(so `="SUM(" & A1` no longer counts as a SUM call), and the OOXML storage
prefixes Excel adds to formulas are stripped: `_xlfn.` (functions newer than
Excel 2007, e.g. `_xlfn.XLOOKUP`), `_xlws.`, `_xlpm.` (LAMBDA parameters) and
`_xludf.` (user-defined / add-in functions that were unresolved when the file
was last saved) and `_xll.` (a call resolved through an XLL add-in such as
MMForExcel -- seen on the real test workbook). See `storage_prefixes`.
"""
from __future__ import annotations

import re

STORAGE_PREFIX_RE = re.compile(r"_xl(?:fn|ws|udf|pm|op|dlm|l)\.", re.IGNORECASE)  # _xll. = XLL add-in call (MMForExcel)
FUNCTION_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.!])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")

REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(?:(?P<sheet>'[^']+'|[A-Za-z0-9_.]+)!)?"
    r"(?P<ref>"
    r"\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?"  # A1 or A1:B2
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"  # A:A (whole columns)
    r"|\$?\d{1,7}:\$?\d{1,7}"  # 1:1 (whole rows)
    r")"
    r"(?![A-Za-z0-9_(])"
)
SINGLE_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
# a bare identifier: not a function call (no '(' after), not sheet-qualified (no '!' around),
# not a structured reference (no '[' after), not part of a reference ('$' before)
# name characters as Excel allows them: letters of any alphabet, digits, '_', '.', '?', '\'
NAME_TOKEN_RE = re.compile(r"(?<![\w.$!'\[#?\\])([^\W\d\\][\w.?\\]*)(?![\w.(!\[?\\])")
ERROR_LITERAL_RE = re.compile(r"#(?:N/A|REF!|DIV/0!|NAME\?|VALUE!|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)", re.IGNORECASE)
LOCAL_SCOPE_FUNCTIONS = {"LET", "LAMBDA"}  # their arguments define names of their own


def name_tokens(formula: str) -> list[str]:
    """Every bare name a formula refers to (defined names, tables, or names
    that do not exist): string literals, quoted sheet names, cell and range
    references, function calls and error literals are left out. A formula
    that calls LET or LAMBDA is skipped: its own variables are not workbook
    names, and RSK-004 reviews LET."""
    if LOCAL_SCOPE_FUNCTIONS & {fn.upper() for fn in called_functions(formula)}:
        return []
    masked = mask_strings(formula)
    masked = ERROR_LITERAL_RE.sub(lambda m: " " * len(m.group(0)), masked)
    masked = REF_RE.sub(lambda m: " " * len(m.group(0)), masked)
    out = []
    for m in NAME_TOKEN_RE.finditer(masked):
        tok = m.group(1)
        if tok.upper() in ("TRUE", "FALSE") or STORAGE_PREFIX_RE.match(tok):
            continue
        out.append(tok)
    return out


def undefined_names(formula: str, known: set[str]) -> list[str]:
    """Names the formula refers to that are not in `known` (upper-cased
    defined names and table names of the workbook), in order, once each."""
    seen: list[str] = []
    for tok in name_tokens(formula):
        if tok.upper() not in known and tok not in seen:
            seen.append(tok)
    return seen


def mask_strings(formula: str, mask_sheet_quotes: bool = True) -> str:
    """Replace the *content* of "..." string literals (and optionally '...'
    quoted sheet names) with spaces, preserving length so positions line up."""
    out = list(formula)
    i = 0
    n = len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':  # escaped quote
                        j += 2
                        continue
                    break
                j += 1
            for k in range(i + 1, min(j, n)):
                out[k] = " "
            i = j + 1
        elif ch == "'" and mask_sheet_quotes:
            j = formula.find("'", i + 1)
            if j == -1:
                break
            for k in range(i + 1, j):
                out[k] = " "
            i = j + 1
        else:
            i += 1
    return "".join(out)


def strip_storage_prefix(name: str) -> str:
    return STORAGE_PREFIX_RE.sub("", name)


def storage_prefixes(formula: str) -> set[str]:
    """Which OOXML storage prefixes (`_xlfn.`, `_xludf.`, ...) appear in a formula."""
    if not formula:
        return set()
    return {m.group(0).lower() for m in STORAGE_PREFIX_RE.finditer(mask_strings(formula))}


def called_functions(formula: str) -> list[str]:
    """Every `NAME(` occurrence in a formula string, upper-cased, storage
    prefixes stripped, string literals ignored, in order."""
    if not formula or not formula.startswith("="):
        return []
    masked = mask_strings(formula)
    return [strip_storage_prefix(m.group(1)).upper() for m in FUNCTION_CALL_RE.finditer(masked)]


def split_top_level_args(arglist: str) -> list[str]:
    """Split a function's argument-list text on top-level commas only,
    respecting nested parens/brackets and quoted strings."""
    args: list[str] = []
    depth = 0
    in_quotes = False
    current: list[str] = []
    i = 0
    while i < len(arglist):
        ch = arglist[i]
        if in_quotes:
            current.append(ch)
            if ch == '"':
                in_quotes = False
        elif ch == '"':
            in_quotes = True
            current.append(ch)
        elif ch in "([{":
            depth += 1
            current.append(ch)
        elif ch in ")]}":
            depth -= 1
            current.append(ch)
        elif ch in ",;" and depth == 0:
            args.append("".join(current).strip())
            current = []
        else:
            current.append(ch)
        i += 1
    tail = "".join(current).strip()
    if tail or args:
        args.append(tail)
    return args


def find_calls(formula: str, function_name: str) -> list[list[str]]:
    """All call sites of `function_name` in a formula, each as its split argument list."""
    return [c["args"] for c in find_calls_detailed(formula, function_name)]


def find_calls_detailed(formula: str, function_name: str) -> list[dict]:
    """Like find_calls but also returns each call's text span and the exact
    call text, which the resize/syntax validators need."""
    if not formula or not formula.startswith("="):
        return []
    masked = mask_strings(formula)
    calls = []
    pattern = re.compile(rf"(?<![A-Za-z0-9_])(?:_xl[a-z]+\.)?{re.escape(function_name)}\s*\(", re.IGNORECASE)
    for m in pattern.finditer(masked):
        start = m.end()
        depth = 1
        i = start
        while i < len(masked) and depth > 0:
            ch = masked[i]
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            i += 1
        args_text = formula[start : i - 1]
        calls.append(
            {
                "args": split_top_level_args(args_text) if args_text.strip() else [],
                "start": m.start(),
                "end": i,
                "text": formula[m.start() : i],
            }
        )
    return calls


def unquote(arg: str | None) -> str | None:
    """The value of a "..." string literal argument, or None if it isn't one."""
    if arg is None:
        return None
    s = arg.strip()
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        return s[1:-1].replace('""', '"')
    return None


def as_number(arg: str | None) -> float | None:
    """The value of a numeric literal argument, or None."""
    if arg is None:
        return None
    s = arg.strip()
    if re.fullmatch(r"[+-]?\d+(\.\d+)?([eE][+-]?\d+)?", s):
        return float(s)
    return None


def col_to_num(col: str) -> int:
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def parse_ref(text: str | None) -> dict | None:
    """Parse `A1`, `$A$1:B2`, `Sheet!A1:B2`, `'My Sheet'!A1`, `A:A`, `1:1` into
    {sheet, c1, r1, c2, r2, whole_column, whole_row, cells}. None if not a
    literal reference (a defined name, an expression, a function call...)."""
    if not text:
        return None
    s = text.strip()
    if s.startswith("="):
        s = s[1:].strip()
    m = REF_RE.fullmatch(s)
    if not m:
        return None
    sheet = m.group("sheet")
    if sheet and sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return _ref_dict(sheet, m.group("ref"))


def _ref_dict(sheet: str | None, ref: str) -> dict:
    ref_clean = ref.replace("$", "").upper()
    whole_column = whole_row = False
    if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", ref_clean):
        a, b = ref_clean.split(":")
        c1, c2 = sorted((col_to_num(a), col_to_num(b)))
        r1, r2 = 1, 1048576
        whole_column = True
    elif re.fullmatch(r"\d+:\d+", ref_clean):
        a, b = ref_clean.split(":")
        r1, r2 = sorted((int(a), int(b)))
        c1, c2 = 1, 16384
        whole_row = True
    else:
        parts = ref_clean.split(":")
        m1 = SINGLE_REF_RE.match(parts[0])
        c1, r1 = col_to_num(m1.group(1)), int(m1.group(2))
        if len(parts) == 2:
            m2 = SINGLE_REF_RE.match(parts[1])
            c2, r2 = col_to_num(m2.group(1)), int(m2.group(2))
        else:
            c2, r2 = c1, r1
        c1, c2 = sorted((c1, c2))
        r1, r2 = sorted((r1, r2))
    cells = None if (whole_column or whole_row) else (c2 - c1 + 1) * (r2 - r1 + 1)
    return {
        "sheet": sheet,
        "ref": ref_clean,
        "c1": c1,
        "r1": r1,
        "c2": c2,
        "r2": r2,
        "whole_column": whole_column,
        "whole_row": whole_row,
        "cells": cells,
    }


def cell_refs_in_formula(formula: str) -> list[dict]:
    """Every A1-style reference in a formula (string literals ignored), each as
    a parse_ref()-shaped dict. Sheet is None when the reference is local."""
    if not formula or not formula.startswith("="):
        return []
    masked = mask_strings(formula, mask_sheet_quotes=False)
    # Mask double-quoted strings only; quoted sheet names must survive for REF_RE.
    out = []
    for m in REF_RE.finditer(masked):
        sheet = m.group("sheet")
        if sheet and sheet.startswith("'"):
            # Recover the original (unmasked) quoted sheet name from the source text.
            sheet = formula[m.start("sheet") + 1 : m.end("sheet") - 1].replace("''", "'")
        out.append(_ref_dict(sheet, m.group("ref")))
    return out


def range_length(ref: str) -> int | None:
    """Cell count of a simple A1 / A1:A10 / Sheet!A1:B2 reference. None if
    unparseable (e.g. a defined name or an expression rather than a literal range)."""
    parsed = parse_ref(ref.strip().strip("'\"") if ref else ref)
    return parsed["cells"] if parsed else None


def ref_text(c1: int, r1: int, c2: int | None = None, r2: int | None = None) -> str:
    a = f"{num_to_col(c1)}{r1}"
    if c2 is None or r2 is None or (c2 == c1 and r2 == r1):
        return a
    return f"{a}:{num_to_col(c2)}{r2}"
