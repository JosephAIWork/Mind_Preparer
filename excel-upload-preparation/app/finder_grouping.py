"""Grouping  columns, the  macro's key chain and the cells that read its paste area (the instance finder's P1).

The instance finder (app.instance_finder) points to where Mind instances would remove repetition. This
module holds the parts that  read the looped table and the macro:

- the table the macro's  loop steps through (the builder's own  helpers), ONE sweep over its cells with
  the  cached value of every formula cell, and the candidate key columns  by the builder's own rule;
- the macro  chain: VLookup(loop row, table, k) -> Match(value, list) -> Range(anchor).Offset(0, w * (n - 1)),
  read with balanced parentheses (a `[^)]+` regex stops inside the arithmetic) and without the
  commented-out lines;
- lookups keyed by the current row's value, three levels deep (the current-row cells, their aliases, a
  key mapped through a table), and formulas that branch on that value with typed-in literals;
- the cells  that read the macro's paste area, by what each needs in Mind: adders of every slot, readers
  of one slot, rows that read a computed slot, the block that reads its own key's slot.

Everything is read from the analysis and the workbook copy; nothing is written. No Excel, no numpy.
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from typing import Any

from .finder_patterns import _computed_arg
from .formula_utils import NAME_TOKEN_RE, cell_refs_in_formula, find_calls, find_calls_detailed, mask_strings, num_to_col, parse_ref, split_top_level_args
from .instance_builder import KEY_WORDS_RE, MAX_DISTINCT, MIN_DISTINCT, _choose_key, _empty, _filled_columns_left, _find_table, _header_rows, _json_value, _macro_key_column, _norm_arg, _resolve_range, _vba_text
from .inventory import values_cell
from .loop_builder import _held, _qs, _split

KEY_FILLED = 0.80  # a  key column  is filled in at least this share of the rows that hold anything (the builder's rule)
MIN_BRANCHES = 5  # formulas  branching on  the key before the branches count as evidence
MAX_COUNTS = 12  # key  values listed  with their row counts
MAX_CELLS = 6  # example  cells per  line
MAX_LIST_CELLS = 2000  # a  list the  macro matches into is read whole up to this size
MAX_SLOTS = 1000  # a  paste area  cut into more slots than this is not classified
MAX_LITERALS = 12
STEER_FUNCS = {"VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "LOOKUP", "CHOOSE", "SWITCH", "INDIRECT", "INDEX", "OFFSET", "SUMIF", "SUMIFS", "COUNTIFS"}
# the  argument a  lookup looks FOR (0-based): a formula whose key argument is the current row's value is keyed by it
KEY_ARGS = {"VLOOKUP": (0,), "HLOOKUP": (0,), "XLOOKUP": (0,), "MATCH": (0,), "LOOKUP": (0,), "CHOOSE": (0,), "SWITCH": (0,), "INDIRECT": (0,),
            "INDEX": (1, 2), "OFFSET": (1, 2), "SUMIF": (1,), "SUMIFS": (2, 4, 6, 8, 10), "COUNTIFS": (1, 3, 5, 7, 9)}
# the  arguments that  hold the table a keyed lookup reads
TABLE_ARGS = {"VLOOKUP": (1,), "HLOOKUP": (1,), "XLOOKUP": (1, 2), "MATCH": (1,), "LOOKUP": (1, 2), "INDEX": (0,), "SUMIF": (0, 2), "SUMIFS": (0, 1, 3, 5, 7, 9), "COUNTIFS": (0, 2, 4, 6, 8)}
BRANCH_FUNCS = {"IF", "IFS", "OR", "AND", "SWITCH", "CHOOSE"}
_KEY0_FUNCS = {fn for fn, pos in KEY_ARGS.items() if pos == (0,)}  # the  key is  the first argument
_POS_FUNCS = {fn for fn, pos in KEY_ARGS.items() if pos != (0,)}  # the  key sits  further in: the arguments are split
_PAREN_RE = re.compile(r"[()]")
_WORD_RE = re.compile(r"[A-Za-z_\\$][A-Za-z0-9_.$\\?]*")  # the  words of  a formula: names, addresses, functions
_LIT_CMP_RE = re.compile(r'(?P<op1>=|<>)\s*"(?P<l1>(?:[^"]|"")*)"|"(?P<l2>(?:[^"]|"")*)"\s*(?P<op2>=|<>)')  # a  comparison with  a literal
_KEY_CHARS = set("_.$\\?")
_BRANCH_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])(?:_xl[a-z]+\.)?(?:IF|IFS|OR|AND|SWITCH|CHOOSE)\s*\(", re.I)
# the  slot width  inside the macro's Offset(...): 10 * (n - 1) | (n - 1) * 10 | n - 1 (width 1), n the slot variable
W_RE = re.compile(r"(?P<w>\d+)\s*\*\s*\(\s*(?P<v>\w+)\s*-\s*1\s*\)|\(\s*(?P<v2>\w+)\s*-\s*1\s*\)\s*\*\s*(?P<w2>\d+)|^\s*(?P<v3>\w+)\s*-\s*1\s*$")
# the  macro's own  statements (copies of the builder's patterns with the assigned variable in front; the builder's are not edited)
_VBA_RANGE = r"""(?:(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*)?Range\(\s*"(?P<rng>[^"]+)"\s*\)(?:\.Value2?)?"""
_VBA_VLOOKUP_RE = re.compile(r"(?P<lhs>[A-Za-z_]\w*)\s*=\s*(?:[A-Za-z_][\w.]*\.)?VLookup\(\s*(?P<key>[A-Za-z_]\w*)\s*,\s*" + _VBA_RANGE + r"\s*,\s*(?P<k>\d+)", re.I)
_VBA_ANY_VLOOKUP_RE = re.compile(r"VLookup\(\s*(?P<key>[A-Za-z_]\w*)\s*,\s*" + _VBA_RANGE + r"\s*,\s*(?P<k>\d+)", re.I)
_VBA_MATCH_RE = re.compile(r"(?P<lhs>[A-Za-z_]\w*)\s*=\s*(?:[A-Za-z_][\w.]*\.)?Match\(\s*(?P<key>[A-Za-z_]\w*)\s*,\s*" + _VBA_RANGE, re.I)
_ANCHOR_RE = re.compile(r"""(?:(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*)?(?:Range\(\s*"(?P<rng>[^"]+)"\s*\)|\[(?P<br>[^\]]+)\])\s*$""", re.I)
# an  alias: a  formula that is exactly one reference or one defined name
_ALIAS_RE = re.compile(r"^=\s*\+?\s*(?P<ref>(?:'[^']+'|[A-Za-z0-9_.]+)!\$?[A-Za-z]{1,3}\$?\d{1,7}|\$?[A-Za-z]{1,3}\$?\d{1,7}|[A-Za-z_\\][\w.?\\]*)\s*$")
_INT_RE = re.compile(r"^\s*[+-]?\d+\s*$")


# ---  the macro's  chain ------------------------------------------------------------------------------
def _strip_vba_comments(vba: str) -> str:
    """The  VBA without  its comments (from the first apostrophe outside a "..." literal to the end of the
    line) and with `_` continuations joined: a commented-out paste line must not count as the macro's."""
    out: list[str] = []
    buf = ""
    for raw in str(vba or "").splitlines():
        kept: list[str] = []
        in_q = False
        for ch in raw:
            if ch == '"':
                in_q = not in_q
            elif ch == "'" and not in_q:
                break
            kept.append(ch)
        s = "".join(kept).rstrip()
        if s.endswith("_"):
            buf += s[:-1] + " "
            continue
        out.append(buf + s)
        buf = ""
    if buf:
        out.append(buf)
    return "\n".join(out)


def _line_of(text: str, pos: int) -> str:
    a = text.rfind("\n", 0, pos) + 1
    b = text.find("\n", pos)
    return text[a:b if b >= 0 else len(text)].strip()


def _offset_spans(line: str) -> list[tuple[str | None, str | None, str]]:
    """Every  `.Offset(` of  a line: (sheet, the anchor written before it, the text inside its balanced
    parentheses). String literals are skipped by the depth counter."""
    out: list[tuple[str | None, str | None, str]] = []
    low = line.lower()
    pos = 0
    while True:
        i = low.find(".offset(", pos)
        if i < 0:
            return out
        start = i + len(".offset(")
        depth, j, in_q = 1, start, False
        while j < len(line):
            ch = line[j]
            if ch == '"':
                in_q = not in_q
            elif not in_q:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        break
            j += 1
        m = _ANCHOR_RE.search(line[:i])
        anchor = (m.group("rng") or m.group("br")) if m else None
        out.append((m.group("sheet") if m else None, anchor, line[start:j]))
        pos = j + 1


def _slot_width(args: list[str]) -> tuple[int, str, str] | None:
    """(slot  width, the  variable, the axis) of an Offset(rows, columns) argument list whose rows or
    columns argument has the shape w * (var - 1), (var - 1) * w or var - 1."""
    for i, arg in enumerate(args[:2]):
        m = W_RE.search(arg)
        if not m:
            continue
        w = int(m.group("w") or m.group("w2") or 1)
        var = m.group("v") or m.group("v2") or m.group("v3")
        return w, var, "rows" if i == 0 else "columns"
    return None


def _offset_slots(vba: str, allowed: set[str] | None = None) -> list[dict[str, Any]]:
    """Every  paste statement  of the macro that places a block in slot n of an anchor:
    [{anchor, sheet, w, axis, var, line}], comments dropped. `allowed`: the variables (lower case) the
    slot index may be; None accepts  any."""
    out: list[dict[str, Any]] = []
    for line in _strip_vba_comments(vba).splitlines():
        if ".offset(" not in line.lower():
            continue
        for sheet, anchor, span in _offset_spans(line):
            hit = _slot_width(split_top_level_args(span))
            if not hit or not anchor:
                continue
            w, var, axis = hit
            if allowed is not None and var.lower() not in allowed:
                continue
            out.append({"anchor": anchor, "sheet": sheet, "w": w, "axis": axis, "var": var, "line": line.strip()[:160]})
    return out


def _range_text(got: tuple[str, dict[str, Any], str | None]) -> str:
    sheet, r, _name = got
    return f"{_qs(sheet)}!{r['ref']}"


def macro_chain(vba: str, loop_var: str | None, table: dict[str, Any] | None, names: Any) -> dict[str, Any] | None:
    """The  macro's key  chain: key = VLookup(<loop var>, Range("<rng>"), k) gives the key column;
    n = Match(key, Range("<list>"), 0) the list of keys; Range("<anchor>").Offset(0, w * (n - 1)) the
    slot each pass's block goes to.  Any further VLookup(key, Range("X"), k) is a table keyed by the
    value.  None when the macro has none  of it."""
    text = _strip_vba_comments(vba or "")
    if not text.strip():
        return None
    lv = str(loop_var or "").lower()
    chain: dict[str, Any] = {"col_k": None, "col": None, "rng": None, "rng_name": None, "list": None, "list_name": None, "anchor": None, "anchor_name": None,
                             "w": None, "axis": None, "var1": None, "var2": None, "keyed_tables": [], "lines": []}
    lines: list[str] = []
    default_sheet = (table or {}).get("sheet") or ""
    if lv:
        for m in _VBA_VLOOKUP_RE.finditer(text):
            if m.group("key").lower() != lv:
                continue
            got = _resolve_range(m.group("rng"), m.group("sheet") or default_sheet, names)
            if not got:
                continue
            sheet, r, name = got
            col = r["c1"] + int(m.group("k")) - 1
            if table and (sheet != table["sheet"] or r["r2"] < table["r1"] or r["r1"] > table["r2"] or r["c2"] < table["c1"] or r["c1"] > table["c2"] or not table["c1"] <= col <= table["c2"]):
                continue
            chain.update(var1=m.group("lhs"), col_k=int(m.group("k")), col=col, rng=_range_text(got), rng_name=name)
            lines.append(_line_of(text, m.start()))
            break
    var1 = (chain["var1"] or "").lower()
    if var1:
        for m in _VBA_MATCH_RE.finditer(text):
            if m.group("key").lower() != var1:
                continue
            got = _resolve_range(m.group("rng"), m.group("sheet") or default_sheet, names)
            if not got or got[1]["whole_column"] or got[1]["whole_row"]:
                continue
            chain.update(var2=m.group("lhs"), list=_range_text(got), list_name=got[2])
            lines.append(_line_of(text, m.start()))
            break
        for m in _VBA_ANY_VLOOKUP_RE.finditer(text):
            if m.group("key").lower() != var1:
                continue
            got = _resolve_range(m.group("rng"), m.group("sheet") or default_sheet, names)
            if got and _range_text(got) not in {t["ref"] for t in chain["keyed_tables"]}:
                chain["keyed_tables"].append({"name": got[2], "ref": _range_text(got), "column": int(m.group("k"))})
                lines.append(_line_of(text, m.start()))
    allowed = {chain["var2"].lower()} if chain["var2"] else None
    for slot in _offset_slots(text, allowed):
        got = _resolve_range(slot["anchor"], slot["sheet"] or default_sheet, names)
        if not got:
            continue
        chain.update(anchor=_range_text(got), anchor_name=got[2], w=slot["w"], axis=slot["axis"], var_known=allowed is not None)
        lines.append(slot["line"])
        break
    if chain["col_k"] is None and chain["list"] is None and chain["anchor"] is None:
        return None
    seen: list[str] = []
    for ln in lines:
        if ln and ln not in seen:
            seen.append(ln)
    chain["lines"] = [s[:160] for s in seen[:6]]
    return chain


# --- the table and its columns  --------------------------------------------------------------------
def sweep_table(analysis: dict[str, Any], wb: Any, sheet: str, first_row: int, r2: int, c1: int, c2: int) -> dict[str, Any]:
    """One  pass over  the table sheet's cells (`ws._cells`: a cell per miss is never created): per column
    a Counter of values (stopped past MAX_DISTINCT distinct), the filled count and the type flags. A
    formula cell counts  with the value Excel saved for it (`ws.mind_ready_cached`, what a data_only  load
    gives; `values_cell` when the load kept  none), and with no saved value it  counts as empty."""
    cols = range(c1, c2 + 1)
    counters: dict[int, Counter] = {c: Counter() for c in cols}
    filled: Counter = Counter()
    no_value: Counter = Counter()
    texty = {c: True for c in cols}
    numeric = {c: True for c in cols}
    over: set[int] = set()
    rows_seen: set[int] = set()
    ws = wb[sheet] if wb is not None and sheet in wb.sheetnames else None  # no workbook: no values, no candidate (said in a  note)
    cells = getattr(ws, "_cells", None) if ws is not None else None
    cached = getattr(ws, "mind_ready_cached", None) if ws is not None else None
    for (row, col), cell in (cells.items() if hasattr(cells, "items") else []):
        if row < first_row or row > r2 or col < c1 or col > c2:
            continue
        v = cell.value
        if getattr(cell, "data_type", None) == "f" or (isinstance(v, str) and v.startswith("=")):
            if cached is not None:
                v = cached.get((row, col))
            else:
                try:
                    v = values_cell(analysis, sheet, row, col)
                except Exception:
                    v = None
            if _empty(v):
                no_value[col] += 1
                continue
        if _empty(v):
            continue
        rows_seen.add(row)
        filled[col] += 1
        if isinstance(v, str):
            numeric[col] = False
        else:
            texty[col] = False
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                v = _json_value(v)
        cnt = counters[col]
        if v in cnt or len(cnt) <= MAX_DISTINCT:
            cnt[v] += 1
        else:
            over.add(col)
    return {"counters": counters, "filled": filled, "no_value": no_value, "texty": texty, "numeric": numeric, "over": over, "rows": len(rows_seen),
            "last_row": max(rows_seen) if rows_seen else None}


def _driver_entries(loop_plan: dict[str, Any] | None, mp: dict[str, Any] | None) -> list[dict[str, Any]]:
    """The  loops whose  driver may index the table: the plan's counter loops, else the map's drivers."""
    out = [le for le in (loop_plan or {}).get("loops") or [] if le.get("kind") == "counter" and le.get("driver")]
    if not out:
        for d in ((mp or {}).get("loop") or {}).get("drivers") or []:
            if d.get("cell"):
                out.append({"driver": d["cell"], "driver_name": d.get("name"), "loop_var": d.get("loop_var"), "name": d.get("name") or d.get("loop_var")})
    return out


# --- the key cells: the current row, its aliases, the key mapped through a table -----------------------
def current_row_cells(analysis: dict[str, Any], le: dict[str, Any], table: dict[str, Any], names: Any) -> dict[tuple[str, int, int], int]:
    """Every  cell that  reads the current row of the table through the loop driver (INDEX / OFFSET whose
    row argument is the driver's address or defined name) -> the table column it reads: INDEX's literal
    third  argument, else the cell's own column when it sits on the  table sheet inside the table."""
    wb0 = analysis["workbooks"][0]
    drv_sheet, drv_ref = _split(str(le.get("driver") or ""))
    if not drv_sheet or not parse_ref(drv_ref):
        return {}
    drv_ref = drv_ref.replace("$", "").upper()
    wanted = {f"{drv_sheet}!{drv_ref}".upper()}
    if le.get("driver_name"):
        wanted.add(str(le["driver_name"]).upper())
    cell = parse_ref(drv_ref)
    tokens = ([re.escape(str(le["driver_name"]))] if le.get("driver_name") else []) + [rf"\$?{num_to_col(cell['c1'])}\$?{cell['r1']}"]
    quick = re.compile(r"(?<![A-Za-z0-9_.])(?:" + "|".join(tokens) + r")(?![A-Za-z0-9_(])", re.I)
    out: dict[tuple[str, int, int], int] = {}
    for f in wb0.get("formulas") or []:
        formula = str(f.get("formula") or "")
        up = formula.upper()
        if ("INDEX(" not in up and "OFFSET(" not in up) or not quick.search(formula):
            continue
        for fn in ("INDEX", "OFFSET"):
            for args in find_calls(formula, fn):
                if len(args) < 2:
                    continue
                row_arg = _norm_arg(args[1])
                if row_arg not in wanted and not (row_arg == drv_ref and str(f["sheet"]).upper() == drv_sheet.upper()):
                    continue
                got = _resolve_range(args[0], f["sheet"], names)
                if not got:
                    continue
                sheet, r, _name = got
                if sheet != table["sheet"] or r["r2"] < table["r1"] or r["r1"] > table["r2"]:
                    continue
                col = None
                if len(args) >= 3 and _INT_RE.match(args[2]):
                    col = r["c1"] + int(args[2]) - (1 if fn == "INDEX" else 0)
                elif f["sheet"] == table["sheet"]:
                    own = parse_ref(f["cell"])
                    if own and table["c1"] <= own["c1"] <= table["c2"]:
                        col = own["c1"]
                if col is None or not table["c1"] <= col <= table["c2"]:
                    continue
                own = parse_ref(f["cell"])
                if own:
                    out[(f["sheet"], own["r1"], own["c1"])] = col
    return out


def reverse_names(names: Any) -> dict[tuple[str, str], list[str]]:
    """("SHEET",  "C2") ->  the defined names that point at that one cell."""
    out: dict[tuple[str, str], list[str]] = defaultdict(list)
    for key, ref in (getattr(names, "by_name", None) or {}).items():
        r = parse_ref(ref)
        if r and r["sheet"] and r["cells"] == 1:
            out[(r["sheet"].upper(), r["ref"])].append(names.pretty(key))
    return out


def key_tokens(keys: dict[tuple[str, int, int], int], names: Any, rev: dict[tuple[str, str], list[str]] | None = None) -> dict[str, Any]:
    """How  the key  cells can be written in a formula: by address (with or without '$' and a sheet) or by
    defined name. {"addr": {("SHEET", "C2"): col}, "names": {"NAME": col}, "re": quick filter, "cmp": comparisons  with a literal}."""
    rev = reverse_names(names) if rev is None else rev
    addr: dict[tuple[str, str], int] = {}
    nm: dict[str, int] = {}
    words: set[str] = set()  # every way a key cell is written, upper case:  C2, $C$2, C$2, $C2, its names
    parts: list[str] = []
    for (sheet, row, col), k in keys.items():
        letters, ref = num_to_col(col), f"{num_to_col(col)}{row}"
        addr[(sheet.upper(), ref)] = k
        words |= {ref, f"${letters}${row}", f"{letters}${row}", f"${letters}{row}"}
        parts.append(rf"\$?{letters}\$?{row}")
        for n in rev.get((sheet.upper(), ref), []):
            nm[n.upper()] = k
            words.add(n.upper())
            parts.append(re.escape(n))
    if not parts:
        return {"addr": addr, "names": nm, "words": words, "key0": None}
    alts = "|".join(sorted(set(parts), key=len, reverse=True))
    # the  lookups whose key is their first argument,  with a key cell there (the positional ones are split below)
    key0 = re.compile(rf"(?<![A-Za-z0-9_.])(?:_xl[a-z]+\.)?(?P<fn>{'|'.join(sorted(_KEY0_FUNCS))})\s*\(\s*@?(?P<k>(?:(?:'[^']+'|[A-Za-z0-9_.]+)!)?(?:{alts}))\s*[,;)]", re.I)
    return {"addr": addr, "names": nm, "words": words, "key0": key0}


def has_key_word(formula_upper: str, words: set[str]) -> bool:
    """Does  the formula  (upper case) write any key cell's address or name? A quick test before the exact ones."""
    return bool(words) and not words.isdisjoint(_WORD_RE.findall(formula_upper))


def _literal_comparisons(formula: str) -> list[tuple[int, str, str]]:
    """Every  `X =  "text"` / `"text" <> X` of a formula: (position, the text X as written, the literal)."""
    out: list[tuple[int, str, str]] = []
    n = len(formula)
    for m in _LIT_CMP_RE.finditer(formula):
        if m.group("op1") is not None:
            end = m.start()
            while end > 0 and formula[end - 1] == " ":
                end -= 1
            start = end
            while start > 0 and (formula[start - 1].isalnum() or formula[start - 1] in _KEY_CHARS):
                start -= 1
            if start < end:
                out.append((start, _qualified(formula, start, formula[start:end]), m.group("l1")))
        else:
            start = m.end()
            while start < n and formula[start] == " ":
                start += 1
            j = start
            if j < n and formula[j] == "'":  # a quoted sheet name first
                k = formula.find("'", j + 1)
                while k != -1 and k + 1 < n and formula[k + 1] == "'":
                    k = formula.find("'", k + 2)
                if k != -1 and k + 1 < n and formula[k + 1] == "!":
                    j = k + 2
            else:
                k = j
                while k < n and (formula[k].isalnum() or formula[k] in "_."):
                    k += 1
                if k < n and formula[k] == "!":
                    j = k + 1
            end = j
            while end < n and (formula[end].isalnum() or formula[end] in _KEY_CHARS):
                end += 1
            if end > start and formula[start] != '"':
                out.append((m.start(), formula[start:end], m.group("l2")))
    return out


def _qualified(formula: str, start: int, token: str) -> str:
    """The  token with  the sheet written before it (`Data!C2`, `'My sheet'!C7`), when there is one."""
    if start < 2 or formula[start - 1] != "!":
        return token
    j = start - 2
    if formula[j] == "'":
        k = formula.rfind("'", 0, j)
        while k > 0 and formula[k - 1] == "'":  # an escaped quote inside the name
            k = formula.rfind("'", 0, k - 1)
        return formula[k:start] + token if k >= 0 else token
    while j >= 0 and (formula[j].isalnum() or formula[j] in "_."):
        j -= 1
    return formula[j + 1:start] + token


def key_col_of(arg: str, sheet: str, tokens: dict[str, Any]) -> int | None:
    """The  table column  a formula argument stands for when it names a key cell (address or defined name), else None."""
    s = _norm_arg(str(arg or ""))
    if not s:
        return None
    if "!" in s:
        sh, ref = s.rsplit("!", 1)
        return tokens["addr"].get((sh, ref))
    if parse_ref(s):
        return tokens["addr"].get((str(sheet).upper(), s))
    return tokens["names"].get(s)


def alias_cells(analysis: dict[str, Any], keys: dict[tuple[str, int, int], int], names: Any, depth: int = 3) -> dict[tuple[str, int, int], int]:
    """Cells  whose formula  is exactly a reference to a key cell, or its defined name, followed up to `depth`
    levels: they carry the current row's value under another address or name."""
    wb0 = analysis["workbooks"][0]
    simple: list[tuple[str, int, int, str]] = []
    for f in wb0.get("formulas") or []:
        m = _ALIAS_RE.match(str(f.get("formula") or ""))
        if m:
            own = parse_ref(f["cell"])
            if own:
                simple.append((f["sheet"], own["r1"], own["c1"], m.group("ref")))
    found = dict(keys)
    rev = reverse_names(names)
    for _ in range(depth):
        tokens = key_tokens(found, names, rev)
        new: dict[tuple[str, int, int], int] = {}
        for sheet, row, col, ref in simple:
            if (sheet, row, col) in found:
                continue
            k = key_col_of(ref, sheet, tokens)
            if k is not None:
                new[(sheet, row, col)] = k
        if not new:
            break
        found.update(new)
    return {k: v for k, v in found.items() if k not in keys}


def _table_label(arg: str, sheet: str, names: Any) -> str | None:
    """What  a lookup  table argument is called: its defined name, else its address, else the first defined
    name inside it; None when it is none of these (NA(), a computed array)."""
    got = _resolve_range(arg, sheet, names)
    if got:
        return got[2] or _range_text(got)
    for m in NAME_TOKEN_RE.finditer(mask_strings(arg)):
        if names.ref(m.group(1)):
            return names.pretty(m.group(1))
    return None


def mapped_keys(analysis: dict[str, Any], keys: dict[tuple[str, int, int], int], names: Any, tokens: dict[str, Any] | None = None) -> dict[tuple[str, int, int], dict[str, Any]]:
    """Cells  whose whole  formula is one lookup of a key cell through a table -- VLOOKUP(key, table, k, 0),
    XLOOKUP(key, col, col) or INDEX(col, MATCH(key, col, 0)) -- are the key mapped through that table:
    {(sheet,  row, col): {"col": table column, "table": label, "table_arg": text, "name": the  cell's defined name}}.
    One level only: a  mapped key of a mapped key is  not followed."""
    wb0 = analysis["workbooks"][0]
    tokens = tokens or key_tokens(keys, names)
    if not tokens["words"]:
        return {}
    rev = reverse_names(names)
    out: dict[tuple[str, int, int], dict[str, Any]] = {}
    for f in wb0.get("formulas") or []:
        formula = str(f.get("formula") or "").strip()
        up = formula.upper().replace("_XLFN.", "")
        fn = next((n for n in ("VLOOKUP", "XLOOKUP", "INDEX") if up.startswith("=" + n + "(")), None)
        if fn is None or not has_key_word(up, tokens["words"]):
            continue
        calls = find_calls_detailed(formula, fn)
        if not calls or calls[0]["end"] != len(formula):
            continue
        args = calls[0]["args"]
        if len(args) < 2:
            continue
        out_arg, out_k = None, None  # where the mapped values come from: a column of the  table
        if fn == "INDEX":
            inner = find_calls_detailed("=" + args[1].strip(), "MATCH")
            if not inner or inner[0]["start"] != 1 or inner[0]["end"] != len(args[1].strip()) + 1 or not inner[0]["args"]:
                continue
            key_arg, table_arg, out_arg = inner[0]["args"][0], args[0], args[0]
        elif fn == "XLOOKUP":
            key_arg, table_arg, out_arg = args[0], args[1], args[2] if len(args) > 2 else None
        else:
            key_arg, table_arg = args[0], args[1]
            out_arg, out_k = args[1], int(args[2]) if len(args) > 2 and _INT_RE.match(args[2]) else None
        k = key_col_of(key_arg, f["sheet"], tokens)
        if k is None:
            continue
        own = parse_ref(f["cell"])
        if not own:
            continue
        output = None
        got = _resolve_range(out_arg, f["sheet"], names) if out_arg else None
        if got and not got[1]["whole_column"] and not got[1]["whole_row"]:
            osheet, r, _n = got
            ocol = r["c1"] + out_k - 1 if out_k else r["c1"]
            if r["c1"] <= ocol <= r["c2"]:
                output = {"sheet": osheet, "col": ocol, "r1": r["r1"], "r2": r["r2"]}
        cell_names = rev.get((f["sheet"].upper(), own["ref"]), [])
        out[(f["sheet"], own["r1"], own["c1"])] = {"col": k, "table": _table_label(table_arg, f["sheet"], names) or table_arg.strip()[:40], "table_arg": table_arg.strip(), "name": cell_names[0] if cell_names else None, "output": output}
    return out


def _call_spans(masked: str) -> list[tuple[int, int]]:
    """(start,  end) of  every IF / IFS / OR / AND / SWITCH / CHOOSE call of a masked formula (no argument split)."""
    closes: dict[int, int] = {}
    stack: list[int] = []
    for m in _PAREN_RE.finditer(masked):
        if m.group() == "(":
            stack.append(m.start())
        elif stack:
            closes[stack.pop()] = m.end()
    return [(m.start(), closes.get(m.end() - 1, len(masked))) for m in _BRANCH_CALL_RE.finditer(masked)]


def keyed_and_branches(analysis: dict[str, Any], keys: dict[tuple[str, int, int], int], names: Any, tokens: dict[str, Any] | None = None) -> tuple[dict[int, dict[str, Any]], dict[int, dict[str, Any]]]:
    """Per  table column:  the formulas whose lookup key (the argument a lookup looks FOR) is one of the key
    cells, with the tables they read; and the formulas that compare a key cell with a  typed-in literal
    inside IF / IFS / OR / AND / SWITCH  / CHOOSE, with the literals. One  pass over the formulas: a quick
    text test  for a key token first, the arguments split only for the  formulas that pass."""
    wb0 = analysis["workbooks"][0]
    tokens = tokens or key_tokens(keys, names)
    words, key0 = tokens["words"], tokens["key0"]
    keyed: dict[int, dict[str, Any]] = defaultdict(lambda: {"count": 0, "cells": [], "tables": Counter()})
    branches: dict[int, dict[str, Any]] = defaultdict(lambda: {"count": 0, "cells": [], "literals": Counter(), "keys": Counter()})
    if key0 is None:
        return {}, {}
    for f in wb0.get("formulas") or []:
        formula = str(f.get("formula") or "")
        up = formula.upper()
        if not has_key_word(up, words):
            continue
        sheet = f["sheet"]
        where = f"{_qs(sheet)}!{f['cell']}"
        hit: dict[int, set[str]] = {}  # table column -> the lookup  functions keyed by it
        for m in key0.finditer(formula):
            k = key_col_of(m.group("k"), sheet, tokens)
            if k is not None:
                hit.setdefault(k, set()).add(m.group("fn").upper())
        for fn in _POS_FUNCS:
            if fn + "(" not in up:
                continue
            for args in find_calls(formula, fn):
                for i in KEY_ARGS[fn]:
                    if i < len(args):
                        k = key_col_of(args[i], sheet, tokens)
                        if k is not None:
                            hit.setdefault(k, set()).add(fn)
        for k, fns in hit.items():
            keyed[k]["count"] += 1
            if len(keyed[k]["cells"]) < MAX_CELLS:
                keyed[k]["cells"].append(where)
            for fn in fns:
                for args in find_calls(formula, fn):
                    if not any(i < len(args) and key_col_of(args[i], sheet, tokens) == k for i in KEY_ARGS[fn]):
                        continue
                    for j in TABLE_ARGS.get(fn, ()):
                        label = _table_label(args[j], sheet, names) if j < len(args) and j not in KEY_ARGS[fn] else None
                        if label:
                            keyed[k]["tables"][label] += 1
        if '"' in formula and ("=" in formula[1:] or "<>" in formula):
            spans: list[tuple[int, int]] | None = None
            lits: dict[int, set[str]] = {}
            key_texts: dict[int, set[str]] = {}  # how this formula writes the key it compares (a name, an address)
            for pos, key_text, lit in _literal_comparisons(formula):
                k = key_col_of(key_text, sheet, tokens)
                if k is None:
                    continue
                if spans is None:
                    spans = _call_spans(mask_strings(formula))
                if not any(a <= pos < b for a, b in spans):
                    continue
                lits.setdefault(k, set()).add(lit.replace('""', '"'))
                key_texts.setdefault(k, set()).add(key_text.strip().lstrip("@"))
            for k, texts in lits.items():
                branches[k]["count"] += 1
                if len(branches[k]["cells"]) < MAX_CELLS:
                    branches[k]["cells"].append(where)
                for t in texts:
                    branches[k]["literals"][t] += 1
                for t in key_texts.get(k, ()):
                    branches[k]["keys"][t] += 1
    return dict(keyed), dict(branches)


def distinct_literals(literals: Counter) -> list[str]:
    """The  literals compared,  one per spelling case-blind (Excel's = on text is case-blind: "Open" and
    "open" are one literal), most used first, the most used spelling kept."""
    out: dict[str, str] = {}
    for t, _n in literals.most_common():
        out.setdefault(str(t).strip().lower(), str(t).strip())
    return list(out.values())


def branch_key_name(keys: Counter, mapped_here: list[dict[str, Any]], fallback: str) -> str:
    """How  the branching  formulas write the key, for the evidence line: the defined name most of them
    compare (never an address: a header or a name reads better), else the mapped key's name, else `fallback`."""
    for text, _n in keys.most_common():
        t = str(text).strip()
        if "!" not in t and not parse_ref(t.replace("$", "")):
            return t
    for m in mapped_here:
        if m.get("name"):
            return m["name"]
    return fallback


def branches_text(count: int, key_name: str, n_literals: int, are_values: int) -> str:
    """The  branches evidence  line: how many formulas choose a value by comparing the key with typed-in
    texts, and how many of those texts are values of the key (the rest are not, and are said)."""
    head = f'{count} formulas choose a value with IF({key_name}="…")'
    others = n_literals - are_values
    if not are_values:
        return head + f" comparing {n_literals} text{'s' if n_literals != 1 else ''} that {'are' if n_literals != 1 else 'is'} not among its values"
    text = head + f" on {are_values} of its values"
    if others > 0:
        text += f"; {others} other text{'s' if others != 1 else ''} compared {'are' if others != 1 else 'is'} not among its values"
    return text


def column_values(analysis: dict[str, Any], sheet: str, col: int, r1: int, r2: int, limit: int = 500) -> set[str]:
    """The  texts one  column of a range holds (cached values for formula cells), up to `limit` cells."""
    out: set[str] = set()
    for row in range(r1, min(r2, r1 + limit - 1) + 1):
        v = _held(analysis, sheet, col, row)
        if isinstance(v, str) and v.startswith("="):
            try:
                v = values_cell(analysis, sheet, row, col)
            except Exception:
                v = None
        if not _empty(v):
            out.add(str(_json_value(v)).strip())
    return out


# --- P1: the grouping columns of the looped table ------------------------------------------------------
def _header_text(h: Any) -> str | None:
    if isinstance(h, str):
        return None if h.startswith("=") or not h.strip() else " ".join(h.split())
    return None if _empty(h) else str(_json_value(h))


def _counts_block(cnt: Counter, rows: int) -> dict[str, Any]:
    """distinct,  counts (top  MAX_COUNTS + the rest as '…'), biggest, smallest, balance (1 = even groups)."""
    items = [(str(_json_value(k)), int(n)) for k, n in cnt.most_common()]
    distinct = len(items)
    counts = {k: n for k, n in items[:MAX_COUNTS]}
    if len(items) > MAX_COUNTS:
        counts["…"] = sum(n for _k, n in items[MAX_COUNTS:])
    biggest = {"value": items[0][0], "rows": items[0][1]} if items else None
    smallest = {"value": items[-1][0], "rows": items[-1][1]} if items else None
    balance = round((rows / distinct) / biggest["rows"], 3) if biggest and distinct and biggest["rows"] else None
    return {"distinct": distinct, "counts": counts, "biggest": biggest, "smallest": smallest, "balance": balance}


def find_grouping(analysis: dict[str, Any], mp: dict[str, Any] | None, loop_plan: dict[str, Any] | None, names: Any, wb: Any, budget: Any = None) -> dict[str, Any]:
    """P1:  which column  of the table the macro's loop steps through splits its rows into copies.

    -> {"table", "loop", "candidates", "macro" (the chain), "macro_col", "builder" (the builder's rule on the
    full sweep), "keys" (the key cells  by table column), "key_set" (their tokens, for the paste-area pass),
    "notes"}. `candidates`  is empty, with a note, when  no table is found."""
    out: dict[str, Any] = {"table": None, "loop": None, "candidates": [], "macro": None, "macro_col": None, "builder": None, "keys": {}, "key_set": set(), "notes": [], "identifier_columns": 0}
    entries = _driver_entries(loop_plan, mp)
    table = None
    le_used: dict[str, Any] | None = None
    for le in entries:
        table = _find_table(analysis, le, names)
        if table:
            le_used = le
            break
    if not table or le_used is None:
        out["notes"].append("No formula reads a row of a table through the macro's loop: no grouping column to look for." if entries
                            else "No macro loop sets a cell on every pass: no looped table, so no grouping column to look for.")
        return out
    sheet = table["sheet"]
    header_row, first_row = _header_rows(analysis, sheet, table["r1"], table["c1"], table["c2"])
    n_left = _filled_columns_left(analysis, sheet, table["c1"], header_row or first_row, max(table["r2"], first_row))
    c1, c2, r2 = table["c1"] - n_left, table["c2"], table["r2"]
    table = {**table, "c1": c1, "widened_by": n_left, "header_row": header_row, "first_row": first_row}
    sweep = sweep_table(analysis, wb, sheet, first_row, r2, c1, c2)
    rows = sweep["rows"]
    cols = range(c1, c2 + 1)
    headers = {c: _held(analysis, sheet, c, header_row) for c in cols} if header_row else {}
    distinct = {c: len(sweep["counters"][c]) for c in cols}
    need = max(1, int(KEY_FILLED * rows))
    rule_cands = [c for c in cols if sweep["texty"][c] and sweep["filled"][c] >= need and c not in sweep["over"] and MIN_DISTINCT <= distinct[c] <= MAX_DISTINCT]
    last_row = sweep["last_row"] or r2  # the last filled row,  as the builder reports it: the sheet's used range may run  on below the data
    out["table"] = {"sheet": sheet, "r1": table["r1"], "r2": last_row, "last_row": last_row, "c1": c1, "c2": c2, "first_row": first_row, "header_row": header_row,
                    "ref": f"{num_to_col(c1)}{first_row}:{num_to_col(c2)}{last_row}", "rows": rows, "name": table.get("name"), "indexed_by": table.get("cells")}
    out["loop"] = {"name": le_used.get("name"), "driver": le_used.get("driver"), "driver_name": le_used.get("driver_name"), "loop_var": le_used.get("loop_var"), "kind": le_used.get("kind"), "status": le_used.get("status")}
    # (a) the macro's own  key: VLookup(loop row, table, k) -> Match -> Offset slots
    vba = _vba_text(analysis)
    loop_var = str(le_used.get("loop_var") or "")
    chain = macro_chain(vba, loop_var, table, names) if vba else None
    macro_col = _macro_key_column(vba, loop_var, table, names) if vba else None
    if macro_col is None and chain and chain.get("col") is not None:
        macro_col = chain["col"]
    out["macro"], out["macro_col"] = chain, macro_col
    # (b) the key cells: the current row, its aliases, the key mapped through a table
    out_of_time = budget is not None and hasattr(budget, "out") and budget.out()
    if out_of_time:
        budget.skip("lookups and IFs keyed by the table's columns not counted: the time budget ran out")
    cur = current_row_cells(analysis, le_used, table, names) if not out_of_time else {}
    aliases = alias_cells(analysis, cur, names) if cur else {}
    keys = {**cur, **aliases}
    mapped = mapped_keys(analysis, keys, names) if keys else {}
    if mapped:
        # a key mapped through a table is a cell other formulas look values up BY (or compare with literals);
        # a value merely looked up by the key (a rate) is not one
        probe_k, probe_b = keyed_and_branches(analysis, {cell: i for i, cell in enumerate(mapped)}, names)
        used = set(probe_k) | set(probe_b)
        mapped = {cell: info for i, (cell, info) in enumerate(mapped.items()) if i in used}
    all_keys = {**keys, **{k: v["col"] for k, v in mapped.items()}}
    tokens = key_tokens(all_keys, names) if all_keys else {"addr": {}, "names": {}, "re": None, "cmp": None}
    keyed, branches = keyed_and_branches(analysis, all_keys, names, tokens) if all_keys else ({}, {})
    mapping_values: dict[int, set[str]] = defaultdict(set)
    for info in mapped.values():
        o = info.get("output")
        if o:
            mapping_values[info["col"]] |= column_values(analysis, o["sheet"], o["col"], o["r1"], o["r2"])
    out["keys"] = {c: sorted(k for k, v in all_keys.items() if v == c) for c in set(all_keys.values())}
    out["key_set"] = {f"{s.upper()}!{num_to_col(cc)}{r}" for (s, r, cc) in all_keys} | set(tokens["names"])
    # the candidates: the builder's rule, plus any column  the macro or the formulas point at
    pointed = {c for c in [macro_col] + [c for c, k in keyed.items() if k["count"]] if c is not None and c1 <= c <= c2 and (distinct[c] >= MIN_DISTINCT or c in sweep["over"])}
    cands = sorted(set(rule_cands) | pointed)
    for c in cands:
        letter = num_to_col(c)
        header = _header_text(headers.get(c))
        cnt = sweep["counters"][c]
        block = _counts_block(cnt, rows)
        what = header or f"column {letter}"
        kind = "text" if sweep["texty"][c] else "number" if sweep["numeric"][c] else "mixed"
        cand: dict[str, Any] = {"column": letter, "col": c, "sheet": sheet, "header": header, "rows": rows, "filled": int(sweep["filled"][c]), "type": kind, "over": c in sweep["over"], **block,
                                "values": [str(_json_value(k)).strip() for k in cnt], "identifier": block["distinct"] >= 2 and block["distinct"] == rows and rows > 2,
                                "keyed_formulas": 0, "keyed_tables": [], "mapped_through": [], "branch_formulas": 0, "branch_literals": [], "branch_literals_count": 0, "literals_are_values": 0,
                                "macro": None, "label_hint": bool(header and KEY_WORDS_RE.search(header)), "by_rule": c in rule_cands, "evidence": []}
        ev = cand["evidence"]
        first_cell = f"{_qs(sheet)}!{letter}{first_row}"
        if chain and macro_col == c:
            cand["macro"] = {"column_k": chain["col_k"], "range": chain["rng"], "range_name": chain["rng_name"], "list": chain["list"], "list_name": chain["list_name"],
                             "slots": {"anchor": chain["anchor"], "anchor_name": chain["anchor_name"], "w": chain["w"], "axis": chain["axis"]} if chain.get("anchor") else None,
                             "keyed_tables": chain["keyed_tables"], "lines": chain["lines"]}
            text = f"the macro looks the loop row up in column {chain['col_k']} of {chain['rng_name'] or chain['rng']} (VLookup)"
            if chain["list"]:
                text += f", finds the value in {chain['list_name'] or chain['list']} (Match)"
            if chain.get("anchor"):
                a = chain["anchor_name"] or chain["anchor"]
                off = f"Offset({chain['w']}*(n-1), 0)" if chain["axis"] == "rows" else f"Offset(0, {chain['w']}*(n-1))"
                text += f" and pastes each pass's block into slot n of {a} ({off})"
            for t in chain["keyed_tables"]:
                text += f"; a table keyed by the value: {t['name'] or t['ref']}"
            ev.append({"kind": "macro", "text": text, "cells": [first_cell], "vba": chain["lines"]})
        elif macro_col == c:
            ev.append({"kind": "macro", "text": f"the macro looks the loop row up in column {letter} of the table (VLookup)", "cells": [first_cell], "vba": []})
        mapped_here = [{"cell": f"{_qs(s)}!{num_to_col(cc)}{r}", "name": info["name"], "table": info["table"]} for (s, r, cc), info in sorted(mapped.items()) if info["col"] == c]
        cand["mapped_through"] = mapped_here[:MAX_CELLS]
        kd = keyed.get(c)
        if kd and kd["count"]:
            tables = [t for t, _n in kd["tables"].most_common(6)]
            cand["keyed_formulas"], cand["keyed_tables"] = kd["count"], tables
            maps = list(dict.fromkeys(m["table"] for m in mapped_here if m.get("table")))
            through = (f" through the mapping table{'s' if len(maps) > 1 else ''} {' and '.join(maps[:3])}") if maps else ""
            text = f"{kd['count']} formulas look a value up by the current row's {what}{through}" + (f" in {', '.join(tables[:4])}" if tables else "")
            ev.append({"kind": "lookups", "text": text, "cells": kd["cells"][:MAX_CELLS], "count": kd["count"]})
        bd = branches.get(c)
        if bd and bd["count"]:
            lits = distinct_literals(bd["literals"])
            values = {str(_json_value(k)).strip().lower() for k in cnt} | {v.lower() for v in mapping_values.get(c, set())}
            cand["branch_formulas"], cand["branch_literals"], cand["branch_literals_count"] = bd["count"], lits[:MAX_LITERALS], len(lits)
            cand["literals_are_values"] = sum(1 for t in lits if t.lower() in values)
            if bd["count"] >= MIN_BRANCHES:
                key_name = branch_key_name(bd.get("keys") or Counter(), mapped_here, what)
                ev.append({"kind": "branches", "text": branches_text(bd["count"], key_name, len(lits), cand["literals_are_values"]), "cells": bd["cells"][:MAX_CELLS], "count": bd["count"]})
        if block["biggest"]:
            ev.append({"kind": "counts", "text": f"column {letter} ({what}) holds {block['distinct']} values over {rows} rows; the biggest group has {block['biggest']['rows']} rows", "cells": [first_cell], "count": block["distinct"]})
        if cand["label_hint"]:
            ev.append({"kind": "header", "text": f"the header '{header}' names a grouping (a hint only)", "cells": [f"{_qs(sheet)}!{letter}{header_row}"]})
        if cand["identifier"] and not (cand["macro"] or cand["keyed_formulas"]):
            out["identifier_columns"] += 1  # every row its own value: a row number or a contract number, not a grouping
            continue
        out["candidates"].append(cand)
    for c in cols:
        if sweep["no_value"].get(c) and c not in cands:
            out["notes"].append(f"column {num_to_col(c)} of the table holds formulas without saved values: it could not be checked as a key (save the workbook with its values first).")
    # the builder's rule, re-run on the full sweep
    key_col, chosen_by, _note = _choose_key(table, headers, rule_cands, macro_col, None)
    if key_col is None and chosen_by == "heuristic" and rule_cands:
        key_col = min(rule_cands, key=lambda c: (distinct[c], c))
    out["builder"] = {"column": num_to_col(key_col) if key_col else None, "header": _header_text(headers.get(key_col)) if key_col else None, "chosen_by": chosen_by if key_col else None,
                      "candidates": [num_to_col(c) for c in rule_cands], "matches_top": None, "source": "rule"}
    return out


# --- lists and their values -------------------------------------------------------------------------------
def list_texts(analysis: dict[str, Any], sheet: str, r: dict[str, Any]) -> list[str]:
    """The  texts of  a list range in order (cached values for formula cells); [] for a whole row / column
    or a range over MAX_LIST_CELLS cells."""
    if not r or r.get("whole_column") or r.get("whole_row") or (r["r2"] - r["r1"] + 1) * (r["c2"] - r["c1"] + 1) > MAX_LIST_CELLS:
        return []
    out: list[str] = []
    for row in range(r["r1"], r["r2"] + 1):
        for col in range(r["c1"], r["c2"] + 1):
            v = _held(analysis, sheet, col, row)
            if isinstance(v, str) and v.startswith("="):
                try:
                    v = values_cell(analysis, sheet, row, col)
                except Exception:
                    v = None
            if not _empty(v):
                out.append(str(_json_value(v)).strip())
    return out


def same_values(a: list[str], b: list[str]) -> bool:
    """Do  two lists  hold the same set of texts (trimmed, order and repeats ignored)?"""
    sa, sb = {str(x).strip() for x in a}, {str(x).strip() for x in b}
    return bool(sa) and sa == sb


def share_covered(values: list[str], keys: Any) -> float:
    """The  share of  a key column's distinct values a list holds (1.0 = every value has a list entry)."""
    ks = {str(_json_value(k)).strip() for k in keys}
    return len(ks & {str(x).strip() for x in values}) / len(ks) if ks else 0.0


# --- computed displacements inside OFFSET / INDEX -------------------------------------------------------
def _tight(text: str) -> str:
    """The  text without  spaces, except inside quoted sheet names and string literals."""
    out: list[str] = []
    q: str | None = None
    for ch in str(text or ""):
        if q:
            out.append(ch)
            if ch == q:
                q = None
        elif ch in "'\"":
            q = ch
            out.append(ch)
        elif ch != " ":
            out.append(ch)
    return "".join(out)


def displacement_form(arg: str | None) -> dict[str, Any] | None:
    """How  an OFFSET  / INDEX displacement is computed, when it is one of the known idioms:
    (ROW()-n)*w, w*(ROW()-n), ROW()-n (w = 1), the same with COLUMN() -> form "row" | "column";
    w*(COUNTA($X$n:$Xn)-1) -> "counta" (steps along  the expanding range's axis);
    w*(MATCH(key, list, 0)-1) -> "match" (reads the slot  of a key, not a step).
    ->  {form, w, n, rng, step_axis, key, list,  text} or None. The idioms are read by the finder's one
    parser,  app.finder_patterns._computed_arg (balanced parentheses: ROW()-21*10 is none of  them), so the
    readers pass and the step target see the same forms."""
    t = _tight(arg or "")
    if not t:
        return None
    c = _computed_arg(t)
    if c is None:
        return None
    return {"form": c["form"], "w": c["w"], "n": c.get("n"), "rng": c.get("range"), "step_axis": c.get("step_axis"), "key": c.get("key"), "list": c.get("list"), "text": t}


# --- the cells that read the macro's paste area, by class ---------------------------------------------
def _intersects(a: tuple[int, int, int, int], b: tuple[int, int, int, int], grow: int = 0) -> bool:
    """Do  two (r1,  r2, c1, c2) rectangles overlap, `b` grown by `grow` cells on every side?"""
    return a[0] <= b[1] + grow and a[1] >= b[0] - grow and a[2] <= b[3] + grow and a[3] >= b[2] - grow


def _rect(r: dict[str, Any]) -> tuple[int, int, int, int]:
    return (r["r1"], r["r2"], r["c1"], r["c2"])


def _expand_names(text: str, names: Any) -> str:
    """The  formula with  every defined name replaced by its reference (names outside string literals only)."""
    if not getattr(names, "by_name", None):
        return text
    masked = mask_strings(text)
    out: list[str] = []
    pos = 0
    for m in NAME_TOKEN_RE.finditer(masked):
        ref = names.ref(m.group(1))
        if ref:
            out.append(text[pos:m.start(1)])
            out.append(ref)
            pos = m.end(1)
    if not out:
        return text
    out.append(text[pos:])
    return "".join(out)


def _rects_on_sheet(formula: str, own_sheet: str, sheet: str, names: Any) -> tuple[list[tuple[int, int, int, int]], list[dict[str, Any]]]:
    """(the  rectangles the  formula reads on `sheet`, by address or defined name; its OFFSET / INDEX calls whose
    anchor is on `sheet`, each {fn, args, anchor (parsed)})."""
    rects: list[tuple[int, int, int, int]] = []
    for r in cell_refs_in_formula(_expand_names(formula, names)):
        if (r.get("sheet") or own_sheet) == sheet:
            rects.append((r["r1"] or 1, r["r2"] or 1_048_576, r["c1"] or 1, r["c2"] or 16_384))
    calls: list[dict[str, Any]] = []
    up = formula.upper()
    for fn in ("OFFSET", "INDEX"):
        if fn + "(" not in up:
            continue
        for args in find_calls(formula, fn):
            if len(args) < 2:
                continue
            got = _resolve_range(args[0], own_sheet, names)
            if got and got[0] == sheet and not got[1]["whole_column"] and not got[1]["whole_row"]:
                calls.append({"fn": fn, "args": args, "anchor": got[1], "anchor_name": got[2]})
    return rects, calls


def _slot_indices(rect: tuple[int, int, int, int], area: tuple[int, int, int, int], w: int, axis: str) -> set[int]:
    """Which  slots (0-based)  of the area a rectangle touches."""
    if axis == "columns":
        lo, hi = max(rect[2], area[2]) - area[2], min(rect[3], area[3]) - area[2]
    else:
        lo, hi = max(rect[0], area[0]) - area[0], min(rect[1], area[1]) - area[0]
    if hi < lo:
        return set()
    return set(range(lo // w, hi // w + 1))


def _is_key(arg: str, own_sheet: str, keys: set[str]) -> bool:
    s = _norm_arg(str(arg or ""))
    if not s:
        return False
    if "!" in s:
        return s in keys
    if parse_ref(s):
        return f"{own_sheet.upper()}!{s}" in keys
    return s in keys


def _class(n: int = 0) -> dict[str, Any]:
    return {"count": n, "example": None, "cells": [], "text": ""}


def _add(cls: dict[str, Any], where: str) -> None:
    cls["count"] += 1
    if cls["example"] is None:
        cls["example"] = where
    if len(cls["cells"]) < MAX_CELLS:
        cls["cells"].append(where)


def _sheet_filter(sheet: str, names: Any) -> re.Pattern | None:
    """A  quick text  filter for formulas that may read `sheet`: its quoted or plain name, or a defined name
    that points into it (a name hides the sheet: the pass must not depend on the  sheet's word)."""
    parts = [re.escape(_qs(sheet)), re.escape(f"'{sheet}'")]
    for key, ref in (getattr(names, "by_name", None) or {}).items():
        r = parse_ref(ref)
        if r and r["sheet"] == sheet:
            parts.append(re.escape(names.pretty(key)))
    return re.compile("|".join(sorted(set(parts), key=len, reverse=True)), re.I) if parts else None


def paste_area_readers(analysis: dict[str, Any], book: Any, names: Any, cleared: dict[str, Any], w: int | None, keys: set[str], axis: str = "columns") -> dict[str, Any] | None:
    """The  cells that  read the macro's paste area (`cleared`: sheet, r1, r2, c1, c2, cut into slots of width
    `w` along `axis`), in four classes, each with what it needs in Mind, plus the cells  that read the adders:
    all_slots (adders), fixed_slot, computed_slot (a slot per row  or column: ROW / COLUMN /  COUNTA),
    keyed_slot (the slot of a key: MATCH).  Defined names are expanded and OFFSET / INDEX anchors followed, so
    a  formula that never writes the sheet's name  is seen. `book` is accepted for the documented interface;
    the pass reads the analysis' formulas. None when the slot width is unknown or the area too finely cut."""
    if not cleared or not w or w < 1:
        return None
    sheet = cleared["sheet"]
    area = (cleared["r1"], cleared["r2"], cleared["c1"], cleared["c2"])
    length = (area[3] - area[2] + 1) if axis == "columns" else (area[1] - area[0] + 1)
    n_slots = -(-length // w)
    if n_slots > MAX_SLOTS:
        return None
    wb0 = analysis["workbooks"][0]
    quick = _sheet_filter(sheet, names)
    classes = {k: _class() for k in ("all_slots", "fixed_slot", "computed_slot", "keyed_slot", "follow_adders")}
    groups: dict[tuple, dict[str, Any]] = {}
    cross: dict[tuple, dict[str, Any]] = {}
    adders: dict[str, list[tuple[int, int]]] = defaultdict(list)
    readers: set[tuple[str, int, int]] = set()
    seen: list[tuple[dict[str, Any], list[tuple[int, int, int, int]], list[dict[str, Any]]]] = []
    other_axis = "rows" if axis == "columns" else "columns"
    for f in wb0.get("formulas") or []:
        text = str(f.get("formula") or "")
        own = f["sheet"]
        if own != sheet and (quick is None or not quick.search(text)):
            continue
        pos = parse_ref(f["cell"])
        if not pos:
            continue
        here = (pos["r1"], pos["r1"], pos["c1"], pos["c1"])
        if own == sheet and _intersects(here, area):
            continue  # inside the paste area itself
        rects, calls = _rects_on_sheet(text, own, sheet, names)
        if not rects and not calls:
            continue
        seen.append((f, rects, calls))
        where = f"{_qs(own)}!{f['cell']}"
        kind = None
        found: dict[str, Any] | None = None
        for call in calls:
            arect = _rect(call["anchor"])
            if not _intersects(arect, area, grow=1):
                continue
            args = call["args"]
            dr, dc = args[1], (args[2] if len(args) > 2 else "")
            forms = {"rows": displacement_form(dr), "columns": displacement_form(dc)}
            slot_form, attr_form = forms[axis], forms[other_axis]
            if slot_form and slot_form["form"] == "match":
                kind, found = "keyed_slot", {**slot_form, "anchor": call["anchor"], "anchor_name": call["anchor_name"], "attr": attr_form}
                break
            if slot_form and slot_form["form"] in ("row", "column", "counta") and (slot_form["w"] == w or w == 1):
                kind, found = "computed_slot", {**slot_form, "anchor": call["anchor"], "anchor_name": call["anchor_name"], "attr": attr_form}
                break
            if _INT_RE.match(dr or "0") and _INT_RE.match(dc or "0"):
                r0, c0 = call["anchor"]["r1"] + int(dr or 0), call["anchor"]["c1"] + int(dc or 0)
                rects.append((r0, r0, c0, c0))
            else:
                rects.append(arect)
        slots: set[int] = set()
        for rect in rects:
            if _intersects(rect, area):
                slots |= _slot_indices(rect, area, w, axis)
        if kind is None:
            if not slots:
                continue
            if len(slots) >= n_slots:
                kind = "all_slots"
                adders[own].append((pos["r1"], pos["c1"]))
            else:
                kind = "fixed_slot"
                if len(slots) > 1:
                    entry = cross.setdefault((own, tuple(sorted(slots))), {"from": where, "to": [f"slot {i + 1}" for i in sorted(slots)], "cells": 0})
                    entry["cells"] += 1
        readers.add((own, pos["r1"], pos["c1"]))
        _add(classes[kind], where)
        if found is not None:
            step = found.get("step_axis")
            key_is_key = bool(found.get("key")) and _is_key(found["key"], own, keys)
            list_text = None
            if found.get("list"):
                got = _resolve_range(found["list"], own, names)
                list_text = _range_text(got) if got else found["list"]
            gkey = (own, kind, found["form"], found["w"], axis, list_text, (found.get("key") or "").upper())
            g = groups.get(gkey)
            if g is None:
                g = groups[gkey] = {"kind": "computed" if kind == "computed_slot" else "keyed", "sheet": own, "axis": step, "positions": set(), "cells": 0, "form": found["form"], "text": found["text"], "w": found["w"], "n": found.get("n"),
                                    "anchor": f"{_qs(sheet)}!{found['anchor']['ref']}", "anchor_name": found.get("anchor_name"), "anchors": set(), "slot_axis": axis,
                                    "list": list_text, "key": found.get("key"), "key_is_key": key_is_key, "example": where, "fn": "OFFSET" if "OFFSET(" in text.upper() else "INDEX"}
            g["anchors"].add(_rect(found["anchor"]))
            g["cells"] += 1
            o = g.get("own")
            g["own"] = [pos["r1"], pos["r1"], pos["c1"], pos["c1"]] if o is None else [min(o[0], pos["r1"]), max(o[1], pos["r1"]), min(o[2], pos["c1"]), max(o[3], pos["c1"])]
            if step == "rows":
                g["positions"].add(pos["r1"])
            elif step == "columns":
                g["positions"].add(pos["c1"])
    # the cells that read the adding block: they move with it
    hulls: dict[str, tuple[int, int, int, int]] = {}
    followers: dict[str, list[tuple[int, int]]] = defaultdict(list)
    for s, cells in adders.items():
        hulls[s] = (min(r for r, _c in cells), max(r for r, _c in cells), min(c for _r, c in cells), max(c for _r, c in cells))
    for s, hull in hulls.items():
        if s == sheet:
            pool = seen
        else:  # the adders sit on another sheet: the formulas that may read that sheet
            q = _sheet_filter(s, names)
            pool = []
            for f in wb0.get("formulas") or []:
                text = str(f.get("formula") or "")
                if f["sheet"] == s or q is None or q.search(text):
                    rects, calls = _rects_on_sheet(text, f["sheet"], s, names)
                    if rects or calls:
                        pool.append((f, rects, calls))
        for f, rects, calls in pool:
            pos = parse_ref(f["cell"])
            if not pos or (f["sheet"], pos["r1"], pos["c1"]) in readers:
                continue
            if any(_intersects(r, hull) for r in rects) or any(_intersects(_rect(c["anchor"]), hull) for c in calls):
                _add(classes["follow_adders"], f"{_qs(f['sheet'])}!{f['cell']}")
                followers[f["sheet"]].append((pos["r1"], pos["c1"]))
    # words for a person
    n = classes["all_slots"]["count"]
    classes["all_slots"]["text"] = (f"{n} cells add every slot up: in a copy they see only that copy; move the adding row to the main workbook as plain links to the copied model "
                                    "(there a link is the sum over the copies), or read the main workbook's All copies grid") if n else ""
    n = classes["fixed_slot"]["count"]
    classes["fixed_slot"]["text"] = (f"{n} cells read one copy's value: MM_INSTANCE(cell, key) in the main workbook" + (f"; {sum(c['cells'] for c in cross.values())} of them read several slots: one MM_INSTANCE per slot" if cross else "")) if n else ""
    n = classes["computed_slot"]["count"]
    steps = sorted({(g["axis"] or "cells", len(g["positions"])) for g in groups.values() if g["kind"] == "computed"})
    classes["computed_slot"]["text"] = (f"{n} cells in " + " and ".join(f"{k} {ax}" for ax, k in steps) + " each read one slot: one row per copy in the main workbook (a plain link per copy, or MM_INSTANCE)") if n else ""
    n = classes["keyed_slot"]["count"]
    classes["keyed_slot"]["text"] = (f"{n} cells read the current key's slot and add each pass to it: in a copy it is the copy's own block, "
                                     "and MM_RESULT does the adding; nothing to write") if n else ""
    n = classes["follow_adders"]["count"]
    classes["follow_adders"]["text"] = f"{n} cells read the adding block: they move with it" if n else ""
    out_groups = []
    for g in groups.values():
        positions = sorted(g.pop("positions"))
        g["positions"] = positions
        g["per_copy"] = g["cells"] // len(positions) if positions else g["cells"]
        anchors = g.pop("anchors")
        ar = (min(a[0] for a in anchors), max(a[1] for a in anchors), min(a[2] for a in anchors), max(a[3] for a in anchors))  # the anchors' union
        g["anchor_rect"] = ar
        if g["kind"] == "computed":
            if g["form"] == "counta":
                ds = [g["w"] * k for k in range(len(positions))]
            else:
                ds = [(p - g["n"]) * g["w"] for p in positions]
            lo, hi = (min(ds), max(ds)) if ds else (0, 0)
        else:
            got = _resolve_range(g["list"], g["sheet"], names) if g["list"] else None
            count = len(list_texts(analysis, got[0], got[1])) if got else n_slots
            lo, hi = 0, g["w"] * max(0, count - 1)
            g["list_length"] = count
        g["hull"] = (ar[0] + lo, ar[1] + hi, ar[2], ar[3]) if axis == "rows" else (ar[0], ar[1], ar[2] + lo, ar[3] + hi)
        g["inside_area"] = g["hull"][0] >= area[0] and g["hull"][1] <= area[1] and g["hull"][2] >= area[2] and g["hull"][3] <= area[3]
        out_groups.append(g)
    out_groups.sort(key=lambda g: (-g["cells"], g["sheet"], g["example"]))
    return {**classes, "cross_reads": sorted(cross.values(), key=lambda c: -c["cells"])[:10], "groups": out_groups, "n_slots": n_slots, "w": w, "axis": axis,
            "area": {"sheet": sheet, "r1": area[0], "r2": area[1], "c1": area[2], "c2": area[3]}, "adders_hull": {s: list(h) for s, h in hulls.items()},
            "followers_hull": {s: [min(r for r, _c in cells), max(r for r, _c in cells), min(c for _r, c in cells), max(c for _r, c in cells)] for s, cells in followers.items()}}
