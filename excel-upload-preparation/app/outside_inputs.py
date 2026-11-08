"""Outside input files -- a workbook whose macros open other files and paste their data in.

Mind runs no VBA and opens no outside file, so data a macro pulls in from other workbooks
must reach Mind another way. This module, for any workbook and any language:

1. reads the VBA for the calls that open outside files (`Workbooks.Open`, `OpenText`,
   `Open ... For Input`, `QueryTables.Add`, `GetObject`) and picks the macro that does the
   import (the procedure without parameters closest to the opening call: in the Tokyo IFRS 17
   tool `calcCF`, which `OneScenarioRun` calls);
2. finds the cells that hold the files' paths (a formula that builds a name ending in .xlsm,
   .xlsx, .csv ... or a typed path) -- the import table -- with, for each row, the cells its
   path formula reads on the same row (the row's key, e.g. "BBA_01") and the sheet it fills;
3. matches the files the user picked to those rows (exact name first, else the row's key as a
   whole word of the file name; Unicode NFKC + case-folding, so full-width （） and Japanese
   names compare as Excel shows them);
4. runs the workbook's OWN import macro on a throw-away copy, in an Excel of its own
   (DispatchEx: never the user's Excel), each path cell pointing at the picked file, and
   watches it through Excel's events: which file it opens and which cells it writes while that
   file is open -- the file's import area. Message boxes are answered and reported. Values
   therefore arrive exactly as the client's own tool puts them (no re-implementation of
   the macro, no openpyxl round trip: openpyxl and Excel parse some decimals 1 ulp apart);
5. makes each import area a Mind input: a sheet IN_<name> with a `#IN_<name> /Input` grid
   holding the data, the formulas that read the area pointed at that grid (same shape, so
   every value stays the same), the old area cleared, and one CSV per grid that Mind's Input
   Manager matches by name (proven in real Mind 2026-11-08: grid `#CF_Test /Input`, CSV
   `CF_Test.csv` with Japanese headers, matched, used by the next run). Each quarter the Mind
   team imports new CSVs instead of uploading the model again.

An area whose readers cannot all be pointed at the new grid (a range that only partly
covers it, OFFSET from inside it, a whole-column reference across it, a name or an array
formula over part of it) stays where it is, inside the model, and is reported.
"""
from __future__ import annotations

import gc
import re
import shutil
import threading
import time
import unicodedata
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable, Iterable

from .formula_utils import REF_RE, col_to_num, find_calls, mask_strings, num_to_col, parse_ref

ProgressFn = Callable[[str], None]

FILE_EXTENSIONS = ("xlsx", "xlsm", "xlsb", "xls", "csv", "txt", "prn")
_EXT = "|".join(FILE_EXTENSIONS)
FILE_NAME_RE = re.compile(rf"\.(?:{_EXT})$", re.IGNORECASE)
PATH_LITERAL_RE = re.compile(rf'"[^"]*\.(?:{_EXT})"', re.IGNORECASE)
INPUT_FILE_EXTENSIONS = {"." + e for e in FILE_EXTENSIONS}

# --- 1. VBA ----------------------------------------------------------------------------------------
IMPORT_CALLS: tuple[tuple[str, re.Pattern], ...] = (
    ("Workbooks.Open", re.compile(r"\bWorkbooks\s*\.\s*Open\b", re.IGNORECASE)),
    ("Workbooks.OpenText", re.compile(r"\bWorkbooks\s*\.\s*OpenText\b", re.IGNORECASE)),
    ("Workbooks.OpenXML", re.compile(r"\bWorkbooks\s*\.\s*OpenXML\b", re.IGNORECASE)),
    ("Open ... For Input", re.compile(r"^\s*Open\s+.+?\s+For\s+(?:Input|Binary)\b", re.IGNORECASE)),
    ("QueryTables.Add", re.compile(r"\bQueryTables\s*\.\s*Add\b", re.IGNORECASE)),
    ("GetObject", re.compile(r"\bGetObject\s*\(", re.IGNORECASE)),
)
# names in any alphabet (a Japanese Sub name is still a Sub)
_PROC_RE = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\W\d]\w*)\s*\((.*)\)", re.IGNORECASE)
_END_PROC_RE = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
_MSGBOX_RE = re.compile(r"\b(?:MsgBox|InputBox)\b", re.IGNORECASE)
_YESNO_RE = re.compile(r"\bvbYesNo|\bvbOKCancel|\bvbYesNoCancel|\bvbRetryCancel", re.IGNORECASE)


def _strip_comment(line: str) -> str:
    out, in_q = [], False
    for ch in line:
        if ch == '"':
            in_q = not in_q
        if ch == "'" and not in_q:
            break
        out.append(ch)
    s = "".join(out).rstrip()
    return "" if re.match(r"^\s*Rem\b", s, re.IGNORECASE) else s


def _mask_vba_strings(line: str) -> str:
    return re.sub(r'"[^"]*"', '""', line)


def vba_procedures(modules: Iterable[tuple[str, str]]) -> dict[str, dict[str, Any]]:
    """{key: {name, module, kind, params, lines}} for every Sub/Function/Property (comments dropped,
    ` _` continuations joined). Names in any alphabet. The key is the name, or `Module.Name` when
    another module has a procedure of the same name (CommandButton1_Click on two sheets): that is
    also how Application.Run reaches it."""
    procs: dict[str, dict[str, Any]] = {}
    for module, code in modules:
        joined: list[str] = []
        buf = ""
        for raw in (code or "").splitlines():
            s = _strip_comment(raw)
            if s.endswith(" _") or s == "_":
                buf += s[:-1] + " "
                continue
            joined.append(buf + s)
            buf = ""
        cur: dict[str, Any] | None = None
        for ln in joined:
            m = _PROC_RE.match(ln)
            if m:
                cur = {"name": m.group(2), "module": module, "kind": m.group(1).split()[0].lower(), "params": m.group(3).strip(), "lines": []}
                key = m.group(2) if m.group(2) not in procs else f"{module}.{m.group(2)}"
                procs[key] = cur
                continue
            if _END_PROC_RE.match(ln):
                cur = None
                continue
            if cur is not None and ln.strip():
                cur["lines"].append(ln)
    return procs


def scan_vba(modules: Iterable[tuple[str, str]]) -> dict[str, Any]:
    """What the macros open, and which macro runs the import.

    {"imports": [{procedure, module, call, line}], "import_macro": name | None,
     "chain": [names from the import macro down to the opening procedure],
     "procedures": {name: {calls, opens, params, msgbox, yes_no}},
     "messages": [{procedure, line}] -- every MsgBox/InputBox line as the source writes it}"""
    modules = list(modules)
    procs = vba_procedures(modules)
    module_alt = "|".join(re.escape(m) for m, _c in modules if m) or "(?!)"
    by_bare: dict[str, list[str]] = {}
    for key, p in procs.items():
        by_bare.setdefault(p["name"], []).append(key)
    info: dict[str, dict[str, Any]] = {}
    imports: list[dict[str, Any]] = []
    msgboxes: list[dict[str, Any]] = []
    for name, p in procs.items():
        masked = [_mask_vba_strings(ln) for ln in p["lines"]]
        opens = []
        for ln, raw in zip(masked, p["lines"]):
            for label, rx in IMPORT_CALLS:
                if rx.search(ln):
                    opens.append(label)
                    imports.append({"procedure": name, "module": p["module"], "call": label, "line": raw.strip()[:200]})
        for raw in p["lines"]:
            if _MSGBOX_RE.search(_mask_vba_strings(raw)):
                msgboxes.append({"procedure": name, "line": raw.strip()[:300]})
        body = "\n".join(masked)
        calls_set: set[str] = set()
        for bare, keys in by_bare.items():
            # `Name`, `Call Name(...)` or `Module.Name` (a call through its module)
            if not re.search(rf"(?<![\w.])(?:(?:{module_alt})\.)?{re.escape(bare)}(?!\w)", body):
                continue
            targets = [k for k in keys if k != name]
            same_module = [k for k in targets if procs[k]["module"] == p["module"]]
            calls_set.update(same_module or targets)
        calls = sorted(calls_set)
        info[name] = {
            "module": p["module"],
            "kind": p["kind"],
            "params": p["params"],
            "calls": calls,
            "opens": sorted(set(opens)),
            "msgbox": bool(_MSGBOX_RE.search(body)),
            "yes_no": bool(_YESNO_RE.search(body)),
        }
    openers = {n for n, i in info.items() if i["opens"]}
    best: tuple[int, int, int, str] | None = None
    best_chain: list[str] = []
    for name, i in info.items():
        if i["kind"] != "sub" or i["params"]:
            continue
        # breadth-first down the calls to the nearest opening procedure
        seen, frontier, parent = {name}, [name], {name: None}
        found = name if name in openers else None
        depth = 0
        while frontier and not found:
            depth += 1
            nxt = []
            for f in frontier:
                for c in info[f]["calls"]:
                    if c not in seen:
                        seen.add(c)
                        parent[c] = f
                        nxt.append(c)
                        if c in openers and not found:
                            found = c
            frontier = nxt
        if not found:
            continue
        chain, x = [], found
        while x is not None:
            chain.append(x)
            x = parent[x]
        chain.reverse()
        key = (len(chain), int(i["yes_no"]), int(i["msgbox"]), name)
        if best is None or key < best:
            best, best_chain = key, chain
    return {
        "imports": imports,
        "import_macro": best_chain[0] if best_chain else None,
        "chain": best_chain,
        "procedures": info,
        "messages": msgboxes,
    }


# --- 2. the import table ---------------------------------------------------------------------------
def normalize_name(name: str) -> str:
    """How two file names compare: Unicode NFKC (full-width （１） = (1)), case-folded, trimmed."""
    return unicodedata.normalize("NFKC", str(name or "")).casefold().strip()


def _split_concat(expr: str) -> list[str]:
    """Top-level operands of `a & b & c` (quotes and parentheses respected)."""
    out, cur, depth, in_q = [], [], 0, False
    i = 0
    while i < len(expr):
        ch = expr[i]
        if in_q:
            cur.append(ch)
            if ch == '"':
                if i + 1 < len(expr) and expr[i + 1] == '"':
                    cur.append('"')
                    i += 1
                else:
                    in_q = False
        elif ch == '"':
            in_q = True
            cur.append(ch)
        elif ch in "([{":
            depth += 1
            cur.append(ch)
        elif ch in ")]}":
            depth -= 1
            cur.append(ch)
        elif ch == "&" and depth == 0:
            out.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur).strip())
    return out


def _literal(op: str) -> str | None:
    if len(op) >= 2 and op[0] == op[-1] == '"' and '"' not in op[1:-1].replace('""', ""):
        return op[1:-1].replace('""', '"')
    return None


def file_name_tail(formula: str) -> str | None:
    """The part of a path formula that builds the FILE NAME: everything after its last "\\"
    (or "/"). `=Folder&"\\"&IF(..)&"_"&D35&".xlsm"` -> `IF(..)&"_"&D35&".xlsm"`. None when the
    formula has no folder separator to cut at."""
    body = formula[1:] if formula.startswith("=") else formula
    ops = _split_concat(body)
    cut = None
    for i, op in enumerate(ops):
        lit = _literal(op)
        if lit is not None and ("\\" in lit or "/" in lit):
            cut = i
    if cut is None:
        return None
    after = re.split(r"[\\/]", _literal(ops[cut]) or "")[-1]
    tail = ([f'"{after}"'] if after else []) + ops[cut + 1:]
    return "&".join(tail) if tail else None


def _same_row_refs(formula: str, sheet: str, row: int) -> list[tuple[str, int, int]]:
    """Single cells on `row` of `sheet` that the formula reads (the row's key cells)."""
    out = []
    for x in _refs(formula, sheet):
        if x["external"] or x["three_d"] or normalize_name(x["sheet"]) != normalize_name(sheet):
            continue
        r = x["ref"]
        if not r["whole_column"] and not r["whole_row"] and r["cells"] == 1 and r["r1"] == row:
            out.append((sheet, r["r1"], r["c1"]))
    return out


def _builds_file_name(formula: str) -> bool:
    """Whether a formula BUILDS a file name: a text with a name or a folder before the extension
    ("data.csv", "C:\\x\\a.xlsm"), or the extension joined on with & (&".xlsm") -- not a check such as
    RIGHT(H35,5)=".xlsm"."""
    for m in re.finditer(rf'"([^"]*)\.(?:{_EXT})"', formula, re.IGNORECASE):
        if m.group(1).strip():
            return True
        if re.search(r"&\s*$", formula[:m.start()]) or re.match(r"\s*&", formula[m.end():]):
            return True
    return False


def find_import_table(cells: Iterable[dict[str, Any]], value_of: Callable[[str, int, int], Any], sheet_names: Iterable[str]) -> list[dict[str, Any]]:
    """The cells that hold an outside file's path, one dict per cell:
    {sheet, cell, row, col, formula, cached, expected (exact name when known), tail (formula text
    that builds the name), keys (values of the same-row cells the formula reads), target_sheet}.

    `cells`: {sheet, cell, formula} for formula cells and {sheet, cell, value} for typed text.
    `value_of(sheet, row, col)`: the value Excel last stored there."""
    sheets = {normalize_name(s): s for s in sheet_names}
    rows: list[dict[str, Any]] = []
    for c in cells:
        formula = c.get("formula")
        value = c.get("value")
        ref = parse_ref(c["cell"])
        if not ref:
            continue
        row, col = ref["r1"], ref["c1"]
        if formula:
            if not _builds_file_name(formula):
                continue
        elif not (isinstance(value, str) and FILE_NAME_RE.search(value.strip()) and re.search(r"[\\/]", value)):
            continue
        cached = value_of(c["sheet"], row, col) if formula else value
        if formula and isinstance(cached, str) and cached.strip() and not cached.startswith("#") and not FILE_NAME_RE.search(cached.strip()):
            continue  # it computes something else ("OK" / "NG" of a check on a path)
        expected = None
        if isinstance(cached, str) and FILE_NAME_RE.search(cached.strip()):
            expected = re.split(r"[\\/]", cached.strip())[-1]
        keys, target = [], None
        if formula:
            for s, r, cc in _same_row_refs(formula, c["sheet"], row):
                v = value_of(s, r, cc)
                if isinstance(v, (str, int, float)) and str(v).strip():
                    keys.append(str(v).strip())
                    if target is None and normalize_name(str(v)) in sheets:
                        target = sheets[normalize_name(str(v))]
        rows.append({
            "sheet": c["sheet"], "cell": c["cell"], "row": row, "col": col,
            "formula": formula, "cached": cached if isinstance(cached, (str, int, float)) else None,
            "expected": expected, "tail": file_name_tail(formula) if formula else None,
            "keys": keys, "target_sheet": target,
        })
    rows.sort(key=lambda r: (r["sheet"], r["col"], r["row"]))
    return rows


def import_table_from_analysis(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    """find_import_table over a scan's analysis (formula cells, plus typed paths)."""
    from .inventory import open_cached, values_cell

    wb_entry = analysis["workbooks"][0]
    cells: list[dict[str, Any]] = [
        {"sheet": f["sheet"], "cell": f["cell"], "formula": f["formula"]}
        for f in wb_entry.get("formulas", [])
        if isinstance(f.get("formula"), str) and PATH_LITERAL_RE.search(f["formula"]) and _builds_file_name(f["formula"])
    ]
    try:
        wb = open_cached(analysis)
        for ws in wb.worksheets:
            for (_r, _c), cell in getattr(ws, "_cells", {}).items():
                v = cell.value
                if isinstance(v, str) and not v.startswith("=") and FILE_NAME_RE.search(v.strip()) and re.search(r"[\\/]", v):
                    cells.append({"sheet": ws.title, "cell": cell.coordinate, "value": v})
        sheet_names = wb.sheetnames
    except Exception:
        sheet_names = [s.get("name") for s in wb_entry.get("sheets", []) if isinstance(s, dict)]

    def value_of(sheet: str, row: int, col: int) -> Any:
        return values_cell(analysis, sheet, row, col)

    return find_import_table(cells, value_of, sheet_names)


# --- 3. matching -----------------------------------------------------------------------------------
def _key_in_name(key: str, name: str) -> bool:
    k, n = normalize_name(key), normalize_name(Path(name).stem)
    if not k:
        return False
    return re.search(rf"(?<![^\W_]){re.escape(k)}(?![^\W_])", n) is not None


def match_files(table: list[dict[str, Any]], uploaded: Iterable[str]) -> dict[str, Any]:
    """Which picked file goes with which import-table row.

    {"rows": [row + {"file": name | None, "how": "exact" | "key" | None, "candidates": [...]}],
     "matched": n, "missing": [rows without a file], "extra": [files no row wants],
     "ambiguous": [files that fit several rows, left for the user to place]}"""
    files = list(dict.fromkeys(uploaded))
    by_norm = {normalize_name(f): f for f in files}
    used: set[str] = set()
    out_rows = []
    for r in table:
        row = dict(r)
        row.update({"file": None, "how": None, "candidates": []})
        if r.get("expected") and normalize_name(r["expected"]) in by_norm:
            f = by_norm[normalize_name(r["expected"])]
            row.update({"file": f, "how": "exact"})
        else:
            cands = [f for f in files if any(_key_in_name(k, f) for k in r.get("keys") or [])]
            ext = Path(r["expected"]).suffix.lower() if r.get("expected") else None
            if ext:
                cands = [f for f in cands if Path(f).suffix.lower() == ext] or cands
            row["candidates"] = cands
            if len(cands) == 1:
                row.update({"file": cands[0], "how": "key"})
        out_rows.append(row)
    # a file claimed by two rows: keep the exact claim, else leave both open
    claims: dict[str, list[dict[str, Any]]] = {}
    for row in out_rows:
        if row["file"]:
            claims.setdefault(row["file"], []).append(row)
    for f, rows in claims.items():
        if len(rows) > 1:
            exact = [r for r in rows if r["how"] == "exact"]
            for r in rows:
                if not (len(exact) == 1 and r is exact[0]):
                    r["candidates"] = sorted(set(r["candidates"]) | {f})
                    r.update({"file": None, "how": None})
    used = {r["file"] for r in out_rows if r["file"]}
    wanted = used | {f for r in out_rows for f in r["candidates"]}
    return {
        "rows": out_rows,
        "matched": len(used),
        "missing": [r for r in out_rows if not r["file"]],
        "extra": [f for f in files if f not in wanted],
        "ambiguous": [f for f in files if f in wanted and f not in used],
    }


def describe_match(match: dict[str, Any]) -> str:
    """One plain sentence: '6 of 6 files matched' / '5 of 6 matched; missing: ...; not used: ...'."""
    n = len(match["rows"])
    parts = [f"{match['matched']} of {n} input files matched"]
    if match["missing"]:
        parts.append("missing: " + ", ".join(
            (r.get("expected") or "/".join(r.get("keys") or []) or f"{r['sheet']}!{r['cell']}") for r in match["missing"]))
    if match["extra"]:
        parts.append("not used: " + ", ".join(match["extra"]))
    return "; ".join(parts) + "."


# --- 4. running the workbook's own import ----------------------------------------------------------
xlCalculationManual = -4135


class _Events:
    """Application events of the import Excel (WithEvents on a DispatchEx object)."""

    log: list[tuple]

    def OnWorkbookOpen(self, wb):  # noqa: N802 -- COM event names
        try:
            self.log.append(("open", str(wb.FullName)))
        except Exception:
            self.log.append(("open", "?"))

    def OnWorkbookBeforeClose(self, wb, cancel):  # noqa: N802
        try:
            self.log.append(("close", str(wb.FullName)))
        except Exception:
            self.log.append(("close", "?"))

    def OnSheetChange(self, sh, target):  # noqa: N802
        try:
            rng = target
            # Columns("AB:DP").ClearContents / Cells.ClearContents: the part of the sheet in use
            if int(target.Rows.Count) >= 1048576 or int(target.Columns.Count) >= 16384:
                inter = sh.Application.Intersect(target, sh.UsedRange)
                if inter is not None:
                    rng = inter
            self.log.append(("change", str(sh.Parent.Name), str(sh.Name), str(rng.Address).replace("$", "")))
        except Exception:
            pass


def repair_dialog_text(text: str) -> str:
    """A VBA MsgBox goes through the ANSI window functions: on a machine whose system code page is
    not the macro's (a Japanese tool on a Hebrew Windows) its text arrives garbled. Undo that when
    it can be undone: re-encode with the system code page, decode with an East Asian one. Bytes the
    system code page cannot hold are lost: compare with source_of_message."""
    if not text or text.isascii():
        return text
    try:
        import codecs
        import ctypes

        acp = int(ctypes.windll.kernel32.GetACP())
        # Windows' own conversion back to bytes: the bytes the code page leaves undefined came in as
        # U+0081 / U+F88E ...; Python's cp1255 codec refuses those, Windows maps them back exactly
        raw = codecs.code_page_encode(acp, text, "strict")[0]
    except Exception:
        return text
    # garbled double-byte text is full of the characters the system code page puts on bytes
    # 0x80-0x9F (lead bytes of Shift-JIS & co.): ‚ ƒ „ … † ‡ ˆ ‰ € ... -- real Hebrew or French text is not
    c1 = set()
    for b in range(0x80, 0xA0):
        try:
            c1.add(codecs.code_page_decode(acp, bytes([b]), "strict", True)[0])
        except Exception:
            pass
    non_ascii = [ch for ch in text if ord(ch) > 127]
    if not non_ascii or sum(ch in c1 for ch in non_ascii) / len(non_ascii) < 0.2:
        return text
    best, best_key = text, None
    for cp in ("cp932", "cp936", "cp949", "cp950"):
        try:
            cand = raw.decode(cp, errors="replace")
        except Exception:
            continue
        cjk = sum(1 for ch in cand if "\u3040" <= ch <= "\u9fff" or "\uac00" <= ch <= "\ud7af")
        if not cjk:
            continue
        # a reading that is valid all through beats one with holes; then the script that only that
        # code page has (kana for Japanese, Hangul for Korean: GBK reads Shift-JIS bytes as Chinese
        # without a single hole); then the most East Asian characters
        # (Korean read from Japanese bytes is all Hangul without a hole too: any kana in a hole-free
        # Japanese reading wins; real Korean bytes leave holes when read as Japanese)
        kana = sum(1 for ch in cand if "\u3040" <= ch <= "\u30ff")
        hangul = sum(1 for ch in cand if "\uac00" <= ch <= "\ud7af")
        own = 2 if (cp == "cp932" and kana) else 1 if (cp == "cp949" and hangul) else 0
        key = (-cand.count("\ufffd"), own, cjk)
        if best_key is None or key > best_key:
            best, best_key = cand, key
    return best


def source_of_message(text: str, messages: list[dict[str, Any]]) -> dict[str, Any] | None:
    """The MsgBox line of the VBA that most likely showed `text` (its string pieces found in it)."""
    best, best_n = None, 0
    for m in messages or []:
        pieces = [x for x in re.findall(r'"([^"]*)"', m.get("line", "")) if len(x) >= 2]
        n = sum(len(x) for x in pieces if x in text)
        if n > best_n:
            best, best_n = m, n
    return best


def _dialog_watch(pid: int, stop: threading.Event, found: list[dict[str, Any]], deadline: float, on_timeout: Callable[[], None]) -> None:
    """Answer the message boxes of Excel `pid` (a VBA MsgBox waits forever in a hidden Excel):
    record the text, press OK -- or No / Cancel on a question, so nothing goes ahead unasked."""
    try:
        import win32api
        import win32con
        import win32gui
        import win32process
    except Exception:
        return

    def texts(hwnd: int) -> list[str]:
        out: list[str] = []

        def child(h, _):
            try:
                t = win32gui.GetWindowText(h)
                if t and win32gui.GetClassName(h) in ("Static", "Button"):
                    out.append(f"{win32gui.GetClassName(h)}:{t}")
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, child, None)
        except Exception:
            pass
        return out

    seen: set[int] = set()
    while not stop.is_set():
        if time.time() > deadline:
            on_timeout()
            return
        hwnds: list[int] = []

        def top(h, _):
            try:
                if win32gui.GetClassName(h) == "#32770" and win32process.GetWindowThreadProcessId(h)[1] == pid and win32gui.IsWindowVisible(h):
                    hwnds.append(h)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(top, None)
        except Exception:
            pass
        for h in hwnds:
            if h in seen:
                continue
            seen.add(h)
            parts = texts(h)
            message = " ".join(p.split(":", 1)[1] for p in parts if p.startswith("Static:")).strip()
            # the buttons by their control id, the same in every language:
            # IDOK 1, IDCANCEL 2, IDABORT 3, IDRETRY 4, IDIGNORE 5, IDYES 6, IDNO 7
            btns: dict[int, tuple[int, str]] = {}

            def btn(hh, _):
                try:
                    if win32gui.GetClassName(hh) == "Button" and win32gui.IsWindowVisible(hh):
                        btns.setdefault(int(win32gui.GetDlgCtrlID(hh)), (hh, win32gui.GetWindowText(hh)))
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumChildWindows(h, btn, None)
            except Exception:
                pass
            # one button: press it; a question: No, else Cancel -- nothing goes ahead unasked
            pick = next(iter(btns)) if len(btns) == 1 else next((i for i in (7, 2, 3, 1) if i in btns), None)
            found.append({"title": win32gui.GetWindowText(h), "message": message,
                          "buttons": [t for _hh, t in btns.values()],
                          "pressed": btns[pick][1] if pick is not None else None, "pressed_id": pick})
            try:
                if pick is not None:
                    win32api.PostMessage(btns[pick][0], win32con.BM_CLICK, 0, 0)
                else:
                    win32api.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
        stop.wait(0.4)


def _addr_rect(address: str) -> tuple[int, int, int, int] | None:
    r = parse_ref(address.split(",")[0])
    if not r or r["whole_column"] or r["whole_row"]:
        return None
    return r["r1"], r["c1"], r["r2"], r["c2"]


def _union(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    return min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])


def _overlap_or_touch(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return not (a[2] + 1 < b[0] or b[2] + 1 < a[0] or a[3] + 1 < b[1] or b[3] + 1 < a[1])


def rect_ref(rect: tuple[int, int, int, int]) -> str:
    r1, c1, r2, c2 = rect
    a = f"{num_to_col(c1)}{r1}"
    return a if (r1, c1) == (r2, c2) else f"{a}:{num_to_col(c2)}{r2}"


def _xl_file_format(path: Path) -> int:
    return {".xlsm": 52, ".xlsx": 51, ".xlsb": 50, ".xls": 56}.get(path.suffix.lower(), 52)


def _new_excel(macros: bool):
    """A hidden Excel of our own -- DispatchEx always starts a new process (Dispatch could
    attach to the user's open Excel)."""
    import win32com.client
    import win32process

    xl = win32com.client.DispatchEx("Excel.Application")
    pid = int(win32process.GetWindowThreadProcessId(int(xl.Hwnd))[1])
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.AutomationSecurity = 1 if macros else 3
    return xl, pid


def _close_excel(xl: Any, ev: Any, pid: int | None) -> None:
    try:
        if ev is not None and hasattr(ev, "close"):
            ev.close()
    except Exception:
        pass
    try:
        if xl is not None:
            xl.Quit()
    except Exception:
        pass
    ev = xl = None
    gc.collect()
    if pid:
        for _ in range(20):
            if not _pid_alive(pid):
                return
            time.sleep(0.25)
        try:
            from .excel_com import kill_excel

            kill_excel(pid)
        except Exception:
            pass


def _pid_alive(pid: int) -> bool:
    try:
        import ctypes

        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))  # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False
        code = ctypes.c_ulong()
        ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
        ctypes.windll.kernel32.CloseHandle(h)
        return code.value == 259  # STILL_ACTIVE
    except Exception:
        return False


def run_import(workbook: Path, rows: list[dict[str, Any]], files: dict[str, Path], macro: str, out_path: Path,
               work_dir: Path, timeout_s: int = 900, progress: ProgressFn | None = None,
               messages: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """Run the workbook's own import macro on a copy with the picked files, watch it, save
    the result to `out_path` (path cells restored to their formulas).

    rows: the import table (match_files rows); files: {row "sheet!cell": picked file path}.
    Rows without a file keep the workbook's own data (the macro clears their area and writes
    "file not found"; both are put back from the upload).

    messages: scan_vba(...)["messages"] -- to show a message box's text as the VBA source writes it.

    Returns {ok, macro, seconds, files: [{cell, file, expected, exact, opened, areas, cleared}],
    kept: [...], dialogs: [{title, message, message_repaired, source, buttons, pressed}],
    stopped_early, error, saved}. ok: no error, no early stop, every picked file opened and written."""
    import pythoncom
    import win32com.client

    say = progress or (lambda _m: None)
    workbook, out_path, work_dir = Path(workbook), Path(out_path), Path(work_dir)
    run_dir = work_dir / "import_run"
    if run_dir.exists():
        shutil.rmtree(run_dir, ignore_errors=True)
    run_dir.mkdir(parents=True, exist_ok=True)
    copy = run_dir / workbook.name
    shutil.copy2(workbook, copy)
    report: dict[str, Any] = {"macro": macro, "seconds": None, "files": [], "kept": [], "dialogs": [],
                              "stopped_early": False, "error": None, "saved": None, "ok": False}
    pythoncom.CoInitialize()
    xl = ev = wb = None
    pid = None
    stop = threading.Event()
    timed_out = {"v": False}
    try:
        xl, pid = _new_excel(macros=True)
        found: list[dict[str, Any]] = []

        def on_timeout() -> None:
            timed_out["v"] = True
            from .excel_com import kill_excel

            kill_excel(pid)

        # from here on every message box is answered and the whole run has a time limit -- the open too
        watcher = threading.Thread(target=_dialog_watch, args=(pid, stop, found, time.time() + timeout_s, on_timeout), daemon=True)
        watcher.start()
        ev = win32com.client.WithEvents(xl, _Events)
        ev.log = []
        say("Opening the workbook to run its import macro")
        xl.EnableEvents = False  # the model's own Workbook_Open does not run: only its import macro does
        wb = xl.Workbooks.Open(str(copy), UpdateLinks=0)
        try:
            xl.Calculation = xlCalculationManual
        except Exception:
            pass
        xl.EnableEvents = True
        xl.AutomationSecurity = 3  # files the macro opens: their own macros stay off
        originals: dict[str, str] = {}
        row_cells: dict[str, Any] = {}
        for r in rows:
            key = f"{r['sheet']}!{r['cell']}"
            ws = wb.Worksheets(r["sheet"])
            cell = ws.Range(r["cell"])
            originals[key] = cell.Formula
            expected = r.get("expected")
            if not expected and r.get("tail"):
                try:
                    v = ws.Evaluate(r["tail"])
                    if isinstance(v, str) and v.strip():
                        expected = v.strip()
                except Exception:
                    pass
            picked = files.get(key)
            entry = {"cell": key, "file": picked.name if picked else None, "expected": expected,
                     "exact": bool(picked and expected and normalize_name(picked.name) == normalize_name(expected)),
                     "opened": False, "areas": [], "cleared": []}
            report["files"].append(entry)
            row_cells[key] = (ws, r["row"])
            # a missing file: a path that does not exist, so the macro takes its own "not found" branch
            cell.Value2 = str(picked.resolve()) if picked else str(run_dir / "not-uploaded" / (expected or "missing.xlsm"))
        ev.log.clear()
        say(f"Running the workbook's import macro {macro}")
        t0 = time.time()
        try:
            xl.Run(f"'{wb.Name}'!{macro}")
        except Exception as exc:  # `End` in the macro (after its message box), or a VBA error
            report["run_exception"] = str(exc)[:300]
        report["seconds"] = round(time.time() - t0, 1)
        stop.set()
        if report.get("run_exception") and not found:
            # no message box explains the stop: the macro itself failed
            report["error"] = "The workbook's import macro stopped with an error: " + report["run_exception"]
        for d in found:
            d["message_repaired"] = repair_dialog_text(d.get("message") or "")
            src = source_of_message(d["message_repaired"], messages or [])
            if src:
                d["source"] = src
        report["dialogs"] = found
        if timed_out["v"]:
            report["error"] = f"the import macro ran longer than {timeout_s} s and was stopped"
            return report
        # --- what happened, file by file
        files_by_name = {normalize_name(Path(e["file"]).name): e for e in report["files"] if e["file"]}
        seen = analyse_events(ev.log, wb.Name, set(files_by_name))
        for nname, info in seen["files"].items():
            e = files_by_name[nname]
            e["opened"] = True
            e["areas"] = info["areas"]
            e["cleared"] = info["cleared"]
        # a box with one button is information ("完了しました"): pressing it changes nothing (its OK
        # button carries id 2 in an MB_OK box); a question answered No / Cancel / Abort is an early stop
        report["stopped_early"] = any(e["file"] and not e["opened"] for e in report["files"]) or any(
            len(d.get("buttons") or []) > 1 and d.get("pressed_id") in (2, 3, 7) for d in found)
        picked = [e for e in report["files"] if e["file"]]
        wrote = any(item[0] == "change" and item[1] == wb.Name for item in ev.log)
        if picked and not any(e["opened"] for e in picked) and wrote and not report["error"]:
            report["error"] = ("The macro read the files without opening them as workbooks (Open ... For Input, a query table ...): "
                               "the app cannot yet tell which data came from which file.")
        # --- put back what must not change: the path formulas; for a file nobody picked, the
        # row of the import table (status cell) and the area the macro cleared get the upload's own content
        for e in report["files"]:
            ws, _row = row_cells[e["cell"]]
            ws.Range(e["cell"].split("!", 1)[1]).Formula = originals[e["cell"]]
        missing = [e for e in report["files"] if not e["file"]]
        if missing or seen["orphans"]:
            orig_copy = run_dir / ("upload_" + workbook.name)
            shutil.copy2(workbook, orig_copy)
            orig = xl.Workbooks.Open(str(orig_copy), UpdateLinks=0, ReadOnly=True)
            try:
                for e in missing:
                    ws, row = row_cells[e["cell"]]
                    last = max(int(ws.UsedRange.Column) + int(ws.UsedRange.Columns.Count) - 1, 1)
                    _restore(orig.Worksheets(ws.Name), ws, (row, 1, row, last))
                    ws.Range(e["cell"].split("!", 1)[1]).Formula = originals[e["cell"]]
                for sheet, rc in seen["orphans"]:
                    _restore(orig.Worksheets(sheet), wb.Worksheets(sheet), rc)
                    report["kept"].append({"sheet": sheet, "ref": rect_ref(rc)})
            finally:
                orig.Close(SaveChanges=False)
        say("Recalculating after the import")
        try:
            xl.CalculateFull()
        except Exception:
            pass
        xl.EnableEvents = False
        out_path.parent.mkdir(parents=True, exist_ok=True)
        if out_path.exists():
            out_path.unlink()
        wb.SaveAs(str(out_path), FileFormat=_xl_file_format(out_path))
        report["saved"] = str(out_path)
        report["ok"] = not report["error"] and not report["stopped_early"] and all(
            e["opened"] and e["areas"] for e in report["files"] if e["file"])
        wb.Close(SaveChanges=False)
        wb = None
        return report
    finally:
        stop.set()
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        wb = None
        _close_excel(xl, ev, pid)
        xl = ev = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def analyse_events(log: list[tuple], book: str, file_names: set[str]) -> dict[str, Any]:
    """Read the macro's run from its event log.

    log items: ("open", full name) / ("close", full name) / ("change", book, sheet, address).
    file_names: the picked files (normalize_name of the file name).

    Returns {"files": {name: {"areas": [...], "cleared": [...]}}, "orphans": [(sheet, rect)]}:
    a file's areas are the ranges written while it was open, joined with the range the macro
    cleared just before opening it on the same sheet (its whole input area, even when this
    file fills less of it); an orphan is a range cleared for no opened file (a file not picked)."""
    windows: list[dict[str, Any]] = []  # one per opened picked file, in order
    outside: list[tuple[int, str, tuple[int, int, int, int]]] = []  # (position, sheet, rect) written while no picked file was open
    current: dict[str, Any] | None = None
    for pos, item in enumerate(log):
        kind = item[0]
        if kind == "open":
            nm = normalize_name(Path(item[1]).name)
            current = {"name": nm, "pos": pos, "writes": []} if nm in file_names else None
            if current is not None:
                windows.append(current)
        elif kind == "close":
            current = None
        elif kind == "change" and item[1] == book:
            rect = _addr_rect(item[3])
            if rect is None:
                continue
            if current is not None:
                current["writes"].append((item[2], rect))
            elif (rect[2] - rect[0] + 1) * (rect[3] - rect[1] + 1) >= 2:
                outside.append((pos, item[2], rect))
    files: dict[str, dict[str, Any]] = {}
    used_clears: set[int] = set()
    for i, w in enumerate(windows):
        areas: list[dict[str, Any]] = []
        for sheet, rect in w["writes"]:
            for a in areas:
                if a["sheet"] == sheet and _overlap_or_touch(a["_rect"], rect):
                    a["_rect"] = _union(a["_rect"], rect)
                    break
            else:
                areas.append({"sheet": sheet, "_rect": rect})
        w["areas"] = areas
    # each range cleared outside a file's window belongs to the FIRST later file that writes there
    # (a macro may clear every area first and import afterwards)
    owner: dict[int, int] = {}
    for k, (pos, sheet, rect) in enumerate(outside):
        for i, w in enumerate(windows):
            if w["pos"] > pos and any(s2 == sheet and _overlap_or_touch(r2, rect) for s2, r2 in w["writes"]):
                owner[k] = i
                break
    for i, w in enumerate(windows):
        areas = w["areas"]
        cleared = []
        for k, (pos, sheet, rect) in enumerate(outside):
            if owner.get(k) != i:
                continue
            for a in areas:
                if a["sheet"] == sheet and _overlap_or_touch(a["_rect"], rect):
                    a["_rect"] = _union(a["_rect"], rect)
                    cleared.append({"sheet": sheet, "ref": rect_ref(rect)})
                    used_clears.add(k)
                    break
        # pieces written in any order (E, C, D) or joined by the cleared range: one area each
        merged = True
        while merged:
            merged = False
            for i1 in range(len(areas)):
                for i2 in range(i1 + 1, len(areas)):
                    a1, a2 = areas[i1], areas[i2]
                    if a1["sheet"] == a2["sheet"] and _overlap_or_touch(a1["_rect"], a2["_rect"]):
                        a1["_rect"] = _union(a1["_rect"], a2["_rect"])
                        del areas[i2]
                        merged = True
                        break
                if merged:
                    break
        for a in areas:
            r1, c1, r2, c2 = a.pop("_rect")
            a.update({"ref": rect_ref((r1, c1, r2, c2)), "r1": r1, "c1": c1, "r2": r2, "c2": c2,
                      "rows": r2 - r1 + 1, "cols": c2 - c1 + 1})
        files[w["name"]] = {"areas": areas, "cleared": cleared}
    # a range cleared for no picked file (a file not picked): put back from the upload -- but never
    # over cells a picked file wrote
    written = [(s2, r2) for w in windows for s2, r2 in w["writes"]]
    orphans = [(sheet, rect) for k, (pos, sheet, rect) in enumerate(outside)
               if k not in used_clears and not any(s2 == sheet and _intersects(r2, rect) for s2, r2 in written)]
    return {"files": files, "orphans": orphans}


def _restore(src_ws: Any, dst_ws: Any, rect: tuple[int, int, int, int]) -> None:
    """Put back `rect` from the upload: values exactly (Value2), formulas as their text --
    never Range.Copy across workbooks (it would turn sheet references into outside links)."""
    ref = rect_ref(rect)
    src = src_ws.Range(ref)
    dst = dst_ws.Range(ref)
    dst.Value2 = src.Value2
    try:
        has = src.HasFormula
    except Exception:
        has = None
    if has is False:
        return
    formulas = src.Formula
    if not isinstance(formulas, tuple):
        formulas = ((formulas,),)
    for i, row in enumerate(formulas):
        for j, f in enumerate(row):
            if isinstance(f, str) and f.startswith("="):
                dst_ws.Cells(rect[0] + i, rect[1] + j).Formula = f


# --- 5. Mind inputs ---------------------------------------------------------------------------------
def _ascii_name(text: str) -> str:
    s = unicodedata.normalize("NFKC", text)
    s = re.sub(r"[^A-Za-z0-9_]+", "_", s).strip("_")
    return s


def plan_mind_inputs(report: dict[str, Any], sheet_names: Iterable[str]) -> list[dict[str, Any]]:
    """One Mind input per import area: {name, sheet, ref, r1, c1, r2, c2, rows, cols, file}.
    Names are `IN_<sheet>` (ASCII, unique; `IN_<sheet>_2` when one sheet holds two areas)."""
    taken = {normalize_name(s) for s in sheet_names}
    out = []
    k = 0
    seen_areas: dict[tuple[str, str], int] = {}
    for e in report.get("files", []):
        for a in e.get("areas", []):
            key = (normalize_name(a["sheet"]), a["ref"])
            if key in seen_areas:  # a later file overwrote the same area: its data is what stays
                out[seen_areas[key]]["file"] = e.get("file")
                continue
            k += 1
            base = "IN_" + (_ascii_name(a["sheet"]) or f"FILE{k}")
            name, n = base[:31], 2
            while normalize_name(name) in taken:
                name = f"{base[:28]}_{n}"
                n += 1
            taken.add(normalize_name(name))
            seen_areas[key] = len(out)
            out.append({"name": name, "file": e.get("file"), **{x: a[x] for x in ("sheet", "ref", "r1", "c1", "r2", "c2", "rows", "cols")}})
    return out


def _refs(formula: str, own_sheet: str) -> Iterable[dict[str, Any]]:
    """Every A1 reference of a formula: {start, end, raw, sheet, ref, external, three_d}.
    external: a reference into another workbook (`[1]Sheet!A1`, `'[Book.xlsx]Sheet'!A1`) -- never
    one of this workbook's areas. three_d: a span of sheets (`Jan:Dec!A1`, `'BBA_04:合計'!A1`) --
    which sheets it covers depends on the tab order, so it is treated as covering every sheet."""
    masked = mask_strings(formula, mask_sheet_quotes=False)
    for m in REF_RE.finditer(masked):
        ref = parse_ref(m.group("ref"))
        if not ref:
            continue
        token = formula[m.start("sheet"):m.end("sheet")] if m.group("sheet") else None
        before = formula[m.start() - 1] if m.start() > 0 else ""
        external = bool(token and ("[" in token)) or before == "]"
        inner = token[1:-1] if token and token.startswith("'") else token
        three_d = bool(inner and ":" in inner) or (before == ":" and token is not None)
        yield {"start": m.start(), "end": m.end(), "raw": m.group("ref"), "token": token,
               "sheet": None if three_d else _sheet_of(token, own_sheet), "ref": ref,
               "external": external, "three_d": three_d}


def _sheet_of(token_sheet: str | None, own: str) -> str:
    if not token_sheet:
        return own
    return token_sheet[1:-1].replace("''", "'") if token_sheet.startswith("'") else token_sheet


def _quote_sheet(name: str) -> str:
    return name if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name) else "'" + name.replace("'", "''") + "'"


def _contains(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return a[0] <= b[0] and a[1] <= b[1] and b[2] <= a[2] and b[3] <= a[3]


def _intersects(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def redirect_formula(formula: str, own_sheet: str, areas: list[dict[str, Any]]) -> tuple[str, int, str | None]:
    """Point every reference that lies inside an import area at the same cells of its Mind input
    grid. Returns (new formula, references changed, why-not) -- why-not is set when the formula
    reads an area in a way that cannot be moved (partly covered, whole columns, OFFSET from inside)."""
    if not formula or not formula.startswith("="):
        return formula, 0, None
    pieces: list[str] = []
    last = 0
    changed = 0
    for r in _refs(formula, own_sheet):
        if r["external"]:
            continue
        ref = r["ref"]
        rect = (ref["r1"], ref["c1"], ref["r2"], ref["c2"])
        for a in areas:
            arect = (a["r1"], a["c1"], a["r2"], a["c2"])
            if not _intersects(arect, rect):
                continue
            if r["three_d"]:
                return formula, 0, f"{r['token']}!{r['raw']} spans several sheets over the import area {a['sheet']}!{a['ref']}"
            if normalize_name(a["sheet"]) != normalize_name(r["sheet"]):
                continue
            if ref["whole_column"] or ref["whole_row"] or not _contains(arect, rect):
                return formula, 0, f"{r['sheet']}!{r['raw']} covers only part of the import area {a['sheet']}!{a['ref']}"
            # same cells on the input grid: data starts at row 3, column A of IN_<name>
            new_ref = _shift_ref(r["raw"], row_shift=3 - a["r1"], col_shift=1 - a["c1"])
            pieces.append(formula[last:r["start"]])
            pieces.append(f"{_quote_sheet(a['name'])}!{new_ref}")
            last = r["end"]
            changed += 1
            break
    if not changed:
        return formula, 0, None
    pieces.append(formula[last:])
    new = "".join(pieces)
    # OFFSET from a cell of the area can land outside it: not provable, keep the area
    for args in find_calls(new, "OFFSET"):
        if args and any(a["name"].lower() + "!" in args[0].lower().replace("'", "") for a in areas):
            return formula, 0, f"OFFSET({args[0]}, ...) may step outside the import area"
    return new, changed, None


_SUMIF_CALL_RE = re.compile(r"(?<![\w.])(?:_xlfn\.)?(SUMIFS?)\s*\(", re.IGNORECASE)
_RANGE_ARG_RE = re.compile(r"^(?P<sheet>'(?:[^']|'')+'!|[^\W\d][\w.]*!)?(?P<a>\$?[A-Za-z]{1,3}\$?\d+):(?P<b>\$?[A-Za-z]{1,3}\$?\d+)$")


def _split_range_cols(arg: str, first: int, last: int, offsets: tuple[int, int]) -> str:
    """`$AD4:$ABS4` sliced to the columns `offsets` (0-based, relative to the range's first column)."""
    m = _RANGE_ARG_RE.match(arg.strip())
    a, b = m.group("a"), m.group("b")
    ma = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", a)
    mb = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", b)
    c0 = min(col_to_num(ma.group(2)), col_to_num(mb.group(2)))
    start = f"{ma.group(1)}{num_to_col(c0 + offsets[0])}{ma.group(3)}{ma.group(4)}"
    end = f"{mb.group(1)}{num_to_col(c0 + offsets[1])}{mb.group(3)}{mb.group(4)}"
    return f"{m.group('sheet') or ''}{start}:{end}"


def split_partial_sums(formula: str, own_sheet: str, areas: list[dict[str, Any]]) -> str | None:
    """SUMIF / SUMIFS whose ranges run past the side of an import area, split into one call over
    the columns outside the area plus one over the columns inside it:
    SUMIF($AD$1:$ABS$1,N$2,$AD4:$ABS4) with the area at AE:ABS ->
    (SUMIF($AD$1:$AD$1,N$2,$AD4:$AD4)+SUMIF($AE$1:$ABS$1,N$2,$AE4:$ABS4)).
    A sum over two sets of columns is the sum over each set, so the value is the same (to the
    last bits of the addition order). Only column splits of ranges that share one shape and whose
    rows lie inside the area. None when nothing could be split."""
    if not formula or not formula.startswith("="):
        return None
    masked = mask_strings(formula, mask_sheet_quotes=False)
    calls = []
    for m in _SUMIF_CALL_RE.finditer(masked):
        start = m.end()
        depth, i = 1, start
        while i < len(masked) and depth > 0:
            depth += {"(": 1, ")": -1}.get(masked[i], 0)
            i += 1
        calls.append((m.start(), i, m.group(1).upper(), formula[start:i - 1]))
    if not calls:
        return None
    from .formula_utils import split_top_level_args

    out = formula
    changed = False
    for c_start, c_end, fn, inner in sorted(calls, reverse=True):
        args = split_top_level_args(inner)
        if fn == "SUMIF":
            if len(args) not in (2, 3):
                continue
            range_pos = [0] + ([2] if len(args) == 3 else [])
        else:
            if len(args) < 3 or len(args) % 2 == 0:
                continue
            range_pos = [0] + list(range(1, len(args), 2))
        parsed = []
        for k in range_pos:
            m = _RANGE_ARG_RE.match(args[k].strip())
            if not m:
                break
            ref = parse_ref(f"{m.group('a')}:{m.group('b')}")
            sheet = _sheet_of(m.group("sheet")[:-1] if m.group("sheet") else None, own_sheet)
            parsed.append((k, sheet, ref))
        if len(parsed) != len(range_pos):
            continue
        shapes = {(r["r2"] - r["r1"], r["c2"] - r["c1"]) for _k, _s, r in parsed}
        if len(shapes) != 1:
            continue
        width = parsed[0][2]["c2"] - parsed[0][2]["c1"] + 1
        cuts: set[int] = set()
        refused = any(":" in (sh or "") or "[" in (sh or "") for _k, sh, _r in parsed)  # sheet spans, outside links
        for _k, sheet, r in parsed:
            for a in areas:
                if normalize_name(a["sheet"]) != normalize_name(sheet):
                    continue
                if not _intersects((a["r1"], a["c1"], a["r2"], a["c2"]), (r["r1"], r["c1"], r["r2"], r["c2"])):
                    continue
                if not (a["r1"] <= r["r1"] and r["r2"] <= a["r2"]):
                    refused = True  # rows outside the area too: a column split would not be enough
                if r["c1"] < a["c1"]:
                    cuts.add(a["c1"] - r["c1"])
                if r["c2"] > a["c2"]:
                    cuts.add(a["c2"] - r["c1"] + 1)
        cuts = {x for x in cuts if 0 < x < width} if not refused else set()
        if not cuts:
            continue
        bounds = [0] + sorted(cuts) + [width]
        pieces = []
        for lo, hi in zip(bounds, bounds[1:]):
            new_args = list(args)
            for k, _sheet, _r in parsed:
                new_args[k] = _split_range_cols(args[k], 0, 0, (lo, hi - 1))
            pieces.append(f"{fn}({','.join(new_args)})")
        out = out[:c_start] + "(" + "+".join(pieces) + ")" + out[c_end:]
        changed = True
    return out if changed else None


def _shift_ref(ref: str, row_shift: int, col_shift: int) -> str:
    def one(part: str) -> str:
        m = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", part)
        if not m:
            return part
        c = col_to_num(m.group(2)) + col_shift
        r = int(m.group(4)) + row_shift
        return f"{m.group(1)}{num_to_col(c)}{m.group(3)}{r}"

    return ":".join(one(p) for p in ref.split(":"))


def csv_number(v: float) -> str:
    """A number as plain decimal text that reads back as exactly the same double
    (shortest round-trip digits, never an exponent: 5.5826458e-05 -> 0.000055826458)."""
    if v != v or v in (float("inf"), float("-inf")):
        return ""
    if float(v).is_integer() and abs(v) < 1e16:
        return str(int(v))
    d = Decimal(repr(float(v)))
    s = format(d, "f")
    return s


def write_input_csv(path: Path, header: list[Any], data: list[list[Any]]) -> dict[str, Any]:
    """The grid as Mind's Input Manager takes it: UTF-8 (no BOM), one header row, then the rows.
    Comma-separated `.csv`; when a value holds a comma or a quote, TAB-separated `.txt` instead
    (Mind reads .txt as TAB by default), so no quoting rule of Mind's is relied on."""
    def cell(v: Any) -> str:
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, (int, float)):
            return csv_number(float(v))
        return str(v)

    rows = [[cell(v) for v in header]] + [[cell(v) for v in r] for r in data]
    if any(("\r" in t or "\n" in t) for r in rows for t in r):
        raise ValueError(f"{Path(path).name}: a value holds a line break; Mind's input format cannot carry it")
    if any(("," in t or '"' in t) for r in rows for t in r):
        if any("\t" in t for r in rows for t in r):
            raise ValueError(f"{Path(path).name}: a value holds a TAB and a comma; Mind's input format cannot carry it")
        path, sep = Path(path).with_suffix(".txt"), "\t"
    else:
        path, sep = Path(path).with_suffix(".csv"), ","
    with open(path, "w", encoding="utf-8", newline="") as fh:
        for r in rows:
            fh.write(sep.join(r) + "\n")
    numbery_text = sum(1 for r in data for v in r if isinstance(v, str) and re.fullmatch(r"\s*-?\d+(?:\.\d+)?\s*", v))
    return {"path": str(path), "rows": len(rows) - 1, "cols": len(header), "separator": "TAB" if sep == "\t" else "comma",
            "text_that_looks_like_numbers": numbery_text}


def _all_text(values: list[Any]) -> bool:
    return bool(values) and all(isinstance(v, str) and v.strip() for v in values)


def _partly_covered(formula: str, own_sheet: str, areas: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """The areas a formula reads in a way that cannot be moved: a reference that covers only part of
    the area (or whole columns / rows across it), or OFFSET starting inside it."""
    out: list[dict[str, Any]] = []
    for r in _refs(formula, own_sheet):
        if r["external"]:
            continue
        ref = r["ref"]
        rect = (ref["r1"], ref["c1"], ref["r2"], ref["c2"])
        for a in areas:
            if a in out:
                continue
            arect = (a["r1"], a["c1"], a["r2"], a["c2"])
            if not _intersects(arect, rect):
                continue
            if r["three_d"] or (normalize_name(a["sheet"]) == normalize_name(r["sheet"]) and (
                    ref["whole_column"] or ref["whole_row"] or not _contains(arect, rect))):
                out.append(a)
    for args in find_calls(formula, "OFFSET"):
        if args:
            for a in areas:
                if a not in out and _reads_area("=" + args[0], own_sheet, a):
                    out.append(a)
    return out


def plan_rewrites(formulas: list[dict[str, Any]], areas: list[dict[str, Any]], names: Iterable[tuple[str, str]] = ()) -> dict[str, Any]:
    """Which import areas can move, and the new text of every formula (and defined name) that
    reads them. Repeats until stable: an area some reader cannot follow stays in the model, and
    the formulas are planned again without it (a formula reading a moving and a staying area
    moves its references to the first only).

    names: (name, refers-to text) pairs. Returns {"areas": movable areas, "kept": [{name, why}],
    "rewrites": {(sheet, cell): formula}, "names": {name: refers-to}, "readers": {area: n}, "sums_split": n}."""
    active = [dict(a) for a in areas]
    kept: list[dict[str, Any]] = []
    while True:
        blocked: dict[str, str] = {}
        rewrites: dict[tuple[str, str], str] = {}
        name_rewrites: dict[str, str] = {}
        readers = {a["name"]: 0 for a in active}
        split_count = 0
        for f in formulas:
            text = f["formula"]
            touched = [a for a in active if _reads_area(text, f["sheet"], a)]
            if not touched:
                continue
            if f.get("array_ref") or f.get("data_table"):
                for a in touched:
                    blocked.setdefault(a["name"], f"the array formula at {f['sheet']}!{f['cell']} reads it")
                continue
            new, n, why = redirect_formula(text, f["sheet"], active)
            split = None
            if why:
                split = split_partial_sums(text, f["sheet"], active)
                if split:
                    new, n, why = redirect_formula(split, f["sheet"], active)
                    if not why and n:
                        split_count += 1
            if why:
                culprits = _partly_covered(split or text, f["sheet"], active) if split else _partly_covered(text, f["sheet"], active)
                for a in culprits or touched:
                    blocked.setdefault(a["name"], f"{f['sheet']}!{f['cell']}: {why}")
                continue
            if n:
                rewrites[(f["sheet"], f["cell"])] = new
                for a in touched:
                    readers[a["name"]] += 1
        for nm, refers in names:
            touched = [a for a in active if _reads_area(refers, "", a)]
            if not touched:
                continue
            new, n, why = redirect_formula(refers if refers.startswith("=") else "=" + refers, "", active)
            if why:
                for a in _partly_covered(refers if refers.startswith("=") else "=" + refers, "", active) or touched:
                    blocked.setdefault(a["name"], f"the name {nm} covers part of it")
            elif n:
                name_rewrites[nm] = new
        if not blocked:
            return {"areas": active, "kept": kept, "rewrites": rewrites, "names": name_rewrites,
                    "readers": readers, "sums_split": split_count}
        kept.extend({"sheet": a["sheet"], "ref": a["ref"], "name": a["name"], "why": blocked[a["name"]]} for a in active if a["name"] in blocked)
        active = [a for a in active if a["name"] not in blocked]


def apply_mind_inputs(workbook: Path, plan: list[dict[str, Any]], formulas: Iterable[dict[str, Any]],
                      out_path: Path, csv_dir: Path, progress: ProgressFn | None = None) -> dict[str, Any]:
    """Make each planned import area a Mind input (see the module text). `formulas`: the
    scan's formula cells {sheet, cell, formula} of `workbook` -- to find the readers.

    Returns {inputs: [{name, sheet, ref, file, rows, cols, csv, separator, readers}],
    kept: [{sheet, ref, name, why}], formulas_changed, names_changed, sums_split, saved}."""
    import pythoncom

    say = progress or (lambda _m: None)
    workbook, out_path, csv_dir = Path(workbook), Path(out_path), Path(csv_dir)
    csv_dir.mkdir(parents=True, exist_ok=True)
    formulas = [f for f in formulas if isinstance(f.get("formula"), str) and f["formula"].startswith("=")]
    areas = [dict(p) for p in plan]
    result: dict[str, Any] = {"inputs": [], "kept": [], "formulas_changed": 0, "names_changed": 0, "sums_split": 0, "saved": None}
    pythoncom.CoInitialize()
    xl = wb = None
    pid = None
    try:
        xl, pid = _new_excel(macros=False)
        try:
            xl.EnableEvents = False
        except Exception:
            pass
        say("Opening the workbook to make the Mind inputs")
        wb = xl.Workbooks.Open(str(workbook), UpdateLinks=0)
        try:
            xl.Calculation = xlCalculationManual
        except Exception:
            pass
        # 1. an area holding error values or formulas cannot become a values-only Mind input
        movable = []
        for a in areas:
            rng = wb.Worksheets(a["sheet"]).Range(a["ref"])
            why = None
            try:
                rng.SpecialCells(2, 16)  # xlCellTypeConstants, xlErrors
                why = "it holds error values (#N/A, #VALUE! ...), which a CSV cannot carry"
            except Exception:
                pass
            try:
                if rng.HasFormula is not False:
                    why = why or "it holds formulas, and a Mind input holds values only"
            except Exception:
                pass
            if why:
                result["kept"].append({"sheet": a["sheet"], "ref": a["ref"], "name": a["name"], "why": why})
            else:
                movable.append(a)
        # 2. who reads the areas (formulas and defined names), planned until stable
        names: list[tuple[str, str]] = []
        name_objs: dict[str, Any] = {}
        for nm in list(wb.Names):
            try:
                names.append((str(nm.Name), str(nm.RefersTo)))
                name_objs[str(nm.Name)] = nm
            except Exception:
                continue
        plan_r = plan_rewrites(formulas, movable, names)
        result["kept"].extend(plan_r["kept"])
        result["sums_split"] = plan_r["sums_split"]
        keep = plan_r["areas"]
        # 3. the input sheets
        for a in keep:
            src = wb.Worksheets(a["sheet"])
            data = src.Range(rect_ref((a["r1"], a["c1"], a["r2"], a["c2"]))).Value2
            data = [list(r) for r in (data if isinstance(data, tuple) else ((data,),))]
            header: list[Any] = []
            if a["r1"] > 1:
                h = src.Range(rect_ref((a["r1"] - 1, a["c1"], a["r1"] - 1, a["c2"]))).Value2
                header = list(h[0]) if isinstance(h, tuple) else [h]
            if not _all_text(header):
                header = [f"C{i + 1}" for i in range(a["cols"])]
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = a["name"]
            ws.Range("A1").Value2 = f"#{a['name']} /Input"
            ws.Range(ws.Cells(2, 1), ws.Cells(2, a["cols"])).Value2 = [header]
            ws.Range(ws.Cells(3, 1), ws.Cells(2 + a["rows"], a["cols"])).Value2 = data
            info = write_input_csv(csv_dir / a["name"], header, data)
            result["inputs"].append({"name": a["name"], "sheet": a["sheet"], "ref": a["ref"], "file": a.get("file"),
                                     "rows": a["rows"], "cols": a["cols"], "csv": info["path"], "separator": info["separator"],
                                     "readers": plan_r["readers"].get(a["name"], 0), "header_from_workbook": header[:1] != ["C1"],
                                     "text_that_looks_like_numbers": info["text_that_looks_like_numbers"]})
        # 4. defined names over a moving area
        for nm, refers in plan_r["names"].items():
            name_objs[nm].RefersTo = refers
            result["names_changed"] += 1
        # 5. the readers, in runs of neighbouring cells of one column
        say(f"Pointing {len(plan_r['rewrites'])} formulas at the Mind inputs")
        by_sheet_col: dict[tuple[str, int], list[tuple[int, str]]] = {}
        for (sheet, cell), new in plan_r["rewrites"].items():
            r = parse_ref(cell)
            by_sheet_col.setdefault((sheet, r["c1"]), []).append((r["r1"], new))
        for (sheet, col), items in by_sheet_col.items():
            items.sort()
            ws = wb.Worksheets(sheet)
            run: list[tuple[int, str]] = []
            for row, new in items + [(None, None)]:
                if run and (row is None or row != run[-1][0] + 1):
                    rng = ws.Range(ws.Cells(run[0][0], col), ws.Cells(run[-1][0], col))
                    rng.Formula = [[x[1]] for x in run] if len(run) > 1 else run[0][1]
                    result["formulas_changed"] += len(run)
                    run = []
                if row is not None:
                    run.append((row, new))
        # 6. the old areas, now read by nobody
        for a in keep:
            wb.Worksheets(a["sheet"]).Range(a["ref"]).ClearContents()
        say("Recalculating")
        try:
            xl.CalculateFull()
        except Exception:
            pass
        out_path.parent.mkdir(parents=True, exist_ok=True)
        if out_path.exists():
            out_path.unlink()
        wb.SaveAs(str(out_path), FileFormat=_xl_file_format(out_path))
        result["saved"] = str(out_path)
        wb.Close(SaveChanges=False)
        wb = None
        return result
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        wb = None
        _close_excel(xl, None, pid)
        xl = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _reads_area(formula: str, own_sheet: str, area: dict[str, Any], partly: bool = True) -> bool:
    """Whether any reference of the formula touches the area (fully or partly); a span of sheets
    over the area's cells counts, a reference into another workbook does not."""
    arect = (area["r1"], area["c1"], area["r2"], area["c2"])
    for r in _refs(formula, own_sheet):
        if r["external"]:
            continue
        ref = r["ref"]
        if not _intersects(arect, (ref["r1"], ref["c1"], ref["r2"], ref["c2"])):
            continue
        if r["three_d"] or normalize_name(r["sheet"]) == normalize_name(area["sheet"]):
            return True
    return False
