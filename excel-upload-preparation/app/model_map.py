"""Model map -- where a workbook's inputs, loop driver, run settings and outputs are.

Everything here is read off the workbook itself; nothing is guessed:

- its VBA (through olevba): each `For` loop, the cell the macro sets on every pass
  (the loop driver), the cells that give the first and last value (the bounds), the
  ranges it copies results into (the outputs), the named cells it tests (`If
  Range("Full_Run?") = "No"`: the switches) and the areas it clears before a run;
- its defined names: a single-cell name on a constant that formulas read is a setting,
  a name the macros use says what a range is for;
- its data validations: a list dropdown is a switch, with its options;
- its sheet names (Inputs, Data, Assumptions, Output, Results, Reporting ...);
- its formulas: the constant cells the calculation reads most through absolute
  references (the key of a model that has no macro loop, e.g. a line-of-business selector cell).

Each item carries its evidence (the macro line, the name, the dropdown) so a person can
check it. The web app shows the map on its "Model map" screen and jumps to any cell.

    python -m app.model_map <workbook.xlsm> [--json]
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from .formula_utils import REF_RE, col_to_num, num_to_col, parse_ref

# --- words that say what a sheet is for -----------------------------------------------------------
INPUT_WORDS = ("input", "data", "base", "assumption", "hypoth", "param", "tarif", "rate", "critere", "criteria", "yield", "curve", "table", "mapping", "population", "list")
OUTPUT_WORDS = ("output", "result", "report", "summary", "exhibit", "dashboard", "saved", "paste")
CALC_WORDS = ("calc", "projection", "model", "working", "cashflow", "cash flow", "decrement", "macro", "temp", "coc", "lob", "provider", "adjustment", "baseline", "cost")
SWITCH_WORDS = ("boolean", "choice", "model", "switch", "flag", "run", "scenario", "method", "option", "type", "basis", "select", "active", "full", "delete", "del_", "send", "mode", "indicator", "version")

# --- VBA ------------------------------------------------------------------------------------------
_FOR_RE = re.compile(r"^\s*For\s+(?P<var>[A-Za-z_]\w*)\s*=\s*(?P<start>.+?)\s+To\s+(?P<end>.+?)(?:\s+Step\s+.+)?\s*$", re.IGNORECASE)
_NEXT_RE = re.compile(r"^\s*Next\b", re.IGNORECASE)
_PROC_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function)\s+(?P<name>[A-Za-z_]\w*)\s*\(", re.IGNORECASE)
_END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function)\b", re.IGNORECASE)
_RANGE_NAME_RE = re.compile(r"""Range\(\s*"(?P<name>[^"]+)"\s*\)""", re.IGNORECASE)
_SHEET_RANGE_RE = re.compile(r"""(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*Range\(\s*(?P<arg>[^()]*(?:\([^()]*\)[^()]*)*)\)""", re.IGNORECASE)
_SHEET_CELLS_RE = re.compile(r"""(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*Cells\(\s*(?P<row>\d+)\s*,\s*(?P<col>\d+)\s*\)""", re.IGNORECASE)
_IF_RANGE_RE = re.compile(r"""^\s*(?:If|ElseIf)\s+Range\(\s*"(?P<name>[^"]+)"\s*\)(?:\.Value2?)?\s*(?P<op>=|<>|>|<|>=|<=)\s*(?P<lit>"[^"]*"|[-\d.]+)""", re.IGNORECASE)
_CLEAR_RE = re.compile(r"""Range\(\s*"(?P<name>[^"]+)"\s*\)\s*\.ClearContents""", re.IGNORECASE)
_SHEET_CLEAR_RE = re.compile(r"""(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*Range\(\s*(?P<arg>[^()]*)\)\s*\.ClearContents""", re.IGNORECASE)
_COPY_RE = re.compile(r"""Range\(\s*"(?P<name>[^"]+)"\s*\)\s*\.Copy\s*$""", re.IGNORECASE)
_SET_RE = re.compile(r"""^\s*Set\s+(?P<var>[A-Za-z_]\w*)\s*=\s*Range\(\s*"(?P<name>[^"]+)"\s*\)""", re.IGNORECASE)
_VAR_FROM_RANGE_RE = re.compile(r"""^\s*(?:Dim\s+[^:]+:\s*)?(?P<var>[A-Za-z_]\w*)\s*=\s*(?P<expr>(?:Worksheets|Sheets|Range)\(.+?\)(?:\.Value2?)?)\s*$""", re.IGNORECASE)
_CALL_RE = re.compile(r"^\s*(?:Call\s+)?(?P<name>[A-Za-z_]\w*)\s*$", re.IGNORECASE)
_SAVEAS_RE = re.compile(r"SaveAs\s+Filename\s*:=\s*(?P<expr>[^,]+)", re.IGNORECASE)
_ASSIGN_RE = re.compile(r"^(?P<lhs>[^=<>]+?)\s*=\s*(?P<rhs>.+)$")


def read_vba(path: str | Path) -> list[tuple[str, str]]:
    """(module name, code) for every VBA module in the file; [] without macros or olevba."""
    try:
        from oletools.olevba import VBA_Parser
    except Exception:
        return []
    try:
        vp = VBA_Parser(str(path))
        out: list[tuple[str, str]] = []
        if vp.detect_vba_macros():
            for _fn, _stream, vba_filename, code in vp.extract_macros():
                code = "\n".join(ln for ln in (code or "").splitlines() if not ln.startswith("Attribute "))
                if code.strip():
                    out.append((Path(vba_filename).stem, code))
        vp.close()
        return out
    except Exception:
        return []


def _split_statements(line: str) -> list[str]:
    """`Dim x As Long: x = Range("A").Value` -> both statements (quotes respected)."""
    parts, cur, in_q = [], [], False
    for i, ch in enumerate(line):
        if ch == '"':
            in_q = not in_q
        if ch == ":" and not in_q and line[i + 1:i + 2] != "=":  # `:=` is a named argument, not a separator
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return [p for p in (x.strip() for x in parts) if p]


def _strip_comment(line: str) -> str:
    out, in_q = [], False
    for ch in line:
        if ch == '"':
            in_q = not in_q
        if ch == "'" and not in_q:
            break
        out.append(ch)
    return "".join(out).rstrip()


def _procedures(code: str) -> list[dict[str, Any]]:
    """Each Sub/Function with its statements (comments dropped, `:` split, continuations joined)."""
    lines = code.splitlines()
    joined: list[str] = []
    buf = ""
    for raw in lines:
        s = _strip_comment(raw)
        if s.rstrip().endswith("_"):
            buf += s.rstrip()[:-1] + " "
            continue
        joined.append(buf + s)
        buf = ""
    procs: list[dict[str, Any]] = []
    cur: dict[str, Any] | None = None
    for ln in joined:
        m = _PROC_RE.match(ln)
        if m:
            cur = {"name": m.group("name"), "kind": m.group(1).lower(), "statements": []}
            procs.append(cur)
            continue
        if _END_PROC_RE.match(ln):
            cur = None
            continue
        if cur is not None:
            cur["statements"].extend(_split_statements(ln))
    return procs


class _Names:
    """Defined names -> 'Sheet!Ref', and the reverse."""

    def __init__(self, defined_names: list[dict[str, Any]]):
        self.by_name: dict[str, str] = {}
        self.canonical: dict[str, str] = {}
        for d in defined_names or []:
            n, v = d.get("name"), str(d.get("value") or "").strip()
            if n and v and "#REF!" not in v and not v.startswith("="):
                self.by_name[str(n).upper()] = v
                self.canonical[str(n).upper()] = str(n)
        self.by_ref: dict[str, str] = {}
        for n, v in self.by_name.items():
            self.by_ref.setdefault(v.replace("$", "").upper(), self.canonical[n])

    def ref(self, name: str) -> str | None:
        return self.by_name.get(name.upper())

    def pretty(self, name: str) -> str:
        return self.canonical.get(name.upper(), name)

    def name_of(self, sheet: str, ref: str) -> str | None:
        key = f"{sheet}!{ref}".upper()
        n = self.by_ref.get(key) or self.by_ref.get(f"'{sheet}'!{ref}".upper())
        return n


def _sheet_ref(sheet: str, ref: str) -> str:
    sheet = sheet.strip("'")
    q = f"'{sheet}'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet
    return f"{q}!{ref.replace('$', '')}"


def _range_text_pieces(arg: str, var: str | None) -> tuple[str, bool]:
    """`"A" & ins & ":NI" & ins` -> ('A<ins>:NI<ins>', True): the literal pieces with the loop
    variable shown in angle brackets; False when the variable is not in it."""
    pieces = re.findall(r'"([^"]*)"|([A-Za-z_]\w*)', arg)
    text, has_var = [], False
    for lit, ident in pieces:
        if lit:
            text.append(lit)
        elif ident:
            if var and ident.lower() == var.lower():
                text.append(f"<{var}>")
                has_var = True
            elif ident.lower() not in ("value", "value2", "offset", "cells", "range", "worksheets", "sheets", "resize"):
                text.append(f"<{ident}>")
    return "".join(text), has_var


def _range_refs(expr: str, names: _Names, var: str | None, local: dict[str, str]) -> list[dict[str, Any]]:
    """Every range the expression touches, in order: {'text','cell','name','sheet','per_item'}."""
    out: list[dict[str, Any]] = []
    for m in _SHEET_CELLS_RE.finditer(expr):
        ref = f"{num_to_col(int(m.group('col')))}{m.group('row')}"
        out.append({"text": m.group(0), "cell": _sheet_ref(m.group("sheet"), ref), "name": names.name_of(m.group("sheet"), ref), "sheet": m.group("sheet"), "per_item": False})
    for m in _SHEET_RANGE_RE.finditer(expr):
        text, has_var = _range_text_pieces(m.group("arg"), var)
        if '"' in m.group("arg") and not has_var and "<" not in text:
            cell = _sheet_ref(m.group("sheet"), text)
            out.append({"text": m.group(0), "cell": cell, "name": names.name_of(m.group("sheet"), text), "sheet": m.group("sheet"), "per_item": False})
        else:
            out.append({"text": m.group(0), "cell": _sheet_ref(m.group("sheet"), text), "name": None, "sheet": m.group("sheet"), "per_item": has_var})
    for m in _RANGE_NAME_RE.finditer(expr):
        if any(m.group(0) in o["text"] for o in out):
            continue
        nm = m.group("name")
        ref = names.ref(nm)
        if ref is None and re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?", nm, re.IGNORECASE):
            ref = nm  # a plain address on the active sheet
            nm = None
        after = expr[m.end():m.end() + 40]
        per_item = bool(var) and bool(re.match(r"\s*\.\s*(Offset|Cells)\(", after)) and var.lower() in expr[m.end():].lower()
        out.append({"text": m.group(0), "cell": ref.replace("$", "") if ref else None, "name": nm, "sheet": ref.split("!")[0].strip("'") if ref and "!" in ref else None, "per_item": per_item})
    # variables that stand for a range (Set x = Range("...")); the target of `x.Resize(...).Value2 = ...` is x, so it goes first
    for v, nm in local.items():
        if re.search(rf"(?<![\w.]){re.escape(v)}(?![\w(])", expr):
            ref = names.ref(nm)
            item = {"text": v, "cell": ref.replace("$", "") if ref else None, "name": nm, "sheet": ref.split("!")[0].strip("'") if ref and "!" in ref else None, "per_item": bool(var) and var.lower() in expr.lower()}
            if expr.strip().lower().startswith(v.lower()):
                out.insert(0, item)
            else:
                out.append(item)
    return out


def _bound(expr: str, names: _Names, vars_from: dict[str, str]) -> dict[str, Any]:
    """What a loop bound is: a literal, a named cell, a sheet cell, or a variable set from one."""
    e = expr.strip()
    while e.startswith("(") and e.endswith(")"):
        e = e[1:-1].strip()
    if re.fullmatch(r"-?\d+", e):
        return {"text": e, "kind": "literal", "value": int(e)}
    refs = _range_refs(e, names, None, {})
    if refs:
        r = refs[0]
        return {"text": e, "kind": "cell", "cell": r["cell"], "name": r["name"]}
    src = vars_from.get(e.lower())
    if src:
        b = _bound(src, names, {})
        b["text"] = f"{e} = {src}"
        return b
    return {"text": e, "kind": "expression"}


def parse_macros(modules: list[tuple[str, str]], defined_names: list[dict[str, Any]]) -> dict[str, Any]:
    """What the macros do with the workbook: loops (driver cells, bounds, outputs), switches,
    cleared areas, copy/paste outputs, save paths. Pure text analysis, no Excel."""
    names = _Names(defined_names)
    procs_by_name: dict[str, dict[str, Any]] = {}
    modules_of: dict[str, str] = {}
    for mod, code in modules:
        for p in _procedures(code):
            procs_by_name[p["name"].lower()] = p
            modules_of[p["name"].lower()] = mod
    loops: list[dict[str, Any]] = []
    written: dict[str, set[str]] = defaultdict(set)  # name -> procedures that assign it
    switches: dict[str, dict[str, Any]] = {}
    cleared: list[dict[str, Any]] = []
    pastes: list[dict[str, Any]] = []
    saves: list[dict[str, Any]] = []
    calls: dict[str, list[str]] = defaultdict(list)

    for pname, proc in procs_by_name.items():
        stmts = proc["statements"]
        vars_from: dict[str, str] = {}
        local: dict[str, str] = {}
        stack: list[dict[str, Any]] = []
        last_copy: str | None = None
        for st in stmts:
            m = _SET_RE.match(st)
            if m:
                local[m.group("var")] = m.group("name")
                continue
            m = _VAR_FROM_RANGE_RE.match(st)
            if m and not _FOR_RE.match(st):
                vars_from[m.group("var").lower()] = m.group("expr")
            mw = re.match(r"""^\s*Range\(\s*"(?P<name>[^"]+)"\s*\)(?:\.Value2?)?\s*=\s*(?P<rhs>.+)$""", st, re.IGNORECASE)
            if mw and not stack:
                written[mw.group("name").lower()].add(proc["name"])
            m = _IF_RANGE_RE.match(st)
            if m:
                nm = m.group("name")
                sw = switches.setdefault(nm.lower(), {"name": nm, "cell": (names.ref(nm) or "").replace("$", "") or None, "compared_with": [], "in": set()})
                lit = m.group("lit").strip('"')
                if lit not in sw["compared_with"]:
                    sw["compared_with"].append(lit)
                sw["in"].add(proc["name"])
            for m in _CLEAR_RE.finditer(st):
                nm = m.group("name")
                cleared.append({"name": nm, "cell": (names.ref(nm) or "").replace("$", "") or None, "in": proc["name"]})
            m = _SHEET_CLEAR_RE.search(st)
            if m:
                text, _ = _range_text_pieces(m.group("arg"), None)
                cleared.append({"name": None, "cell": _sheet_ref(m.group("sheet"), text), "in": proc["name"]})
            m = _COPY_RE.search(st)
            if m:
                last_copy = m.group("name")
            if last_copy and re.search(r"PasteSpecial", st, re.IGNORECASE):
                # the paste goes where the last Range(...).Select pointed
                pastes.append({"from_name": last_copy, "from": (names.ref(last_copy) or "").replace("$", "") or None, "in": proc["name"], "to_name": None, "to": None})
                last_copy = None
            if re.search(r"""Range\(\s*"([^"]+)"\s*\)\s*\.Select""", st, re.IGNORECASE) and pastes and pastes[-1]["to"] is None and pastes[-1]["in"] == proc["name"]:
                pass
            m = re.search(r"""Range\(\s*"([^"]+)"\s*\)\s*\.Select""", st, re.IGNORECASE)
            if m:
                last_select = m.group(1)
                # a PasteSpecial right after a Select pastes there
                proc.setdefault("_last_select", None)
                proc["_last_select"] = last_select
            if re.search(r"PasteSpecial", st, re.IGNORECASE) and pastes and pastes[-1]["to"] is None and proc.get("_last_select"):
                sel = proc["_last_select"]
                pastes[-1]["to_name"] = sel
                pastes[-1]["to"] = (names.ref(sel) or "").replace("$", "") or None
            m = _SAVEAS_RE.search(st)
            if m:
                refs = _range_refs(m.group("expr"), names, None, local)
                src = vars_from.get(m.group("expr").strip().lower())
                if not refs and src:
                    refs = _range_refs(src, names, None, local)
                saves.append({"in": proc["name"], "path_from": refs[0]["cell"] if refs else None, "name": refs[0]["name"] if refs else None})
            m = _CALL_RE.match(st)
            if m and m.group("name").lower() in procs_by_name and m.group("name").lower() != pname:
                calls[proc["name"]].append(m.group("name"))
            m = _FOR_RE.match(st)
            if m:
                loop = {"procedure": proc["name"], "module": modules_of.get(pname), "var": m.group("var"), "from": _bound(m.group("start"), names, vars_from), "to": _bound(m.group("end"), names, vars_from),
                        "drivers": [], "outputs": [], "depth": len(stack), "lines": 0, "copies_sheet": None, "builds_sheets": False}
                stack.append(loop)
                loops.append(loop)
                continue
            mc = re.search(r"""(?:Worksheets|Sheets)\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.Copy\b""", st, re.IGNORECASE)
            if mc and stack:
                stack[-1]["copies_sheet"] = mc.group("sheet")
                stack[-1]["builds_sheets"] = True
                continue
            if stack and re.search(r"""(?:Worksheets|Sheets)\(\s*"[^"]+"\s*\)\s*\.Name\s*=""", st, re.IGNORECASE):
                stack[-1]["builds_sheets"] = True
                continue
            if _NEXT_RE.match(st):
                if stack:
                    stack.pop()
                continue
            if not stack:
                continue
            # inside one or more loops: what does this statement do with the loop variable?
            a = _ASSIGN_RE.match(st)
            if not a or st.lower().startswith(("if ", "elseif ", "for ", "dim ")):
                continue
            lhs, rhs = a.group("lhs").strip(), a.group("rhs").strip()
            if not re.search(r"Range\(|Cells\(|Worksheets\(|Sheets\(", lhs, re.IGNORECASE) and lhs not in local:
                continue
            for loop in stack:
                loop["lines"] += 1
            innermost = stack[-1]
            var = innermost["var"]
            vars_in_scope = [l["var"] for l in stack]
            l_refs = _range_refs(lhs, names, var, local)
            r_refs = _range_refs(rhs, names, var, local)
            rhs_uses_var = any(re.search(rf"(?<!\w){re.escape(v)}(?!\w)", rhs) for v in vars_in_scope)
            lhs_uses_var = any(re.search(rf"(?<!\w){re.escape(v)}(?!\w)", lhs) for v in vars_in_scope)
            which = next((l for l in reversed(stack) if re.search(rf"(?<!\w){re.escape(l['var'])}(?!\w)", lhs + " " + rhs)), innermost)
            if l_refs and not lhs_uses_var and rhs_uses_var:
                # the macro sets this cell from the loop value (or from a list indexed by it): the driver
                cell0 = l_refs[0]["cell"]
                if cell0 and "!" not in cell0 and which.get("copies_sheet"):
                    cell0 = _sheet_ref(which["copies_sheet"], cell0)  # on the sheet just copied (and on every copy of it)
                d = {"cell": cell0, "name": l_refs[0]["name"], "set_from": None, "set_from_name": None, "statement": st.strip(), "template_sheet": which.get("copies_sheet")}
                src = next((r for r in r_refs if r["cell"] and r["cell"] != d["cell"]), None)
                if src:
                    d["set_from"], d["set_from_name"] = src["cell"], src["name"]
                    d["how"] = "the loop value picks one row/column of " + (src["name"] or src["cell"]) if src["per_item"] else "each loop value, through " + (src["name"] or src["cell"])
                elif re.fullmatch(rf"\s*{re.escape(var)}\s*", rhs):
                    d["how"] = "the loop value itself"
                else:
                    d["how"] = "an expression of the loop value"
                which["drivers"].append(d)
            elif l_refs and lhs_uses_var and r_refs and not any(r["per_item"] for r in r_refs):
                # one row/block per loop value on the left, a fixed block on the right: an output
                src = r_refs[0]
                which["outputs"].append({"to": l_refs[0]["cell"], "to_name": l_refs[0]["name"], "from": src["cell"], "from_name": src["name"], "statement": st.strip(), "per_item": True})
            elif l_refs and r_refs and not rhs_uses_var and not lhs_uses_var and any(r["cell"] for r in r_refs) and l_refs[0]["cell"] != r_refs[0]["cell"]:
                # a fixed block copied to a fixed block inside the loop: an output written each pass (same place)
                which["outputs"].append({"to": l_refs[0]["cell"], "to_name": l_refs[0]["name"], "from": r_refs[0]["cell"], "from_name": r_refs[0]["name"], "statement": st.strip(), "per_item": False})
            elif l_refs and lhs_uses_var and r_refs and any(r["per_item"] for r in r_refs):
                # both sides move with the loop value: a block per item on both sides (an output table)
                src = r_refs[0]
                which["outputs"].append({"to": l_refs[0]["cell"], "to_name": l_refs[0]["name"], "from": src["cell"], "from_name": src["name"], "statement": st.strip(), "per_item": True})
    for sw in switches.values():
        sw["in"] = sorted(sw["in"])
    # which procedure is the entry point: the one that calls others and is called by none
    called = {c.lower() for lst in calls.values() for c in lst}
    runners = {l["procedure"].lower() for l in loops if any(d["cell"] for d in l["drivers"]) and not l["builds_sheets"]}

    def reaches_runner(name: str, seen: set[str]) -> bool:
        if name.lower() in runners:
            return True
        if name.lower() in seen:
            return False
        seen.add(name.lower())
        return any(reaches_runner(c, seen) for c in calls.get(name, []))

    entry = [p["name"] for p in procs_by_name.values() if p["name"].lower() not in called and reaches_runner(p["name"], set())]
    status = []
    real_switches = []
    for sw in switches.values():
        procs_writing = written.get(sw["name"].lower(), set())
        if procs_writing:
            status.append({**sw, "written_in": sorted(procs_writing)})
        else:
            real_switches.append(sw)
    return {"procedures": [{"name": p["name"], "module": modules_of.get(n), "statements": len(p["statements"]), "calls": calls.get(p["name"], [])} for n, p in procs_by_name.items()],
            "entry_points": entry, "loops": loops, "switches": real_switches, "status": status, "cleared": cleared, "pastes": pastes, "saves": saves,
            "written": {k: sorted(v) for k, v in written.items()}}  # keys in lower case


# --- the workbook side ----------------------------------------------------------------------------
def sheet_role(name: str, dims: str | None = None) -> str:
    n = name.lower().strip()
    if n.startswith(">>") or n.endswith("-->") or n.endswith(">>>") or n.endswith(">>") or (dims and dims.upper() == "A1:A1"):
        return "divider"
    if any(w in n for w in OUTPUT_WORDS):
        return "output"
    if any(w in n for w in INPUT_WORDS):
        return "input"
    if any(w in n for w in CALC_WORDS):
        return "calculation"
    return "other"


def _validations(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    """Data validations of the analysis copy: list dropdowns with their options, number limits."""
    from .inventory import open_cached

    out: list[dict[str, Any]] = []
    try:
        wb = open_cached(analysis)
    except Exception:
        return out
    names = _Names(analysis["workbooks"][0].get("defined_names", []))
    for ws in wb.worksheets:
        dvs = getattr(getattr(ws, "data_validations", None), "dataValidation", None) or []
        for dv in dvs:
            try:
                kind = dv.type
                f1 = str(dv.formula1 or "")
                ranges = str(dv.sqref)
            except Exception:
                continue
            if not kind:
                continue
            item: dict[str, Any] = {"sheet": ws.title, "ranges": ranges, "type": kind, "formula": f1}
            if kind == "list":
                if f1.startswith('"') and f1.endswith('"'):
                    item["options"] = [x.strip() for x in f1.strip('"').split(",") if x.strip()]
                    item["source"] = "a list written into the validation"
                else:
                    ref = f1.lstrip("=")
                    target = names.ref(ref) or (ref if "!" in ref else f"{ws.title}!{ref}")
                    item["source"] = target.replace("$", "")
                    item["options"] = _cells_text(wb, target, analysis)[:40]
            out.append(item)
    return out


def _cells_text(wb, target: str, analysis: dict[str, Any] | None = None) -> list[str]:
    """The texts in a range: constants as they are, formula cells by the value Excel stored."""
    from .inventory import values_cell

    r = parse_ref(target)
    if not r or r["whole_column"] or r["whole_row"]:
        return []
    sheet = r["sheet"]
    if not sheet or sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    out = []
    for row in range(r["r1"], min(r["r2"], r["r1"] + 200) + 1):
        for col in range(r["c1"], min(r["c2"], r["c1"] + 200) + 1):
            v = ws.cell(row=row, column=col).value
            if isinstance(v, str) and v.startswith("=") and analysis is not None:
                try:
                    v = values_cell(analysis, sheet, row, col)
                except Exception:
                    v = None
            if v is not None and str(v) != "" and not str(v).startswith("="):
                out.append(str(v))
    return out


def _formula_counts(analysis: dict[str, Any]) -> tuple[Counter, Counter, dict[str, int]]:
    """(absolute single-cell references counted, defined names used by formulas, formulas per sheet)."""
    wb0 = analysis["workbooks"][0]
    names = _Names(wb0.get("defined_names", []))
    abs_refs: Counter = Counter()
    name_uses: Counter = Counter()
    per_sheet: dict[str, int] = Counter()
    name_tokens = {n for n in names.by_name}
    tok_re = re.compile(r"(?<![A-Za-z0-9_.!'\[])([A-Za-z_][A-Za-z0-9_.?]*)(?![A-Za-z0-9_(!])")
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        per_sheet[f["sheet"]] += 1
        for m in REF_RE.finditer(formula):
            ref = m.group("ref")
            if ":" in ref or "$" not in ref:
                continue
            sheet = (m.group("sheet") or f["sheet"]).strip("'") if "sheet" in m.groupdict() else f["sheet"]
            abs_refs[(sheet, ref.replace("$", ""))] += 1
        if name_tokens:
            for m in tok_re.finditer(formula):
                t = m.group(1).upper()
                if t in name_tokens:
                    name_uses[t] += 1
    return abs_refs, name_uses, dict(per_sheet)


def _constant_at(analysis: dict[str, Any], sheet: str, ref: str) -> Any:
    from .inventory import cell_value

    r = parse_ref(ref)
    if not r:
        return None
    try:
        v = cell_value(analysis, sheet, r["r1"], r["c1"])
    except Exception:
        return None
    return None if isinstance(v, str) and v.startswith("=") else v


def _held(analysis: dict[str, Any], sheet: str, ref: str) -> Any:
    """What the cell holds: a formula text, a constant, or None."""
    from .inventory import cell_value

    r = parse_ref(ref)
    if not r:
        return None
    try:
        return cell_value(analysis, sheet, r["r1"], r["c1"])
    except Exception:
        return None


def _label_near(analysis: dict[str, Any], sheet: str, ref: str) -> str | None:
    """The text to the left of (or above) a cell: what a person calls it."""
    from .inventory import cell_value

    r = parse_ref(ref)
    if not r:
        return None
    for dc in (1, 2, 3):
        if r["c1"] - dc >= 1:
            v = cell_value(analysis, sheet, r["r1"], r["c1"] - dc)
            if isinstance(v, str) and v.strip() and not v.startswith("="):
                return v.strip()[:80]
    if r["r1"] > 1:
        v = cell_value(analysis, sheet, r["r1"] - 1, r["c1"])
        if isinstance(v, str) and v.strip() and not v.startswith("="):
            return v.strip()[:80]
    return None


def _labelled_settings(analysis: dict[str, Any], sheets: list[dict[str, Any]], per_sheet: dict[str, int], taken: set[str]) -> list[dict[str, Any]]:
    """On small input/control sheets (few formulas, few cells): 'label | value' pairs -- a text cell with a
    constant right next to it (a parameters sheet, a macro control sheet, a model inputs sheet)."""
    from .inventory import open_cached

    out: list[dict[str, Any]] = []
    try:
        wb = open_cached(analysis)
    except Exception:
        return out
    doc_words = ("read", "readme", "instruction", "cover", "index", "caveat", "version", "reliance", "limitation", "note", "contents", "log")
    for s in sheets:
        if s["role"] not in ("input", "other", "calculation") or s["name"] not in wb.sheetnames:
            continue
        if any(w in s["name"].lower() for w in doc_words):
            continue
        ws = wb[s["name"]]
        if ws.max_row > 400 or ws.max_column > 60 or per_sheet.get(s["name"], 0) > 400:
            continue
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), max_col=min(ws.max_column, 60)))
        filled = sum(1 for row in rows for c in row if c.value is not None and str(c.value) != "")
        if filled > 400 or filled == 0:
            continue  # a data table or a big input sheet, not a control panel
        pairs: list[dict[str, Any]] = []

        tiny = filled <= 40  # a control panel: a few label | value lines, text values count too

        def ok_value(w: Any) -> bool:
            if isinstance(w, bool) or isinstance(w, (int, float)):
                return True
            return tiny and isinstance(w, str) and w.strip() != "" and not w.startswith("=") and len(w.strip()) <= 25

        grid = {(c.row, c.column): c.value for row in rows for c in row if c.value is not None and str(c.value) != ""}
        for row in rows:
            cells = [c for c in row if c.value is not None and str(c.value) != ""]
            for i, c in enumerate(cells):
                v = c.value
                if not (isinstance(v, str) and not v.startswith("=") and 2 <= len(v.strip()) <= 80):
                    continue
                nxt = cells[i + 1] if i + 1 < len(cells) else None
                if nxt is not None and nxt.column - c.column <= 3 and ok_value(nxt.value):
                    pairs.append({"sheet": s["name"], "cell": _sheet_ref(s["name"], nxt.coordinate), "label": v.strip(), "value": nxt.value if not isinstance(nxt.value, float) or abs(nxt.value) < 1e15 else str(nxt.value)})
                    break
                if nxt is None or nxt.column - c.column > 3:
                    # a label with the value under it (the text on one row, the number on the next)
                    for dc in (0, 1, 2):
                        w = grid.get((c.row + 1, c.column + dc))
                        if isinstance(w, (int, float)) and not isinstance(w, bool):
                            pairs.append({"sheet": s["name"], "cell": _sheet_ref(s["name"], f"{num_to_col(c.column + dc)}{c.row + 1}"), "label": v.strip(), "value": w if abs(w) < 1e15 else str(w)})
                            break
        counts = Counter(p["label"].lower() for p in pairs)
        for pr in pairs:
            if counts[pr["label"].lower()] == 1 and pr["cell"].upper() not in taken:
                out.append(pr)
    return out


def build_model_map(analysis: dict[str, Any]) -> dict[str, Any]:
    """The map: loop, inputs, run settings, outputs -- each item with sheet!ref and its evidence."""
    wb0 = analysis["workbooks"][0]
    copy_path = wb0["copy_path"]
    defined_names = wb0.get("defined_names", []) or []
    names = _Names(defined_names)
    sheets = [{"name": s["name"], "dims": s.get("dimensions") or s.get("dims"), "state": s.get("state") or s.get("sheet_state"), "grids": len(s.get("grids", []))} for s in wb0.get("sheets", [])]
    for s in sheets:
        s["role"] = sheet_role(s["name"], s["dims"])
    modules = read_vba(copy_path)
    vba = parse_macros(modules, defined_names) if modules else {"procedures": [], "entry_points": [], "loops": [], "switches": [], "status": [], "cleared": [], "pastes": [], "saves": [], "written": {}}
    validations = _validations(analysis)
    abs_refs, name_uses, per_sheet = _formula_counts(analysis)

    notes: list[str] = []
    out: dict[str, Any] = {"workbook": Path(copy_path).name, "has_macros": bool(modules), "sheets": sheets, "notes": notes}

    # --- the loop ----------------------------------------------------------------------------------
    drivers: list[dict[str, Any]] = []
    bounds: list[dict[str, Any]] = []
    loop_outputs: list[dict[str, Any]] = []
    seen_cells: set[str] = set()
    builders: list[dict[str, Any]] = []
    for loop in vba["loops"]:
        if loop["builds_sheets"]:
            for d in loop["drivers"]:
                if d["cell"]:
                    tpl = d.get("template_sheet")
                    present = tpl in {x["name"] for x in sheets} if tpl else False
                    builders.append({"cell": d["cell"], "name": d["name"], "macro": loop["procedure"], "loop_var": loop["var"], "template_sheet": tpl, "template_present": present,
                                     "how": (f"the macro {loop['procedure']} copies the sheet {tpl!r} once per loop value and writes the item into this cell on the copy" + ("" if present else f" (no sheet {tpl!r} in the workbook now: the macro was written for an earlier layout)")) if tpl else f"set by the sheet-building macro {loop['procedure']}",
                                     "evidence": f"{loop['procedure']}: For {loop['var']} = {loop['from']['text']} To {loop['to']['text']} ... {d['statement'][:120]}"})
            continue
        for d in loop["drivers"]:
            if not d["cell"] or d["cell"] in seen_cells:
                continue
            seen_cells.add(d["cell"])
            sheet = d["cell"].split("!")[0].strip("'") if "!" in d["cell"] else None
            ref = d["cell"].split("!")[-1]
            drivers.append({**d, "label": _label_near(analysis, sheet, ref) if sheet and ":" not in ref else None,
                            "macro": loop["procedure"], "loop_var": loop["var"], "nested_depth": loop["depth"],
                            "from": loop["from"], "to": loop["to"],
                            "evidence": f"{loop['procedure']}: For {loop['var']} = {loop['from']['text']} To {loop['to']['text']} ... {d['statement'][:140]}"})
        for b, role in ((loop["from"], "first"), (loop["to"], "last")):
            if b["kind"] == "cell" and b.get("cell") and (b["cell"], role) not in {(x["cell"], x["role"]) for x in bounds}:
                sheet = b["cell"].split("!")[0].strip("'")
                ref_only = b["cell"].split("!")[-1]
                held = _held(analysis, sheet, ref_only)
                bounds.append({"role": role, "cell": b["cell"], "name": b.get("name"), "label": _label_near(analysis, sheet, ref_only), "value": _constant_at(analysis, sheet, ref_only),
                               "formula": held if isinstance(held, str) and held.startswith("=") else None, "macro": loop["procedure"], "loop_var": loop["var"]})
        for o in loop["outputs"]:
            key = (o["to"], o["from"])
            if o["to"] and key not in {(x["to"], x["from"]) for x in loop_outputs}:
                loop_outputs.append({**o, "macro": loop["procedure"], "loop_var": loop["var"], "when": "each loop pass"})
    # a driver whose own sheet is a per-item row (Calc!A2:NI2 <- row n of a policy table)
    per_item_inputs: list[dict[str, Any]] = []
    for d in drivers:
        if d.get("set_from") and "<" in str(d["set_from"]):
            per_item_inputs.append({"cell": d["set_from"], "what": f"one row per loop value, copied into {d['cell']}", "macro": d["macro"]})
    entry = vba["entry_points"]
    out["loop"] = {
        "found": bool(drivers),
        "drivers": drivers, "bounds": bounds, "entry_points": entry,
        "macros": [{"procedure": l["procedure"], "module": l["module"], "var": l["var"], "from": l["from"]["text"], "to": l["to"]["text"], "depth": l["depth"], "drivers": [d["cell"] for d in l["drivers"]], "outputs": len(l["outputs"])} for l in vba["loops"] if l["drivers"] or l["outputs"]],
        "how_to_run": None, "key_cells": [], "item_cells": builders,
    }
    if drivers:
        starts = [b for b in bounds if b["role"] == "first"]
        ends = [b for b in bounds if b["role"] == "last"]
        how = f"Run the macro {entry[0] if entry else drivers[0]['macro']}" + (f" (it calls {', '.join(vba['procedures'][[p['name'] for p in vba['procedures']].index(entry[0])]['calls'])})" if entry and vba["procedures"][[p["name"] for p in vba["procedures"]].index(entry[0])]["calls"] else "") + "."
        if starts or ends:
            how += " It goes from " + " / ".join(f"{b['cell']}" + (f" ({b['name']})" if b["name"] else "") for b in starts) + " to " + " / ".join(f"{b['cell']}" + (f" ({b['name']})" if b["name"] else "") for b in ends) + ": change those cells to run another range."
        how += " Each pass sets " + ", ".join(f"{d['cell']}" + (f" ({d['name']})" if d["name"] else "") for d in drivers) + "."
        out["loop"]["how_to_run"] = how
    else:
        # no macro loop: the key cell is the constant the calculation reads most through $A$1 references
        key = []
        taken_cols: dict[tuple[str, int], dict[str, Any]] = {}
        for (sheet, ref), n in abs_refs.most_common(400):
            v = _constant_at(analysis, sheet, ref)
            if v is None or v == "" or isinstance(v, bool):
                continue
            r = parse_ref(ref)
            col_key = (sheet, r["c1"]) if r else None
            # consecutive keys in one column (a table's row keys) are one entry
            prev = taken_cols.get(col_key) if col_key else None
            if prev and r and prev["_last_row"] + 1 == r["r1"] and abs(prev["readers"] - n) <= max(2, n // 10):
                prev["_last_row"] = r["r1"]
                prev["cells"] = prev["cells"] + 1
                prev["cell"] = _sheet_ref(sheet, f"{num_to_col(r['c1'])}{prev['_first_row']}:{num_to_col(r['c1'])}{r['r1']}")
                continue
            item = {"cell": _sheet_ref(sheet, ref), "value": v if not isinstance(v, (int, float)) or abs(v) < 1e15 else str(v), "readers": n, "label": _label_near(analysis, sheet, ref), "name": names.name_of(sheet, ref),
                    "cells": 1, "_first_row": r["r1"] if r else 0, "_last_row": r["r1"] if r else 0}
            key.append(item)
            if col_key:
                taken_cols[col_key] = item
            if len(key) >= 10:
                break
        for k in key:
            k.pop("_first_row", None)
            k.pop("_last_row", None)
            if k["cells"] > 1:
                k["label"] = (k["label"] + " ... " if k.get("label") else "") + f"a column of {k['cells']} keys, each read by about {k['readers']} formulas"
        out["loop"]["key_cells"] = key
        if modules:
            notes.append("The macros hold no loop that sets a cell on every pass; the key cells below are the constants the formulas read most (by absolute reference).")
        else:
            notes.append("No macro in the workbook; the key cells below are the constants the formulas read most (by absolute reference).")

    # --- run settings: switches ----------------------------------------------------------------------
    switches: list[dict[str, Any]] = []
    seen_sw: set[str] = set()
    computed: list[dict[str, Any]] = []
    for sw in vba["switches"]:
        if not sw["cell"]:
            continue
        sheet, ref = sw["cell"].split("!")[0].strip("'"), sw["cell"].split("!")[-1]
        held = _held(analysis, sheet, ref)
        if isinstance(held, str) and held.startswith("="):
            computed.append({**sw, "formula": held})
            continue
        switches.append({"cell": sw["cell"], "name": sw["name"], "label": _label_near(analysis, sheet, ref), "value": _constant_at(analysis, sheet, ref), "options": sw["compared_with"],
                         "how": "typed or chosen in the cell; the macro tests it", "evidence": f"If Range(\"{sw['name']}\") = {', '.join(sw['compared_with'])} in {', '.join(sw['in'])}"})
        seen_sw.add(sw["cell"].upper())
    for v in validations:
        if v["type"] != "list":
            continue
        for rng in v["ranges"].split():
            r0 = parse_ref(rng)
            if r0 and not (r0["r1"] == r0["r2"] and r0["c1"] == r0["c2"]) and (r0["r2"] - r0["r1"] + 1) * (r0["c2"] - r0["c1"] + 1) <= 60:
                # a dropdown over a block: one switch per cell when the cells carry names (one named flag per row of G51:G60)
                cells = [f"{num_to_col(c)}{rw}" for rw in range(r0["r1"], r0["r2"] + 1) for c in range(r0["c1"], r0["c2"] + 1)]
                named = [(c, names.name_of(v["sheet"], c)) for c in cells]
                if sum(1 for _, n in named if n) >= len(cells) // 2:
                    for c, n in named:
                        cell = _sheet_ref(v["sheet"], c)
                        if cell.upper() in seen_sw:
                            for sw0 in switches:
                                if sw0["cell"].upper() == cell.upper():
                                    sw0["how"] = "a dropdown in the cell; the macro tests it"
                                    sw0["options"] = v.get("options") or sw0["options"]
                            continue
                        seen_sw.add(cell.upper())
                        switches.append({"cell": cell, "name": n, "label": _label_near(analysis, v["sheet"], c), "value": _constant_at(analysis, v["sheet"], c), "options": v.get("options", []),
                                         "how": "a dropdown in the cell", "evidence": f"data validation list on {v['sheet']}!{rng}: {v.get('source')}"})
                    continue
            cell = _sheet_ref(v["sheet"], rng)
            if cell.upper() in seen_sw:
                for sw0 in switches:
                    if sw0["cell"].upper() == cell.upper():
                        sw0["how"] = "a dropdown in the cell; the macro tests it"
                        sw0["options"] = v.get("options") or sw0["options"]
                continue
            seen_sw.add(cell.upper())
            r = parse_ref(rng)
            single = r and r["r1"] == r["r2"] and r["c1"] == r["c2"]
            switches.append({"cell": cell, "name": names.name_of(v["sheet"], rng) if single else None, "label": _label_near(analysis, v["sheet"], rng) if single else None,
                             "value": _constant_at(analysis, v["sheet"], rng) if single else None, "options": v.get("options", []), "how": "a dropdown in the cell" + ("" if single else "s"),
                             "evidence": f"data validation list on {v['sheet']}!{rng}: {v.get('source')}"})
    # named single constants that formulas read, with a switch-like name
    named_settings: list[dict[str, Any]] = []
    for n, ref in names.by_name.items():
        r = parse_ref(ref)
        if not r or not (r["r1"] == r["r2"] and r["c1"] == r["c2"]) or not r["sheet"]:
            continue
        uses = name_uses.get(n, 0)
        v = _constant_at(analysis, r["sheet"], f"{num_to_col(r['c1'])}{r['r1']}")
        if v is None or (isinstance(v, str) and (uses == 0 or len(v) > 40)):
            continue  # a text nobody reads is a caption, not a setting
        cell = _sheet_ref(r["sheet"], f"{num_to_col(r['c1'])}{r['r1']}")
        item = {"cell": cell, "name": names.pretty(n), "label": _label_near(analysis, r["sheet"], f"{num_to_col(r['c1'])}{r['r1']}"),
                "value": v if not isinstance(v, float) or abs(v) < 1e15 else str(v), "readers": uses, "switch_like": any(w in n.lower() for w in SWITCH_WORDS)}
        if cell.upper() in seen_sw:
            for s in switches:
                if s["cell"].upper() == cell.upper() and not s.get("name"):
                    s["name"] = item["name"]
            continue
        named_settings.append(item)
    named_settings.sort(key=lambda x: (-int(x["switch_like"]), -x["readers"], x["name"]))
    # --- assumption tables: input sheets by name, the names that cover a block, read by formulas ------
    tables: list[dict[str, Any]] = []
    for n, ref in names.by_name.items():
        r = parse_ref(ref)
        if not r or not r["sheet"] or r["whole_column"] or r["whole_row"]:
            continue
        if (r["r2"] - r["r1"] + 1) * (r["c2"] - r["c1"] + 1) < 4:
            continue
        uses = name_uses.get(n, 0)
        used_by_macro = any(x.get("to_name") == n or x.get("from_name") == n for x in loop_outputs) or any(c.get("name") == n for c in vba["cleared"])
        if uses == 0 and not used_by_macro:
            continue
        tables.append({"name": names.pretty(n), "cell": ref.replace("$", ""), "readers": uses, "role": sheet_role(r["sheet"])})
    tables.sort(key=lambda x: -x["readers"])
    input_sheets = [{"sheet": s["name"], "formulas": per_sheet.get(s["name"], 0), "state": s["state"], "why": "its name"} for s in sheets if s["role"] == "input"]
    for pi in per_item_inputs:
        sheet = pi["cell"].split("!")[0].strip("'")
        if not any(x["sheet"] == sheet for x in input_sheets):
            input_sheets.insert(0, {"sheet": sheet, "formulas": per_sheet.get(sheet, 0), "state": next((s["state"] for s in sheets if s["name"] == sheet), None), "why": "the macro reads one row of it per loop value"})
        else:
            for x in input_sheets:
                if x["sheet"] == sheet:
                    x["why"] = "the macro reads one row of it per loop value"
    for d in drivers:
        if d.get("set_from") and d.get("set_from_name"):
            sheet = d["set_from"].split("!")[0].strip("'") if "!" in d["set_from"] else None
            if sheet and not any(x["sheet"] == sheet for x in input_sheets):
                input_sheets.insert(0, {"sheet": sheet, "formulas": per_sheet.get(sheet, 0), "state": None, "why": f"the loop picks its values from {d['set_from_name']} there"})
    out["inputs"] = {"per_item": per_item_inputs, "sheets": input_sheets, "tables": [t for t in tables if t["role"] in ("input", "other")][:60]}
    labelled = _labelled_settings(analysis, sheets, per_sheet, {sw["cell"].upper() for sw in switches} | {x["cell"].upper() for x in named_settings})
    status_cells = [{"cell": x["cell"], "name": x["name"], "label": _label_near(analysis, x["cell"].split("!")[0].strip("'"), x["cell"].split("!")[-1]) if x["cell"] else None,
                     "value": _constant_at(analysis, x["cell"].split("!")[0].strip("'"), x["cell"].split("!")[-1]) if x["cell"] else None,
                     "written_in": x["written_in"], "evidence": f"the macro writes it ({', '.join(x['written_in'])}) and reads it back"} for x in vba["status"] if x["cell"]]
    status_names = {x["name"] for x in vba["status"]}
    output_names = {o.get("to_name") for o in loop_outputs} | {p.get("to_name") for p in vba["pastes"]} | {c.get("name") for c in vba["cleared"]}
    for c in computed:
        status_cells.append({"cell": c["cell"], "name": c["name"], "label": _label_near(analysis, c["cell"].split("!")[0].strip("'"), c["cell"].split("!")[-1]), "value": None,
                             "written_in": [], "evidence": f"holds the formula {c['formula'][:80]}: computed, the macro only reads it ({', '.join(c['in'])})"})
    written_lower = {k for k in vba["written"]}
    named_settings = [n for n in named_settings if n["name"] not in status_names and n["name"].lower() not in written_lower and n["name"] not in output_names]
    out["run_settings"] = {"switches": switches, "named_settings": named_settings[:80], "labelled": labelled, "status": status_cells, "tables": [t for t in tables if t["role"] == "input"][:60]}

    # --- outputs -------------------------------------------------------------------------------------
    outputs: list[dict[str, Any]] = []
    for o in loop_outputs:
        outputs.append({"cell": o["to"], "name": o["to_name"], "from": o["from"], "from_name": o["from_name"], "written_by": o["macro"], "when": o["when"] + (" (one block per loop value)" if o["per_item"] else " (same place, overwritten)"), "evidence": o["statement"][:160]})
    for p in vba["pastes"]:
        if p["to"] and not any(x["cell"] == p["to"] for x in outputs):
            outputs.append({"cell": p["to"], "name": p["to_name"], "from": p["from"], "from_name": p["from_name"], "written_by": p["in"], "when": "copy / paste values", "evidence": f"Range(\"{p['from_name']}\").Copy ... Range(\"{p['to_name']}\") PasteSpecial"})
    for c in vba["cleared"]:
        if c["cell"] and not any(x["cell"] == c["cell"] for x in outputs):
            outputs.append({"cell": c["cell"], "name": c["name"], "from": None, "from_name": None, "written_by": c["in"], "when": "cleared before a run (results area)", "evidence": "ClearContents"})
    output_sheets = [{"sheet": s["name"], "formulas": per_sheet.get(s["name"], 0), "why": "its name"} for s in sheets if s["role"] == "output"]
    for o in outputs:
        sheet = o["cell"].split("!")[0].strip("'") if o["cell"] and "!" in o["cell"] else None
        if sheet and not any(x["sheet"] == sheet for x in output_sheets):
            output_sheets.insert(0, {"sheet": sheet, "formulas": per_sheet.get(sheet, 0), "why": f"the macro {o['written_by']} writes results there"})
    out["outputs"] = {"cells": outputs, "sheets": output_sheets, "saves": vba["saves"]}
    out["macros"] = {"procedures": vba["procedures"], "entry_points": entry}
    return out


# --- words for a person -----------------------------------------------------------------------------
def describe(mp: dict[str, Any]) -> str:
    lines = [f"Model map of {mp['workbook']}" + ("" if mp["has_macros"] else " (no macros)")]
    lp = mp["loop"]
    lines.append("")
    lines.append("THE LOOP")
    if lp["found"]:
        for d in lp["drivers"]:
            lines.append(f"  driver  {d['cell']}" + (f"  [{d['name']}]" if d["name"] else "") + (f"  '{d['label']}'" if d.get("label") else "") + f"  <- {d['how']} (macro {d['macro']}, For {d['loop_var']})")
        for b in lp["bounds"]:
            lines.append(f"  {b['role']:5s}   {b['cell']}" + (f"  [{b['name']}]" if b["name"] else "") + (f"  '{b['label']}'" if b.get("label") else "") + (f"  holds {b['formula']}" if b.get("formula") else f"  = {b['value']!r}"))
        if lp["how_to_run"]:
            lines.append("  " + lp["how_to_run"])
    else:
        for n in mp["notes"]:
            lines.append("  " + n)
        for b in lp.get("item_cells", []):
            lines.append(f"  item cell  {b['cell']}" + (f"  [{b['name']}]" if b.get("name") else "") + f"  <- {b['how']}")
        for k in lp["key_cells"]:
            lines.append(f"  key cell  {k['cell']}" + (f"  [{k['name']}]" if k.get("name") else "") + f"  = {k['value']!r}  read by {k['readers']} formulas" + (f"  '{k['label']}'" if k.get("label") else ""))
    lines.append("")
    lines.append("INPUTS")
    for pi in mp["inputs"]["per_item"]:
        lines.append(f"  per item  {pi['cell']}  ({pi['what']})")
    for s in mp["inputs"]["sheets"]:
        lines.append(f"  sheet     {s['sheet']}  ({s['formulas']} formulas; {s['why']})")
    lines.append("")
    lines.append("RUN SETTINGS")
    for s in mp["run_settings"]["switches"]:
        lines.append(f"  switch    {s['cell']}" + (f"  [{s['name']}]" if s.get("name") else "") + (f"  '{s['label']}'" if s.get("label") else "") + f"  = {s.get('value')!r}  options {s['options']}  ({s['how']})")
    for n in mp["run_settings"]["named_settings"][:25]:
        lines.append(f"  setting   {n['cell']}  [{n['name']}]  = {n['value']!r}  read by {n['readers']} formulas")
    for n in mp["run_settings"].get("status", []):
        lines.append(f"  status    {n['cell']}  [{n['name']}]  = {n['value']!r}  ({n['evidence']})")
    for n in mp["run_settings"].get("labelled", [])[:40]:
        lines.append(f"  labelled  {n['cell']}  '{n['label']}'  = {n['value']!r}")
    for t in mp["run_settings"]["tables"][:25]:
        lines.append(f"  table     {t['cell']}  [{t['name']}]  read by {t['readers']} formulas")
    lines.append("")
    lines.append("OUTPUTS")
    for o in mp["outputs"]["cells"]:
        lines.append(f"  {o['cell']}" + (f"  [{o['name']}]" if o.get("name") else "") + (f"  <- {o['from']}" + (f" [{o['from_name']}]" if o.get("from_name") else "") if o.get("from") else "") + f"  ({o['written_by']}: {o['when']})")
    for s in mp["outputs"]["sheets"]:
        lines.append(f"  sheet     {s['sheet']}  ({s['formulas']} formulas; {s['why']})")
    for sv in mp["outputs"]["saves"]:
        lines.append(f"  saved as  the path in {sv['path_from']}" + (f" [{sv['name']}]" if sv.get("name") else "") + f"  ({sv['in']})")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    import argparse
    import tempfile

    from .inventory import build_analysis

    ap = argparse.ArgumentParser(description="Where a model's inputs, loop driver, run settings and outputs are.")
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--json", action="store_true", help="print the map as JSON")
    ap.add_argument("--work", type=Path, default=None, help="work directory (default: a temp folder)")
    args = ap.parse_args(argv)
    work = args.work or Path(tempfile.mkdtemp(prefix="model_map_"))
    analysis = build_analysis(args.workbook.resolve(), work, "map")
    mp = build_model_map(analysis)
    text = json.dumps(mp, indent=1, ensure_ascii=False, default=str) if args.json else describe(mp)
    stream = sys.__stdout__  # olevba wraps sys.stdout in a console-codepage writer; sheet names hold accents
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="replace")
    stream.write(text + chr(10))
    stream.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
