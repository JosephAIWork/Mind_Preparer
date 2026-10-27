"""Problem reports for the Milliman Tel Aviv team: in codes, with nothing from the user's workbook.

When a step fails, `record()` writes an *incident* into the data folder (feedback\\incidents.jsonl):
the step, the stage and the functions the app was in (as short codes), the error's kind, the rule
being worked on, and the *shape* of the formulas involved -- function names and keywords kept,
every reference, text and number taken out. `compose()` turns an incident (or none, for a message
sent from the header) into the text of the email, with the recent stage timings and the machine,
never a file name, a sheet name, a cell address or a cell's content.

The codes are short hashes (`code()`): what they stand for is worked out from the source by the
team's own decoder, not kept in the app.
"""
from __future__ import annotations

import ctypes
import hashlib
import json
import os
import platform
import re
import secrets
import sys
import threading
import time
from pathlib import Path
from typing import Any, Iterable

from .settings import SETTINGS

MAX_KEPT = 300  # incidents kept in the file
_LOCK = threading.Lock()
_REF_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"

# Text arguments that are Excel keywords, not the user's text: CELL / INFO info types, mostly
KEYWORDS = frozenset(
    "address col color contents directory filename format parentheses prefix protect row type width "
    "numfile origin osversion recalc release system".split()
)
# Local variables a failing function may hold the formula it was working on in, innermost first
_FORMULA_LOCALS = ("formula", "formula_text", "text", "before", "after", "target", "f", "o", "op", "item", "cell")


def code(kind: str, name: str) -> str:
    """Two letters and three digits standing for `name` (a stage, a function, an error type)."""
    h = hashlib.sha256(f"mind-ready:{kind}:{name}".encode("utf-8")).digest()
    return chr(65 + h[0] % 26) + chr(65 + h[1] % 26) + f"{int.from_bytes(h[2:6], 'big') % 1000:03d}"


def _new_ref() -> str:
    return "".join(secrets.choice(_REF_CHARS) for _ in range(4)) + "-" + "".join(secrets.choice("0123456789") for _ in range(4))


def _version() -> str:
    path = Path(__file__).resolve().parents[1] / "VERSION"
    try:
        return path.read_text(encoding="utf-8").strip()
    except OSError:
        return "dev"


def _how_run() -> str:
    return "exe" if "__compiled__" in globals() or getattr(sys, "frozen", False) else "source"


# --- formula shapes -----------------------------------------------------------------------
_A1 = r"\$?[A-Za-z]{1,3}\$?\d{1,7}"
_LOCAL_REF = re.compile(rf"^(?:{_A1}(?::{_A1})?|\$?[A-Za-z]{{1,3}}:\$?[A-Za-z]{{1,3}}|\$?\d{{1,7}}:\$?\d{{1,7}})#?$")


def _ref_kind(value: str) -> str:
    """A reference as what kind of reference it is, without its sheet, its cells or its name."""
    v = value.strip()
    if v.upper() == "#REF!":
        return "#REF!"
    sheet = ""
    if "!" in v:
        before, _, v = v.rpartition("!")
        sheet = "<external>!" if "[" in before else "<sheet>!"
        if v.upper() == "#REF!":
            return sheet + "#REF!"
    elif "[" in v:
        return "<table ref>" if not v.startswith("[") else "<external>"
    spill = "#" if v.endswith("#") else ""
    core = v.rstrip("#")
    if re.fullmatch(_A1, core):
        kind = "<cell>"
    elif re.fullmatch(rf"{_A1}:{_A1}", core):
        kind = "<range>"
    elif re.fullmatch(r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}", core):
        kind = "<columns>"
    elif re.fullmatch(r"\$?\d{1,7}:\$?\d{1,7}", core):
        kind = "<rows>"
    else:
        kind = "<name>"
    return sheet + kind + spill


def _function_name(value: str) -> str:
    """A function token ("MID(") kept when it is Excel's or Mind's; a macro's own function is not named."""
    from .validators.excel_functions import EXCEL_FUNCTIONS

    name = value[:-1]
    bare = re.sub(r"^_xl(?:fn|ws|pm)\.", "", name, flags=re.I).upper()
    bare = bare.replace("_XLFN.", "").replace("_XLWS.", "")
    if bare in EXCEL_FUNCTIONS or bare.startswith("MM_") or bare in ("SINGLE", "ANCHORARRAY", "LAMBDA", "LET"):
        return name + "("
    return "<function>("


def shape(formula: Any) -> str:
    """The formula with every reference, text and number taken out: what kind of formula it is, not
    what it reads. Function names, operators and Excel's own keywords stay."""
    text = str(formula or "").strip()
    if not text:
        return ""
    if not text.startswith("="):
        return "<value>"
    # the tokenizer does not know spill references (A1#)
    text = re.sub(r"(\$?[A-Za-z]{1,3}\$?\d{1,7})#", r"\1", text)
    try:
        from openpyxl.formula.tokenizer import Token, Tokenizer

        tokens = Tokenizer(text).items
    except Exception:
        return "<a formula the app could not read>"
    out: list[str] = []
    for t in tokens:
        if t.type == Token.OPERAND:
            if t.subtype == Token.RANGE:
                out.append(_ref_kind(t.value))
            elif t.subtype == Token.TEXT:
                inner = t.value[1:-1] if len(t.value) >= 2 else ""
                # an Excel keyword, or text with no letter or digit in it ("]", "#"), says nothing of the workbook
                out.append(t.value if inner.lower() in KEYWORDS or not re.search(r"[^\W_]", inner) else '"<text>"')
            elif t.subtype == Token.NUMBER:
                out.append(t.value if t.value in ("0", "1", "2", "-1") else "<n>")
            else:
                out.append(t.value)  # TRUE / FALSE / #N/A ...
        elif t.type == Token.FUNC and t.subtype == Token.OPEN:
            out.append(_function_name(t.value))
        elif t.type == Token.WSPACE:
            out.append(" ")
        else:
            out.append(t.value)
    return ("=" + "".join(out))[:300]


def formula_facts(item: dict[str, Any]) -> list[str]:
    """What kind of cell a planned change or a scanned formula is (no address)."""
    facts = []
    ref = str(item.get("array_ref") or "")
    if item.get("one_cell_array") or (ref and ":" not in ref):
        facts.append("one-cell array")
    elif ref:
        facts.append("array over several cells")
    if item.get("block_cells"):
        facts.append(f"filled-down block of {item['block_cells']} cells")
    if item.get("data_table"):
        facts.append("data table")
    return facts


# --- masking ------------------------------------------------------------------------------
_MASKS = [
    (re.compile(r"(?:[A-Za-z]:\\|\\\\)[^\s'\"<>|]*"), "<path>"),
    (re.compile(r"[^\s'\"\\/]+\.(?:xls[xmb]?|xlam|xla|csv|py|json|tmp)\b", re.I), "<file>"),
    (re.compile(r"'[^']*'!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"), "<cell>"),
    (re.compile(r"\b[\w.]+!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"), "<cell>"),
    (re.compile(r"(?<![<>!=])=(?=[A-Za-z_$'(+-])\S*"), "<formula>"),
    (re.compile(r"'[^']*'|\"[^\"]*\"|‘[^’]*’|“[^”]*”"), "<text>"),
    (re.compile(r"\b\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?\b"), "<cell>"),
    (re.compile(r"(?<![\w.])-?\d+\.\d+\b"), "<n>"),
    (re.compile(r"(?<![\w.x])\d{5,}\b"), "<n>"),
]


def mask(text: Any, names: Iterable[str] = ()) -> str:
    """A message with paths, file names, quoted text, cell addresses and the workbook's own names
    (sheets, defined names) taken out. Excel's and Python's wording stays."""
    out = str(text or "")
    for name in sorted({n for n in names if n and len(n) >= 3}, key=len, reverse=True):
        out = re.sub(rf"(?<![\w]){re.escape(name)}(?![\w])", "<name>", out, flags=re.I)
    for pattern, repl in _MASKS:
        out = pattern.sub(repl, out)
    return re.sub(r"\s+", " ", out).strip()[:240]


# --- what an exception says ---------------------------------------------------------------
def _frames(tb: Any) -> list[Any]:
    out = []
    while tb is not None:
        out.append(tb)
        tb = tb.tb_next
    return out


def exc_facts(exc: BaseException) -> dict[str, Any]:
    """The error as codes: its type, Excel's code when Excel raised it, the app functions it went
    through (code:line, innermost last) and the library it came out of."""
    facts: dict[str, Any] = {"error": code("error", f"{type(exc).__module__}.{type(exc).__qualname__}")}
    info = getattr(exc, "excepinfo", None)
    hr = getattr(exc, "hresult", None)
    scode = info[5] if info and len(info) > 5 else None
    if isinstance(scode, int) and scode:
        facts["excel_code"] = f"0x{scode & 0xFFFFFFFF:08X}"
    elif isinstance(hr, int):
        facts["excel_code"] = f"0x{hr & 0xFFFFFFFF:08X}"
    description = info[2] if info and len(info) > 2 and info[2] else None
    facts["said"] = description or (str(exc) if str(exc) else type(exc).__name__)
    trace, lib = [], None
    formula_seen: dict[str, Any] | None = None
    for tb in _frames(getattr(exc, "__traceback__", None)):
        frame = tb.tb_frame
        module = str(frame.f_globals.get("__name__") or "")
        found = _formula_in(frame)  # a library's frame may hold it too: only its shape is kept
        if found is not None:
            formula_seen = found
        if module == "app" or module.startswith("app."):
            trace.append(f"{code('fn', module + '.' + frame.f_code.co_name)}:{tb.tb_lineno}")
            lib = None
        elif trace:
            lib = lib or module.split(".")[0]
    facts["trace"] = trace[-6:]
    if lib:
        facts["library"] = lib
    if formula_seen is not None:
        facts["formula"] = formula_seen
    stage = getattr(exc, "mind_ready_stage", None)
    if stage:
        facts["stage"] = code("stage", stage)
    return facts


def _formula_in(frame: Any) -> dict[str, Any] | None:
    """The formula the function was working on, read from its local variables (best effort)."""
    try:
        local = frame.f_locals
    except Exception:
        return None
    for name in _FORMULA_LOCALS:
        v = local.get(name)
        if isinstance(v, str) and v.startswith("=") and len(v) > 1:
            return {"shape": shape(v), "facts": []}
        if isinstance(v, dict):
            for key in ("after", "formula", "before"):
                f = v.get(key)
                if isinstance(f, str) and f.startswith("="):
                    return {"shape": shape(f), "facts": formula_facts(v)}
    return None


# --- incidents ----------------------------------------------------------------------------
def _path() -> Path:
    return SETTINGS.data_folder / "feedback" / "incidents.jsonl"


def _write(incident: dict[str, Any]) -> None:
    path = _path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with _LOCK:
            lines = path.read_text(encoding="utf-8").splitlines() if path.is_file() else []
            lines = lines[-(MAX_KEPT - 1):] + [json.dumps(incident, default=str)]
            path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except OSError:
        pass  # a report that cannot be written must never break the step that failed


def recent(n: int = 20) -> list[dict[str, Any]]:
    try:
        lines = _path().read_text(encoding="utf-8").splitlines()
    except OSError:
        return []
    out = []
    for line in lines[-n:]:
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return out


def find(ref: str) -> dict[str, Any] | None:
    return next((i for i in reversed(recent(MAX_KEPT)) if i.get("ref") == ref), None)


def record(
    exc: BaseException | None = None,
    *,
    step: str,
    stage: str | None = None,
    rule_id: str | None = None,
    what: str = "interrupted",
    failed_ops: list[dict[str, Any]] | None = None,
    workbook: dict[str, Any] | None = None,
    names: Iterable[str] = (),
) -> dict[str, Any]:
    """Write one incident and return it (its `ref` is what the screen and the email show).
    `exc`: the error, when there is one. `failed_ops`: changes Excel did not write. `workbook`:
    counts only (`workbook_counts`). `names`: the workbook's sheet and defined names, masked
    wherever they appear."""
    names = list(names)
    incident: dict[str, Any] = {
        "ref": _new_ref(),
        "at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "version": _version(),
        "run": _how_run(),
        "step": step,
        "what": what,
    }
    if rule_id:
        incident["rule_id"] = rule_id
    if stage:
        incident["stage"] = code("stage", stage)
    if exc is not None:
        facts = exc_facts(exc)
        facts["said"] = mask(facts.get("said"), names)
        if "stage" in facts and "stage" not in incident:
            incident["stage"] = facts["stage"]
        facts.pop("stage", None)
        incident.update(facts)
    if failed_ops:
        incident["failed"] = _failed_summary(failed_ops, names)
        if not rule_id:
            rules = sorted({str(o.get("rule_id")) for o in failed_ops if o.get("rule_id")})
            if rules:
                incident["rule_id"] = ", ".join(rules[:4])
    if workbook:
        incident["workbook"] = workbook
    _write(incident)
    return incident


def _failed_summary(ops: list[dict[str, Any]], names: list[str]) -> dict[str, Any]:
    by: dict[str, int] = {}
    for o in ops:
        key = f"{o.get('rule_id') or '?'} {o.get('op') or '?'}"
        by[key] = by.get(key, 0) + 1
    examples = []
    for o in ops[:3]:
        ex: dict[str, Any] = {"rule_id": o.get("rule_id"), "op": o.get("op"), "said": mask(o.get("error"), names)}
        for key in ("error_code", "excel_code", "trace"):
            if o.get(key):
                ex[key] = o[key]
        if isinstance(o.get("before"), str) and o["before"].startswith("="):
            ex["before"] = shape(o["before"])
        if isinstance(o.get("after"), str) and o["after"].startswith("="):
            ex["after"] = shape(o["after"])
        facts = formula_facts(o)
        if facts:
            ex["facts"] = facts
        examples.append(ex)
    return {"count": len(ops), "by": by, "examples": examples}


def failed_op_facts(exc: BaseException) -> dict[str, Any]:
    """The codes to keep with one change Excel refused (prep's apply loop)."""
    f = exc_facts(exc)
    out = {"error_code": f["error"], "trace": f["trace"]}
    if f.get("excel_code"):
        out["excel_code"] = f["excel_code"]
    return out


def workbook_counts(summary: dict[str, Any] | None, size: dict[str, Any] | None = None) -> dict[str, Any]:
    """The workbook in numbers only (no name of anything)."""
    if not summary and not size:
        return {}
    s = summary or {}
    out = {
        "type": s.get("file_type"),
        "sheets": s.get("total_sheet_count") or s.get("sheet_count"),
        "formulas": s.get("formula_count"),
        "array_formulas": s.get("array_formula_count"),
        "grids": s.get("grid_count"),
        "names": s.get("defined_name_count"),
        "macros": s.get("has_vba"),
    }
    if size:
        if isinstance(size.get("file_mb"), (int, float)):
            out["file_mb"] = size["file_mb"]
        if isinstance(size.get("decompressed_mb"), (int, float)):
            out["mb"] = size["decompressed_mb"]
    return {k: v for k, v in out.items() if v not in (None, "")}


# --- the message --------------------------------------------------------------------------
def _incident_lines(i: dict[str, Any]) -> list[str]:
    lines = []
    head = f"{i.get('rule_id') + ' ' if i.get('rule_id') else ''}{i.get('what', 'interrupted')}"
    if i.get("stage"):
        head += f" at stage {i['stage']}"
    head += f', during "{i.get("step", "?")}"'
    lines.append(head)
    if i.get("error"):
        err = f"Error {i['error']}"
        if i.get("excel_code"):
            err += f" · Excel code {i['excel_code']}"
        if i.get("said"):
            err += f" · {i['said']}"
        lines.append(err)
    if i.get("trace"):
        lines.append("Trace " + " > ".join(i["trace"]) + (f" (in {i['library']})" if i.get("library") else ""))
    if i.get("formula"):
        f = i["formula"]
        lines.append("Formula shape: " + f["shape"] + "".join(" · " + x for x in f.get("facts", [])))
    failed = i.get("failed")
    if failed:
        lines.append(f"{failed['count']} change(s) not written: " + ", ".join(f"{k} x{n}" for k, n in failed["by"].items()))
        for ex in failed.get("examples", []):
            parts = [f"  {ex.get('rule_id')} {ex.get('op')}"]
            if ex.get("error_code"):
                parts.append(f"error {ex['error_code']}")
            if ex.get("excel_code"):
                parts.append(f"Excel code {ex['excel_code']}")
            if ex.get("said"):
                parts.append(ex["said"])
            lines.append(" · ".join(parts))
            if ex.get("trace"):
                lines.append("    trace " + " > ".join(ex["trace"]))
            if ex.get("before"):
                lines.append("    before: " + ex["before"])
            if ex.get("after"):
                lines.append("    after:  " + ex["after"] + "".join(" · " + x for x in ex.get("facts", [])))
    return lines


def _machine() -> str:
    parts = [f"Windows {platform.version()}", f"Python {platform.python_version()}", f"{os.cpu_count()} cores"]
    try:
        class _Mem(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong), ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong), ("ullTotalVirtual", ctypes.c_ulonglong),
                        ("ullAvailVirtual", ctypes.c_ulonglong), ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]

        m = _Mem()
        m.dwLength = ctypes.sizeof(_Mem)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m)):  # type: ignore[attr-defined]
            parts.append(f"{m.ullTotalPhys / 2**30:.0f} GB memory ({m.ullAvailPhys / 2**30:.0f} GB free)")
    except Exception:
        pass
    return " · ".join(parts)


def _timings(n: int = 25) -> list[str]:
    from .perf import log_path

    path = log_path()
    if path is None or not path.is_file():
        return []
    try:
        with path.open("rb") as fh:
            fh.seek(0, 2)
            start = max(0, fh.tell() - 200_000)
            fh.seek(start)
            lines = fh.read().decode("utf-8", "replace").splitlines()[1 if start else 0:]  # a cut first line is dropped
    except OSError:
        return []
    since = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time() - 24 * 3600))  # the last day only
    out = []
    for line in lines[-n:]:
        try:
            t = json.loads(line)
        except json.JSONDecodeError:
            continue
        at = str(t.get("at", ""))
        if at < since:
            continue
        facts = [f"{k} {t[k]}" for k in ("operations", "errors", "worker") if k in t]
        out.append(f"{at[5:]} {code('stage', str(t.get('stage')))} {t.get('s')} s" + (f" ({', '.join(facts)})" if facts else ""))
    return out


def _workbook_line(w: dict[str, Any]) -> str:
    labels = [("type", "."), ("sheets", " sheets"), ("formulas", " formulas"), ("array_formulas", " array formulas"), ("grids", " grids"), ("names", " names"), ("file_mb", " MB file"), ("mb", " MB unpacked")]
    parts = []
    for key, label in labels:
        if key in w:
            v = w[key]
            parts.append(f".{v}" if key == "type" else f"{v:,}{label}" if isinstance(v, int) else f"{v}{label}")
    if "macros" in w:
        parts.append("macros" if w["macros"] else "no macros")
    return " · ".join(parts)


def assistant_lines(incident: dict[str, Any] | None, note: str, rule_title: str | None = None) -> str | None:
    """Two or three plain-English sentences opening the message, written by the assistant from the
    same facts the message holds (never the workbook). None when the assistant is off or fails."""
    from .llm import chat_completion, llm_available

    if not llm_available():
        return None
    facts = []
    if incident:
        facts.append(f"The step the user ran: {incident.get('step')}.")
        if incident.get("rule_id"):
            facts.append(f"The Mind rule being worked on: {incident['rule_id']}" + (f" ({rule_title})." if rule_title else "."))
        if incident.get("said"):
            facts.append(f"What the app reported: {incident['said']}.")
        if incident.get("failed"):
            facts.append(f"{incident['failed']['count']} change(s) to the workbook could not be written.")
        if incident.get("formula"):
            facts.append(f"The kind of formula involved: {incident['formula']['shape']}.")
    else:
        facts.append("The user is sending feedback from the app's header; no step failed.")
    if note.strip():
        facts.append(f"The user's own words: {note.strip()[:600]}")
    system = (
        "You write the opening lines of a problem report that a user of the Mind Ready app (it prepares Excel models "
        "for Milliman Mind) sends to the app's developers. Write two or three short sentences in plain English, in the "
        "first person ('I was ...'). Use only the facts given. Do not guess the cause, do not give advice, do not mention codes."
    )
    res = chat_completion([{"role": "user", "content": "\n".join(facts)}], system, max_tokens=220, timeout=45)
    text = (res.get("text") or "").strip()
    return text or None


def compose(incident: dict[str, Any] | None, note: str = "", ai_text: str | None = None, workbook: dict[str, Any] | None = None) -> dict[str, Any]:
    """The email: subject, the text to paste, the file name to save it under, the address."""
    version, run = _version(), _how_run()
    ref = incident["ref"] if incident else _new_ref()
    at = time.strftime("%d %b %Y %H:%M")
    lines = [f"Mind Ready {version} ({run}) · ref {ref} · {at}", ""]
    if ai_text:
        lines += [ai_text, ""]
    if incident:
        lines += ["What happened"] + ["  " + x for x in _incident_lines(incident)] + [""]
    else:
        lines += ["Feedback (no step failed)", ""]
    if note.strip():
        lines += ["The user's note"] + ["  " + x for x in note.strip().splitlines()] + [""]
    lines += ["Log", "  Machine: " + _machine()]
    w = (incident or {}).get("workbook") or workbook
    if w:
        lines.append("  Workbook: " + _workbook_line(w))
    others = [i for i in recent(6) if not incident or i.get("ref") != incident.get("ref")][-5:]
    if others:
        lines.append("  Earlier problems:")
        for o in others:
            first = _incident_lines(o)[:2]
            lines.append(f"    {o.get('at')} ref {o.get('ref')} · " + " · ".join(first))
    timings = _timings()
    if timings:
        lines.append("  Last stages (oldest first):")
        lines += ["    " + t for t in timings]
    text = "\n".join(lines).rstrip() + "\n"
    return {
        "ref": ref,
        "text": text,
        "subject": f"Mind Ready {version} feedback · ref {ref}",
        "email": SETTINGS.feedback_email,
        "file_name": f"MindReady-feedback-{ref}.txt",
        "assistant": bool(ai_text),
    }
