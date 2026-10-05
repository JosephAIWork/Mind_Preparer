"""Workbook chatbot (1.4.0; changes 1.5.0): grounds a multi-turn conversation
in the analysis and the findings, and lets the user ask for changes.

The model gets, on every turn:
  * a bounded *context pack* -- workbook facts, every sheet with its grids
    and flags, formula/MM_ statistics, every non-PASS finding in full and
    the PASS rules by id, plus the prep actions the app can apply;
  * *retrieved detail* for whatever the question names -- cell references
    (the formula / value there), sheet or grid names (the grid's rows and
    header), rule ids (the finding's observed data), MM_ function names
    (their call sites);
  * the conversation so far.

Changing the workbook (1.5.0): the assistant may (a) request cell contents
it hasn't seen with a ```lookup block -- the app answers within the same
turn, up to MAX_LOOKUPS times -- and (b) end its answer with a ```changes
block listing operations. The app validates the block against the real
workbook (app/prep.py::validate_proposal), shows it as a before/after table,
and applies it through Excel only when the user clicks Apply. The model's
text is never written into the workbook by itself.
"""
from __future__ import annotations

import json
import re
from typing import Any, Callable

from .formula_utils import STORAGE_PREFIX_RE, parse_ref, ref_text
from .grids import all_grids
from .inventory import cell_value
from .llm import chat_completion
from .prep import validate_proposal

MAX_CONTEXT_CHARS = 14000
MAX_RETRIEVAL_CHARS = 6000
MAX_GRIDS_PER_SHEET = 40
MAX_LOOKUPS = 2
MAX_PROPOSAL_RETRIES = 2  # a rejected ```changes block is sent back (twice at most) with the validation errors
PROPOSAL_MAX_TOKENS = 4000  # room for ~120 operations in one ```changes block (1.7.2)
CELL_REF_RE = re.compile(r"(?<![A-Za-z0-9_])(?:(?P<sheet>'[^']+'|[A-Za-z0-9_]+)!)?(?P<ref>\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)(?![A-Za-z0-9_(])")
RULE_ID_RE = re.compile(r"\b([A-Z]{2,7})-(\d{3})\b")
MM_FN_RE = re.compile(r"\bMM_[A-Za-z0-9_]+\b", re.IGNORECASE)
# Replacements the tool vouches for (every function on the right is on the Mind list); when the user
# asks to replace a function, the assistant proposes these rather than inventing one, and keeps the
# original's behaviour on blanks/errors.
KNOWN_EQUIVALENTS: dict[str, str] = {
    "NUMBERVALUE": 'IF(x="",0,VALUE(x)) -- NUMBERVALUE("") is 0 while VALUE("") is #VALUE!',
    "IFNA": "IF(ISNA(x),alt,x) -- IFERROR would also swallow #REF!/#DIV/0!, which changes behaviour",
    "XLOOKUP": "INDEX(return_range, MATCH(key, lookup_range, 0)) wrapped in IF(ISNA(...),if_not_found,...) when a default was given",
    "CONCAT": "CONCATENATE(...) or the & operator",
    "TEXTJOIN": "CONCATENATE with the delimiter written between the parts",
    "SWITCH": "nested IF(...)",
    "ANCHORARRAY": "the fixed range the spill covers today (Prep action 'freeze_spill_refs' does it for every A1# reference); MM_RANGE if the range must grow",
}
FN_WORD_RE = re.compile(r"(?<![A-Za-z0-9_])(?:_xl(?:fn|ws|udf|pm|l)\.)?([A-Za-z][A-Za-z0-9_.]{1,30})", re.IGNORECASE)
BLOCK_RE = {kind: re.compile(rf"```{kind}\s*\n(.*?)```", re.DOTALL | re.IGNORECASE) for kind in ("changes", "lookup")}

SYSTEM_PROMPT = """You are the assistant inside "Excel Upload Preparation", a tool that checks Excel workbooks before upload to Milliman Mind (an actuarial modelling platform that imports Excel models; grids are blocks of cells titled with a '#Name /Flag' cell; MM_-prefixed functions come from the MMForExcel add-in).

You are given the tool's analysis of ONE workbook: its sheets, grids and flags, formula statistics, and the findings of 99 readiness rules (PASS / WARNING / ERROR / REQUIRES_USER_INPUT / NOT_SUPPORTED). Answer the user's questions about this workbook from that material, and make the changes the user asks for.

Rules for you:
- Anything you state about Mind -- what it accepts, refuses, requires, computes or displays -- must come from the Mind documentation the tool holds: the rule findings, the "## Function support" block, the KB extracts and the rule texts in the context. Your general knowledge of Excel or of "actuarial platforms" is NOT evidence about Mind: never conclude "Mind will accept this" or "this is a non-issue" from it. If the tool's material says nothing about a point, say exactly that ("the Mind documentation held by the tool does not cover X") and stop there.
- Ground every answer in the context; quote rule ids, sheet!cell locations and grid names. If the context doesn't contain what is needed, look it up (below) or say what is missing instead of guessing.
- Explain findings in plain language and say concretely what to change in Excel to fix them.
- Whether Mind accepts an Excel function is decided by the tool's own lists (the Mind supported-function list, the functions accepted on other evidence, and the MM_ registry, applied by rule FRM-002), never by general knowledge. When the context has a "## Function support" block, state its verdict as the answer ("IFERROR is on the Mind supported list" / "the Mind documentation does not list NUMBERVALUE: it neither confirms nor refuses it"); never hedge with "most platforms" or "should be fine". A function the Mind documentation does not list is NOT unsupported and NOT a blocker: say the documentation is silent, and do not propose to replace it unless the user asks for a replacement.
- `_xlfn.`, `_xlws.`, `_xludf.` and `_xll.` in front of a function name are Excel's internal storage prefixes (Excel writes them for functions newer than Excel 2007 and for add-in calls). They are not a compatibility-mode artefact, not a defect and not something to fix; the rules strip them before checking the function name. If a user asks about one, say exactly that in one sentence.
- When you replace a function with another, keep the formula's behaviour on blank cells, text and errors: e.g. NUMBERVALUE("") returns 0 but VALUE("") returns #VALUE!, so wrap the replacement (IF(x="",0,VALUE(x))) when the original tolerated blanks; the app recalculates the file in Excel afterwards and every new error is counted against the fix.
- The app has these sections: "Findings" (rule results), "Prep workbook" (applies the listed prep actions and user-edited formula replacements), "Ask the assistant" (you), "Recalculate" (real Excel recalculation -- the only way READY-001 clears), "Reports" (downloads).
- PASS for upload is only ever granted after a real recalculation; never tell the user the workbook is ready.
- If a <recalculation> block is present it is the result of a real Excel recalculation of the current file: a genuine formula error (#NAME?, #REF!, #VALUE!, #DIV/0!, #N/A ...) at a cell must be explained from that cell's formula and, when the user asks, fixed with a changes block; a #NAME? on an MM_ function listed as an add-in gap is NOT a workbook defect (the MMForExcel add-in is missing on this machine) and must not be "fixed" by rewriting the formula.
- Be concise: short paragraphs or bullets, no preamble.

## Making changes
When the user asks you to change something, do it: explain the change in one or two sentences, then end your answer with exactly one fenced block that the app turns into an Apply button (the app writes it through Excel to a fresh copy after the user confirms, verifies the file, and re-analyzes it):
```changes
{"summary": "<what this does, one line>", "operations": [ ... ]}
```
Operation objects (use exact sheet names from the context; cells may be a single cell "B3" or a range "B3:D3"):
- {"op": "set_value", "sheet": "S", "cell": "B3", "value": "text or number or true/false"}  -- a grid title is a value starting with '#', e.g. "#Premiums /Input"
- {"op": "set_formula", "sheet": "S", "cell": "C5", "formula": "=..."}  -- on a range, relative references adjust like fill-down
- {"op": "clear_cell", "sheet": "S", "cell": "A1"}  -- removes content; only when the user asked for it
- {"op": "set_array_formula", "sheet": "S", "range": "C7:C9", "formula": "=..."}  -- ONE array formula (Ctrl+Shift+Enter) over the whole range; use it for cells reported with an `array`, because Excel cannot change part of an array (otherwise cover every cell of the array with one operation each)
- {"op": "rename_sheet", "sheet": "Old name", "new_name": "New name"}  -- Excel updates every reference
- {"op": "insert_row", "sheet": "S", "row": 7}  /  {"op": "insert_column", "sheet": "S", "column": "C"}
- {"op": "set_sheet_visibility", "sheet": "S", "visible": true}  /  {"op": "unprotect_sheet", "sheet": "S"}
Only propose what the user asked for; never turn a formula into a value unless asked; if a request is ambiguous or risky, ask one short question instead of proposing. If the user only asks a question, do not include a changes block.

## Looking things up
If you need cell contents or grid rows you don't have (to answer or to write a correct change), reply with ONLY this block and nothing else; the app sends the data back and you then answer normally:
```lookup
{"cells": ["Sheet!A1:D10"], "grids": ["Grid name"], "sheets": ["Sheet"], "functions": ["MM_LOOP"]}
```
You may do this at most twice per question."""


def _trim(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 20] + "... [truncated]"


def _location(f: dict[str, Any]) -> str:
    loc = f.get("location") or {}
    if "sheet" in loc and "cell" in loc:
        return f"{loc['sheet']}!{loc['cell']}"
    return ", ".join(f"{k}={v}" for k, v in loc.items()) if loc else "-"


def build_context_pack(analysis: dict[str, Any], validation_report: dict[str, Any], prep_actions: list[dict[str, Any]] | None = None, max_chars: int = MAX_CONTEXT_CHARS) -> str:
    wb = analysis["workbooks"][0]
    f = analysis["features"]
    lines = [
        "# Workbook",
        f"- File: {wb.get('file_name')}.{wb.get('file_type')} | last saved by {f.get('application')} {f.get('app_version') or ''} | VBA project: {f.get('has_vba')}",
        f"- {f.get('sheet_count')} sheet(s) ({f.get('hidden_sheet_count')} hidden), {f.get('grid_count')} grid(s) ({f.get('flagged_grid_count')} flagged), {f.get('formula_count')} formula(s) incl. {f.get('array_formula_count')} array formulas, {f.get('defined_name_count')} defined names, {f.get('external_link_part_count')} external link part(s)",
    ]
    mm = f.get("mm_functions_used", {})
    lines.append(f"- MM_ functions used: {', '.join(f'{k}×{v}' for k, v in sorted(mm.items())) or 'none'}")
    native = sorted(f.get("native_function_usage", {}).items(), key=lambda kv: -kv[1])[:15]
    lines.append(f"- Most used native functions: {', '.join(f'{k}×{v}' for k, v in native) or 'none'}")
    if f.get("non_openpyxl_parts"):
        lines.append(f"- Package parts only Excel can preserve (data model / customXml / ...): {len(f['non_openpyxl_parts'])}")

    lines.append("\n# Sheets and grids (grid = 'name (range) [flags] headers'; 'untitled X' = no '#' title cell)")
    if wb.get("ignored_sheets"):
        lines.append(
            "NOT SCANNED (the user chose to skip these sheets: they exist in the file but nothing on them was read or checked; "
            "say so if asked about them, and never propose changes on them): "
            + ", ".join(f"{s['name']} ({s['state']})" for s in wb["ignored_sheets"])
        )
    for s in wb["sheets"]:
        lines.append(f"## {s['name']} -- state {s['state']}, used range {s.get('dimensions')}, {len(s['grids'])} grid(s), {s.get('formula_count', 0)} formula(s), {len(s.get('standalone_text_cells', []))} standalone text cell(s)")
        for g in s["grids"][:MAX_GRIDS_PER_SHEET]:
            headers = [str(h)[:18] for h in g.get("header_values", [])[:8] if h is not None]
            flags = " ".join(fl["raw"] for fl in g.get("flags", []))
            lines.append(f"- {g['display_name']} ({g['ref']}) [{flags}] headers: {headers}")
        if len(s["grids"]) > MAX_GRIDS_PER_SHEET:
            lines.append(f"- ... {len(s['grids']) - MAX_GRIDS_PER_SHEET} more grid(s)")
        standalone = s.get("standalone_text_cells", [])[:10]
        if standalone:
            lines.append("- standalone text: " + "; ".join(f"{c['cell']}='{str(c['text'])[:24]}'" for c in standalone))

    rep = validation_report
    lines.append(f"\n# Findings -- overall {rep['status']}; counts {rep['summary']['status_counts']}")
    non_pass = [x for x in rep["findings"] if x["status"] != "PASS"]
    order = {"ERROR": 0, "REQUIRES_USER_INPUT": 1, "NOT_SUPPORTED": 2, "WARNING": 3}
    for x in sorted(non_pass, key=lambda x: (order.get(x["status"], 9), x["rule_id"])):
        lines.append(f"- [{x['status']}] {x['rule_id']} @ {_location(x)}: {_trim(x['message'], 420)}")
    lines.append("- PASS: " + ", ".join(x["rule_id"] for x in rep["findings"] if x["status"] == "PASS"))

    if prep_actions:
        lines.append("\n# Prep actions the app can apply (user selects and approves them in 'Prep workbook')")
        for a in prep_actions:
            lines.append(f"- {a['title']} ({a['id']}, rules {', '.join(a['rule_ids'])}): {a['count']} change(s) planned{'; skipped: ' + '; '.join(a['skipped'][:3]) if a.get('skipped') else ''}")
    return _trim("\n".join(lines), max_chars)


def _cells_block(analysis: dict[str, Any], sheet: str, ref: dict[str, Any], grids: list[dict[str, Any]]) -> str:
    cells = []
    for r in range(ref["r1"], min(ref["r2"], ref["r1"] + 40) + 1):
        for c in range(ref["c1"], min(ref["c2"], ref["c1"] + 12) + 1):
            v = cell_value(analysis, sheet, r, c)
            if v is not None:
                cells.append(f"{ref_text(c, r)} = {_trim(str(v), 200)}")
    g = next((g for g in grids if g["sheet"] == sheet and g["first_row"] <= ref["r1"] <= g["last_row"] and g["first_col"] <= ref["c1"] <= g["last_col"]), None)
    return f"## Cells {sheet}!{ref['ref']}" + (f" (in grid {g['display_name']} {g['ref']})" if g else " (not inside a detected grid)") + "\n" + ("\n".join(cells) if cells else "(empty)")


def _grid_block(analysis: dict[str, Any], g: dict[str, Any]) -> str:
    rows = g.get("rows")
    if rows is None:
        rows = [[cell_value(analysis, g["sheet"], r, c) for c in range(g["first_col"], min(g["last_col"], g["first_col"] + 11) + 1)] for r in range(g["first_row"], min(g["last_row"], g["first_row"] + 15) + 1)]
    preview = "\n".join(" | ".join(_trim(str(v), 24) if v is not None else "" for v in row) for row in rows[:16])
    return f"## Grid {g['display_name']} at {g['sheet']}!{g['ref']} flags {[fl['raw'] for fl in g['flags']]}\nheaders {g.get('header_values')}\n{preview}"


def _sheet_block(s: dict[str, Any]) -> str:
    return f"## Sheet {s['name']}: {len(s['grids'])} grid(s)\n" + "\n".join(
        f"- {g['display_name']} {g['ref']} {[fl['raw'] for fl in g['flags']]} headers {[str(h)[:16] for h in g.get('header_values', [])[:6]]}" for g in s["grids"][:60]
    )


def function_support(names: list[str], analysis: dict[str, Any]) -> str:
    """The tool's verdict on Excel / MM_ function names (the same lists FRM-002
    applies), with the workbook's usage, so the assistant answers "is X
    supported by Mind?" from fact rather than from general knowledge."""
    from .formula_utils import called_functions
    from .validators.excel_functions import EXCEL_FUNCTIONS
    from .validators.formula import DYNAMIC_ARRAY_ARTIFACTS, _known_mm_functions, _supported_native_functions, confirmed_functions, evidence_text, kb_page_note

    supported, known_mm, confirmed = _supported_native_functions(), _known_mm_functions(), confirmed_functions()
    wb = analysis["workbooks"][0]
    lines = []
    for raw in names:
        fn = STORAGE_PREFIX_RE.sub("", raw).upper()
        is_mm = fn.startswith("MM_")
        if not is_mm and fn not in EXCEL_FUNCTIONS:
            continue
        sites = [x for x in wb["formulas"] if fn in {c.upper() for c in called_functions(x["formula"])}]
        usage = f"used in {len(sites)} formula(s)" + (f", first at {sites[0]['sheet']}!{sites[0]['cell']}" if sites else "")
        if is_mm:
            entry = known_mm.get(fn)
            verdict = ("registered MM_ function" + (" (deprecated)" if entry.get("deprecated") else "")) if entry else "MM_ name not in either MM_ registry: FRM-002 flags it (do not invent syntax)"
        elif fn in DYNAMIC_ARRAY_ARTIFACTS:
            verdict = "Excel dynamic-array storage token, reviewed by FRM-003 (Mind resizes arrays with MM_RANGE, not Office 365 spills)"
        elif fn in confirmed:
            verdict = f"accepted by Mind ({evidence_text(confirmed[fn])}: {confirmed[fn].get('detail')}) although {kb_page_note()} omits it: FRM-002 passes it, leave it as it is"
        elif fn in supported:
            verdict = "on the Mind supported-function list: accepted by Mind as is (FRM-002 passes it)"
        else:
            verdict = (
                f"not listed on {kb_page_note()}, which says nothing about the functions it omits: the Mind documentation neither confirms nor refuses it. "
                "FRM-002 reports it as a warning, not a blocker; a conversion in Mind is the proof. Do not call it unsupported"
            )
        prefix_note = " The `_xlfn.`/`_xl*.` prefix seen in stored formulas is Excel's storage prefix, not a defect." if STORAGE_PREFIX_RE.search(raw) else ""
        equivalent = f" If the user asks for a replacement: {KNOWN_EQUIVALENTS[fn]}." if fn in KNOWN_EQUIVALENTS and fn not in supported else ""
        lines.append(f"- {fn}: {verdict}; {usage}.{prefix_note}{equivalent}")
    return "## Function support (tool verdict, rule FRM-002)" + chr(10) + chr(10).join(lines) if lines else ""


def retrieve(question: str, analysis: dict[str, Any], validation_report: dict[str, Any], max_chars: int = MAX_RETRIEVAL_CHARS) -> str:
    """Detail for the things the question names."""
    wb = analysis["workbooks"][0]
    sheet_names = {s["name"]: s for s in wb["sheets"]}
    grids = all_grids(analysis)
    findings = {x["rule_id"]: x for x in validation_report["findings"]}
    out: list[str] = []
    q_low = question.lower()

    for m in RULE_ID_RE.finditer(question.upper()):
        rid = f"{m.group(1)}-{m.group(2)}"
        x = findings.get(rid)
        if x:
            out.append(f"## Finding {rid}\nstatus {x['status']} | severity {x.get('severity')} | location {_location(x)}\n{x['message']}\nobserved: {_trim(json.dumps(x.get('observed'), default=str), 1800)}")

    mentioned_sheets = [n for n in sorted(sheet_names, key=len, reverse=True) if n.lower() in q_low]
    for m in CELL_REF_RE.finditer(question):
        ref = parse_ref(m.group(0))
        if not ref:
            continue
        sheet = ref["sheet"] or (mentioned_sheets[0] if mentioned_sheets else wb["sheets"][0]["name"])
        if sheet not in sheet_names:
            continue
        out.append(_cells_block(analysis, sheet, ref, grids))

    for g in grids:
        name = (g.get("name") or "").strip()
        if len(name) >= 3 and name.lower() in q_low:
            out.append(_grid_block(analysis, g))

    for name in mentioned_sheets[:2]:
        out.append(_sheet_block(sheet_names[name]))

    # any Excel / MM_ function the question names: the tool's support verdict first
    named: list[str] = []
    for m in FN_WORD_RE.finditer(question):
        tok = m.group(0)
        if tok.upper() not in {n.upper() for n in named}:
            named.append(tok)
    block = function_support(named, analysis)
    if block:
        out.append(block)

    for m in {t.upper() for t in MM_FN_RE.findall(question)}:
        sites = [x for x in wb["formulas"] if m in x["formula"].upper()][:15]
        if sites:
            out.append(f"## {m} call sites\n" + "\n".join(f"- {x['sheet']}!{x['cell']}: {_trim(STORAGE_PREFIX_RE.sub('', x['formula']), 160)}" for x in sites))

    if "flag" in q_low and not out:
        flagged = [g for g in grids if g["flags"]]
        out.append("## Flagged grids\n" + "\n".join(f"- {g['sheet']}!{g['ref']} {g['display_name']} {[fl['raw'] for fl in g['flags']]}" for g in flagged[:80]))
    return _trim("\n\n".join(out), max_chars)


def retrieve_structured(spec: Any, analysis: dict[str, Any], max_chars: int = MAX_RETRIEVAL_CHARS) -> str:
    """Answer a ```lookup request: {"cells": [...], "grids": [...], "sheets": [...], "functions": [...]}."""
    wb = analysis["workbooks"][0]
    sheet_names = {s["name"]: s for s in wb["sheets"]}
    lower_sheets = {k.lower(): k for k in sheet_names}
    grids = all_grids(analysis)
    out: list[str] = []
    if not isinstance(spec, dict):
        return "lookup request was not a JSON object"
    for item in spec.get("cells", []) or []:
        ref = parse_ref(str(item))
        if not ref:
            out.append(f"## {item}: not a cell reference")
            continue
        sheet = lower_sheets.get((ref["sheet"] or "").lower()) or (wb["sheets"][0]["name"] if not ref["sheet"] else None)
        if sheet is None:
            out.append(f"## {item}: sheet not found (sheets: {list(sheet_names)})")
            continue
        out.append(_cells_block(analysis, sheet, ref, grids))
    for name in spec.get("grids", []) or []:
        hits = [g for g in grids if (g.get("name") or g["display_name"]).strip().lower() == str(name).strip().lower()]
        if not hits:
            hits = [g for g in grids if str(name).strip().lower() in (g.get("name") or g["display_name"]).lower()][:3]
        out.extend(_grid_block(analysis, g) for g in hits[:3]) if hits else out.append(f"## Grid '{name}': not found")
    for name in spec.get("sheets", []) or []:
        s = sheet_names.get(lower_sheets.get(str(name).lower(), ""))
        out.append(_sheet_block(s) if s else f"## Sheet '{name}': not found (sheets: {list(sheet_names)})")
    for fn in spec.get("functions", []) or []:
        fn_u = str(fn).upper()
        sites = [x for x in wb["formulas"] if fn_u in x["formula"].upper()][:20]
        out.append(f"## {fn_u} call sites\n" + ("\n".join(f"- {x['sheet']}!{x['cell']}: {_trim(STORAGE_PREFIX_RE.sub('', x['formula']), 160)}" for x in sites) if sites else "(none)"))
    return _trim("\n\n".join(out) or "(nothing requested)", max_chars)


def extract_block(text: str | None, kind: str) -> tuple[Any | None, str | None]:
    """(parsed JSON, error) for the first ```<kind> block in `text`."""
    if not text:
        return None, None
    m = BLOCK_RE[kind].search(text)
    if not m:
        return None, None
    raw = m.group(1).strip()
    try:
        return json.loads(raw), None
    except json.JSONDecodeError as exc:
        return None, f"{kind} block is not valid JSON: {exc}"


def strip_block(text: str, kind: str) -> str:
    return BLOCK_RE[kind].sub("", text or "").strip()


def recalculation_context(recalc: dict[str, Any] | None, max_sites: int = 40) -> str:
    """A compact rendering of the last real Excel recalculation for the model:
    status, message, genuine formula-error cells, and the #NAME? cells that
    only fail because the MMForExcel add-in is missing on this machine."""
    if not recalc:
        return ""
    lines = [f"status {recalc.get('status')}: {recalc.get('message', '')}"]
    groups = recalc.get("groups") or []
    if groups:
        lines.append("root causes (errors that can be fixed together):")
        for g in groups[:12]:
            lines.append(f"- {g['count']} cell(s): {g['cause']} -- " + ", ".join(f"{c['sheet']}!{c['cell']}" for c in g["cells"][:12]) + (" ..." if len(g["cells"]) > 12 else ""))
    errors = recalc.get("formula_errors") or []
    if errors:
        lines.append(f"genuine formula errors ({len(errors)}):")
        lines += [f"- {e.get('sheet')}!{e.get('cell')} {e.get('error')}: {_trim(str(e.get('formula', '')), 160)}" for e in errors[:max_sites]]
        if len(errors) > max_sites:
            lines.append(f"- ... {len(errors) - max_sites} more")
    gaps = recalc.get("addin_gap_errors") or []
    if gaps:
        lines.append(f"#NAME? on MM_ functions because the MMForExcel add-in is not installed on this machine ({len(gaps)}) -- not workbook defects:")
        lines += [f"- {e.get('sheet')}!{e.get('cell')}: {_trim(str(e.get('formula', '')), 120)}" for e in gaps[:15]]
    return "\n".join(lines)


def answer_question(
    history: list[dict[str, str]],
    question: str,
    analysis: dict[str, Any],
    validation_report: dict[str, Any],
    prep_actions: list[dict[str, Any]] | None = None,
    model_id: str | None = None,
    extra_context: str | None = None,
    on_event: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """Run one turn (with up to MAX_LOOKUPS lookup round-trips). Returns the
    display text, the raw text, and any validated change proposal.
    `extra_context` (e.g. the last recalculation) is appended to the system
    context for this turn. `on_event` (1.7.2) streams the turn: a
    {"type": "phase", "phase": "answer"|"lookup"|"repair", "n"} before every
    model call, then {"type": "delta", "text"} for each chunk of its reply."""
    context = build_context_pack(analysis, validation_report, prep_actions)
    detail = retrieve(question, analysis, validation_report)
    system = SYSTEM_PROMPT + "\n\n<workbook_context>\n" + context + "\n</workbook_context>"
    if extra_context:
        system += "\n\n<recalculation>\n" + _trim(extra_context, 6000) + "\n</recalculation>"
    turn = question if not detail else f"{question}\n\n<retrieved_detail>\n{detail}\n</retrieved_detail>"
    messages = [{"role": m["role"], "content": m["content"]} for m in history[-12:]] + [{"role": "user", "content": turn}]
    # 1.7.2: a ```changes block costs ~30 tokens per cell; the old 1200-token
    # default silently truncated proposals at ~35 operations (the rest of the
    # finding's cells never made it into the proposal and nothing said so).
    kwargs: dict[str, Any] = {"max_tokens": PROPOSAL_MAX_TOKENS}
    if model_id:
        kwargs["model_id"] = model_id
    if on_event is not None:
        kwargs["on_delta"] = lambda text: on_event({"type": "delta", "text": text})

    def _phase(phase: str, n: int) -> None:
        if on_event is not None:
            on_event({"type": "phase", "phase": phase, "n": n})

    lookups = 0
    result: dict[str, Any] = {"available": False, "text": None, "message": "no call made"}
    while True:
        _phase("lookup" if lookups else "answer", lookups)
        result = chat_completion(messages, system, **kwargs)
        text = result.get("text") or ""
        spec, err = extract_block(text, "lookup")
        if (spec is not None or err) and lookups < MAX_LOOKUPS and result.get("available"):
            lookups += 1
            data = retrieve_structured(spec, analysis) if spec is not None else err
            messages += [{"role": "assistant", "content": text}, {"role": "user", "content": f"<lookup_result>\n{data}\n</lookup_result>\nNow answer the original question (and propose the changes if asked)."}]
            continue
        break

    def _validated(text: str) -> tuple[list[dict[str, Any]], list[str], Any]:
        proposal, perr = extract_block(text, "changes")
        if proposal is not None:
            ops, errors = validate_proposal(proposal, analysis)
            return ops, errors, (proposal.get("summary") if isinstance(proposal, dict) else None)
        return [], ([perr] if perr else []), None

    raw_text = result.get("text") or ""
    ops, errors, summary = _validated(raw_text)
    retries = 0
    while errors and retries < MAX_PROPOSAL_RETRIES and result.get("available"):
        # Give the model one chance to repair a rejected proposal (e.g. a change to part of an array formula).
        retries += 1
        messages += [
            {"role": "assistant", "content": raw_text},
            {"role": "user", "content": "<proposal_rejected>\n" + "\n".join(f"- {e}" for e in errors) + "\n</proposal_rejected>\n"
             "Propose the changes again in a ```changes block, fixing every point above (keep the explanation short)."},
        ]
        _phase("repair", retries)
        result = chat_completion(messages, system, **kwargs)
        raw_text = result.get("text") or ""
        ops, errors, summary = _validated(raw_text)
    display = strip_block(strip_block(raw_text, "changes"), "lookup") or ("(the assistant proposed changes without explanation)" if ops else raw_text)
    return {
        **result,
        "text": display,
        "raw_text": raw_text,
        "proposal": ops,
        "proposal_summary": summary,
        "proposal_errors": errors,
        "lookups": lookups,
        "proposal_retries": retries,
        "context_chars": len(context),
        "detail_chars": len(detail),
    }
