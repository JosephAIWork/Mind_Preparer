"""The app runs itself against real Milliman Mind.

One call -- ``run_loop(LoopConfig(...))`` -- takes a raw model workbook and,
with nobody watching, repeats:

    prepare a fresh copy  ->  local gates  ->  upload to Mind  ->  convert
    ->  add template  ->  run  ->  read what Mind shows  ->  decide

until everything looks good and the numbers match, or until it can say
precisely why it cannot get there. Every decision is a rule in this file;
no model, no person and no assistant is consulted while it runs (the
assistant is used exactly once, before the first iteration, to name grids
-- and only if it is reachable).

"Everything looks good" means, per iteration:
  prepared        the prep plan applied through Excel and the copy opens
  numbers         a full Excel recalculation of the prepared copy matches a
                  full recalculation of the source cell for cell. Cells the
                  plan wrote as *content* (titles, moved labels) are the
                  plan's intent and are not compared; cells whose *formula*
                  the plan rewrote are compared under the conservative rule:
                  the value is unchanged, or an error became another error
                  (never a valid value into anything else); every other
                  cell must be identical, an error may become another error
  structure       the grid count did not balloon (a shattered workbook is
                  the failure mode of a bad row insert)
  names           every grid the plan could title is titled; what is left is
                  the list of grids the app cannot title, with reasons
  convert         Mind's Upload / Convert / Test wizard all Complete
  run             Mind ran the model and reports it consistent with the
                  audit trail
  counts          the number of 'Untitled(r,c)' entries Mind lists equals
                  the number of untitled grids the app predicted

A gate that could not be checked is a failure, never a pass. When a gate
fails the loop changes *one* thing about the next iteration -- enable an
opt-in fix Mind's error log asks for, disable the action whose edits moved
a number (found by walking the changed cell's precedents back to the cell
an action wrote) -- and starts again from the ORIGINAL source. It never
re-preps its own output (that grows a workbook for ever fewer names). When
no rule applies it stops and says so: that report is the developer's cue.

Guardrails: the source is never modified (immutable copies everywhere);
only sandbox projects (the project prefix and sandbox folder in LoopConfig /
mind_client) are created or deleted; a project that has been run cannot be deleted by
automation and is listed for manual cleanup.
"""
from __future__ import annotations

import datetime as _dt
import json
import re
import time
from collections import deque
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable

import openpyxl

from . import mind_client as mc
from . import prep
from .change_apply import convert_output_format
from .formula_utils import cell_refs_in_formula, parse_ref, ref_text
from .grid_naming import build_grid_context, suggest_names
from .grids import ERROR_LITERALS, all_grids
from .inventory import has_vba_project, make_immutable_copy
from .modes import plan_mode
from .recalc import recalculate
from .rules_engine import RulesEngine

Progress = Callable[[dict[str, Any]], None]

VOLATILE = re.compile(r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(", re.I)
NUMERIC_TOL = 1e-9
FRAGMENTATION_SLACK = 0.10  # grids may grow by this share (+5) before it counts as shattering
RELAXED_OPS = {"set_formula", "set_array_formula"}  # a rewrite may keep the value or turn an error into an error
PRECEDENT_WALK_LIMIT = 500  # cells visited when tracing a changed value back to what the plan wrote

# What Mind's Convert error log asks for, mapped to the prep action that answers it.
MIND_ERROR_TO_ACTION: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"#REF!|formula compilation|not a function|could not compile|cannot compile", re.I), "fix_broken_refs"),
    (re.compile(r"setting does not exist", re.I), "project_settings"),  # no action: PRJ-006 finding; needs a human
]


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@dataclass
class LoopConfig:
    source: Path
    work_dir: Path
    max_iterations: int = 4
    project_prefix: str = mc.PROJECT_PREFIX
    folder: str = mc.SANDBOX_FOLDER
    headed: bool = False
    use_assistant: bool = True
    run_model: bool = True
    check_numbers: bool = True
    enable: list[str] = field(default_factory=list)  # opt-in actions forced on from iteration 1
    disable: list[str] = field(default_factory=list)  # actions never used
    delete_projects: bool = True
    skip_mind: bool = False  # local gates only (tests, dry runs)


# --- numbers: does the prepared workbook compute the same values? -----------------
def _is_error(v: Any) -> bool:
    return isinstance(v, str) and v.strip().upper() in ERROR_LITERALS


def _same(a: Any, b: Any) -> bool:
    if a is None and b is None:
        return True
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= NUMERIC_TOL * max(1.0, abs(a), abs(b))
    if _is_error(a) and _is_error(b):
        return True  # an error may become another error; no valid value is involved
    if isinstance(a, _dt.datetime) and isinstance(b, _dt.datetime):
        return a == b
    return a == b


def _rewrite_ok(a: Any, b: Any) -> bool:
    """A cell whose formula the plan rewrote passes only if its value is
    unchanged, or an error became another error (or nothing)."""
    return _same(a, b) or (_is_error(a) and (_is_error(b) or b is None))


def row_map(inserts: list[int]) -> Callable[[int], int]:
    """Source row -> prepared row after whole-row inserts at `inserts`
    (source coordinates): inserting at k pushes rows >= k down by one."""
    ins = sorted(int(k) for k in inserts)

    def f(r: int) -> int:
        return r + sum(1 for k in ins if k <= r)

    return f


def _name_targets(defined_names: list[dict[str, Any]]) -> dict[str, tuple[str, str] | None]:
    names: dict[str, tuple[str, str] | None] = {}
    for d in defined_names:
        val = str(d.get("value") or "")
        m = re.match(r"^=?'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$", val.replace("$", ""))
        names[str(d.get("name") or "").upper()] = (m.group(1), f"{m.group(2)}{m.group(3)}") if m else None
    return names


def volatile_cells(formulas: list[dict[str, Any]], defined_names: list[dict[str, Any]]) -> set[tuple[str, str]]:
    """Cells whose value legitimately differs between two recalculations:
    NOW()/TODAY()/RAND() and, transitively, every cell whose formula reads one
    of them (directly, or through a defined name)."""
    by_cell = {(f["sheet"], f["cell"].replace("$", "").upper()): f["formula"] for f in formulas}
    vol: set[tuple[str, str]] = {k for k, fx in by_cell.items() if VOLATILE.search(fx or "")}
    names = _name_targets(defined_names)
    if not vol:
        return vol

    # a reference -- even a whole column like A:A -- is tested against the
    # small volatile set instead of being expanded cell by cell
    def coords(key: tuple[str, str]) -> tuple[str, int, int]:
        r = parse_ref(key[1])
        return (key[0], r["r1"], r["c1"]) if r else (key[0], 0, 0)

    changed = True
    rounds = 0
    while changed and rounds < 50:
        changed = False
        rounds += 1
        vol_coords = [coords(k) for k in vol]
        for key, fx in by_cell.items():
            if key in vol or not fx:
                continue
            sheet = key[0]
            hit = False
            for ref in cell_refs_in_formula(fx):
                target_sheet = ref.get("sheet") or sheet
                r1, r2, c1, c2 = ref["r1"], ref["r2"], ref["c1"], ref["c2"]
                if any(s == target_sheet and r1 <= rr <= r2 and c1 <= cc <= c2 for s, rr, cc in vol_coords):
                    hit = True
                    break
            if not hit:
                for token in re.findall(r"[A-Za-z_][A-Za-z0-9_.]*", fx):
                    tgt = names.get(token.upper())
                    if tgt and tgt in vol:
                        hit = True
                        break
            if hit:
                vol.add(key)
                changed = True
    return vol


def _allowed_cells(applied: list[dict[str, Any]]) -> tuple[dict[tuple[str, str], str], dict[tuple[str, str], str], dict[str, list[int]], dict[str, str]]:
    """What the plan wrote, keyed by the plan's (source) sheet name:
    (source-coordinate cells -> mode, prepared-coordinate cells written after
    inserts -> mode, row inserts per sheet, sheet renames).
    mode 'skip'    set_value / clear_cell -- the new content IS the intent
    mode 'relaxed' set_formula / set_array_formula -- compared, error->error allowed"""
    src: dict[tuple[str, str], str] = {}
    prepared: dict[tuple[str, str], str] = {}
    inserts: dict[str, list[int]] = {}
    renames: dict[str, str] = {}
    for o in applied:
        if o["op"] == "insert_row":
            inserts.setdefault(o["sheet"], []).append(int(o["row"]))
        elif o["op"] == "rename_sheet":
            renames[o["sheet"]] = str(o["after"])
    for o in prep._ordered(applied):
        if o["op"] not in prep.CELL_OPS:
            continue
        mode = "relaxed" if o["op"] in RELAXED_OPS else "skip"
        target = prepared if o.get("after_inserts") else src
        for c in prep._cells_of(o.get("range") or o.get("cell")):
            key = (o["sheet"], c)
            if target.get(key) != "skip":  # a deliberate content write on the same cell wins
                target[key] = mode
    return src, prepared, inserts, renames


def compare_values(source_path: Path, prepared_path: Path, applied: list[dict[str, Any]], formulas: list[dict[str, Any]], defined_names: list[dict[str, Any]], max_report: int = 40) -> dict[str, Any]:
    """Cell-for-cell comparison of two workbooks' stored values, mapping the
    prepared workbook's coordinates back through the plan's row inserts.
    Both files should have been recalculated by Excel first."""
    src_allowed, prep_allowed, inserts, renames = _allowed_cells(applied)
    vol = volatile_cells(formulas, defined_names)
    wb_s = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    wb_p = openpyxl.load_workbook(prepared_path, data_only=True, read_only=True)
    diffs: list[dict[str, Any] | None] = []
    compared = 0
    skipped_volatile = 0
    inserted_cells_checked = 0

    def record(d: dict[str, Any]) -> None:
        diffs.append(d if sum(1 for x in diffs if x) < max_report else None)  # counted, not listed

    try:
        for ws_s in wb_s.worksheets:
            if not hasattr(ws_s, "iter_rows"):  # chart sheets carry no cells
                continue
            sheet = ws_s.title
            p_name = renames.get(sheet, sheet)
            if p_name not in wb_p.sheetnames or not hasattr(wb_p[p_name], "iter_rows"):
                record({"sheet": sheet, "cell": "*", "source": "sheet", "prepared": "missing"})
                continue
            ws_p = wb_p[p_name]
            rmap = row_map(inserts.get(sheet, []))
            # read-only sheets yield EmptyCell objects without coordinates, so
            # coordinates come from the position in the (1-based) iteration
            p_values: dict[tuple[int, int], Any] = {}
            for pr_, row in enumerate(ws_p.iter_rows(min_row=1, min_col=1, values_only=True), start=1):
                for pc_, v in enumerate(row, start=1):
                    if v is not None:
                        p_values[(pr_, pc_)] = v
            seen_p: set[tuple[int, int]] = set()
            for r, row in enumerate(ws_s.iter_rows(min_row=1, min_col=1, values_only=True), start=1):
                for col, a in enumerate(row, start=1):
                    pr = rmap(r)
                    seen_p.add((pr, col))
                    ref_s, ref_p = ref_text(col, r), ref_text(col, pr)
                    mode = src_allowed.get((sheet, ref_s)) or prep_allowed.get((sheet, ref_p))
                    if mode == "skip":
                        continue
                    if (sheet, ref_s) in vol:
                        skipped_volatile += 1
                        continue
                    b = p_values.get((pr, col))
                    if a is None and b is None:
                        continue
                    compared += 1
                    ok = _rewrite_ok(a, b) if mode == "relaxed" else _same(a, b)
                    if not ok:
                        record({"sheet": sheet, "cell": ref_s, "prepared_cell": ref_p, "source": _short(a), "prepared": _short(b), "rewritten": mode == "relaxed"})
            # anything in the prepared sheet that maps to no source cell must be a title the plan wrote
            for (pr, col), val in p_values.items():
                if (pr, col) in seen_p:
                    continue
                inserted_cells_checked += 1
                ref_p = ref_text(col, pr)
                if prep_allowed.get((sheet, ref_p)) == "skip":
                    continue
                record({"sheet": sheet, "cell": None, "prepared_cell": ref_p, "source": None, "prepared": _short(val)})
    finally:
        wb_s.close()
        wb_p.close()
    listed = [d for d in diffs if d]
    return {
        "match": not diffs,
        "compared": compared,
        "differences": len(diffs),
        "listed": listed,
        "volatile_skipped": skipped_volatile,
        "inserted_cells_checked": inserted_cells_checked,
    }


def _short(v: Any) -> Any:
    if isinstance(v, (int, float, bool)) or v is None:
        return v
    return str(v)[:60]


def numbers_gate(source: Path, prepared: Path, applied: list[dict[str, Any]], analysis: dict[str, Any], work: Path) -> dict[str, Any]:
    """Recalculate fresh copies of both workbooks in Excel, then compare.
    Always returns a complete shape: ran=False means the gate could not be
    checked (which the loop treats as a failure, never a pass)."""
    wb0 = analysis["workbooks"][0]
    out: dict[str, Any] = {"ran": False, "requested": True}
    try:
        s_copy, _ = make_immutable_copy(source, work / "source")
        p_copy, _ = make_immutable_copy(prepared, work / "prepared")
        rs = recalculate(s_copy)
        rp = recalculate(p_copy)
        out["source_recalc"] = {"status": rs["status"], "ran": rs["ran"], "errors": len(rs["formula_errors"])}
        out["prepared_recalc"] = {"status": rp["status"], "ran": rp["ran"], "errors": len(rp["formula_errors"])}
        if not (rs["ran"] and rp["ran"]):
            out["message"] = f"Excel recalculation did not run ({rs['message'] if not rs['ran'] else rp['message']})"
            return out
        before = {(x["sheet"], x["cell"]) for x in rs["formula_errors"]}
        out["new_formula_errors"] = [e for e in rp["formula_errors"] if (e["sheet"], e["cell"]) not in before]
        cmp = compare_values(s_copy, p_copy, applied, wb0.get("formulas", []), wb0.get("defined_names", []))
        out.update(cmp)
        out["ran"] = True
        out["match"] = bool(cmp["match"])
    except Exception as exc:  # the gate must report, never crash the loop
        out["ran"] = False
        out["message"] = f"numbers gate failed: {exc}"[:300]
        out.setdefault("match", False)
        out.setdefault("differences", 0)
        out.setdefault("listed", [])
    return out


# --- names and structure ---------------------------------------------------------
def names_gate(after_analysis: dict[str, Any], after_report: dict[str, Any], baseline_grid_count: int) -> dict[str, Any]:
    grids = all_grids(after_analysis)
    unnamed = [f"{g['sheet']}!{g['ref']}" for g in grids if not g.get("name")]
    unnamed_anchors = [g["anchor"] for g in grids if not g.get("name")]  # what Mind will show as Untitled(row,col)
    plan = {a["id"]: a for a in prep.plan_actions(after_analysis, after_report)}
    titles = plan["create_grid_titles"]
    blocked = [s for s in titles["skipped"] if "standalone text" not in s]
    limit = baseline_grid_count * (1 + FRAGMENTATION_SLACK) + 5
    return {
        "grids": len(grids),
        "baseline_grids": baseline_grid_count,
        "fragmented": len(grids) > limit,
        "unnamed": len(unnamed),
        "unnamed_grids": unnamed,
        "unnamed_anchors": unnamed_anchors,
        "still_titleable": titles["count"],
        "blocked": blocked,
        "ok": len(grids) <= limit and titles["count"] == 0,
    }


# --- deciding ----------------------------------------------------------------------
def active_actions(plan: list[dict[str, Any]], enabled: set[str], disabled: set[str]) -> set[str]:
    """The actions an iteration actually applies: default-on ones plus enabled
    opt-ins, minus disabled, with something to do. One definition, used both
    to build the iteration and to judge it."""
    return {a["id"] for a in plan if (a["default_on"] or a["id"] in enabled) and a["id"] not in disabled and a["count"] > 0}


def auto_enable(plan: list[dict[str, Any]], report: dict[str, Any]) -> list[str]:
    """Opt-in actions the loop switches on by itself: those whose rule is an
    ERROR finding (a blocker) and that have something to do."""
    errors = {f["rule_id"] for f in report["findings"] if f["status"] == "ERROR"}
    return sorted(a["id"] for a in plan if not a["default_on"] and a["count"] > 0 and set(a["rule_ids"]) & errors)


def _writers(plan: list[dict[str, Any]]) -> dict[tuple[str, str], set[str]]:
    """(sheet, cell) in source coordinates -> actions whose operations write it."""
    out: dict[tuple[str, str], set[str]] = {}
    for a in plan:
        for o in a["operations"]:
            if o["op"] in prep.CELL_OPS and not o.get("after_inserts"):
                for c in prep._cells_of(o.get("range") or o.get("cell")):
                    out.setdefault((o["sheet"], c), set()).add(a["id"])
    return out


def _rect_writers(writers: dict[tuple[str, str], set[str]], sheet: str, r1: int, r2: int, c1: int, c2: int) -> set[str]:
    hit: set[str] = set()
    for (s, cell), acts in writers.items():
        if s != sheet:
            continue
        r = parse_ref(cell)
        if r and r1 <= r["r1"] <= r2 and c1 <= r["c1"] <= c2:
            hit |= acts
    return hit


def actions_touching(plan: list[dict[str, Any]], diffs: list[dict[str, Any]], formulas: list[dict[str, Any]] | None = None, defined_names: list[dict[str, Any]] | None = None) -> list[str]:
    """Actions to blame for changed cells. A cell an action wrote is blamed
    directly; otherwise the changed cell's precedents are walked (bounded)
    back to a cell an action wrote -- so a formula rewrite that moved a value
    downstream is blamed before any row insert on the same sheet. Only when
    no writer is reachable does an insert on the sheet take the blame."""
    writers = _writers(plan)
    by_cell = {(f["sheet"], f["cell"].replace("$", "").upper()): f["formula"] for f in (formulas or [])}
    names = _name_targets(defined_names or [])
    blamed: set[str] = set()
    unexplained_sheets: set[str] = set()
    for d in diffs:
        sheet, cell = d.get("sheet"), d.get("cell")
        found: set[str] = set()
        if cell:
            found |= writers.get((sheet, cell), set())
            if not found:
                seen: set[tuple[str, str]] = set()
                queue: deque[tuple[str, str]] = deque([(sheet, cell)])
                while queue and len(seen) < PRECEDENT_WALK_LIMIT and not found:
                    key = queue.popleft()
                    if key in seen:
                        continue
                    seen.add(key)
                    fx = by_cell.get(key)
                    if not fx:
                        continue
                    for ref in cell_refs_in_formula(fx):
                        ts = ref.get("sheet") or key[0]
                        r1, r2, c1, c2 = ref["r1"], ref["r2"], ref["c1"], ref["c2"]
                        found |= _rect_writers(writers, ts, r1, r2, c1, c2)
                        if found:
                            break
                        if (r2 - r1 + 1) * (c2 - c1 + 1) <= 2000:
                            for rr in range(r1, r2 + 1):
                                for cc in range(c1, c2 + 1):
                                    k2 = (ts, ref_text(cc, rr))
                                    if k2 in by_cell and k2 not in seen:
                                        queue.append(k2)
                    if found:
                        break
                    for token in re.findall(r"[A-Za-z_][A-Za-z0-9_.]*", fx):
                        tgt = names.get(token.upper())
                        if tgt:
                            found |= writers.get(tgt, set())
                            if tgt in by_cell and tgt not in seen:
                                queue.append(tgt)
        if found:
            blamed |= found
        else:
            unexplained_sheets.add(sheet)
    if not blamed:
        for a in plan:
            if any(o["op"] in ("insert_row", "insert_column") and o["sheet"] in unexplained_sheets for o in a["operations"]):
                blamed.add(a["id"])
    return sorted(blamed)


def insert_heaviest(plan: list[dict[str, Any]], candidates: set[str]) -> str | None:
    counts = {a["id"]: sum(1 for o in a["operations"] if o["op"] in ("insert_row", "insert_column")) for a in plan if a["id"] in candidates}
    counts = {k: v for k, v in counts.items() if v}
    return max(counts, key=counts.get) if counts else None


def _stuck(reason: str) -> dict[str, Any]:
    return {"verdict": "stuck", "reason": reason, "enable": [], "disable": []}


def decide(iteration: dict[str, Any], plan: list[dict[str, Any]], enabled: set[str], disabled: set[str], formulas: list[dict[str, Any]] | None = None, defined_names: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """The one rule that fires for an iteration -> {'verdict', 'reason',
    'enable', 'disable'}. verdict: 'converged' | 'retry' | 'stuck'.
    A gate that could not be checked is never treated as passed."""
    apply = iteration.get("apply") or {}
    if apply.get("status") not in ("APPLIED", "NOT_APPLICABLE"):
        return _stuck(f"prep could not be applied: {apply.get('message')}")
    active = active_actions(plan, enabled, disabled)
    nums = iteration.get("numbers") or {}
    if nums.get("requested") and not nums.get("ran"):
        return _stuck(f"numbers could not be verified: {nums.get('message') or 'Excel recalculation did not run'}")
    if nums.get("ran") and not nums.get("match"):
        culprits = [a for a in actions_touching(plan, nums.get("listed", []), formulas, defined_names) if a in active]
        if culprits:
            return {"verdict": "retry", "reason": f"values changed in {nums.get('differences', 0)} cell(s); disabling {', '.join(culprits)}", "enable": [], "disable": culprits}
        return _stuck(f"values changed in {nums.get('differences', 0)} cell(s) and no active action can be blamed for it")
    names = iteration.get("names") or {}
    if names.get("fragmented"):
        culprit = insert_heaviest(plan, active)
        if culprit:
            return {"verdict": "retry", "reason": f"grid count {names.get('grids')} vs {names.get('baseline_grids')} -- the workbook is being fragmented; disabling {culprit}", "enable": [], "disable": [culprit]}
        return _stuck("grid count ballooned with no row-inserting action active")
    mind = iteration.get("mind") or {}
    if mind.get("skipped"):
        return {"verdict": "converged", "reason": "local gates passed (Mind skipped)", "enable": [], "disable": []}
    if mind.get("error"):
        return _stuck(f"Mind automation stopped before the model was verified: {mind['error']}")
    conv = mind.get("convert") or {}
    if not conv.get("success"):
        wanted: set[str] = set()
        unmapped: list[str] = []
        for msg in conv.get("errors", []):
            for pat, action in MIND_ERROR_TO_ACTION:
                if pat.search(msg):
                    wanted.add(action)
                    break
            else:
                unmapped.append(msg)
        real = sorted(a for a in wanted if any(p["id"] == a for p in plan) and a not in enabled and a not in disabled)
        if real:
            return {"verdict": "retry", "reason": f"Mind's Convert asked for {', '.join(real)}", "enable": real, "disable": []}
        detail = " | ".join(conv.get("errors") or unmapped) or conv.get("text", "")[:160] or "no error log captured"
        return _stuck("Mind's Convert failed for a reason the app has no fix for: " + detail)
    run = mind.get("run") or {}
    if iteration.get("run_model", True) and not run:
        return _stuck("Mind never ran the model (no run result recorded)")
    if run and not run.get("completed"):
        return _stuck(f"Mind ran the model but did not complete: {run.get('text', '')[:160]}")
    if run and not run.get("audit_consistent"):
        return _stuck("Mind ran the model but does not report it consistent with the audit trail")
    counts = mind.get("counts") or {}
    if counts.get("mind_untitled") is not None and (counts.get("unexplained") or counts.get("app_only")):
        # Mind reads some blocks differently from the app (a detection gap, not
        # something any prep action can change) -- reported, never a block
        gap = f"Mind lists {counts['mind_untitled']} Untitled grid(s), the app predicted {counts['app_untitled']}: {len(counts.get('unexplained') or [])} Mind-only, {len(counts.get('app_only') or [])} app-only (see detection_gaps)"
        return {"verdict": "converged", "reason": "every gate passed; " + gap, "enable": [], "disable": []}
    return {"verdict": "converged", "reason": "every gate passed", "enable": [], "disable": []}


def untitled_counts(mind_cells: list[str], app_anchors: list[str]) -> dict[str, Any]:
    """Mind's 'Untitled(row,col)' entries against the app's untitled anchors,
    as multisets of cell references (the selector does not say which sheet).
    'unexplained' = Mind shows an untitled grid the app did not know about --
    a detection gap and a failure. 'app_only' = the app split a block Mind
    reads as part of a larger grid -- reported, not a failure: nothing is
    shown untitled that the app could have titled."""
    from collections import Counter

    mind = Counter(c.upper() for c in mind_cells)
    app = Counter(a.upper() for a in app_anchors)
    unexplained = sorted((mind - app).elements())
    app_only = sorted((app - mind).elements())
    return {"mind_untitled": sum(mind.values()), "app_untitled": sum(app.values()), "unexplained": unexplained, "app_only": app_only}


# --- the loop ---------------------------------------------------------------------
def _write(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    tmp.replace(path)


def _assistant_names(analysis: dict[str, Any], progress: Progress) -> dict[str, str]:
    grids = all_grids(analysis)
    labels = prep.standalone_labels(analysis)
    targets = []
    for g in grids:
        if g.get("name") and not prep.is_weak_name(g.get("name"), g["sheet"]):
            continue
        det, _ = prep.deterministic_name(analysis, g, labels)
        if g.get("name") or prep.is_weak_name(det, g["sheet"]):
            targets.append(g)
    if not targets:
        return {}
    progress({"event": "names", "message": f"asking the assistant to name {len(targets)} grid(s)"})
    out = suggest_names([build_grid_context(analysis, g) for g in targets])
    if not out["available"]:
        progress({"event": "names", "message": f"assistant unavailable ({out.get('message')}); deterministic names stand"})
        return {}
    progress({"event": "names", "message": f"assistant named {len(out['names'])} grid(s)"})
    return dict(out["names"])


def run_loop(cfg: LoopConfig, progress: Progress | None = None) -> dict[str, Any]:
    progress = progress or (lambda e: None)
    cfg.work_dir.mkdir(parents=True, exist_ok=True)
    source = cfg.source.resolve()
    report: dict[str, Any] = {
        "source": str(source),
        "started": _now(),
        "config": {k: (str(v) if isinstance(v, Path) else v) for k, v in asdict(cfg).items()},
        "iterations": [],
        "projects": [],
        "verdict": None,
        "reason": None,
    }
    report_path = cfg.work_dir / "loop_report.json"
    engine = RulesEngine()

    def finish(verdict: str, reason: str) -> dict[str, Any]:
        report["verdict"], report["reason"] = verdict, reason
        report["finished"] = _now()
        report["leftover_projects"] = [p["name"] for p in report["projects"] if p.get("status") not in ("deleted", "not created")]
        _write(report_path, report)
        progress({"event": "done", "verdict": verdict, "message": f"{verdict}: {reason}"})
        return report

    if source.suffix.lower() == ".xlsb":
        # Mind accepts .xlsx/.xlsm only; the same Excel Save-As the upload endpoint uses.
        target = "xlsm" if has_vba_project(source) else "xlsx"
        progress({"event": "log", "message": f"converting {source.name} to .{target} through Excel (Mind does not accept .xlsb)"})
        conv = convert_output_format(source, cfg.work_dir / "source", target)
        if conv.get("status") != "APPLIED":
            return finish("stuck", f"could not convert .xlsb: {conv.get('message')}")
        source = Path(conv["output_path"]).resolve()
        report["converted_source"] = str(source)

    enabled: set[str] = set(cfg.enable)
    disabled: set[str] = set(cfg.disable)
    names: dict[str, str] | None = None
    baseline_grids: int | None = None
    stamp = time.strftime("%Y%m%d_%H%M%S")

    def log(msg: str, **extra: Any) -> None:
        progress({"event": "log", "message": msg, **extra})

    for n in range(1, cfg.max_iterations + 1):
        it: dict[str, Any] = {"n": n, "enabled": sorted(enabled), "disabled": sorted(disabled), "run_model": cfg.run_model, "started": _now()}
        it_dir = cfg.work_dir / f"iter_{n}"
        it_dir.mkdir(parents=True, exist_ok=True)
        progress({"event": "iteration", "n": n, "message": f"iteration {n}: analysing the source"})

        # 1. analyse the ORIGINAL source
        res = plan_mode.run(source, it_dir / "analysis", {}, engine)
        analysis, vreport = res["workbook_analysis"], res["validation_report"]
        wb0 = analysis["workbooks"][0]
        if baseline_grids is None:
            baseline_grids = len(all_grids(analysis))
        if names is None:
            names = _assistant_names(analysis, progress) if cfg.use_assistant else {}
            _write(cfg.work_dir / "assistant_names.json", names)
        plan = prep.plan_actions(analysis, vreport, names or None)
        if n == 1:
            for a in auto_enable(plan, vreport):
                if a not in disabled:
                    enabled.add(a)
                    log(f"enabling opt-in action '{a}': its rule is an ERROR finding")
        active = active_actions(plan, enabled, disabled)
        chosen = [a for a in plan if a["id"] in active]
        ops = [o for a in chosen for o in a["operations"]]
        it["actions"] = [{"id": a["id"], "count": a["count"]} for a in chosen]
        it["findings_before"] = vreport["summary"]["status_counts"]
        log(f"iteration {n}: applying {len(ops)} operation(s) from {', '.join(a['id'] for a in chosen) or 'no actions'}")

        # 2. apply through Excel
        if ops:
            applied = prep.apply_operations(source, it_dir / "prep", ops, prefer_excel=True)
        else:
            copy, _ = make_immutable_copy(source, it_dir / "prep")
            applied = {"status": "NOT_APPLICABLE", "output_path": copy, "applied": [], "failed": [], "message": "no operations", "verified_opens_in_excel": None}
        prepared = Path(applied["output_path"])
        it["apply"] = {"status": applied["status"], "applied": len(applied.get("applied", [])), "failed": len(applied.get("failed", [])), "verified": applied.get("verified_opens_in_excel"), "message": applied.get("message"), "output": str(prepared)}
        it["prepared_path"] = str(prepared)
        if applied["status"] not in ("APPLIED", "NOT_APPLICABLE"):
            d = it["decision"] = decide(it, plan, enabled, disabled)
            it["finished"] = _now()
            report["iterations"].append(it)
            progress({"event": "decision", "n": n, "verdict": d["verdict"], "message": f"iteration {n}: {d['verdict']} -- {d['reason']}"})
            return finish(d["verdict"], d["reason"])

        # 3. local gates
        after = plan_mode.run(prepared, it_dir / "after", {}, engine)
        it["findings_after"] = after["validation_report"]["summary"]["status_counts"]
        it["names"] = names_gate(after["workbook_analysis"], after["validation_report"], baseline_grids)
        log(f"iteration {n}: grids {it['names']['grids']} (baseline {baseline_grids}), unnamed {it['names']['unnamed']}, blocked {len(it['names']['blocked'])}")
        if cfg.check_numbers:
            progress({"event": "numbers", "message": f"iteration {n}: recalculating source and prepared copies in Excel"})
            it["numbers"] = numbers_gate(source, prepared, applied.get("applied", []), analysis, it_dir / "numbers")
            nm = it["numbers"]
            log(f"iteration {n}: numbers {'match' if nm.get('match') else 'DIFFER' if nm.get('ran') else 'NOT CHECKED (' + str(nm.get('message')) + ')'} ({nm.get('compared', 0)} cells compared, {nm.get('differences', 0)} differ, {nm.get('volatile_skipped', 0)} volatile skipped)")
        else:
            it["numbers"] = {"ran": False, "requested": False, "message": "disabled"}
        nums = it["numbers"]
        local_fail = (nums.get("requested") and not nums.get("ran")) or (nums.get("ran") and not nums.get("match")) or it["names"]["fragmented"]
        if local_fail or cfg.skip_mind:
            it["mind"] = {"skipped": True, "reason": "local gate failed" if local_fail else "skip_mind"}
        else:
            # 4. Mind
            it["mind"] = _mind_iteration(cfg, it_dir, prepared, it["names"], [s["name"] for s in wb0["sheets"]], n, stamp, report, progress)
        it["decision"] = decide(it, plan, enabled, disabled, wb0.get("formulas", []), wb0.get("defined_names", []))
        it["finished"] = _now()
        report["iterations"].append(it)
        _write(report_path, report)
        d = it["decision"]
        progress({"event": "decision", "n": n, "verdict": d["verdict"], "message": f"iteration {n}: {d['verdict']} -- {d['reason']}"})
        if d["verdict"] == "converged":
            report["final_workbook"] = str(prepared)
            counts = (it.get("mind") or {}).get("counts") or {}
            report["detection_gaps"] = {"mind_only": counts.get("unexplained") or [], "app_only": counts.get("app_only") or []}
            report["residual_untitled"] = {"app": it["names"]["unnamed"], "mind": counts.get("mind_untitled"), "blocked": it["names"]["blocked"]}
            return finish("converged", d["reason"])
        if d["verdict"] == "stuck":
            return finish("stuck", d["reason"])
        enabled.update(d["enable"])
        disabled.update(d["disable"])
    return finish("exhausted", f"no convergence after {cfg.max_iterations} iteration(s)")


def _mind_iteration(cfg: LoopConfig, it_dir: Path, prepared: Path, names: dict[str, Any], sheet_names: list[str], n: int, stamp: str, report: dict[str, Any], progress: Progress) -> dict[str, Any]:
    out: dict[str, Any] = {"skipped": False}
    shots = it_dir / "mind"
    name = f"{cfg.project_prefix}_{stamp}_i{n}"
    record = {"name": name, "status": "not created", "init_url": None, "ran": False}
    report["projects"].append(record)
    try:
        with mc.mind_session(headed=cfg.headed) as (_ctx, page):
            try:
                page.set_viewport_size({"width": 1680, "height": 1000})
            except Exception:
                pass
            progress({"event": "mind", "message": f"iteration {n}: creating sandbox project {name}"})
            if not mc.open_manager(page):
                out["error"] = "could not reach the Project manager (not logged in?)"
                return out
            mc.ensure_folder(page, cfg.folder)
            init = mc.create_blank_project(page, name)
            record["init_url"] = init
            if not init:
                out["error"] = "project was not created"
                return out
            record["status"] = "created"
            state = mc.open_project(page, init)
            if state != "empty":
                out["error"] = f"project opened in state '{state}', expected an empty drop zone"
                return out
            progress({"event": "mind", "message": f"iteration {n}: uploading {prepared.name} and waiting for Upload / Convert / Test"})
            out["convert"] = mc.upload_and_convert(page, prepared, shots=shots)
            conv = out["convert"]
            progress({"event": "mind", "message": f"iteration {n}: convert {'succeeded' if conv['success'] else 'FAILED'} {conv['steps']}" + (f" -- {' | '.join(conv['errors'])[:200]}" if conv["errors"] else "")})
            if not conv["success"]:
                return out
            if not mc.add_template(page):
                out["error"] = "Add template to Milliman Mind did not lead to a Run button"
                return out
            record["status"] = "model added"
            stem = prepared.stem
            listed = mc.template_names(page, stem)
            # the list is only trusted when it shows the workbook's own sheets
            trusted = bool(listed) and any(s in listed for s in sheet_names)
            untitled = mc.untitled_entries(listed) if trusted else []
            out["templates"] = {"stem": stem, "entries": len(listed), "trusted": trusted, "untitled": untitled}
            out["counts"] = untitled_counts([u["cell"] for u in untitled], names.get("unnamed_anchors", [])) if trusted else {"mind_untitled": None, "app_untitled": names["unnamed"]}
            c = out["counts"]
            progress({"event": "mind", "message": f"iteration {n}: Mind lists {len(listed)} template entries" + (f", {c['mind_untitled']} Untitled; app predicted {c['app_untitled']}, unexplained {len(c['unexplained'])}, app-only {len(c['app_only'])}" if trusted else " (selector could not be read; count skipped)")})
            if cfg.run_model:
                progress({"event": "mind", "message": f"iteration {n}: running the model"})
                record["ran"] = True  # set before the click: a project that was run must never be deleted
                record["status"] = "run"
                out["run"] = mc.run_model(page, shots=shots)
                progress({"event": "mind", "message": f"iteration {n}: run {'completed' if out['run']['completed'] else 'did not complete'}" + (", consistent with the audit trail" if out["run"]["audit_consistent"] else "")})
                try:
                    out["exports"] = mc.export_summary(page)
                except Exception as exc:
                    out["exports"] = {"opened": False, "error": str(exc)[:120]}
            mc.screenshot(page, shots / "final.png")
    except Exception as exc:
        out["error"] = f"Mind automation failed: {exc}"[:300]
    finally:
        if cfg.delete_projects and record.get("status") in ("created", "model added") and not record.get("ran"):
            try:
                with mc.mind_session(headed=cfg.headed) as (_ctx, page):
                    st = mc.delete_project(page, name, cfg.folder)
                    record["delete"] = st
                    if st == "deleted":
                        record["status"] = "deleted"
            except Exception as exc:
                record["delete"] = f"failed: {exc}"[:120]
    return out


def summarize(report: dict[str, Any]) -> str:
    """Human summary for the CLI and the change log."""
    lines = [f"verdict: {report.get('verdict')} -- {report.get('reason')}", f"source: {report.get('source')}"]
    for it in report.get("iterations", []):
        nm, nums, mind = it.get("names") or {}, it.get("numbers") or {}, it.get("mind") or {}
        conv = (mind.get("convert") or {}).get("success")
        run = mind.get("run") or {}
        numbers = "match" if nums.get("match") else (f"differ:{nums.get('differences')}" if nums.get("ran") else ("not checked" if nums.get("requested") else "off"))
        lines.append(
            f"  iter {it['n']}: actions={','.join(a['id'] for a in it.get('actions', []))} | apply={it.get('apply', {}).get('status')} "
            f"| grids {nm.get('baseline_grids')}->{nm.get('grids')} unnamed={nm.get('unnamed')} blocked={len(nm.get('blocked', []))} "
            f"| numbers={numbers} "
            f"| convert={'ok' if conv else ('skipped' if mind.get('skipped') else 'FAIL')} run={'ok' if run.get('completed') else '-'} audit={'ok' if run.get('audit_consistent') else '-'} "
            f"| {it.get('decision', {}).get('verdict')}: {it.get('decision', {}).get('reason')}"
        )
    res = report.get("residual_untitled")
    if res:
        lines.append(f"untitled grids left: app {res['app']} (of which {len(res['blocked'])} the app cannot title safely), Mind shows {res['mind']}")
    gaps = report.get("detection_gaps") or {}
    if gaps.get("mind_only") or gaps.get("app_only"):
        lines.append(f"detection gaps vs Mind: {len(gaps.get('mind_only') or [])} untitled grid(s) only Mind sees ({', '.join(gaps.get('mind_only') or [])[:120]}), {len(gaps.get('app_only') or [])} only the app sees")
    if report.get("leftover_projects"):
        lines.append("leftover Mind projects (delete by hand): " + ", ".join(report["leftover_projects"]))
    if report.get("final_workbook"):
        lines.append(f"final workbook: {report['final_workbook']}")
    return "\n".join(lines)
