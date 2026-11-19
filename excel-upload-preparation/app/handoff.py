"""Never  stop for  a person when the app can fix it: the decisions, as plain functions (17 Nov 2026).

Start's Prep step is one server job, `prep-run` (app/web/server.py): Prep rounds while a round still
changes something,  then what is left goes to the assistant, whose fix is  written by itself only when
it is  narrow and kept only when no number  changed, its problem got smaller and no other blocking
problem got worse.  Only then is a person asked, in  plain words.

This module holds every rule of that job and every sentence it shows, and nothing else: no Excel, no
server, no model call. The job gives it plans, findings and counts, and does what it answers. So each
rule is tested without Excel (tests/unit/test_handoff.py), and the screens and scripts/night_chain.py
get the same answer.

The steps and who decides them:
  1. rounds        pass_ops, left_key, rounds_stop / rounds_go_on,  run_rounds (the loop itself)
  2. what is left problems_left (cells_of, sites_by_rule)
  4. each problem memo_key, question_for, narrow / narrow_problem, judge
  7. the end       stops_start, build_ask (ask_lines, why_line, show_me_for, feedback_note), no_key_ask
The sentences of the job's `message`: round_message, assistant_message, WRITE_MESSAGE, NUMBERS_MESSAGE,
undo_message, CHECK_MESSAGE, done_lines.

Nothing the user reads here holds a rule code, a version or (in the feedback note) a sheet or cell name.
"""
from __future__ import annotations

import hashlib
import re
from typing import Any, Callable, Iterable

from .formula_utils import REF_RE, mask_strings, num_to_col, parse_ref
from .plain_codes import rule_name
from .readiness import OUTSIDE_RULE

# ----  the limits  ---------------------------------------------------------------------------------
# A safety cap only: the rounds stop by themselves when a round changes nothing. Only an endless
# repair gets here (a row insert that moves its own row every round).
# 19 Nov 2026  (Joseph): 10 rounds, down from 20. On the models run so  far the rounds ended after 1  to 4.
MAX_PASSES = 10
# How  many problems  the assistant is given in one run; the others are "not_tried".
MAX_PROBLEMS = 5
# Tries  per problem.  One try is one assistant turn = up to 5 model calls (chat_context.py).
MAX_TRIES = 2
# How  many cells  the question names (the finding's own list is cut at 1800 characters by the chat).
MAX_QUESTION_CELLS = 200
# An  array formula  over more cells than this is never "narrow", whatever the finding lists.
MAX_ARRAY_CELLS = 2000
# 19  Nov 2026  (Joseph: "allow the fix to find other cells cautiously"): how many cells that are NOT the problem's own
# one fix may also rewrite. Only formula cells that read one of the problem's  cells count (see narrow_problem).
MAX_OTHER_CELLS = 5
# The  numbers check  does not compare a cell whose formula uses one of these (autofix.VOLATILE_FUNCS;
# tests/unit/test_handoff.py fails when the two drift apart): such a cell is never one of the "other" cells.
_VOLATILE_RE = re.compile(r"\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(", re.IGNORECASE)

# How  a problem  ended (the job's `problems[].outcome`).
# "not_reached" (17 Nov 2026): the call got no answer (no connection, a timeout, a key the gateway refuses)
OUTCOMES = ("fixed", "no_fix", "too_wide", "not_written", "numbers_changed", "not_cleared", "not_checked", "not_reached", "not_tried")

# Why  the rounds  stopped (run_rounds' `why`).
CLEAN, NOTHING_WRITTEN, SAME_AGAIN, MAX_PASSES_RUN = "clean", "nothing_written", "same_again", "max_passes"

PREP_ROUTES = ("prep", "prep_skipped")


# ----  step 1:  the Prep rounds --------------------------------------------------------------------
def _round_actions(plan: Iterable[dict[str, Any]] | None) -> list[dict[str, Any]]:
    """The  Prep actions  a round writes: blocking, on by default, with work left."""
    return [a for a in plan or [] if a.get("level") == "blocking" and a.get("default_on") and (a.get("count") or 0) > 0]


def pass_ops(plan: Iterable[dict[str, Any]] | None) -> list[dict[str, Any]]:
    """What  one round  writes.

    In: `plan`, the session's Prep plan (`s.plan`: a list of actions, each with `level`, `default_on`,
         `count`, `operations`).
    Out: every operation of every action with level "blocking", `default_on` and `count > 0`, in the
         plan's  order: the list for one `_apply_and_version(s, ops, "prep", "Prep", reanalyze=True)`.
         Empty: nothing  is left for Prep to write."""
    return [o for a in _round_actions(plan) for o in a.get("operations") or []]


def _op_place(o: dict[str, Any]) -> str:
    """Where  an operation  writes: its range or cell; for the few with neither, what stands for it (the
    row or column an insert opens, the file a link names, a sheet's new name)."""
    if o.get("range") or o.get("cell"):
        return str(o.get("range") or o.get("cell"))
    if o.get("row") is not None:
        return f"row {o['row']}"
    if o.get("column"):
        return f"column {o['column']}"
    if isinstance(o.get("cells"), list):
        return ",".join(str(c) for c in o["cells"])
    if o.get("file_name") or o.get("target"):
        return f"{o.get('file_name') or ''}>{o.get('target') or ''}"
    return str(o.get("after") if o.get("after") is not None else "")


def left_key(plan: Iterable[dict[str, Any]] | None) -> frozenset[str]:
    """What  the blocking  Prep actions would still write, as a set that can be compared between rounds.

    In: `plan`, as for pass_ops.
    Out: one text per operation of pass_ops(plan): "action_id | rule_id | op | sheet  | cell or range".
         Equal sets before and after a round mean  the same operations came back: writing  them again
         would change nothing. What an operation  would write is not part of the key on purpose (a
         repair  that rewrites its own cell every round  must not run every round up to the cap)."""
    return frozenset(" | ".join(str(x or "") for x in (o.get("action_id"), o.get("rule_id"), o.get("op"), o.get("sheet"), _op_place(o))) for o in pass_ops(plan))


def rounds_stop(pass_no: int, applied: int, before: frozenset[str], after: frozenset[str], max_passes: int = MAX_PASSES) -> str | None:
    """Why  the rounds  stop after round `pass_no`, or None when another round should run.

    In: `pass_no` (1 for the first round), `applied` = how many changes that round wrote, `before` and
         `after` = left_key of the  plan before and after it.
    Out: "clean"           nothing is left to write;
         "nothing_written" the  round wrote nothing (every change was  refused);
         "same_again"      the same operations came back;
         "max_passes"      `max_passes` rounds were  run;
         None              go on: the round wrote something and left a different plan,  however small.
    A refused change alone never stops  the rounds: it comes back as "nothing_written" or "same_again"."""
    if not after:
        return CLEAN
    if applied <= 0:
        return NOTHING_WRITTEN
    if after == before:
        return SAME_AGAIN
    if pass_no >= max_passes:
        return MAX_PASSES_RUN
    return None


def rounds_go_on(pass_no: int, applied: int, before: frozenset[str], after: frozenset[str], max_passes: int = MAX_PASSES) -> bool:
    """True  when another  Prep round should run after round `pass_no` (see rounds_stop for the inputs)."""
    return rounds_stop(pass_no, applied, before, after, max_passes) is None


def run_rounds(plan_now: Callable[[], Iterable[dict[str, Any]] | None], write_round: Callable[[list[dict[str, Any]], int], tuple[int, int]], max_passes: int = MAX_PASSES) -> dict[str, Any]:
    """Step  1 of  the job: Prep rounds until a round changes nothing.

    In: `plan_now()` gives the session's plan as it is NOW (it is asked before and after every round);
         `write_round(ops, pass_no)` writes one round (the  Apply with its re-analysis) and returns
         `(applied, refused)`, the counts of that  round. It may raise: nothing is  caught here, so a
         Cancel (`cancel.Cancelled`) and an  error both reach the job as they are.
    Out: {"passes": rounds run,  "applied": changes written in all, "refused": changes  refused in all,
          "why": "clean" | "nothing_written" | "same_again" | "max_passes" (rounds_stop),
          "left": how many operations the blocking Prep actions still hold at the end}."""
    passes = applied = refused = 0
    why = CLEAN
    while True:
        plan = list(plan_now() or [])
        ops = pass_ops(plan)
        if not ops:  # (a) nothing left to write: clean
            why = CLEAN
            break
        before = left_key(plan)
        passes += 1
        wrote, not_written = write_round(ops, passes)
        applied += int(wrote or 0)
        refused += int(not_written or 0)
        stop = rounds_stop(passes, int(wrote or 0), before, left_key(plan_now()), max_passes)
        if stop is not None:
            why = stop
            break
    return {"passes": passes, "applied": applied, "refused": refused, "why": why, "left": len(pass_ops(plan_now()))}


def stops_start(plan: Iterable[dict[str, Any]] | None) -> bool:
    """`ask.stops_start`:  a blocking  Prep action that is on by default still has changes in the plan.
    This is the rule by which the Prep step stopped Start before tonight; other blocking problems are
    asked about at  the end of Start and do not stop it."""
    return bool(pass_ops(plan))


# ---- step 2:  what is left -----------------------------------------------------------------------
# The lists a  finding's `observed` may hold its places in,  as readiness._site_count reads them.
_SITE_KEYS = ("sites", "cells", "call_sites", "problems")


def _clean_cell(sheet: Any, cell: Any) -> tuple[str, str] | None:
    """(sheet,  cell) with  the cell as Excel writes it (no "$", capitals), or None when it is no cell or
    range. A cell given as "Sheet!A1" brings its own sheet."""
    ref = parse_ref(str(cell or "").strip())
    if not ref or ref["whole_column"] or ref["whole_row"]:
        return None
    name = str(ref["sheet"] or sheet or "").strip()
    return (name, ref["ref"]) if name else None


def cells_of(finding: dict[str, Any] | None) -> list[dict[str, str]]:
    """The  finding's own  cells: where its problem is.

    In: a finding of the validation report (`location`, `observed`).
    Out: [{"sheet", "cell"}, ...] without repeats: the finding's `location` first, then every place of
         its first list among observed "sites",  "cells", "call_sites", "problems" (the list readiness
         counts as `sites`). A cell is  "W25" or a range "C7:C9", never  with "$". Entries that name no
         cell are  left out; a finding with none gives []."""
    out: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    def add(sheet: Any, cell: Any) -> None:
        got = _clean_cell(sheet, cell)
        if got and (got[0].lower(), got[1]) not in seen:
            seen.add((got[0].lower(), got[1]))
            out.append({"sheet": got[0], "cell": got[1]})

    f = finding or {}
    loc = f.get("location") if isinstance(f.get("location"), dict) else {}
    add(loc.get("sheet"), loc.get("cell"))
    obs = f.get("observed")
    if isinstance(obs, dict):
        for key in _SITE_KEYS:
            if isinstance(obs.get(key), list):
                for item in obs[key]:
                    if isinstance(item, dict):
                        add(item.get("sheet"), item.get("cell") or item.get("ref"))
                    elif isinstance(item, str) and "!" in item:
                        add(None, item)
                break
    if not out:  # a finding that points  at an area only
        add(loc.get("sheet"), loc.get("ref"))
    return out


def _cell_count(finding: dict[str, Any] | None, named: list[dict[str, str]], sites: Any) -> int:
    """How  many cells  the problem is in: the finding's own full count when it gives one (the list it
    shows may be cut), else what readiness counted, else the cells that can be named."""
    obs = (finding or {}).get("observed")
    full = obs.get("sites_count") if isinstance(obs, dict) else None
    for n in (full, sites):
        if isinstance(n, int) and not isinstance(n, bool) and n > 0 and named:
            return max(n, 1)
    return len(named)


def problems_left(blocking: Iterable[dict[str, Any]] | None, report: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    """What  is left  for the assistant (step 2), and at the end what is left to ask about (step 7).

    In: `blocking` = `s.readiness()["blocking"]` (each with `rule_id`, `rule_name`, `location`, `fix`,
         `sites`); `report` = the validation report  (`s.result["validation_report"]`), for the cells.
    Out: one entry per blocking rule except the  outside-files rule (that is step 1  of Start), Prep's
         own leftovers first (route "prep"  or "prep_skipped"), each in the shape of the job's `problems`:
         {"rule_id", "rule_name",  "sheet", "cell", "cells": how many, "route": "prep"  | "prep_skipped" |
          "assistant", "tries": 0, "outcome": "not_tried"}. `sheet` and `cell` are "" when the finding
         names none."""
    findings = {f.get("rule_id"): f for f in (report or {}).get("findings") or []}
    out = []
    for b in blocking or []:
        rule = b.get("rule_id")
        if rule == OUTSIDE_RULE or b.get("fix") == "outside_inputs":
            continue
        finding = findings.get(rule) or {"location": b.get("location") or {}}
        named = cells_of(finding)
        first = named[0] if named else {"sheet": str((b.get("location") or {}).get("sheet") or ""), "cell": ""}
        out.append({
            "rule_id": rule,
            "rule_name": b.get("rule_name") or rule_name(rule),
            "sheet": first["sheet"],
            "cell": first["cell"],
            "cells": _cell_count(findings.get(rule), named, b.get("sites")),
            "route": b.get("fix") if b.get("fix") in PREP_ROUTES else "assistant",
            "tries": 0,
            "outcome": "not_tried",
        })
    return [p for p in out if p["route"] in PREP_ROUTES] + [p for p in out if p["route"] not in PREP_ROUTES]


def sites_by_rule(blocking: Iterable[dict[str, Any]] | None) -> dict[str, int | None]:
    """`{rule_id:  sites}` of  `s.readiness()["blocking"]`: what judge compares before and after a fix
    (None for a rule whose finding lists no places)."""
    return {str(b.get("rule_id")): b.get("sites") for b in blocking or []}


# ---- step 4a: the memo ---------------------------------------------------------------------------
def _range_cells(ref: dict[str, Any], limit: int) -> list[tuple[int, int]]:
    """(row,  col) of  the first `limit` cells of a parsed reference, row by row."""
    out = []
    for r in range(ref["r1"], ref["r2"] + 1):
        for c in range(ref["c1"], ref["c2"] + 1):
            out.append((r, c))
            if len(out) >= limit:
                return out
    return out


def memo_key(finding: dict[str, Any], analysis: dict[str, Any] | None, read: Callable[[str, int, int], Any] | None = None) -> str:
    """The  key under  which a problem the assistant failed on is remembered (`HANDOFF_MEMO[sid][key]`).

    In: the finding, and the analysis it came from (`s.result["analysis"]`). `read(sheet, row, col)`
         gives a cell's content; left out, the content is read  from the analysis' own copy of the
         workbook (inventory.cell_value: openpyxl, no Excel).
    Out:  "<rule id>:<16 hex digits>": the rule  id plus a hash of the finding's  cells with their
         content (and of how many cells the finding counts).  The same problem on an untouched
         workbook gives  the same key, so the assistant is not paid twice; a hand edit of one of those
         cells gives another key, and the assistant is asked again."""
    if read is None:
        def read(sheet: str, row: int, col: int) -> Any:
            from .inventory import cell_value  # here, not at the top: importing this module opens nothing

            return cell_value(analysis, sheet, row, col) if analysis else None

    h = hashlib.sha256()
    obs = finding.get("observed")
    h.update(repr(obs.get("sites_count") if isinstance(obs, dict) else None).encode("utf-8"))
    for c in sorted(cells_of(finding), key=lambda x: (x["sheet"].lower(), x["cell"])):
        ref = parse_ref(c["cell"])
        for row, col in _range_cells(ref, 64) if ref else []:
            try:
                content = read(c["sheet"], row, col)
            except Exception:  # a cell that cannot be read counts as empty: the key must never stop the job
                content = None
            h.update(repr((c["sheet"].lower(), row, col, content)).encode("utf-8", "replace"))
    return f"{finding.get('rule_id')}:{h.hexdigest()[:16]}"


# ---- step 4c: the question -----------------------------------------------------------------------
# The same words app/web/server.py `_chat_turn` puts before  a question about a finding (the Assistant
# tab's "Propose a fix"). tests/unit/test_handoff.py fails when the two drift apart.
_ABOUT_FINDING = (
    "[About finding {rule}{where}. If you propose a replacement formula, it must behave like the original on "
    "blank cells, text and errors (e.g. wrap with IF(x=\"\",0,...) when the original tolerated blanks); every new error "
    "after recalculation counts against the fix.] "
)
PROPOSE = "Propose a fix"
# Said  with the  cells, so the first proposal is one the guard (narrow) can let through.
_ONLY_THESE = (f"Change these cells. When the fix needs it, you may also give a new formula to at most {MAX_OTHER_CELLS} other cells whose formula reads one of these cells, "
               "and you may write a label into one empty cell. Nothing else is written without a person: no typed-in value is changed, no range, nothing that moves cells.")

# Why  the first  try failed, for the second try. For the assistant only (never on screen).
_EARLIER = {
    "no_fix": "Your first answer held no change the app could apply{detail}: end your answer with one changes block.",
    "too_wide": "Your first proposal was too wide to write without a person{detail}: change the cells listed, give a new formula to at most " + str(MAX_OTHER_CELLS) + " other cells whose formula reads one of them, or write a label into one empty cell.",
    "not_written": "Excel refused your first proposal{detail}: propose a different change.",
    "numbers_changed": "Your first proposal changed a computed result{detail}, so it was taken back: propose a change that leaves every result as it is.",
    "not_cleared": "Your first proposal was written, but it did not make this problem smaller or it made another one worse{detail}, so it was taken back: propose a different change.",
    "not_checked": "Your first proposal could not be checked{detail}, so it was taken back: propose a simpler change.",
    "fixed": "Your first proposal was kept and fixed part of the problem{detail}: now fix what is left of it.",
}


def cell_text(sheet: str, cell: str) -> str:
    """A  cell as  Excel writes it in a formula: Sheet!W25, or 'My sheet'!W25 when the name needs quotes
    (the chat's own look-up of the cells a question names reads this form)."""
    if not sheet:
        return cell
    return f"{sheet}!{cell}" if re.fullmatch(r"[\w.]+", sheet) else "'" + sheet.replace("'", "''") + f"'!{cell}"


def question_for(finding: dict[str, Any], cells: list[dict[str, str]] | None = None, earlier: str | None = None, detail: str = "") -> str:
    """The  question the  job asks the assistant about one problem (step 4c).

    In: `finding` = the report's finding, or the job's own problem entry (`rule_id`, and `sheet` and
         `cell` either at the top or in `location`);  `cells` = the finding's cells (cells_of; left out,
         they are taken from  `finding`); on a second try `earlier`  = the outcome of the first
         ("no_fix", "too_wide",  "not_written", "numbers_changed", "not_cleared", "not_checked", or
         "fixed" when what is left of a  problem gets its second try) and `detail`  = a few words on it
         (narrow_problem's answer, changed_sample's, Excel's refusal).
    Out: the text for `answer_question([], question, analysis, report, s.plan)`: the Assistant tab's
         "Propose a fix" for that finding, then the line "Cells: ..." naming every cell (up to
         MAX_QUESTION_CELLS, then how many more), then, on a second try, one sentence on why the first
         failed. It holds the rule id and the cells on purpose:  the chat pulls the finding and the cells'
         content in by them."""
    loc = finding.get("location") if isinstance(finding.get("location"), dict) else {}
    sheet = finding.get("sheet") or loc.get("sheet")
    cell = finding.get("cell") or loc.get("cell")
    where = f" at {sheet}!{cell}" if sheet and cell else ""
    text = _ABOUT_FINDING.format(rule=finding.get("rule_id"), where=where) + PROPOSE
    named = cells_of(finding) if cells is None else cells
    if not named and _clean_cell(sheet, cell):  # the job's own problem entry: its one cell
        named = [dict(zip(("sheet", "cell"), _clean_cell(sheet, cell)))]
    if named:
        shown = ", ".join(cell_text(c["sheet"], c["cell"]) for c in named[:MAX_QUESTION_CELLS])
        more = f", and {len(named) - MAX_QUESTION_CELLS:,} more" if len(named) > MAX_QUESTION_CELLS else ""
        text += f"\nCells: {shown}{more}. {_ONLY_THESE}"
    if earlier in _EARLIER:
        extra = f" ({detail.strip().rstrip('.')})" if detail and detail.strip() else ""
        text += "\n" + _EARLIER[earlier].format(detail=extra)
    return text


def changed_sample(numbers: dict[str, Any] | None) -> str:
    """The  first value  a numbers check found changed, in a few words for the second try's sentence
    ("Calc!D4 showed 100 and would show 0"); "" when the check names none. For the assistant only."""
    samples = (numbers or {}).get("samples") or {}
    # an error  that now shows a value counts too (17 Nov 2026: the  job undoes such a fix: _shown_changed)
    first = next((x for kind in ("good_to_error", "good_to_other", "missing", "error_to_good") for x in samples.get(kind) or [] if isinstance(x, dict)), None)
    if not first or not first.get("original"):
        return ""
    if "is" in first:
        return f"{first['original']} showed {first.get('was')} and would show {first['is']}"
    return f"{first['original']} ({first.get('was')}) would no longer be a formula"


# ----  step 4e: the guard --------------------------------------------------------------------------
def _own_cells(cells: Iterable[dict[str, str]] | None) -> tuple[list[tuple[str, int, int, int, int]], int]:
    """The  finding's cells  as rectangles (sheet in small letters, r1, c1, r2, c2), and how many cells they are."""
    rects, count = [], 0
    for c in cells or []:
        got = _clean_cell(c.get("sheet"), c.get("cell"))
        ref = parse_ref(got[1]) if got else None
        if ref:
            rects.append((got[0].lower(), ref["r1"], ref["c1"], ref["r2"], ref["c2"]))
            count += ref["cells"]
    return rects, count


def _reads_own(formula: str, home: str, rects: list[tuple[str, int, int, int, int]]) -> bool:
    """Does  `formula`, written  on sheet `home`, name one of the problem's own cells -- alone, or inside a range, a
    whole column or a whole row? Only references written out in the formula are followed  (no defined name, no
    3-D reference): what cannot be followed does not  count as reading."""
    masked = mask_strings(formula, mask_sheet_quotes=False)
    for m in REF_RE.finditer(masked):
        name = m.group("sheet")
        if name and name.startswith("'"):
            name = formula[m.start("sheet") + 1 : m.end("sheet") - 1].replace("''", "'")
        if name and ":" in name:
            continue
        ref = parse_ref(m.group("ref"))
        if not ref:
            continue
        sheet = (name or home).lower()
        if any(s == sheet and ref["r1"] <= r2 and r1 <= ref["r2"] and ref["c1"] <= c2 and c1 <= ref["c2"] for s, r1, c1, r2, c2 in rects):
            return True
    return False


def narrow_problem(ops: Iterable[dict[str, Any]] | None, cells: Iterable[dict[str, str]] | None,
                   formula_cells: Iterable[tuple[Any, Any]] | None = None) -> str | None:
    """Why  a proposal  is too wide to be written without a person, or None when it is narrow (step 4e).

    In: `ops` = the assistant's operations as prep.validate_proposal returns them (`op`, `sheet`,
         `cell` or `range`, `before`);  `cells` = the finding's own cells (cells_of).
    Out: None when EVERY operation  is one of
           - set_formula on one  of the finding's own cells;
           - set_array_formula whose  whole range lies in the finding's own cells;
           - set_value or clear_cell  on ONE cell that is one of  the finding's own cells;
           - set_value on ONE cell that is empty now (`before` is None: a label beside the problem, as
             Prep's own lone-text fix writes);
           - 19 Nov 2026 (Joseph: "allow the fix to find other cells cautiously"): set_formula on ONE cell
             that is not the finding's own, when all of this holds:
               . it HOLDS A FORMULA NOW: `before` starts with "=", and  the scan lists it as a formula cell
                 (`formula_cells`, (sheet, cell) pairs; not given: `before` alone is trusted). A text typed
                 as "=..." is a typed-in value;
               . that formula READS ONE OF THE FINDING'S OWN CELLS (the formula that reads a lone text cell
                 is the usual case); a cell with no tie to the problem is not touched;
               . neither that formula nor the new one (`after`) uses NOW, TODAY, RAND, RANDBETWEEN or
                 RANDARRAY: the numbers check does not compare such a cell;
               . there are  at most MAX_OTHER_CELLS such cells,  and the finding names cells of its  own.
             A typed-in value  is never  replaced, by a value or by a formula: the numbers check compares
             formula cells, and would not see an input change. What the numbers check cannot see at all --
             a result that  differs only at other inputs than today's -- is as true  of these cells as of the
             finding's  own: the reason they are few and  tied to the problem;
         and there are no more than 2 x  cells + 2 operations. Else a few  words on the first thing that
         is too wide (a range, an insert, a rename, hide, unprotect, a cell with content elsewhere, too
         many operations): for the second try's sentence, never for the screen.
    Why so narrow: the numbers check sees formula cells at today's inputs only, and until tonight a
    person read every proposal before Apply."""
    ops = list(ops or [])
    if not ops:
        return "there is no change in it"
    rects, count = _own_cells(cells)

    def own(sheet: str, row: int, col: int) -> bool:
        return any(s == sheet and r1 <= row <= r2 and c1 <= col <= c2 for s, r1, c1, r2, c2 in rects)

    others: set[tuple[str, int, int]] = set()  # the formula cells changed that are not the finding's  own
    known: set[tuple[str, str]] | None = None  # the scan's formula cells, read once and only when needed
    if len(ops) > 2 * count + 2:
        return f"{len(ops)} operations for a problem in {count} cell{'' if count == 1 else 's'}"
    for i, o in enumerate(ops, start=1):
        kind = str(o.get("op") or "")
        sheet = str(o.get("sheet") or "")
        if kind == "set_array_formula":
            ref = parse_ref(str(o.get("range") or o.get("cell") or ""))
            if not ref or ref["sheet"] or ref["whole_column"] or ref["whole_row"]:
                return f"operation {i} ({kind}) names no range"
            if ref["cells"] > MAX_ARRAY_CELLS or not all(own(sheet.lower(), r, c) for r, c in _range_cells(ref, MAX_ARRAY_CELLS)):
                return f"operation {i} writes an array formula over {sheet}!{ref['ref']}, which reaches outside the problem's own cells"
        elif kind in ("set_formula", "set_value", "clear_cell"):
            ref = parse_ref(str(o.get("cell") or ""))
            if not ref or ref["sheet"] or ref["whole_column"] or ref["whole_row"]:
                return f"operation {i} ({kind}) names no cell"
            if ref["cells"] != 1:
                return f"operation {i} ({kind}) changes the range {sheet}!{ref['ref']}: one cell at a time"
            if own(sheet.lower(), ref["r1"], ref["c1"]):
                continue
            if kind == "set_value" and "before" in o and o["before"] is None:
                continue  # a label in an empty cell
            before, after = o.get("before"), o.get("after")
            place = f"{sheet}!{num_to_col(ref['c1'])}{ref['r1']}"
            if kind == "set_formula" and count and isinstance(before, str) and before.startswith("="):
                # another cell that holds a formula now: with care (19 Nov 2026, see above)
                if formula_cells is not None:
                    if known is None:
                        known = {(str(sh or "").lower(), str(c or "").replace("$", "").upper()) for sh, c in formula_cells}
                    if (sheet.lower(), f"{num_to_col(ref['c1'])}{ref['r1']}") not in known:
                        return f"operation {i} ({kind}) changes {place}, which is not one of the problem's own cells and holds a typed-in text, not a formula"
                if _VOLATILE_RE.search(before) or (isinstance(after, str) and _VOLATILE_RE.search(after)):
                    return f"operation {i} ({kind}) changes {place}, which is not one of the problem's own cells, and its result changes by itself (NOW, TODAY, RAND): the numbers check cannot compare it"
                if not _reads_own(before, sheet, rects):
                    return f"operation {i} ({kind}) changes {place}, which is not one of the problem's own cells and does not read one of them"
                others.add((sheet.lower(), ref["r1"], ref["c1"]))
                if len(others) > MAX_OTHER_CELLS:
                    return f"operation {i} ({kind}) changes a formula cell outside the problem's own cells, and that makes {len(others)} of them: at most {MAX_OTHER_CELLS}"
                continue
            return (f"operation {i} ({kind}) changes {place}, which is not one of the problem's own cells"
                    + (" and is not empty" if kind == "set_value" else " and holds no formula now" if kind == "set_formula" else ""))
        else:
            return f"operation {i} ({kind or 'no op'}) is not a change to one of the problem's own cells"
    return None


def narrow(ops: Iterable[dict[str, Any]] | None, cells: Iterable[dict[str, str]] | None, formula_cells: Iterable[tuple[Any, Any]] | None = None) -> bool:
    """True  when the  assistant's fix may be written by itself (see narrow_problem for the rule and the
    inputs). False: outcome "too_wide", nothing is written."""
    return narrow_problem(ops, cells, formula_cells) is None


# ---- step 4i: the judge --------------------------------------------------------------------------
def _number(n: Any) -> bool:
    return isinstance(n, int) and not isinstance(n, bool)


def judge(blocking_before: dict[str, int | None], blocking_after: dict[str, int | None] | None, rule_id: str, changed: int | None) -> str:
    """Whether  a fix  that was written is kept (step 4i).

    In: `blocking_before` and `blocking_after` = sites_by_rule(s.readiness()["blocking"]) before the
         fix and after the scan that follows it (`blocking_after` None: that scan did not run);
         `rule_id` = the problem  the fix was for; `changed` = how many computed values the  numbers
         check found changed between the version  before the fix and the version with  it, None when
         the check did not run.
    Out: "not_checked"     the numbers check (or  the scan) did not run: never a  pass;
         "numbers_changed" a value changed;
         "not_cleared"     a rule is blocking now that was not before, or any blocking rule has more
                           sites than before, or this rule is still there and no smaller;
         "fixed"           this rule is gone, or has fewer sites.
    Only "fixed" keeps the version; every other answer means: put the snapshot back."""
    if changed is None:
        return "not_checked"
    if changed > 0:
        return "numbers_changed"  # the scan is not run then: `blocking_after` may be None
    if blocking_after is None:
        return "not_checked"
    for rule, now in blocking_after.items():
        if rule not in blocking_before:
            return "not_cleared"  # another rule turned blocking, or  one fixed earlier came back
        was = blocking_before[rule]
        if _number(now) and _number(was) and now > was:
            return "not_cleared"
    if rule_id not in blocking_after:
        return "fixed"
    was, now = blocking_before.get(rule_id), blocking_after[rule_id]
    return "fixed" if _number(was) and _number(now) and now < was else "not_cleared"


# ---- the job's message, one plain sentence at a time (DESIGN 5.4) ---------------------------------
WRITE_MESSAGE = "Writing the assistant's fix."
NUMBERS_MESSAGE = "Checking that no result changed."
CHECK_MESSAGE = "Checking the workbook again."


def round_message(pass_no: int, changes: int) -> str:
    """Stage  "pass": the  words Custom's step line shows today."""
    return f"Round {pass_no}: writing {changes} changes through Excel…"


def assistant_message(problem_no: int, problem_count: int) -> str:
    """Stage  "assistant"."""
    return f"The assistant is working on what is left ({problem_no} of {problem_count})."


def undo_message(outcome: str | None = "numbers_changed") -> str:
    """Stage  "undo". The  design's sentence is about a changed result; a fix taken back for another
    reason (not cleared, not checked) must not say a result changed."""
    return "A result changed: taking the fix back." if outcome == "numbers_changed" else "Taking the fix back."


def done_lines(applied: int, assistant_fixed: int = 0) -> list[str]:
    """State  "done": "{n}  changes applied." or "Nothing blocked the file." (today's words), and when the
    assistant fixed something, "The assistant fixed {m} more: see Advanced › Versions." (Show me: the
    place `versions`). When Prep wrote nothing  and the assistant did, only the assistant's sentence,
    without "more": "Nothing blocked  the file." would not be true."""
    if assistant_fixed and not applied:
        return [f"The assistant fixed {assistant_fixed}: see Advanced › Versions."]
    first = f"{applied} changes applied." if applied else "Nothing blocked the file."
    return [first, f"The assistant fixed {assistant_fixed} more: see Advanced › Versions."] if assistant_fixed else [first]


# ----  step 7: the question for the person  ----------------------------------------------------------
def _plain_name(problem: dict[str, Any]) -> str:
    """The  rule in  plain words (rules/plain_names.yaml); never its code."""
    return str(problem.get("rule_name") or rule_name(problem.get("rule_id")) or "a problem the app has no name for")


def ask_lines(problems: list[dict[str, Any]], button: str = "Start") -> list[str]:
    """`ask.lines`:  the two  sentences of the question (Basic draws one per line).

    In: `problems` = problems_left of the FINAL blocking list (the first one is asked about);
         `button` = the name of the button on the  screen that started the job: "Start" on Basic,
         "Do everything" on Custom.
    Out:  [sentence 1, sentence 2]; [] when  nothing is left. Sentence 1 has three  forms: one problem in
         one cell, one problem in several cells, several  problems. No rule code in either."""
    if not problems:
        return []
    p = problems[0]
    name, sheet, cell, n = _plain_name(p), str(p.get("sheet") or ""), str(p.get("cell") or ""), int(p.get("cells") or 0)
    at = f", on sheet {sheet}, cell {cell}" if sheet and cell else f", on sheet {sheet}" if sheet else ""
    if len(problems) > 1:
        first = f"{len(problems)} things are left that the app could not fix. The first: {name}{at}."
    elif n > 1 and sheet and cell:
        first = f"One thing is left that the app could not fix: {name}, in {n:,} cells; the first is {cell} on sheet {sheet}."
    else:
        first = f"One thing is left that the app could not fix: {name}{at}."
    return [first, f"Fix it by hand, then press {button} again."]


# `ask.why`: never  a line on Basic; Custom shows it under "View details".
_WHY = {
    "no_fix": "The assistant had no safe fix for it.",
    "too_wide": "The assistant's fix was too wide to apply without you.",
    "numbers_changed": "The assistant's fix changed a result, so it was taken back.",
    "not_cleared": "The assistant's fix did not solve it, so it was taken back.",
    "not_checked": "The assistant's fix could not be checked, so it was taken back.",
    "not_written": "Excel did not accept the assistant's fix.",
    "not_reached": "The assistant could not be reached.",
    "not_tried": "The assistant was not asked.",
    # not  in the  design: a problem that got smaller twice and is still there
    "fixed": "The assistant fixed part of it.",
}


def why_line(outcome: str | None) -> str:
    """`ask.why`  for the  last outcome of the problem asked about ("not_tried" when the assistant never saw it)."""
    return _WHY.get(str(outcome or "not_tried"), _WHY["not_tried"])


def show_me_for(problem: dict[str, Any]) -> dict[str, Any]:
    """`ask.show_me`:  where "Show  me" leads for the problem asked about: {"id", "rule_id", "sheet",
    "cell"}; id "prep-action" when its route is "prep" or "prep_skipped", "assistant-ask" otherwise
    (the screen's ShowMeId; with `open` the pointer goes there itself and  ends on the cell's line)."""
    return {"id": "prep-action" if problem.get("route") in PREP_ROUTES else "assistant-ask", "rule_id": problem.get("rule_id"), "sheet": problem.get("sheet") or "", "cell": problem.get("cell") or ""}


# What the assistant did, for the note  in Send feedback.
_TRIED = {
    "no_fix": "it proposed no change",
    "too_wide": "its change was too wide to apply without a person",
    "numbers_changed": "its change altered a result and was undone",
    "not_cleared": "its change did not solve the problem and was undone",
    "not_checked": "its change could not be checked and was undone",
    "not_written": "Excel refused its change",
    "fixed": "it fixed part of the problem",  # not  in the  design: see _WHY
}


def feedback_note(problem: dict[str, Any], button: str = "Start") -> str:
    """The  note written  into Send feedback when the person is asked (`ask.feedback.note`).

    In: the problem asked about (`rule_name` or `rule_id`, `tries`, `outcome`); `button` as in ask_lines.
    Out: plain words only: "Start could not fix this by  itself: {rule name}. The assistant tried
         {once | twice}: {what happened}.", "...  The assistant could not be reached."  or "... The
         assistant was not asked." No  rule id, no version,
         no sheet or cell name: the person reads  this note, and the bug report the  server writes from
         the incident ref already carries the rule id."""
    head = f"{button} could not fix this by itself: {_plain_name(problem)}."
    if problem.get("outcome") == "not_reached":
        return f"{head} {_WHY['not_reached']}"  # 17 Nov 2026: it never answered, so it did not "try"
    did = _TRIED.get(str(problem.get("outcome") or ""))
    if not did:
        return f"{head} The assistant was not asked."
    return f"{head} The assistant tried {'twice' if int(problem.get('tries') or 0) >= 2 else 'once'}: {did}."


def build_ask(problems: list[dict[str, Any]], plan: Iterable[dict[str, Any]] | None, ref: str | None = None, button: str = "Start") -> dict[str, Any] | None:
    """The  job's `ask`  (state "ask"), whole.

    In: `problems` = problems_left of the FINAL blocking list, each carrying the `tries` and `outcome`
         the run gave it ("not_tried" when the assistant never saw it); `plan` = the final  plan;
         `ref` = the incident ref of `_report_problem(None, "Assistant", s, rule_id=..., what="not  fixed
         by the assistant")`, None when no  incident was written; `button` as in ask_lines.
    Out:  {"lines", "why", "stops_start", "show_me", "feedback": {"ref", "note"}}, or None when nothing
         is  left (state "done")."""
    if not problems:
        return None
    first = problems[0]
    return {
        "lines": ask_lines(problems, button),
        "why": why_line(first.get("outcome")),
        "stops_start": stops_start(plan),
        "show_me": show_me_for(first),
        "feedback": {"ref": ref, "note": feedback_note(first, button)},
    }


def no_key_ask(button: str = "Start") -> dict[str, Any]:
    """The  job's `ask`  for state "no_key": the two sentences, and Show me to the AI key box. Nothing to
    send as feedback: the app is not at fault."""
    return {
        "lines": ["Something is left that needs the assistant, and no AI key is saved.", f"Paste an AI key under Advanced › Assistant, then press {button} again."],
        "why": why_line("not_tried"),
        "stops_start": True,
        "show_me": {"id": "assistant-key", "rule_id": None, "sheet": "", "cell": ""},
        "feedback": None,
    }
