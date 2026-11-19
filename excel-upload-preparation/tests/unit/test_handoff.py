"""17 Nov 2026: never stop for a person when the app can fix it (app/handoff.py, the pure parts).

Start's Prep step goes on while a round still changes something, hands what is left to the assistant,
writes its fix only when it is narrow, keeps it only when no number changed and the problem got
smaller, and only then asks the person, in plain words. Every rule of that is a plain function, tested
here without Excel, without the server and without a model.
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
import pytest

from app import handoff
from app.handoff import (
    MAX_OTHER_CELLS, MAX_PASSES, MAX_PROBLEMS, MAX_QUESTION_CELLS, MAX_TRIES, ask_lines, assistant_message, build_ask, cells_of, changed_sample, done_lines,
    feedback_note, judge, left_key, memo_key, narrow, narrow_problem, no_key_ask, pass_ops, problems_left, question_for, round_message,
    rounds_go_on, rounds_stop, run_rounds, show_me_for, sites_by_rule, stops_start, undo_message, why_line,
)
from app.plain_codes import rule_name

SHEET = "ロールフォワード"  # the Tokyo sheet of the staircase
STAIRS = ["V24", "W25", "X26", "Y27", "Z28", "AA29"]


def _op(cell: str, action: str = "title_lone_text_cells", rule: str = "STR-008", op: str = "set_value", sheet: str = SHEET, **kw) -> dict:
    return {"op": op, "action_id": action, "rule_id": rule, "sheet": sheet, "cell": cell, "before": None, "after": "#x", **kw}


def _action(ops: list[dict], id: str = "title_lone_text_cells", level: str = "blocking", default_on: bool = True) -> dict:
    return {"id": id, "title": id, "rule_ids": ["STR-008"], "level": level, "default_on": default_on, "operations": ops, "count": len(ops), "skipped": []}


class FakePrep:
    """A workbook as the rounds see it: `plan()` is what Prep would write now, `write` is one round."""

    def __init__(self, plan_of, written=None):
        self.done: list[str] = []  # the cells written so far
        self.rounds: list[list[str]] = []  # the cells each round was given
        self._plan_of = plan_of
        self._written = written or (lambda ops, n: (len(ops), 0))

    def plan(self) -> list[dict]:
        return self._plan_of(self)

    def write(self, ops: list[dict], pass_no: int) -> tuple[int, int]:
        self.rounds.append([o["cell"] for o in ops])
        applied, refused = self._written(ops, pass_no)
        self.done += [o["cell"] for o in ops[:applied]]
        return applied, refused


# ---- the limits ---------------------------------------------------------------------------------
def test_the_limits_are_the_designs():
    # 19 Nov 2026 (Joseph): 10 rounds, down from 20; and at most 5 formula cells outside the problem's own
    assert (MAX_PASSES, MAX_PROBLEMS, MAX_TRIES, MAX_OTHER_CELLS) == (10, 5, 2, 5)


# ---- the rounds ---------------------------------------------------------------------------------
def test_a_round_is_every_operation_of_the_blocking_actions_that_are_on_by_default():
    plan = [
        _action([_op("A1"), _op("A2")]),
        _action([_op("B1")], id="create_grid_titles", level="optional"),  # optional: not Start's
        _action([_op("C1")], id="unprotect_sheets", default_on=False),  # off by default
        {**_action([]), "id": "empty"},  # nothing to write
        _action([_op("D1", action="fix_broken_refs", rule="REF-001", op="set_formula")], id="fix_broken_refs"),
    ]
    assert [o["cell"] for o in pass_ops(plan)] == ["A1", "A2", "D1"]
    assert left_key(plan) == {
        f"title_lone_text_cells | STR-008 | set_value | {SHEET} | A1",
        f"title_lone_text_cells | STR-008 | set_value | {SHEET} | A2",
        f"fix_broken_refs | REF-001 | set_formula | {SHEET} | D1",
    }
    assert stops_start(plan) is True
    assert pass_ops([]) == [] and pass_ops(None) == [] and left_key(None) == frozenset() and stops_start([]) is False


def test_the_key_tells_apart_operations_with_no_cell():
    rows = [{"op": "insert_row", "action_id": "separate_merged_grids", "rule_id": "STR-001", "sheet": "S", "row": r} for r in (7, 8)]
    array = {"op": "set_array_formula", "action_id": "offset_ranges", "rule_id": "FRM-014", "sheet": "S", "range": "C7:C9", "cell": "C7:C9"}
    link = {"op": "break_external_link", "action_id": "freeze_external_links", "rule_id": "REF-002", "sheet": "S", "file_name": "a.xlsx", "target": "t"}
    key = left_key([_action(rows + [array, link])])
    assert key == {
        "separate_merged_grids | STR-001 | insert_row | S | row 7",
        "separate_merged_grids | STR-001 | insert_row | S | row 8",
        "offset_ranges | FRM-014 | set_array_formula | S | C7:C9",
        "freeze_external_links | REF-002 | break_external_link | S | a.xlsx>t",
    }


def test_a_rule_that_shows_one_new_cell_per_round_runs_six_rounds_and_ends_clean():
    # the staircase as it was before the lone-text rule was fixed in one pass: each round uncovers the next step
    prep = FakePrep(lambda p: [_action([_op(STAIRS[len(p.done)])])] if len(p.done) < len(STAIRS) else [_action([])])
    out = run_rounds(prep.plan, prep.write)
    assert out == {"passes": 6, "applied": 6, "refused": 0, "why": "clean", "left": 0}
    assert prep.rounds == [[c] for c in STAIRS]


def test_nothing_to_write_runs_no_round():
    prep = FakePrep(lambda p: [_action([])])
    assert run_rounds(prep.plan, prep.write) == {"passes": 0, "applied": 0, "refused": 0, "why": "clean", "left": 0}
    assert prep.rounds == []


def test_the_same_operations_coming_back_stop_after_one_round():
    # Excel says it wrote them, and the scan reports the same cells again: a second round would change nothing
    prep = FakePrep(lambda p: [_action([_op("W25"), _op("X26")])])
    out = run_rounds(prep.plan, prep.write)
    assert out == {"passes": 1, "applied": 2, "refused": 0, "why": "same_again", "left": 2}
    assert len(prep.rounds) == 1


def test_a_round_that_wrote_nothing_stops():
    # every change refused; even a plan that looks different afterwards is no reason for another round
    prep = FakePrep(lambda p: [_action([_op(f"A{len(p.rounds) + 1}")])], written=lambda ops, n: (0, len(ops)))
    out = run_rounds(prep.plan, prep.write)
    assert out == {"passes": 1, "applied": 0, "refused": 1, "why": "nothing_written", "left": 1}


def test_a_refused_change_alone_does_not_stop_the_rounds():
    # C1 is refused every round; the rounds go on while the others still get written, and C1 is counted
    def plan(p: FakePrep) -> list[dict]:
        waiting = [c for c in ("A1", "B1") if c not in p.done][:1]
        return [_action([_op(c) for c in waiting] + [_op("C1")])]

    def written(ops: list[dict], n: int) -> tuple[int, int]:
        return len(ops) - 1, 1  # everything but C1 (the last one)

    prep = FakePrep(plan, written)
    out = run_rounds(prep.plan, prep.write)
    assert prep.rounds == [["A1", "C1"], ["B1", "C1"], ["C1"]]
    # round 1 and 2 wrote something and left a different plan: go on; round 3 wrote nothing: stop
    assert out == {"passes": 3, "applied": 2, "refused": 3, "why": "nothing_written", "left": 1}
    assert rounds_go_on(1, 1, left_key(plan(FakePrep(plan))), frozenset({"x"})) is True


def test_an_endless_repair_stops_at_max_passes():
    # a row insert that moves its own row every round: always written, always a new plan
    prep = FakePrep(lambda p: [_action([{"op": "insert_row", "action_id": "separate_merged_grids", "rule_id": "STR-001", "sheet": "S", "row": 7 + len(p.rounds), "cell": None}])])
    out = run_rounds(prep.plan, prep.write)
    assert out["passes"] == MAX_PASSES == len(prep.rounds)
    assert out["why"] == "max_passes" and out["applied"] == MAX_PASSES and out["left"] == 1


def test_why_the_rounds_stop_in_the_designs_order():
    a, b = frozenset({"a"}), frozenset({"b"})
    assert rounds_stop(1, 3, a, frozenset()) == "clean"
    assert rounds_stop(1, 0, a, frozenset()) == "clean"  # nothing left is clean, whatever was written
    assert rounds_stop(1, 0, a, b) == "nothing_written"
    assert rounds_stop(1, 2, a, a) == "same_again"
    assert rounds_stop(MAX_PASSES, 2, a, b) == "max_passes"
    assert rounds_stop(MAX_PASSES - 1, 1, a, b) is None  # progress, however small
    assert rounds_go_on(2, 1, a, b) is True and rounds_go_on(2, 1, a, a) is False


def test_a_cancel_inside_a_round_is_not_caught():
    class Cancelled(BaseException):
        pass

    def write(ops, n):
        raise Cancelled()

    with pytest.raises(Cancelled):
        run_rounds(lambda: [_action([_op("A1")])], write)


# ---- what is left -------------------------------------------------------------------------------
def _finding(rule: str, cells: list[str], sheet: str = SHEET, **observed) -> dict:
    sites = [{"sheet": sheet, "cell": c, "text": "x"} for c in cells]
    return {"rule_id": rule, "rule_name": rule_name(rule), "status": "ERROR", "location": {"sheet": sheet, "cell": cells[0]} if cells else {}, "observed": {"sites": sites, "sites_count": len(sites), **observed}}


def _blocking(rule: str, fix: str, sites: int | None = 1, sheet: str = SHEET, cell: str = "W25") -> dict:
    return {"rule_id": rule, "rule_name": rule_name(rule), "status": "ERROR", "message": "m", "location": {"sheet": sheet, "cell": cell}, "fix": fix, "action_id": None, "sites": sites}


def test_the_cells_of_a_finding():
    f = _finding("STR-008", ["V24", "$W$25", "x26", "V24"])
    assert cells_of(f) == [{"sheet": SHEET, "cell": "V24"}, {"sheet": SHEET, "cell": "W25"}, {"sheet": SHEET, "cell": "X26"}]
    # the other lists readiness counts, a place written "Sheet!A1", a range, and entries that name no cell
    assert cells_of({"location": {}, "observed": {"problems": [{"sheet": "P", "cell": "B2", "issues": ["x"]}, "no cell here", "Data!C7:C9", {"grid": "G"}]}}) == [
        {"sheet": "P", "cell": "B2"}, {"sheet": "Data", "cell": "C7:C9"}]
    assert cells_of({"location": {"sheet": "S", "cell": "A1"}, "observed": {"sites": [{"sheet": "S", "cell": "B2"}]}}) == [{"sheet": "S", "cell": "A1"}, {"sheet": "S", "cell": "B2"}]
    assert cells_of({"location": {"sheet": "S", "ref": "B2:D9"}, "observed": None}) == [{"sheet": "S", "cell": "B2:D9"}]
    assert cells_of({"location": {}, "observed": {"sites": []}}) == [] and cells_of(None) == [] and cells_of({"location": {"sheet": "S", "cell": "A:A"}}) == []


def test_what_is_left_leaves_out_the_outside_files_and_puts_preps_own_first():
    report = {"findings": [_finding("FRM-007", ["B2"], sheet="Calc"), _finding("STR-008", STAIRS[:3]), _finding("REF-001", ["C3", "C4"], sheet="Calc")]}
    blocking = [
        _blocking("INP-010", "outside_inputs", 6),
        _blocking("FRM-007", "assistant", 1, "Calc", "B2"),
        _blocking("STR-008", "prep", 3, SHEET, "V24"),
        _blocking("REF-001", "prep_skipped", 2, "Calc", "C3"),
    ]
    left = problems_left(blocking, report)
    assert [(p["rule_id"], p["route"]) for p in left] == [("STR-008", "prep"), ("REF-001", "prep_skipped"), ("FRM-007", "assistant")]
    assert left[0] == {"rule_id": "STR-008", "rule_name": "a formula reads a lone text cell", "sheet": SHEET, "cell": "V24", "cells": 3, "route": "prep", "tries": 0, "outcome": "not_tried"}
    assert problems_left([], report) == [] and problems_left(None) == []
    assert sites_by_rule(blocking) == {"INP-010": 6, "FRM-007": 1, "STR-008": 3, "REF-001": 2}


def test_the_count_of_cells_is_the_findings_full_count_not_the_list_it_shows():
    f = _finding("STR-008", STAIRS[:2])
    f["observed"]["sites_count"] = 333  # the list shown is cut at 80
    assert problems_left([_blocking("STR-008", "prep", 2, SHEET, "V24")], {"findings": [f]})[0]["cells"] == 333
    # no report at hand: the place readiness names, and its count
    assert problems_left([_blocking("STR-008", "prep", 4, SHEET, "V24")])[0] | {"rule_name": ""} == {
        "rule_id": "STR-008", "rule_name": "", "sheet": SHEET, "cell": "V24", "cells": 4, "route": "prep", "tries": 0, "outcome": "not_tried"}
    # a finding that names no cell at all
    assert problems_left([{"rule_id": "FILE-002", "fix": "assistant", "location": {}, "sites": None}])[0] | {"rule_name": ""} == {
        "rule_id": "FILE-002", "rule_name": "", "sheet": "", "cell": "", "cells": 0, "route": "assistant", "tries": 0, "outcome": "not_tried"}


# ---- the guard: narrow ---------------------------------------------------------------------------
OWN = [{"sheet": "Calc", "cell": "B2"}, {"sheet": "Calc", "cell": "C7:C9"}]  # 4 cells: at most 10 operations


def _aop(op: str, cell: str, before="=A1", sheet: str = "Calc", **kw) -> dict:
    """An operation as prep.validate_proposal returns it."""
    return {"action_id": "assistant", "rule_id": "CHAT", "sheet": sheet, "note": "n", "op": op, "cell": cell, "before": before, "after": "x", **kw}


@pytest.mark.parametrize("ops", [
    [_aop("set_formula", "B2")],  # a formula on one of the finding's own cells
    [_aop("set_formula", "C8")],  # ... also inside a range the finding names
    [_aop("set_formula", "b2", sheet="CALC")],  # sheets and cells are told apart by name, not by capitals
    [_aop("set_array_formula", "C7:C9", range="C7:C9")],  # an array formula whose whole range is the finding's own
    [_aop("set_array_formula", "C7:C8", range="C7:C8")],
    [_aop("set_value", "B2", before="text")],  # a value on ONE own cell
    [_aop("clear_cell", "C9", before="text")],  # a clear on ONE own cell
    [_aop("set_value", "A1", before=None, sheet="Other")],  # a label in a cell that is empty now, anywhere
    [_aop("set_formula", "B2"), _aop("set_value", "A2", before=None), _aop("clear_cell", "C7")],
    [_aop("set_formula", "B2")] * 10,  # 2 x 4 cells + 2
    # 19 Nov 2026 (Joseph: "allow the fix to find other cells cautiously"): a cell that is not the finding's own but holds a formula now
    [_aop("set_formula", "D4")],
    [_aop("set_formula", "B2", sheet="Other")],  # ... also on another sheet (the formula that reads the problem cell)
    [_aop("set_formula", "B2"), _aop("set_formula", "D4"), _aop("set_formula", "D5", before="=SUM(A1:A9)")],
    [_aop("set_formula", f"D{r}") for r in range(4, 9)],  # five other formula cells: the most
    [_aop("set_formula", "D4")] * 3,  # the same other cell three times is one cell
], ids=["formula-own", "formula-own-in-range", "capitals", "array-own", "array-part", "value-own", "clear-own", "label-empty", "mixed", "at-the-limit",
        "formula-elsewhere", "formula-other-sheet", "own-and-others", "five-others", "one-other-thrice"])
def test_narrow_lets_through_each_allowed_shape(ops):
    assert narrow_problem(ops, OWN) is None
    assert narrow(ops, OWN) is True


@pytest.mark.parametrize("ops, why", [
    ([_aop("set_formula", "B2:B5", before="4 cells")], "range"),  # a range
    ([_aop("set_value", "A1:A3", before="3 cells")], "range"),
    ([_aop("clear_cell", "C7:C9", before="3 cells")], "range"),  # even over own cells: one cell at a time
    ([{"action_id": "assistant", "rule_id": "CHAT", "sheet": "Calc", "op": "insert_row", "row": 7}], "insert_row"),  # an insert
    ([{"action_id": "assistant", "rule_id": "CHAT", "sheet": "Calc", "op": "insert_column", "column": "C"}], "insert_column"),
    ([{"action_id": "assistant", "rule_id": "CHAT", "sheet": "Calc", "op": "rename_sheet", "before": "Calc", "after": "Calc2"}], "rename_sheet"),  # a rename
    ([{"action_id": "assistant", "rule_id": "CHAT", "sheet": "Calc", "op": "set_sheet_visibility", "after": False}], "set_sheet_visibility"),
    ([{"action_id": "assistant", "rule_id": "CHAT", "sheet": "Calc", "op": "unprotect_sheet"}], "unprotect_sheet"),
    ([_aop("set_value", "D4", before=12)], "not one of the problem's own cells"),  # a typed-in value elsewhere: never
    ([_aop("clear_cell", "D4", before="text")], "not one of the problem's own cells"),
    ([_aop("clear_cell", "D4")], "not one of the problem's own cells"),  # another formula cell may be rewritten, not emptied
    ([_aop("set_value", "D4")], "not one of the problem's own cells"),  # ... nor turned into a value
    ([_aop("set_formula", "D4", before=12)], "holds no formula now"),  # a typed-in value is not replaced by a formula either
    ([_aop("set_formula", "D4", before="text")], "holds no formula now"),
    ([_aop("set_formula", "D4", before="part of array formula D4:D6")], "holds no formula now"),
    ([_aop("set_formula", "D4", before=None)], "not one of the problem's own cells"),  # a formula into an empty cell is no label
    ([_aop("set_formula", f"D{r}") for r in range(4, 10)], "at most 5"),  # a sixth formula cell outside the problem's own
    ([_aop("set_value", "D4", before="part of array formula D4:D6")], "not one of the problem's own cells"),  # not empty in Excel
    ([_aop("set_array_formula", "C7:C10", range="C7:C10")], "outside the problem's own cells"),  # the array reaches one cell further
    ([_aop("set_array_formula", "A1:XFD1048576", range="A1:XFD1048576")], "outside the problem's own cells"),
    ([_aop("set_formula", "B2")] * 11, "11 operations"),  # too many operations
    ([_aop("set_formula", "B2"), {"action_id": "assistant", "sheet": "Calc", "op": "insert_row", "row": 2}], "insert_row"),  # one wide operation is enough
    ([], "no change"),
], ids=["formula-range", "value-range", "clear-range", "insert-row", "insert-column", "rename", "hide", "unprotect", "value-elsewhere",
        "clear-elsewhere", "clear-formula-elsewhere", "value-over-formula-elsewhere", "formula-over-value", "formula-over-text", "formula-over-array-member",
        "formula-empty", "six-others", "array-member", "array-wider", "array-huge", "too-many", "one-wide", "nothing"])
def test_narrow_refuses_anything_wider(ops, why):
    assert narrow(ops, OWN) is False
    assert why in narrow_problem(ops, OWN)


def test_narrow_with_no_cells_allows_only_two_labels():
    label = _aop("set_value", "A1", before=None)
    assert narrow([label, label], []) is True
    assert narrow([label, label, label], []) is False
    assert narrow([_aop("set_formula", "B2")], []) is False  # a problem that names no cell: no formula cell is "another" one
    # a missing `before` is not "empty now": only what the validator read as empty counts
    assert narrow([{"op": "set_value", "sheet": "Calc", "cell": "A1", "after": "x"}], []) is False


def test_the_real_pieces_fit_the_finding_the_plan_the_validator_and_the_guard(tmp_path: Path):
    # no fakes: the lone-text rule's finding, readiness, the Prep plan and the assistant's operations as
    # prep.validate_proposal returns them (openpyxl only, no Excel)
    from app.config import load_config
    from app.inventory import build_analysis
    from app.prep import plan_actions, validate_proposal
    from app.readiness import compute_readiness
    from app.rules_engine import RulesEngine
    from app.validators import run_rule

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Roll sheet"
    for cell, value in {"B2": "#Grid", "B3": "h1", "C3": "h2", "B4": 1, "C4": 2, "G9": "Lonely", "H12": '="x"&G9', "J20": "Other", "K22": '="y"&J20'}.items():
        ws[cell] = value
    p = tmp_path / "a.xlsx"
    wb.save(p)
    analysis = build_analysis(p, tmp_path / "w", "s")
    finding = run_rule(RulesEngine().get("STR-008"), analysis, load_config())
    report = {"findings": [finding]}
    plan = plan_actions(analysis, report)
    blocking = compute_readiness(report, plan, None, "ver-001")["blocking"]

    assert cells_of(finding) == [{"sheet": "Roll sheet", "cell": "G9"}, {"sheet": "Roll sheet", "cell": "J20"}]
    assert problems_left(blocking, report) == [
        {"rule_id": "STR-008", "rule_name": NAME, "sheet": "Roll sheet", "cell": "G9", "cells": 2, "route": "prep", "tries": 0, "outcome": "not_tried"}]
    assert left_key(plan) == {"title_lone_text_cells | STR-008 | set_value | Roll sheet | G8", "title_lone_text_cells | STR-008 | set_value | Roll sheet | J19"}
    assert sites_by_rule(blocking) == {"STR-008": 2}
    assert "Cells: 'Roll sheet'!G9, 'Roll sheet'!J20." in question_for(finding)

    def proposed(*operations: dict) -> list[dict]:
        ops, errors = validate_proposal({"summary": "s", "operations": list(operations)}, analysis)
        assert not errors, errors
        return ops

    own = cells_of(finding)
    # a label in the empty cell above (the validator read it as empty) and a change to the cell itself: narrow
    assert narrow(proposed({"op": "set_value", "sheet": "roll sheet", "cell": "$G$8", "value": "#Label"}, {"op": "set_value", "sheet": "Roll sheet", "cell": "G9", "value": "Alone"}), own)
    # 19 Nov 2026: the formula that reads it is another cell that holds a formula now, so it may be rewritten too
    # (whether the fix is kept is the numbers check's and the judge's word); a typed-in value may not
    assert narrow_problem(proposed({"op": "set_formula", "sheet": "Roll sheet", "cell": "H12", "formula": "=G9"}), own) is None
    assert "holds no formula now" in narrow_problem(proposed({"op": "set_formula", "sheet": "Roll sheet", "cell": "B4", "formula": "=1"}), own)
    # an insert moves everything under it
    assert "insert_row" in narrow_problem(proposed({"op": "insert_row", "sheet": "Roll sheet", "row": 9}), own)
    assert "range" in narrow_problem(proposed({"op": "clear_cell", "sheet": "Roll sheet", "cell": "G9:J20"}), own)


# ---- the judge -----------------------------------------------------------------------------------
def test_judge_each_outcome():
    before = {"STR-008": 4, "FRM-007": 2}
    assert judge(before, {"FRM-007": 2}, "STR-008", 0) == "fixed"  # the rule is gone
    assert judge(before, {"STR-008": 1, "FRM-007": 2}, "STR-008", 0) == "fixed"  # fewer sites: kept, the rest gets its second try
    assert judge(before, {"STR-008": 4, "FRM-007": 2}, "STR-008", 0) == "not_cleared"  # still there, no smaller
    assert judge(before, {"STR-008": 5, "FRM-007": 2}, "STR-008", 0) == "not_cleared"
    assert judge(before, {"FRM-007": 2}, "STR-008", 3) == "numbers_changed"  # a value changed, whatever the scan says
    assert judge(before, None, "STR-008", 3) == "numbers_changed"  # the scan is not even run then
    assert judge(before, {"FRM-007": 2}, "STR-008", None) == "not_checked"  # a check that did not run is never a pass
    assert judge(before, None, "STR-008", 0) == "not_checked"  # nor a scan that did not run
    assert judge(before, {}, "STR-008", 0) == "fixed"


def test_judge_another_rule_turning_blocking_is_not_cleared():
    # the fix cleared its own rule and broke something else
    assert judge({"STR-008": 4}, {"FRM-009": 1}, "STR-008", 0) == "not_cleared"
    assert judge({"STR-008": 4, "FRM-007": 2}, {"FRM-007": 3}, "STR-008", 0) == "not_cleared"  # another rule got worse


def test_judge_a_rule_already_fixed_coming_back_is_not_cleared():
    # REF-001 was fixed by an earlier try of this run (it is not in `before`); this fix brings it back
    before = {"STR-008": 4}
    assert "REF-001" not in before
    assert judge(before, {"STR-008": 2, "REF-001": 1}, "STR-008", 0) == "not_cleared"
    assert judge(before, {"REF-001": 1}, "STR-008", 0) == "not_cleared"


def test_judge_with_rules_that_count_no_sites():
    assert judge({"FILE-002": None}, {"FILE-002": None}, "FILE-002", 0) == "not_cleared"
    assert judge({"FILE-002": None}, {}, "FILE-002", 0) == "fixed"
    assert judge({"STR-008": 4, "FILE-002": None}, {"STR-008": 4, "FILE-002": None}, "STR-008", 0) == "not_cleared"
    assert judge({"STR-008": 4, "FILE-002": None}, {"STR-008": 3, "FILE-002": None}, "STR-008", 0) == "fixed"


# ---- the memo ------------------------------------------------------------------------------------
def test_memo_key_changes_when_a_cells_content_changes_and_not_otherwise():
    f = _finding("STR-008", ["V24", "W25"])
    content = {("V24"): "Cohort", ("W25"): "Group"}
    read = lambda sheet, row, col: content.get(f"{openpyxl.utils.get_column_letter(col)}{row}")  # noqa: E731
    key = memo_key(f, None, read)
    assert re.fullmatch(r"STR-008:[0-9a-f]{16}", key)
    assert memo_key(f, None, read) == key  # the same problem on an untouched workbook
    assert memo_key({**f, "message": "another sentence", "status": "ERROR"}, None, read) == key  # only the cells count
    assert memo_key({**f, "observed": {**f["observed"], "sites": list(reversed(f["observed"]["sites"]))}}, None, read) == key  # nor their order
    content["W25"] = "Group name"  # a hand edit of one of the cells
    edited = memo_key(f, None, read)
    assert edited != key
    content["W25"] = "Group"
    assert memo_key(f, None, read) == key
    assert memo_key(_finding("STR-008", ["V24"]), None, read) != key  # one cell fewer is another problem
    assert memo_key({**f, "rule_id": "FRM-007"}, None, read).split(":")[1] == key.split(":")[1]  # the rule id stands in front


def test_memo_key_reads_the_analysis_own_copy_of_the_workbook(tmp_path: Path):
    from app.inventory import build_analysis

    def analysis_of(text: str, name: str) -> dict:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Roll"
        ws["B2"] = text
        ws["C3"] = "=LEN(B2)"
        p = tmp_path / f"{name}.xlsx"
        wb.save(p)
        return build_analysis(p, tmp_path / name, name)

    f = _finding("STR-008", ["B2"], sheet="Roll")
    first = memo_key(f, analysis_of("Cohort", "a"))
    assert memo_key(f, analysis_of("Cohort", "b")) == first  # another copy, the same content
    assert memo_key(f, analysis_of("Cohort year", "c")) != first
    assert memo_key(f, None) != first  # no analysis: the cells count as empty, and nothing stops


def test_memo_key_never_stops_on_a_cell_that_cannot_be_read():
    def read(sheet, row, col):
        raise RuntimeError("no workbook")

    assert memo_key(_finding("STR-008", ["V24"]), None, read).startswith("STR-008:")


# ---- the question for the assistant --------------------------------------------------------------
def test_the_question_names_every_cell():
    cells = ["V24", "W25", "X26", "Y27", "Z28", "AA29", "AB30", "AC31", "AD32", "AE33", "AF34"]  # Tokyo: 11 cells
    f = _finding("STR-008", cells)
    q = question_for(f)
    assert q.startswith(f"[About finding STR-008 at {SHEET}!V24. If you propose a replacement formula")
    assert "] Propose a fix\nCells: " in q
    for c in cells:
        assert f"{SHEET}!{c}" in q
    line = q.split("\n")[1]
    assert line == "Cells: " + ", ".join(f"{SHEET}!{c}" for c in cells) + (". Change these cells; when the fix needs it, also at most 5 other cells that hold a formula now, or write a label into one empty cell: "
                                                                             "a typed-in value, a range or anything that moves cells is not written without a person.")
    assert len(q.split("\n")) == 2  # a first try says nothing about an earlier one


def test_the_question_quotes_a_sheet_name_the_way_excel_does():
    q = question_for(_finding("FRM-007", ["B2", "C7:C9"], sheet="Policy summary"))
    assert "Cells: 'Policy summary'!B2, 'Policy summary'!C7:C9." in q
    assert "'It''s'!A1" in question_for(_finding("FRM-007", ["A1"], sheet="It's"))
    # the chat reads these references itself (chat_context.CELL_REF_RE), sheet and all
    from app.chat_context import CELL_REF_RE

    found = [(m.group("sheet"), m.group("ref")) for m in CELL_REF_RE.finditer(q.split("\n")[1])]
    assert found == [("'Policy summary'", "B2"), ("'Policy summary'", "C7:C9")]
    assert [(m.group("sheet"), m.group("ref")) for m in CELL_REF_RE.finditer(question_for(_finding("STR-008", ["V24"])).split("\n")[1])] == [(SHEET, "V24")]


def test_the_question_names_two_hundred_cells_at_most_and_says_how_many_more():
    cells = [f"A{i}" for i in range(1, 251)]
    q = question_for(_finding("STR-008", cells, sheet="S"))
    assert "S!A200, and 50 more." in q and "S!A201" not in q
    assert MAX_QUESTION_CELLS == 200


def test_the_question_for_the_jobs_own_problem_entry_and_given_cells():
    problem = {"rule_id": "STR-008", "rule_name": "a formula reads a lone text cell", "sheet": "Roll", "cell": "W25", "cells": 3, "route": "prep", "tries": 1, "outcome": "no_fix"}
    assert question_for(problem).split("\n")[1].startswith("Cells: Roll!W25. ")
    left = [{"sheet": "Roll", "cell": "X26"}, {"sheet": "Roll", "cell": "Y27"}]  # what is left of it, on the second try
    assert "Cells: Roll!X26, Roll!Y27." in question_for(problem, left)
    assert "Cells:" not in question_for({"rule_id": "FILE-002", "location": {}})


@pytest.mark.parametrize("earlier, words", [
    ("no_fix", "Your first answer held no change the app could apply"),
    ("too_wide", "Your first proposal was too wide to write without a person (operation 2 (insert_row) is not a change to one of the problem's own cells)"),
    ("not_written", "Excel refused your first proposal"),
    ("numbers_changed", "Your first proposal changed a computed result"),
    ("not_cleared", "it did not make this problem smaller or it made another one worse"),
    ("not_checked", "Your first proposal could not be checked"),
    ("fixed", "now fix what is left of it"),
])
def test_a_second_try_says_in_one_sentence_why_the_first_failed(earlier, words):
    detail = "operation 2 (insert_row) is not a change to one of the problem's own cells" if earlier == "too_wide" else ""
    q = question_for(_finding("STR-008", ["V24"]), earlier=earlier, detail=detail)
    lines = q.split("\n")
    assert len(lines) == 3 and words in lines[2]
    assert lines[2].count(". ") == 0 and lines[2].endswith(".")  # one sentence


def test_the_changed_value_named_in_the_second_try():
    numbers = {"counts": {"good_to_other": 1}, "samples": {"good_to_error": [], "good_to_other": [{"original": "Calc!D4", "now": "Calc!D4", "was": "100", "is": "0"}]}}
    assert changed_sample(numbers) == "Calc!D4 showed 100 and would show 0"
    assert changed_sample({"samples": {"missing": [{"original": "Calc!D4", "now": "Calc!D4", "was": "100"}]}}) == "Calc!D4 (100) would no longer be a formula"
    assert changed_sample({}) == "" and changed_sample(None) == ""
    q = question_for(_finding("STR-008", ["V24"]), earlier="numbers_changed", detail=changed_sample(numbers))
    assert q.endswith("Your first proposal changed a computed result (Calc!D4 showed 100 and would show 0), so it was taken back: propose a change that leaves every result as it is.")


def test_the_question_starts_with_the_words_of_the_assistant_tab():
    # app/web/server.py `_chat_turn` writes the same sentence before "Propose a fix" on the Assistant tab;
    # the two must not drift apart (read as text: importing the server starts the whole app)
    source = (Path(__file__).resolve().parents[2] / "app" / "web" / "server.py").read_text(encoding="utf-8")
    for part in ('f"[About finding {focus[\'rule_id\']}{where}. If you propose a replacement formula, it must behave like the original on "',
                 '"blank cells, text and errors (e.g. wrap with IF(x=\\"\\",0,...) when the original tolerated blanks); every new error "',
                 '"after recalculation counts against the fix.] "'):
        assert part in source, part
    assert question_for({"rule_id": "FRM-009", "location": {"sheet": "Calc", "cell": "B2"}}, cells=[]).startswith(
        '[About finding FRM-009 at Calc!B2. If you propose a replacement formula, it must behave like the original on blank cells, text and errors '
        '(e.g. wrap with IF(x="",0,...) when the original tolerated blanks); every new error after recalculation counts against the fix.] Propose a fix')


# ---- the sentences on screen (DESIGN 5.4, word for word) -----------------------------------------
NAME = "a formula reads a lone text cell"  # STR-008 in rules/plain_names.yaml


def _problem(cells: int = 1, route: str = "prep", outcome: str = "no_fix", tries: int = 1, rule: str = "STR-008", sheet: str = SHEET, cell: str = "W25") -> dict:
    return {"rule_id": rule, "rule_name": rule_name(rule), "sheet": sheet, "cell": cell, "cells": cells, "route": route, "tries": tries, "outcome": outcome}


def test_the_two_question_sentences_in_their_three_forms():
    assert rule_name("STR-008") == NAME
    # one problem, one cell
    assert ask_lines([_problem()]) == [f"One thing is left that the app could not fix: {NAME}, on sheet {SHEET}, cell W25.", "Fix it by hand, then press Start again."]
    # one problem, several cells
    assert ask_lines([_problem(cells=11)]) == [f"One thing is left that the app could not fix: {NAME}, in 11 cells; the first is W25 on sheet {SHEET}.", "Fix it by hand, then press Start again."]
    # several problems
    three = [_problem(cells=11), _problem(rule="FRM-007", sheet="Calc", cell="B2"), _problem(rule="REF-001")]
    assert ask_lines(three) == [f"3 things are left that the app could not fix. The first: {NAME}, on sheet {SHEET}, cell W25.", "Fix it by hand, then press Start again."]
    # Custom's button has another name
    assert ask_lines([_problem()], "Do everything")[1] == "Fix it by hand, then press Do everything again."
    assert ask_lines(three, "Do everything")[1] == "Fix it by hand, then press Do everything again."
    assert ask_lines([]) == []


def test_the_question_never_shows_a_rule_code():
    lines = ask_lines([_problem(), _problem(rule="FRM-007")]) + ask_lines([_problem(cells=1234)]) + ask_lines([{"rule_id": "ZZZ-999", "sheet": "", "cell": "", "cells": 0}])
    for line in lines:
        assert not re.search(r"\b[A-Z]{2,7}-\d{3}\b", line), line
    assert ask_lines([_problem(cells=1234)])[0] == f"One thing is left that the app could not fix: {NAME}, in 1,234 cells; the first is W25 on sheet {SHEET}."
    # a problem that names no cell, and a rule with no plain name: still a sentence, still no code
    assert ask_lines([{"rule_id": "ZZZ-999", "sheet": "", "cell": "", "cells": 0}])[0] == "One thing is left that the app could not fix: a problem the app has no name for."
    assert ask_lines([{"rule_id": "STR-008", "sheet": "Roll", "cell": "", "cells": 0}])[0] == f"One thing is left that the app could not fix: {NAME}, on sheet Roll."


def test_why_under_view_details():
    assert why_line("no_fix") == "The assistant had no safe fix for it."
    assert why_line("too_wide") == "The assistant's fix was too wide to apply without you."
    assert why_line("numbers_changed") == "The assistant's fix changed a result, so it was taken back."
    assert why_line("not_cleared") == "The assistant's fix did not solve it, so it was taken back."
    assert why_line("not_checked") == "The assistant's fix could not be checked, so it was taken back."
    assert why_line("not_written") == "Excel did not accept the assistant's fix."
    assert why_line("not_tried") == "The assistant was not asked."
    assert why_line(None) == why_line("something new") == "The assistant was not asked."
    assert set(handoff.OUTCOMES) == set(handoff._WHY)  # every outcome has its sentence


@pytest.mark.parametrize("outcome, tries, did", [
    ("no_fix", 1, "The assistant tried once: it proposed no change."),
    ("no_fix", 2, "The assistant tried twice: it proposed no change."),
    ("too_wide", 2, "The assistant tried twice: its change was too wide to apply without a person."),
    ("numbers_changed", 1, "The assistant tried once: its change altered a result and was undone."),
    ("not_cleared", 2, "The assistant tried twice: its change did not solve the problem and was undone."),
    ("not_checked", 2, "The assistant tried twice: its change could not be checked and was undone."),
    ("not_written", 1, "The assistant tried once: Excel refused its change."),
    ("not_tried", 0, "The assistant was not asked."),
])
def test_the_feedback_note_word_for_word(outcome, tries, did):
    assert feedback_note(_problem(outcome=outcome, tries=tries)) == f"Start could not fix this by itself: {NAME}. {did}"
    assert feedback_note(_problem(outcome=outcome, tries=tries), "Do everything").startswith("Do everything could not fix this by itself: ")


def test_the_feedback_note_holds_no_rule_id_no_version_no_sheet_or_cell_name():
    for outcome in handoff.OUTCOMES:
        for tries in (0, 1, 2):
            p = {**_problem(cells=11, outcome=outcome, tries=tries), "version": "ver-004", "from_version_id": "ver-003", "label": "v5"}
            note = feedback_note(p)
            assert "STR-008" not in note and not re.search(r"\b[A-Z]{2,7}-\d{3}\b", note), note  # no rule id
            assert not re.search(r"\bver-\d+|\bv\d+\b|version", note, re.IGNORECASE), note  # no version
            assert SHEET not in note and "W25" not in note and not re.search(r"\b[A-Z]{1,3}\d{1,7}\b", note), note  # no sheet, no cell
            assert len(note) <= 500  # it is the first thing in the email's body, which takes 500 characters
    # a rule with no plain name: still no code
    assert feedback_note({"rule_id": "ZZZ-999", "outcome": "no_fix", "tries": 1}) == "Start could not fix this by itself: a problem the app has no name for. The assistant tried once: it proposed no change."


def test_show_me_leads_to_the_prep_action_or_to_the_findings_row():
    assert show_me_for(_problem(route="prep")) == {"id": "prep-action", "rule_id": "STR-008", "sheet": SHEET, "cell": "W25"}
    assert show_me_for(_problem(route="prep_skipped"))["id"] == "prep-action"
    assert show_me_for(_problem(route="assistant")) == {"id": "assistant-ask", "rule_id": "STR-008", "sheet": SHEET, "cell": "W25"}


def test_the_whole_ask():
    left = [_problem(cells=11, outcome="numbers_changed", tries=2), _problem(rule="FRM-007", route="assistant", outcome="not_tried", tries=0)]
    plan = [_action([_op("W25")])]
    ask = build_ask(left, plan, "ABCD-1234")
    assert ask == {
        "lines": [f"2 things are left that the app could not fix. The first: {NAME}, on sheet {SHEET}, cell W25.", "Fix it by hand, then press Start again."],
        "why": "The assistant's fix changed a result, so it was taken back.",
        "stops_start": True,  # a blocking Prep action that is on by default still has changes
        "show_me": {"id": "prep-action", "rule_id": "STR-008", "sheet": SHEET, "cell": "W25"},
        "feedback": {"ref": "ABCD-1234", "note": f"Start could not fix this by itself: {NAME}. The assistant tried twice: its change altered a result and was undone."},
    }
    # nothing left for Prep: the question does not stop Start (it is asked at the end)
    assert build_ask(left[1:], [_action([])], None)["stops_start"] is False
    assert build_ask(left[1:], [_action([])], None, "Do everything")["lines"][1] == "Fix it by hand, then press Do everything again."
    assert build_ask([], plan) is None


def test_no_ai_key():
    ask = no_key_ask()
    assert ask["lines"] == ["Something is left that needs the assistant, and no AI key is saved.", "Paste an AI key under Advanced › Assistant, then press Start again."]
    assert ask["show_me"]["id"] == "assistant-key" and ask["stops_start"] is True and ask["feedback"] is None
    assert no_key_ask("Do everything")["lines"][1] == "Paste an AI key under Advanced › Assistant, then press Do everything again."


def test_the_jobs_messages_word_for_word():
    assert round_message(2, 966) == "Round 2: writing 966 changes through Excel…"
    assert assistant_message(1, 2) == "The assistant is working on what is left (1 of 2)."
    assert handoff.WRITE_MESSAGE == "Writing the assistant's fix."
    assert handoff.NUMBERS_MESSAGE == "Checking that no result changed."
    assert undo_message() == undo_message("numbers_changed") == "A result changed: taking the fix back."
    assert undo_message("not_cleared") == undo_message("not_checked") == "Taking the fix back."  # no result changed there
    assert handoff.CHECK_MESSAGE == "Checking the workbook again."
    assert done_lines(966) == ["966 changes applied."]
    assert done_lines(0) == ["Nothing blocked the file."]
    assert done_lines(966, 2) == ["966 changes applied.", "The assistant fixed 2 more: see Advanced › Versions."]
    assert done_lines(0, 1) == ["The assistant fixed 1: see Advanced › Versions."]  # "Nothing blocked the file." would not be true


def test_no_sentence_for_the_screen_holds_a_rule_code():
    shown = (done_lines(3, 1) + done_lines(0) + no_key_ask()["lines"] + [round_message(1, 2), assistant_message(1, 1), handoff.WRITE_MESSAGE, handoff.NUMBERS_MESSAGE,
             undo_message(), undo_message("not_cleared"), handoff.CHECK_MESSAGE] + [why_line(o) for o in handoff.OUTCOMES])
    for line in shown:
        assert not re.search(r"\b[A-Z]{2,7}-\d{3}\b", line), line


def test_the_module_needs_no_excel_and_no_server():
    # pure functions: what it imports is the standard library and three small modules of the app that
    # read no workbook (inventory only inside memo_key's default reader: openpyxl, never Excel)
    import ast

    tree = ast.parse(Path(handoff.__file__).read_text(encoding="utf-8"))
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported |= {a.name for a in node.names}
        elif isinstance(node, ast.ImportFrom):
            imported.add(("." * node.level) + (node.module or ""))
    assert imported == {"__future__", "hashlib", "re", "typing", ".formula_utils", ".plain_codes", ".readiness", ".inventory"}
