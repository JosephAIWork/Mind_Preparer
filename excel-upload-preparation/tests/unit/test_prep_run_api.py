"""17 Nov 2026: the prep-run job (POST and GET /api/sessions/{id}/prep-run in app/web/server.py, on top
of app/handoff.py), through the test client. Start's Prep step never stops for a person while the
app can fix it: Prep rounds while a round still changes something, then what is left goes to the
assistant, whose fix is kept only when no number changed, its problem got smaller and no other
blocking problem got worse.

No Excel and no model run here. The Apply, the assistant and the numbers check are fakes, and what
the check of the workbook "finds" in each version is written by the test (`World`): the real check
runs on the small workbook, then the findings and the Prep plan the test asked for are put in.

The promises these tests pin: a fix that changes a value, does not help, or cannot be checked never
stays current, and nothing of it is saved before it is judged; the person is asked only about what
is really left at the end; a Cancel puts the workbook back as it was before the run."""
from __future__ import annotations

import importlib.util
import json
import time
from pathlib import Path

import pytest

from app import cancel, session_store
from app.settings import SETTINGS
from app.web import server
from test_cancel_api import STOPPED, WAIT, Gate, _answer, _applied, _bg, _cancel, _clean, _incidents, _jobs, _look, _upload  # the Cancel tests' own tools
from test_cancel_api import client, excels  # noqa: F401  (fixtures: a client on a data folder of its own; no process is ever ended)

ACTION = "title_lone_text_cells"  # a real Prep action: blocking, on by default, for the lone-text rule
PREP_RULE = "STR-008"             # ... its rule ("a formula reads a lone text cell")
RULE = "FRM-011"                  # a rule with no Prep work in these tests: its problem goes to the assistant
OTHER = "REF-001"
MANY = ["FRM-006", "FRM-007", "FRM-011", "FRM-012", "FRM-014", "REF-001", "REF-002"]  # seven rules, in the report's order
LINE_BREAK = "a line break typed inside formula text"  # the plain names (rules/plain_names.yaml)
BROKEN_REF = "a broken reference to a cell or name"
LONE_TEXT = "a formula reads a lone text cell"
AGAIN = "Fix it by hand, then press Start again."
BUSY = "Wait a moment: the formulas and the structure are still being fixed."
_REAL_ANALYZE = server._analyze  # the check itself, taken before any test puts its own in its place
FIELDS = {"state", "stage", "title", "message", "pass_no", "applied", "refused", "assistant_fixed", "problem_no", "problem_count", "problems", "left", "ask",
          "from_version_id", "version", "started", "updated", "finished", "elapsed_s", "error"}  # GET .../prep-run, as the design gives it


# --- the test's own tools ---------------------------------------------------------------------------
class World:
    """What the check of the workbook "finds", written by the test: the states the workbook goes
    through, one step for every Apply that wrote something. A state is {"problems": {rule id: [cells
    on sheet Data]}, "prep": [the changes the blocking Prep action still holds]}. Every file has its
    place in the list, so a version that is taken back takes its state with it, and the next Apply
    starts again from the state of the version it copies."""

    def __init__(self, *states: dict) -> None:
        self.states = [{"problems": dict(st.get("problems") or {}), "prep": list(st.get("prep") or [])} for st in states]
        self.place: dict[str, int] = {}
        self.applies: list[list[dict]] = []  # the operations of every Apply, in order
        self.scans: list[Path] = []          # every check, in order
        self.refuse = lambda ops: []         # the changes "Excel" refuses
        self.scan_fails = False
        self.apply_error: Exception | None = None

    def of(self, path) -> dict:
        return self.states[self.place.get(str(Path(path).resolve()), 0)]

    def written(self, source, out, wrote: bool) -> None:
        here = self.place.get(str(Path(source).resolve()), 0)
        self.place[str(Path(out).resolve())] = min(here + 1, len(self.states) - 1) if wrote else here


def _doctor(s, state: dict) -> None:
    """Put the test's problems and Prep work into the check that just ran (as the rules would report
    them: a failing required rule with its cells, and the action's operations)."""
    report = s.result["validation_report"]
    for f in report["findings"]:
        cells = state["problems"].get(f["rule_id"])
        if cells:
            f.update(status="ERROR", priority="REQUIRED", message="put here by the test", location={"sheet": "Data", "cell": cells[0]},
                     observed={"sites": [{"sheet": "Data", "cell": c} for c in cells], "sites_count": len(cells)})
    for a in s.plan:
        if a["id"] == ACTION:
            a.update(operations=[dict(o) for o in state["prep"]], count=len(state["prep"]))
    server._recount_report(report)


def _install(monkeypatch, world: World, gate: Gate | None = None, hold=lambda n: True) -> None:
    """The fake Apply (app.prep.apply_operations without Excel; with a gate, Apply number n waits
    there when `hold(n)`, between two written changes) and the check with the test's findings."""
    real = _REAL_ANALYZE  # never another test world's check: a test may install twice

    def analyze(s, path, progress=None):
        world.scans.append(Path(path))
        if world.scan_fails:
            raise RuntimeError("the check fell over")
        out = real(s, path, progress)
        _doctor(s, world.of(path))
        return out

    def apply_operations(source_path, work_dir, operations, prefer_excel=True, progress=None, defer_verify=False, run=None, verify=None, excel=None):
        world.applies.append([dict(o) for o in operations])
        progress("apply_copy", "Copying the version", None, done=0, total=len(operations))
        if world.apply_error is not None:
            raise world.apply_error
        if gate is not None and hold(len(world.applies)):
            gate.hold(lambda: progress("apply_write", "Writing", None, done=0, total=len(operations)))
        refused = world.refuse(operations)
        wrote = [o for o in operations if o not in refused]
        res = _applied(source_path, work_dir, wrote)
        res["failed"] = [{**o, "error": "Excel refused the change (a protected sheet)"} for o in refused]
        res["status"] = "APPLIED" if wrote and not refused else "PARTIAL" if wrote else "ERROR"
        world.written(source_path, res["output_path"], bool(wrote))
        return res

    monkeypatch.setattr(server, "_analyze", analyze)
    monkeypatch.setattr(server, "apply_operations", apply_operations)


class Assistant:
    """app.chat_context.answer_question without a model: every call takes the next reply of the list
    (the last one again once the list is used up). A reply is the operations of its proposal, or a
    dict of the answer's own fields, or a function of the question that gives one of the two."""

    def __init__(self, *replies) -> None:
        self.replies = list(replies) or [[]]
        self.questions: list[str] = []

    def __call__(self, history, question, analysis, validation_report, prep_actions=None, **kw):
        assert history == [] and prep_actions is not None
        self.questions.append(question)
        reply = self.replies[min(len(self.questions), len(self.replies)) - 1]
        if callable(reply):
            reply = reply(question)
        answer = {"available": True, "text": "Here is a fix.", "message": None, "proposal": [], "proposal_summary": None, "proposal_errors": []}
        return {**answer, **(reply if isinstance(reply, dict) else {"proposal": [dict(o) for o in reply]})}


def _fix(cell: str = "B4", formula: str = "=A4+B3+0", **more) -> dict:
    """One operation of the assistant's proposal, as app.prep.validate_proposal hands it on."""
    return {"action_id": "assistant", "rule_id": "CHAT", "sheet": "Data", "note": "a fix", "op": "set_formula", "cell": cell, "before": "=A4+B3", "after": formula, **more}


def _prep_op(cell: str) -> dict:
    """One change of the blocking Prep action: a title for a lone text cell."""
    return {"action_id": ACTION, "rule_id": PREP_RULE, "op": "set_value", "sheet": "Data", "cell": cell, "before": None, "after": "#Label"}


def _numbers(changed: int = 0, ran: bool = True, calls: list | None = None, gate: Gate | None = None):
    """app.numbers_check.compare_workbooks without Excel: `changed` computed values differ."""
    def compare_workbooks(original, prepared, moves, progress=None):
        if calls is not None:
            calls.append((Path(original), Path(prepared)))
        if gate is not None:
            gate.hold()
        if not ran:
            return {"ran": False, "clean": None, "verdict": "The numbers could not be compared: Excel did not open the file.", "counts": {}, "regressions": []}
        return {"ran": True, "clean": not changed, "verdict": "compared", "regressions": [],
                "counts": {"compared": 1, "same": 0 if changed else 1, "good_to_error": 0, "good_to_other": changed, "missing_good": 0},
                "samples": {"good_to_error": [], "good_to_other": [{"original": "Data!B4", "now": "Data!B4", "was": "5", "is": "6"}] if changed else [], "missing": []}}

    return compare_workbooks


def _start(client, sid: str, **body) -> None:
    res = client.post(f"/api/sessions/{sid}/prep-run", json=body)
    assert res.status_code == 202 and res.json() == {"state": "running"}, res.text


def _wait(client, sid: str) -> dict:
    deadline = time.time() + WAIT
    while True:
        st = client.get(f"/api/sessions/{sid}/prep-run").json()
        if st["state"] != "running":
            return st
        assert time.time() < deadline, st
        time.sleep(0.02)


def _run(client, sid: str, **body) -> dict:
    _start(client, sid, **body)
    return _wait(client, sid)


def _at_stage(client, sid: str, stage: str) -> dict:
    deadline = time.time() + WAIT
    while True:
        st = client.get(f"/api/sessions/{sid}/prep-run").json()
        if st["state"] == "running" and st["stage"] == stage:
            return st
        assert st["state"] == "running" and time.time() < deadline, st
        time.sleep(0.02)


def _saved_versions(sid: str) -> list[str]:
    return [v["id"] for v in session_store.read_state(sid)["versions"]]


def _last_incident() -> dict:
    path = Path(SETTINGS.data_folder) / "feedback" / "incidents.jsonl"
    return json.loads(path.read_text(encoding="utf-8").splitlines()[-1])


def _free(client, sid: str) -> None:
    """Nothing runs for the workbook any more, and nothing says it does."""
    s = server.SESSIONS[sid]
    assert server._busy_with(s) is None and _jobs(client, sid) == [] and not s.lock.locked() and not s.apply_lock.locked()
    status = client.get(f"/api/sessions/{sid}/status").json()
    assert status["state"] != "running" and (status["apply"] or {}).get("state") != "running", status


@pytest.fixture
def prep(client, monkeypatch):  # noqa: F811
    """The client, with nothing kept from another test and an AI key "saved"."""
    server.PREP_RUNS.clear()
    server.HANDOFF_MEMO.clear()
    monkeypatch.setattr(server, "llm_available", lambda: True)
    try:
        yield client
    finally:
        server.PREP_RUNS.clear()
        server.HANDOFF_MEMO.clear()


# --- the routes -------------------------------------------------------------------------------------
def test_before_any_run_the_job_is_idle_and_an_unknown_workbook_is_404(prep, plain_grid_xlsx, monkeypatch):
    assert prep.get("/api/sessions/nope/prep-run").status_code == 404
    assert prep.post("/api/sessions/nope/prep-run", json={}).status_code == 404
    _install(monkeypatch, World({}))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    assert prep.get(f"/api/sessions/{sid}/prep-run").json() == {"state": "idle"}
    # nothing blocks the file: no round, no assistant, no question
    asked = Assistant()
    monkeypatch.setattr(server, "answer_question", asked)
    st = _run(prep, sid)
    assert set(st) == FIELDS  # every field of the contract, and no incident: that comes only with an error
    assert (st["state"], st["stage"], st["message"]) == ("done", "done", "Nothing blocked the file.")
    assert (st["pass_no"], st["applied"], st["refused"], st["assistant_fixed"], st["problems"], st["left"], st["ask"], st["error"]) == (0, 0, 0, 0, [], 0, None, None)
    assert st["from_version_id"] == "ver-001" and st["version"]["id"] == "ver-001" and st["finished"] >= st["started"]
    assert asked.questions == [] and _saved_versions(sid) == ["ver-001"]
    full = prep.get(f"/api/sessions/{sid}/prep-run?full=1").json()
    assert full["analysis"] == prep.get(f"/api/sessions/{sid}").json()  # the workbook as GET /api/sessions/{id} gives it
    _free(prep, sid)
    assert _incidents() == 0


def test_the_rounds_go_on_while_a_round_still_changes_something(prep, plain_grid_xlsx, monkeypatch):
    """A rule that shows one new cell per round (the staircase of the lone-text rule before it was
    fixed at its root): four rounds, each with other work than the one before, then nothing is left."""
    gate = Gate()
    world = World(*[{"problems": {PREP_RULE: [f"A{n}"]}, "prep": [_prep_op(f"Z{n}")]} for n in (2, 3, 4, 5)], {})
    _install(monkeypatch, world, gate, hold=lambda n: n == 1)
    asked = Assistant()
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _start(prep, sid)
    assert gate.started.wait(WAIT)
    st = prep.get(f"/api/sessions/{sid}/prep-run").json()
    assert (st["state"], st["stage"], st["pass_no"], st["message"]) == ("running", "pass", 1, "Round 1: writing 1 changes through Excel…")
    assert st["title"] == "Fixing the formulas and the structure" and st["from_version_id"] == "ver-001"
    assert [(j["kind"], j["title"]) for j in _jobs(prep, sid)] == [("prep", "Fixing the formulas and the structure")]
    assert prep.get(f"/api/sessions/{sid}/status").json()["apply"]["state"] == "running"  # the Apply's own progress, as before
    gate.go.set()
    st = _wait(prep, sid)
    assert (st["state"], st["pass_no"], st["applied"], st["refused"], st["left"], st["ask"]) == ("done", 4, 4, 0, 0, None)
    assert st["message"] == "4 changes applied." and st["problems"] == [] and st["version"]["id"] == "ver-005"
    assert [[o["cell"] for o in ops] for ops in world.applies] == [["Z2"], ["Z3"], ["Z4"], ["Z5"]]
    assert [v["source"] for v in s.versions] == ["upload", "prep", "prep", "prep", "prep"] and _saved_versions(sid)[-1] == "ver-005"
    assert asked.questions == []  # nothing was left for it
    _free(prep, sid)


def test_the_same_changes_coming_back_stop_the_rounds_and_a_refused_change_alone_does_not(prep, plain_grid_xlsx, monkeypatch):
    # round 1 writes Z2 and is refused Z3; round 2 writes Z4 and is refused Z3 again; round 3 is
    # refused everything: it wrote nothing, so the rounds stop there
    both = {"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z2"), _prep_op("Z3")]}
    world = World(both, {"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z3"), _prep_op("Z4")]}, {"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z3")]})
    world.refuse = lambda ops: [o for o in ops if o["cell"] == "Z3"]
    _install(monkeypatch, world)
    asked = Assistant()
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid, assistant=False)
    assert (st["pass_no"], st["applied"], st["refused"]) == (3, 2, 3) and len(world.applies) == 3
    # what is left is asked about, in the words of the Prep route: Start stopped here before tonight too
    assert st["state"] == "ask" and st["left"] == 1 and asked.questions == []
    assert st["ask"]["stops_start"] is True and st["ask"]["show_me"] == {"id": "prep-action", "rule_id": PREP_RULE, "sheet": "Data", "cell": "A2"}
    assert st["problems"] == [{"rule_id": PREP_RULE, "rule_name": LONE_TEXT, "sheet": "Data", "cell": "A2", "cells": 1, "route": "prep", "tries": 0, "outcome": "not_tried"}]
    # the same work coming back unchanged stops after ONE round
    again = World(both)
    _install(monkeypatch, again)
    sid2 = _upload(prep, plain_grid_xlsx)["sessionId"]
    st2 = _run(prep, sid2, assistant=False)
    assert (st2["state"], st2["pass_no"], st2["applied"]) == ("ask", 1, 2) and len(again.applies) == 1
    _free(prep, sid)
    _free(prep, sid2)


# --- the assistant's fix ----------------------------------------------------------------------------
def test_a_fix_that_is_kept(prep, plain_grid_xlsx, monkeypatch):
    world, compared = World({"problems": {RULE: ["B4"]}}, {}), []
    _install(monkeypatch, world)
    asked = Assistant([_fix()])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers(calls=compared))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    assert s.readiness()["state"] == "blocked"
    st = _run(prep, sid)
    assert (st["state"], st["left"], st["ask"], st["assistant_fixed"], st["applied"], st["pass_no"]) == ("done", 0, None, 1, 0, 0)
    assert st["message"] == "The assistant fixed 1: see Advanced › Versions."
    assert st["problems"] == [{"rule_id": RULE, "rule_name": LINE_BREAK, "sheet": "Data", "cell": "B4", "cells": 1, "route": "assistant", "tries": 1, "outcome": "fixed"}]
    assert (st["problem_no"], st["problem_count"]) == (1, 1)
    # the version is kept, current and saved; the workbook is not blocked any more
    assert [(v["id"], v["source"]) for v in s.versions] == [("ver-001", "upload"), ("ver-002", "assistant")] and s.current_version_id == "ver-002"
    assert st["version"]["id"] == "ver-002" and st["version"]["label"].startswith("v2 — Assistant: 1 change(s)")
    assert _saved_versions(sid) == ["ver-001", "ver-002"] and s.readiness()["blocking"] == []
    # the numbers check itself, between the version just before the fix and the version with it
    assert compared == [(s.paths["ver-001"], s.paths["ver-002"])]
    # the question is the Assistant tab's "Propose a fix" for that finding, and names its cells
    assert len(asked.questions) == 1 and asked.questions[0].startswith(f"[About finding {RULE} at Data!B4. ") and "Propose a fix\nCells: Data!B4." in asked.questions[0]
    assert world.applies == [[_fix()]] and len(world.scans) == 2  # the upload's check, and the check of the fix
    assert server.HANDOFF_MEMO[sid] == {}  # nothing to remember: it did not fail
    _free(prep, sid)
    assert _incidents() == 0


def test_a_fix_that_changes_a_value_is_undone_and_only_then_the_person_is_asked(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}}, {})  # the fix would even clear the problem: the changed value decides
    _install(monkeypatch, world)
    asked = Assistant([_fix()], [_fix(formula="=A4+B3+1")])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers(changed=1))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before, status_before = _look(s), prep.get(f"/api/sessions/{sid}/status").json()
    st = _run(prep, sid)
    # the workbook is back on the version before the fix, exactly: nothing of either try is left
    assert s.current_version_id == "ver-001" and _look(s) == before and _saved_versions(sid) == ["ver-001"]
    assert len(world.applies) == 2 and len(world.scans) == 1  # both tries were written; a changed value needs no check of the workbook
    status = prep.get(f"/api/sessions/{sid}/status").json()
    assert (status["state"], status["version_id"], status["apply"]) == (status_before["state"], "ver-001", None)  # no status speaks of a version that is gone
    # ... and only then is the person asked, in plain words
    assert (st["state"], st["left"], st["assistant_fixed"], st["version"]["id"], st["error"]) == ("ask", 1, 0, "ver-001", None)
    assert st["problems"] == [{"rule_id": RULE, "rule_name": LINE_BREAK, "sheet": "Data", "cell": "B4", "cells": 1, "route": "assistant", "tries": 2, "outcome": "numbers_changed"}]
    ask = st["ask"]
    assert ask["lines"] == [f"One thing is left that the app could not fix: {LINE_BREAK}, on sheet Data, cell B4.", AGAIN]
    assert st["message"] == " ".join(ask["lines"])
    assert ask["why"] == "The assistant's fix changed a result, so it was taken back."
    assert ask["stops_start"] is False  # no Prep fix is left: Start goes on and asks at its end
    assert ask["show_me"] == {"id": "assistant-ask", "rule_id": RULE, "sheet": "Data", "cell": "B4"}
    assert ask["feedback"]["note"] == f"Start could not fix this by itself: {LINE_BREAK}. The assistant tried twice: its change altered a result and was undone."
    assert "incident" not in st and "problem" not in st  # the question carries its feedback itself: no red box beside it
    incident = _last_incident()
    assert _incidents() == 1 and incident["ref"] == ask["feedback"]["ref"]
    assert (incident["step"], incident["rule_id"], incident["what"]) == ("Assistant", RULE, "not fixed by the assistant")
    # the second try was told why the first failed, with the value that changed
    assert len(asked.questions) == 2 and "Your first proposal changed a computed result (Data!B4 showed 5 and would show 6)" in asked.questions[1]
    assert "Your first proposal" not in asked.questions[0]
    _free(prep, sid)


def test_no_proposal_and_a_proposal_with_errors_write_nothing(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    asked = Assistant([], {"proposal": [_fix()], "proposal_errors": ["operation 1 (set_formula): formula must start with '='"]})
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "no_fix" and st["problems"][0]["tries"] == 2
    assert st["ask"]["why"] == "The assistant had no safe fix for it."
    assert st["ask"]["feedback"]["note"] == f"Start could not fix this by itself: {LINE_BREAK}. The assistant tried twice: it proposed no change."
    assert world.applies == [] and _look(s) == before  # a proposal the chat itself found faulty is never written
    assert "Your first answer held no change the app could apply" in asked.questions[1]
    _free(prep, sid)


def test_a_fix_that_is_too_wide_is_not_written(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}}, {})
    _install(monkeypatch, world)
    wide = [  # a cell with content elsewhere; then an inserted row beside a fix of the problem's own cell
        [{"action_id": "assistant", "rule_id": "CHAT", "sheet": "Data", "op": "set_value", "cell": "A3", "before": 1, "after": 2}],
        [_fix(), {"action_id": "assistant", "rule_id": "CHAT", "sheet": "Data", "op": "insert_row", "row": 2, "before": "row 2", "after": "empty row inserted at 2"}],
    ]
    asked = Assistant(*wide)
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "too_wide" and st["problems"][0]["tries"] == 2
    assert st["ask"]["why"] == "The assistant's fix was too wide to apply without you."
    assert world.applies == [] and _look(s) == before and _saved_versions(sid) == ["ver-001"]  # nothing was written, by either try
    assert "Your first proposal was too wide to write without a person (operation 1 (set_value) changes Data!A3" in asked.questions[1]
    # the narrow shapes are written: the problem's own cell, and a label in an empty cell beside it
    server.HANDOFF_MEMO.clear()
    narrow = [_fix(), {"action_id": "assistant", "rule_id": "CHAT", "sheet": "Data", "op": "set_value", "cell": "C4", "before": None, "after": "#Label"}]
    monkeypatch.setattr(server, "answer_question", Assistant(narrow))
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    st = _run(prep, sid)
    assert st["state"] == "done" and world.applies == [narrow] and s.current_version_id == "ver-002"
    _free(prep, sid)


def test_a_fix_that_does_not_make_its_problem_smaller_is_undone(prep, plain_grid_xlsx, monkeypatch):
    # try 1 leaves the problem as it is; try 2 clears it and turns another rule blocking: neither is kept
    world = World({"problems": {RULE: ["B4"]}}, {"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    asked = Assistant([_fix()])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)

    def worse_after_the_first(ops):
        if len(world.applies) == 2:
            world.states[1] = {"problems": {OTHER: ["A3"]}, "prep": []}
        return []

    world.refuse = worse_after_the_first
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "not_cleared" and st["problems"][0]["tries"] == 2
    assert st["ask"]["why"] == "The assistant's fix did not solve it, so it was taken back."
    assert len(world.applies) == 2 and len(world.scans) == 3  # each try was written and checked ...
    assert s.current_version_id == "ver-001" and _look(s) == before and _saved_versions(sid) == ["ver-001"]  # ... and taken back
    assert [b["rule_id"] for b in s.readiness()["blocking"]] == [RULE]  # the other rule's problem went with the fix that made it
    status = prep.get(f"/api/sessions/{sid}/status").json()
    assert status["state"] == "ready" and status["version_id"] == "ver-001" and status["apply"] is None
    _free(prep, sid)


def test_a_comparison_that_did_not_run_is_never_a_pass(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}}, {})  # the fix would clear the problem: without the numbers it is not kept
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([_fix()]))
    monkeypatch.setattr(server, "compare_workbooks", _numbers(ran=False))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "not_checked" and st["problems"][0]["tries"] == 2
    assert st["ask"]["why"] == "The assistant's fix could not be checked, so it was taken back."
    assert st["ask"]["feedback"]["note"].endswith("The assistant tried twice: its change could not be checked and was undone.")
    assert len(world.applies) == 2 and len(world.scans) == 1 and _look(s) == before and _saved_versions(sid) == ["ver-001"]
    _free(prep, sid)


def test_a_change_excel_refuses_is_taken_back(prep, plain_grid_xlsx, monkeypatch):
    world, compared = World({"problems": {RULE: ["B4"]}}, {}), []
    world.refuse = lambda ops: list(ops)
    _install(monkeypatch, world)
    asked = Assistant([_fix()])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers(calls=compared))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "not_written" and st["ask"]["why"] == "Excel did not accept the assistant's fix."
    assert compared == [] and _look(s) == before and _saved_versions(sid) == ["ver-001"]  # the version with nothing in it is gone too
    assert "Excel refused your first proposal (Excel refused the change (a protected sheet))" in asked.questions[1]
    _free(prep, sid)


def test_nothing_of_the_fix_is_saved_while_the_comparison_waits(prep, plain_grid_xlsx, monkeypatch):
    world, gate = World({"problems": {RULE: ["B4"]}}, {}), Gate()
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([_fix()]))
    monkeypatch.setattr(server, "compare_workbooks", _numbers(gate=gate))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _start(prep, sid)
    assert gate.started.wait(WAIT)
    st = prep.get(f"/api/sessions/{sid}/prep-run").json()
    assert (st["state"], st["stage"], st["message"], st["title"]) == ("running", "numbers", "Checking that no result changed.", "Checking the numbers")
    assert (st["problem_no"], st["problem_count"], st["problems"][0]["tries"]) == (1, 1, 1)
    # the fix is written and current in memory, and the state file does not know it: a crash here loses only the fix
    assert s.current_version_id == "ver-002" and s.versions[-1]["source"] == "assistant"
    saved = session_store.read_state(sid)
    assert [v["id"] for v in saved["versions"]] == ["ver-001"] and saved["current_version_id"] == "ver-001"
    assert not any(v.get("source") == "assistant" for v in saved["versions"])
    assert server._busy_with(s) == "the formulas and the structure are still being fixed"
    gate.go.set()
    assert _wait(prep, sid)["state"] == "done" and _saved_versions(sid) == ["ver-001", "ver-002"]  # judged and kept: now it is saved
    _free(prep, sid)


def test_a_check_that_raises_leaves_the_workbook_on_the_version_before_the_fix(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}}, {})
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([_fix()]))
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    world.scan_fails = True  # the check of the workbook with the fix in it falls over, both times
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["error"] is None and st["problems"][0]["outcome"] == "not_checked" and st["problems"][0]["tries"] == 2
    assert s.current_version_id == "ver-001" and _look(s) == before and _saved_versions(sid) == ["ver-001"]
    assert len(world.applies) == 2 and len(world.scans) == 3
    status = prep.get(f"/api/sessions/{sid}/status").json()
    assert status["state"] == "ready" and status["error"] is None and status["version_id"] == "ver-001"  # not the failed check of a version that is gone
    assert _incidents() == 1  # the question's own report, and none for the check: the person is asked, nothing "failed"
    _free(prep, sid)


def test_a_fix_that_makes_its_problem_smaller_is_kept_and_the_rest_gets_the_second_try(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4", "A3", "A4"]}}, {"problems": {RULE: ["A3"]}}, {})
    _install(monkeypatch, world)
    asked = Assistant([_fix()], [_fix(cell="A3", formula="=1")])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    st = _run(prep, sid)
    assert st["state"] == "done" and st["assistant_fixed"] == 1 and st["problems"][0]["tries"] == 2 and st["problems"][0]["outcome"] == "fixed"
    assert st["problems"][0]["cells"] == 3  # the problem as it was when the assistant got it
    assert [v["source"] for v in s.versions] == ["upload", "assistant", "assistant"] and _saved_versions(sid) == ["ver-001", "ver-002", "ver-003"]
    assert "Cells: Data!B4, Data!A3, Data!A4." in asked.questions[0]  # every cell of the problem is named
    assert "Cells: Data!A3." in asked.questions[1] and "Your first proposal was kept and fixed part of the problem" in asked.questions[1]
    _free(prep, sid)


def test_two_tries_at_most_for_a_problem_and_five_problems_at_most(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {rule: ["B4"] for rule in MANY}})
    _install(monkeypatch, world)
    asked = Assistant([])
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert len(asked.questions) == 10  # five problems, two tries each: the cap on what one run may cost
    assert [(p["rule_id"], p["tries"], p["outcome"]) for p in st["problems"]] == [(r, 2, "no_fix") for r in MANY[:5]] + [(r, 0, "not_tried") for r in MANY[5:]]
    assert [q.split(" at ")[0] for q in asked.questions] == [f"[About finding {r}" for r in MANY[:5] for _ in (1, 2)]
    assert (st["state"], st["left"], st["problem_no"], st["problem_count"]) == ("ask", 7, 5, 5)
    assert st["ask"]["lines"][0].startswith("7 things are left that the app could not fix. The first: ") and st["ask"]["lines"][1] == AGAIN
    _free(prep, sid)


def test_a_second_run_on_an_untouched_workbook_makes_no_model_call(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    asked = Assistant([])
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    first = _run(prep, sid)
    assert first["state"] == "ask" and len(asked.questions) == 2 and len(server.HANDOFF_MEMO[sid]) == 1
    again = _run(prep, sid)  # the run after the Mind loop, or Start pressed again: the same problem, the same cells
    assert len(asked.questions) == 2  # not paid twice
    assert again["state"] == "ask" and again["problems"] == first["problems"] and again["ask"]["lines"] == first["ask"]["lines"]
    assert again["ask"]["why"] == first["ask"]["why"] and again["ask"]["feedback"]["note"] == first["ask"]["feedback"]["note"]
    # the memo goes with the workbook
    assert prep.delete(f"/api/saved-workbooks/{sid}").status_code == 200
    assert sid not in server.HANDOFF_MEMO and sid not in server.PREP_RUNS


def test_the_problem_asked_about_comes_from_the_final_blocking_list(prep, plain_grid_xlsx, monkeypatch):
    # two problems after the rounds; the assistant fixes the first and has nothing for the second
    world = World({"problems": {RULE: ["B4"], OTHER: ["A3"]}}, {"problems": {OTHER: ["A3"]}})
    _install(monkeypatch, world)
    asked = Assistant([_fix()], [])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert [(p["rule_id"], p["tries"], p["outcome"]) for p in st["problems"]] == [(RULE, 1, "fixed"), (OTHER, 2, "no_fix")]
    assert (st["state"], st["left"], st["assistant_fixed"]) == ("ask", 1, 1)
    assert st["ask"]["lines"][0] == f"One thing is left that the app could not fix: {BROKEN_REF}, on sheet Data, cell A3."  # not the first of the list it was given
    assert st["ask"]["show_me"]["rule_id"] == OTHER and _last_incident()["rule_id"] == OTHER
    assert st["ask"]["feedback"]["note"] == f"Start could not fix this by itself: {BROKEN_REF}. The assistant tried twice: it proposed no change."
    _free(prep, sid)


def test_prep_runs_once_more_after_a_kept_fix_and_what_it_leaves_is_asked_about_without_the_assistant(prep, plain_grid_xlsx, monkeypatch):
    # the kept fix leaves Prep work; that round turns up a problem nobody saw before: no second turn of the assistant
    world = World({"problems": {RULE: ["B4"]}}, {"prep": [_prep_op("Z9")]}, {"problems": {OTHER: ["A3"]}})
    _install(monkeypatch, world)
    asked = Assistant([_fix()])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    st = _run(prep, sid)
    assert [v["source"] for v in s.versions] == ["upload", "assistant", "prep"] and (st["pass_no"], st["applied"], st["assistant_fixed"]) == (1, 1, 1)
    assert len(asked.questions) == 1
    assert st["state"] == "ask" and st["left"] == 1 and [p["rule_id"] for p in st["problems"]] == [RULE]  # the list the assistant was given
    assert st["ask"]["lines"][0] == f"One thing is left that the app could not fix: {BROKEN_REF}, on sheet Data, cell A3."
    assert st["ask"]["why"] == "The assistant was not asked." and st["ask"]["stops_start"] is False
    assert st["ask"]["feedback"]["note"] == f"Start could not fix this by itself: {BROKEN_REF}. The assistant was not asked."
    _free(prep, sid)


def test_stops_start_is_true_with_a_prep_action_left_and_false_with_none(prep, plain_grid_xlsx, monkeypatch):
    # a Prep change Excel refuses every time, and nothing from the assistant: the rule by which this step stopped Start before tonight
    world = World({"problems": {PREP_RULE: ["A2", "B2"]}, "prep": [_prep_op("Z9")]})
    world.refuse = lambda ops: list(ops)
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([]))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert (st["state"], st["pass_no"], st["applied"], st["refused"]) == ("ask", 1, 0, 1)
    assert st["ask"]["stops_start"] is True and st["ask"]["show_me"] == {"id": "prep-action", "rule_id": PREP_RULE, "sheet": "Data", "cell": "A2"}
    assert st["ask"]["lines"][0] == f"One thing is left that the app could not fix: {LONE_TEXT}, in 2 cells; the first is A2 on sheet Data."
    assert st["problems"][0]["route"] == "prep" and st["problems"][0]["outcome"] == "no_fix"
    # with no Prep action left the same question does not stop Start: it goes on and asks at its end
    other = World({"problems": {RULE: ["B4"]}})
    _install(monkeypatch, other)
    sid2 = _upload(prep, plain_grid_xlsx)["sessionId"]
    st2 = _run(prep, sid2)
    assert st2["state"] == "ask" and st2["ask"]["stops_start"] is False and st2["ask"]["show_me"]["id"] == "assistant-ask"
    _free(prep, sid)
    _free(prep, sid2)


def test_no_ai_key_says_so_and_keeps_the_rounds(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {PREP_RULE: ["A2"], RULE: ["B4"]}, "prep": [_prep_op("Z9")]}, {"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    asked = Assistant([_fix()])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "llm_available", lambda: False)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    st = _run(prep, sid)
    assert st["state"] == "no_key" and asked.questions == []
    assert st["ask"] == {
        "lines": ["Something is left that needs the assistant, and no AI key is saved.", "Paste an AI key under Advanced › Assistant, then press Start again."],
        "why": "The assistant was not asked.", "stops_start": True, "show_me": {"id": "assistant-key", "rule_id": None, "sheet": "", "cell": ""}, "feedback": None,
    }
    # the round that ran is kept
    assert (st["pass_no"], st["applied"], st["left"]) == (1, 1, 1) and [v["source"] for v in s.versions] == ["upload", "prep"]
    assert s.current_version_id == "ver-002" and _saved_versions(sid) == ["ver-001", "ver-002"] and st["version"]["id"] == "ver-002"
    assert [(p["rule_id"], p["outcome"]) for p in st["problems"]] == [(RULE, "not_tried")]
    assert _incidents() == 0  # the app is not at fault: nothing to send
    # with nothing left, no key is no problem at all
    done = World({"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z9")]}, {})
    _install(monkeypatch, done)
    sid2 = _upload(prep, plain_grid_xlsx)["sessionId"]
    assert _run(prep, sid2)["state"] == "done"
    _free(prep, sid)
    _free(prep, sid2)


def test_without_the_assistant_what_is_left_is_asked_about_in_the_words_of_the_button(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    asked = Assistant([_fix()])
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid, assistant=False, button="Do everything")  # Custom's button
    assert st["state"] == "ask" and asked.questions == [] and st["problems"][0]["outcome"] == "not_tried" and st["problem_count"] == 0
    assert st["ask"]["lines"][1] == "Fix it by hand, then press Do everything again." and st["ask"]["why"] == "The assistant was not asked."
    assert st["ask"]["feedback"]["note"] == f"Do everything could not fix this by itself: {LINE_BREAK}. The assistant was not asked."
    assert _last_incident()["what"] == "left for a person: the assistant was not asked"
    assert server.HANDOFF_MEMO.get(sid, {}) == {}
    st = _run(prep, sid, assistant=False, button="Go")  # any other word is Basic's button
    assert st["ask"]["lines"][1] == AGAIN
    _free(prep, sid)


# --- Cancel -----------------------------------------------------------------------------------------
def test_a_cancel_in_a_round_gives_cancelled_and_the_first_snapshot(prep, plain_grid_xlsx, monkeypatch):
    """Round 1 is finished and saved when the Cancel lands in round 2: both go. This is what makes
    "Your workbook is as it was after ..." true for the Prep step of Start."""
    gate = Gate()
    world = World({"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z2")]}, {"problems": {PREP_RULE: ["A3"]}, "prep": [_prep_op("Z3")]}, {})
    _install(monkeypatch, world, gate, hold=lambda n: n == 2)
    monkeypatch.setattr(server, "answer_question", Assistant([]))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    _start(prep, sid)
    assert gate.started.wait(WAIT)
    assert [v["id"] for v in s.versions] == ["ver-001", "ver-002"] and _saved_versions(sid) == ["ver-001", "ver-002"]  # round 1 is in, and saved
    assert prep.get(f"/api/sessions/{sid}/prep-run").json()["pass_no"] == 2
    out = _cancel(prep, sid, "prep")
    assert out["cancelled"] == ["prep"] and out["late"] == [] and out["stopping"] == [] and out["message"] == STOPPED
    assert out["session"]["version"]["id"] == "ver-001" and out["session"]["versions"] == before["versions"]
    st = prep.get(f"/api/sessions/{sid}/prep-run").json()
    assert (st["state"], st["title"], st["message"], st["error"], st["ask"], st["version"]) == ("cancelled", "Stopped", STOPPED, None, None, None)
    assert (st["pass_no"], st["applied"], st["refused"], st["assistant_fixed"], st["problems"]) == (0, 0, 0, 0, []) and "incident" not in st
    assert _look(s) == before and _saved_versions(sid) == ["ver-001"]  # the whole run, not only the round that was running
    status = prep.get(f"/api/sessions/{sid}/status").json()
    assert status["state"] == "cancelled" and status["apply"]["state"] == "cancelled"
    # nothing comes later, and the same step runs again from the same version to the same numbers
    gate.go.set()
    _clean(prep, sid)
    st = _run(prep, sid)
    assert (st["state"], st["pass_no"], st["applied"], st["version"]["id"]) == ("done", 2, 2, "ver-003") and [v["parent"] for v in s.versions[1:]] == ["ver-001", "ver-002"]
    assert _incidents() == 0


def test_a_cancel_during_the_assistants_turn_gives_cancelled_and_the_first_snapshot(prep, plain_grid_xlsx, monkeypatch):
    world, gate = World({"problems": {PREP_RULE: ["A2"], RULE: ["B4"]}, "prep": [_prep_op("Z2")]}, {"problems": {RULE: ["B4"]}}), Gate()
    _install(monkeypatch, world)

    def thinking(question):
        gate.started.set()
        cancel.call(lambda: gate.go.wait(WAIT))  # as app/llm.py makes its request: given up when Cancel is pressed
        return [_fix()]

    asked = Assistant(thinking)
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    _start(prep, sid)
    assert gate.started.wait(WAIT)
    st = prep.get(f"/api/sessions/{sid}/prep-run").json()
    assert (st["stage"], st["message"], st["applied"]) == ("assistant", "The assistant is working on what is left (1 of 1).", 1)
    assert s.current_version_id == "ver-002" and _saved_versions(sid) == ["ver-001", "ver-002"]  # the round before it is kept so far
    out = _cancel(prep, sid)  # Start's own Cancel: no kind
    assert out["cancelled"] == ["prep"] and out["message"] == STOPPED and out["session"]["version"]["id"] == "ver-001"
    st = prep.get(f"/api/sessions/{sid}/prep-run").json()
    assert st["state"] == "cancelled" and st["message"] == STOPPED and st["applied"] == 0 and st["problems"] == []
    assert _look(s) == before and _saved_versions(sid) == ["ver-001"] and world.applies == [[_prep_op("Z2")]]  # the fix was never written
    assert server.HANDOFF_MEMO[sid] == {}  # a stopped turn is no failure of the assistant
    gate.go.set()
    _clean(prep, sid)


def test_a_cancel_while_a_written_fix_is_being_checked_takes_it_back_with_everything_else(prep, plain_grid_xlsx, monkeypatch, excels):  # noqa: F811
    world = World({"problems": {RULE: ["B4"]}}, {})
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([_fix()]))

    def compare(original, prepared, moves, progress=None):
        assert excels.start(4808).wait(WAIT)  # its own Excel, ended by the Cancel
        return {"ran": False, "clean": None, "verdict": "The numbers could not be compared: the Excel was closed.", "counts": {}, "regressions": []}  # what the real one answers then

    monkeypatch.setattr(server, "compare_workbooks", compare)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    _start(prep, sid)
    for _ in range(1500):
        if excels.gone:
            break
        time.sleep(0.02)
    assert s.current_version_id == "ver-002"  # the fix is written, not judged
    out = _cancel(prep, sid, "prep")
    assert out["cancelled"] == ["prep"] and excels.killed == [4808]
    st = prep.get(f"/api/sessions/{sid}/prep-run").json()
    assert st["state"] == "cancelled" and st["problems"] == [] and st["ask"] is None  # "could not be checked" is no outcome of a stopped run
    assert _look(s) == before and _saved_versions(sid) == ["ver-001"] and server.HANDOFF_MEMO[sid] == {}
    _clean(prep, sid)
    # too late: after the run has ended nothing is taken back
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    assert _run(prep, sid)["state"] == "done"
    assert _cancel(prep, sid, "prep")["message"] == "Nothing was running." and s.current_version_id == "ver-002"


# --- busy -------------------------------------------------------------------------------------------
def test_409_when_busy_and_everything_else_waits_for_the_job(prep, plain_grid_xlsx, monkeypatch):
    gate = Gate()
    world = World({"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z2")]}, {})
    _install(monkeypatch, world, gate, hold=lambda n: n == 1)
    # no analysis yet
    pending = _upload(prep, plain_grid_xlsx, defer=True)["sessionId"]
    res = prep.post(f"/api/sessions/{pending}/prep-run", json={})
    assert res.status_code == 409 and res.json()["detail"] == "no analysis yet"
    # another job changes the workbook: an Apply that is still being written
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    applying = _bg(lambda c: c.post(f"/api/sessions/{sid}/apply", json={"operations": [_prep_op("Y1")], "reanalyze": False}))
    assert gate.started.wait(WAIT)
    res = prep.post(f"/api/sessions/{sid}/prep-run", json={})
    assert res.status_code == 409 and res.json()["detail"] == "Wait a moment: changes are still being written."
    assert prep.get(f"/api/sessions/{sid}/prep-run").json() == {"state": "idle"}
    gate.go.set()
    assert _answer(applying).status_code == 200 and s.current_version_id == "ver-002"
    # the job itself: a second one is refused, and so is everything that would land in its middle
    gate2 = Gate()
    _install(monkeypatch, world, gate2, hold=lambda n: True)
    _start(prep, sid)
    assert gate2.started.wait(WAIT)
    assert server._busy_with(s) == "the formulas and the structure are still being fixed"
    again = prep.post(f"/api/sessions/{sid}/prep-run", json={})
    assert again.status_code == 409 and again.json()["detail"] == BUSY
    for route, body in (("apply", {"operations": [_prep_op("Y2")]}), ("reanalyze", {}), ("recalculate", None), ("numbers-check", None), ("mind-loop-plan/apply", {})):
        res = prep.post(f"/api/sessions/{sid}/{route}", json=body) if body is not None else prep.post(f"/api/sessions/{sid}/{route}")
        assert res.status_code == 409 and res.json()["detail"] == BUSY, (route, res.text)
    assert prep.post(f"/api/sessions/{sid}/analyze", json={}).status_code == 409
    assert prep.delete(f"/api/saved-workbooks/{sid}").status_code == 409  # Remove waits too
    gate2.go.set()
    assert _wait(prep, sid)["state"] == "done"
    _free(prep, sid)
    assert prep.post(f"/api/sessions/{sid}/reanalyze", json={}).status_code == 200  # free again


def test_the_end_is_said_only_once_the_workbook_is_free(prep, plain_grid_xlsx, monkeypatch):
    """A screen that reads "done" starts the next step at once: it must not meet "Wait a moment"."""
    world, gate = World({}), Gate()
    _install(monkeypatch, world)
    real_end = cancel.end

    def slow_end(job, outcome="done"):
        if job is not None and job.kind == "prep":
            gate.hold()  # the job's last save takes seconds on a big model
        real_end(job, outcome)

    monkeypatch.setattr(cancel, "end", slow_end)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    _start(prep, sid)
    assert gate.started.wait(WAIT)
    assert server.PREP_RUNS[sid]["state"] == "done"  # the thread has written its last word ...
    st = prep.get(f"/api/sessions/{sid}/prep-run?full=1").json()
    assert st["state"] == "running" and "analysis" not in st  # ... and the route does not say it yet
    assert server._busy_with(s) == "the formulas and the structure are still being fixed"
    assert prep.post(f"/api/sessions/{sid}/recalculate").status_code == 409
    gate.go.set()
    assert _wait(prep, sid)["state"] == "done"
    _free(prep, sid)


# --- errors -----------------------------------------------------------------------------------------
def test_an_error_in_a_round_ends_the_job_with_its_report_and_never_leaves_it_running(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z2")]}, {"problems": {PREP_RULE: ["A3"]}, "prep": [_prep_op("Z3")]}, {})
    _install(monkeypatch, world)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    real_refuse = world.refuse

    def fails_in_round_two(ops):
        if len(world.applies) == 2:
            raise RuntimeError("Excel went away")
        return real_refuse(ops)

    world.refuse = fails_in_round_two
    st = _run(prep, sid)
    assert st["state"] == "error" and "Excel went away" in st["error"] and st["ask"] is None
    assert st["incident"] and st["problem"].startswith("Fix the formulas and the structure stopped on an error")  # the red box and Send feedback
    assert _incidents() == 1 and _last_incident()["step"] == "Fix the formulas and the structure"
    # the round that finished is kept (as before tonight, when each round was a request of its own)
    assert (st["pass_no"], st["applied"], st["version"]["id"], st["left"]) == (2, 1, "ver-002", 1)
    assert s.current_version_id == "ver-002" and _saved_versions(sid) == ["ver-001", "ver-002"]
    _free(prep, sid)
    # a round whose own check did not run: nothing may be decided from the check before it
    world2 = World({"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z2")]}, {})
    _install(monkeypatch, world2)
    sid2 = _upload(prep, plain_grid_xlsx)["sessionId"]
    world2.scan_fails = True
    st2 = _run(prep, sid2)
    assert st2["state"] == "error" and st2["error"].startswith("The changes were written, but the check of the workbook that follows did not run")
    assert st2["incident"] and server.SESSIONS[sid2].current_version_id == "ver-002" and len(world2.applies) == 1  # no second round on an old check
    # the scan status says "error" here, and that is the truth: this check did fail, and its version stays
    assert server._busy_with(server.SESSIONS[sid2]) is None and _jobs(prep, sid2) == []


def test_an_error_while_a_written_fix_is_checked_never_leaves_the_fix_current(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"]}}, {})
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([_fix()]))

    def compare(original, prepared, moves, progress=None):
        raise RuntimeError("the comparison fell over")

    monkeypatch.setattr(server, "compare_workbooks", compare)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "not_checked" and len(world.applies) == 2
    assert s.current_version_id == "ver-001" and _look(s) == before and _saved_versions(sid) == ["ver-001"]
    # the write itself failing is "Excel did not accept it", and the assistant's own call failing is "could not be
    # reached" (17 Nov 2026: it was "no fix", and remembered as such): nothing to take back
    server.HANDOFF_MEMO.clear()
    world.apply_error = RuntimeError("Excel went away")
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "not_written" and _look(s) == before
    server.HANDOFF_MEMO.clear()

    def no_line(*a, **k):
        raise ConnectionError("no connection")

    monkeypatch.setattr(server, "answer_question", no_line)
    st = _run(prep, sid)
    assert st["state"] == "ask" and st["problems"][0]["outcome"] == "not_reached" and _look(s) == before and server.HANDOFF_MEMO[sid] == {}
    _free(prep, sid)


# --- scripts/night_chain.py -------------------------------------------------------------------------
class _Http:
    """`requests` for scripts/night_chain.py, answered by the test client."""

    def __init__(self, client) -> None:
        self.client, self.posted = client, []

    def post(self, url, json=None, timeout=None):
        self.posted.append((url.split("//x", 1)[1], json))
        return self.client.post(url.split("//x", 1)[1], json=json)

    def get(self, url, timeout=None):
        return self.client.get(url.split("//x", 1)[1])


def _night_chain():
    spec = importlib.util.spec_from_file_location("night_chain_under_test", Path(server.PACKAGE_ROOT) / "scripts" / "night_chain.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_night_chain_runs_the_prep_step_through_the_job_and_records_how_it_ended(prep, plain_grid_xlsx, monkeypatch):
    chain = _night_chain()
    said: list = []
    world = World({"problems": {PREP_RULE: ["A2"], RULE: ["B4"]}, "prep": [_prep_op("Z2")]}, {"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    asked = Assistant([])
    monkeypatch.setattr(server, "answer_question", asked)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    http = _Http(prep)
    step = chain.prep_run("http://x/api", sid, assistant=False, every=0.02, http=http, tell=lambda *a: said.append(a))  # --no-assistant
    assert http.posted == [(f"/api/sessions/{sid}/prep-run", {"assistant": False, "button": "Do everything"})] and asked.questions == []
    assert (step["ok"], step["state"], step["passes"], step["applied"], step["refused"]) == (False, "ask", 1, 1, 0)  # not ok: something is left
    assert [(p["rule_id"], p["outcome"]) for p in step["problems"]] == [(RULE, "not_tried")]
    assert step["ask"]["lines"][1] == "Fix it by hand, then press Do everything again." and step["version"].startswith("v2 — Prep") and step["left"] == 1
    assert said  # each new line of the job is printed
    # with the assistant (the default), and a workbook it leaves clean: ok
    world2 = World({"problems": {PREP_RULE: ["A2"]}, "prep": [_prep_op("Z2")]}, {})
    _install(monkeypatch, world2)
    sid2 = _upload(prep, plain_grid_xlsx)["sessionId"]
    step = chain.prep_run("http://x/api", sid2, every=0.02, http=http, tell=lambda *a: None)
    assert http.posted[-1][1] == {"assistant": True, "button": "Do everything"}
    assert (step["ok"], step["state"], step["passes"], step["applied"], step["ask"], step["error"]) == (True, "done", 1, 1, None, None)
    # refused (the workbook is busy): recorded as not ok, and the chain goes on
    gate = Gate()
    _install(monkeypatch, world2, gate)
    applying = _bg(lambda c: c.post(f"/api/sessions/{sid2}/apply", json={"operations": [_prep_op("Y1")], "reanalyze": False}))
    assert gate.started.wait(WAIT)
    step = chain.prep_run("http://x/api", sid2, every=0.02, http=http, tell=lambda *a: None)
    assert (step["ok"], step["state"]) == (False, "refused") and "Wait a moment" in step["detail"]
    gate.go.set()
    assert _answer(applying).status_code == 200
    _free(prep, sid)
    _free(prep, sid2)


# --- 17 Nov 2026, after the review ------------------------------------------------------------------
def _shown(**counts):
    """compare_workbooks without Excel, with the counts given: an error that now shows a value, or a
    formula that showed an error and is gone -- no value that was good changed."""
    def compare_workbooks(original, prepared, moves, progress=None):
        c = {"compared": 1, "same": 0, "good_to_error": 0, "good_to_other": 0, "error_to_good": 0, "missing": 0, "missing_good": 0, **counts}
        return {"ran": True, "clean": True, "verdict": "The preparation changed no computed value.", "regressions": [], "counts": c,
                "samples": {"good_to_error": [], "good_to_other": [], "missing": [],
                            "error_to_good": [{"original": "Data!B4", "now": "Data!B4", "was": "#REF!", "is": "42"}] if c["error_to_good"] else []}}

    return compare_workbooks


@pytest.mark.parametrize("counts", [{"error_to_good": 1}, {"missing": 1}], ids=["an error now shows a number", "a formula that showed an error is gone"])
def test_a_fix_that_changes_what_an_error_cell_shows_is_undone(prep, plain_grid_xlsx, monkeypatch, counts):
    # R2-1 (Joseph: "if any value changed, the fix is undone"): #REF! that now shows 42 is a changed
    # value, even though the numbers check does not call it one (Fix all issues does that on purpose)
    world = World({"problems": {OTHER: ["B4"]}}, {})  # the guess even clears the problem: the changed value decides
    _install(monkeypatch, world)
    asked = Assistant([_fix(formula="=A2+A1")])
    monkeypatch.setattr(server, "answer_question", asked)
    monkeypatch.setattr(server, "compare_workbooks", _shown(**counts))
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    s = server.SESSIONS[sid]
    before = _look(s)
    st = _run(prep, sid)
    assert s.current_version_id == "ver-001" and _look(s) == before and _saved_versions(sid) == ["ver-001"]
    assert [(p["rule_id"], p["tries"], p["outcome"]) for p in st["problems"]] == [(OTHER, 2, "numbers_changed")]
    assert st["state"] == "ask" and st["assistant_fixed"] == 0
    assert st["ask"]["why"] == "The assistant's fix changed a result, so it was taken back."
    if counts.get("error_to_good"):
        assert "Your first proposal changed a computed result (Data!B4 showed #REF! and would show 42)" in asked.questions[1]
    # what the numbers check reports elsewhere is unchanged
    assert server._changed_values(_shown(**counts)(None, None, None)) == 0
    _free(prep, sid)


def _unreachable(monkeypatch, how: str) -> list:
    """The real answer_question and chat_completion, with a key saved; only the network call is a fake
    that fails as `how` says. -> the list of calls tried."""
    import requests

    from app import llm

    posts: list = []

    class Refused:
        status_code = {"401": 401, "500": 500}.get(how, 200)
        text = "refused"
        headers = {"content-type": "application/json"}

        def raise_for_status(self):
            raise requests.HTTPError(f"{self.status_code} Client Error")

        def close(self):
            pass

    def post(*a, **k):
        posts.append(how)
        if how == "connection":
            raise requests.ConnectionError("the gateway did not answer")
        if how == "timeout":
            raise requests.Timeout("read timed out")
        return Refused()

    monkeypatch.setattr(llm, "find_api_key", lambda: ("a-key", "http://127.0.0.1:9"))
    monkeypatch.setattr(requests, "post", post)
    return posts


@pytest.mark.parametrize("how", ["connection", "timeout", "401", "500"])
def test_an_assistant_that_could_not_be_reached_is_not_remembered_and_is_asked_again_next_time(prep, plain_grid_xlsx, monkeypatch, how):
    # R2-2: a call that failed comes back "available" with no text: never "it had no fix"
    world = World({"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    posts = _unreachable(monkeypatch, how)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert st["state"] == "ask" and [(p["tries"], p["outcome"]) for p in st["problems"]] == [(1, "not_reached")]  # no second try
    assert server.HANDOFF_MEMO[sid] == {}  # never remembered as "no fix"
    assert st["ask"]["why"] == "The assistant could not be reached."
    assert st["ask"]["feedback"]["note"] == f"Start could not fix this by itself: {LINE_BREAK}. The assistant could not be reached."
    assert _last_incident()["what"] == "left for a person: the assistant could not be reached"
    first = len(posts)
    assert first >= 1
    _run(prep, sid)  # the connection is back, or not: the next Start asks again
    assert len(posts) > first
    _free(prep, sid)


def test_after_an_assistant_that_could_not_be_reached_the_next_problems_are_not_asked(prep, plain_grid_xlsx, monkeypatch):
    world = World({"problems": {RULE: ["B4"], OTHER: ["A3"]}})
    _install(monkeypatch, world)
    posts = _unreachable(monkeypatch, "connection")
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert [p["outcome"] for p in st["problems"]] == ["not_reached", "not_tried"] and len(posts) == 1
    _free(prep, sid)


@pytest.mark.parametrize("body", [{"content": [], "stop_reason": "refusal"}, {"id": "msg_1", "type": "message"}], ids=["no text", "a shape not known"])
def test_a_model_that_answered_with_no_fix_is_no_fix_and_is_not_asked_again(prep, plain_grid_xlsx, monkeypatch, body):
    # the refuter's check of R2-2: the gateway answered (200), the model gave no text, or a reply in a
    # shape not known. That is an answer: "no fix", remembered, never "could not be reached"
    import requests

    from app import llm

    posts: list = []

    class Answered:
        status_code = 200
        text = "{}"
        headers = {"content-type": "application/json"}

        def raise_for_status(self):
            pass

        def json(self):
            return dict(body)

        def close(self):
            pass

    def post(*a, **k):
        posts.append(1)
        return Answered()

    monkeypatch.setattr(llm, "find_api_key", lambda: ("a-key", "http://127.0.0.1:9"))
    monkeypatch.setattr(llm, "meter", lambda *a, **k: None)  # never a row in the real token meter
    monkeypatch.setattr(requests, "post", post)
    world = World({"problems": {RULE: ["B4"]}})
    _install(monkeypatch, world)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert st["state"] == "ask" and [(p["tries"], p["outcome"]) for p in st["problems"]] == [(2, "no_fix")]
    assert st["ask"]["why"] != "The assistant could not be reached." and len(server.HANDOFF_MEMO[sid]) == 1
    asked = len(posts)
    st = _run(prep, sid)  # remembered: the next Start does not pay for the same question again
    assert len(posts) == asked and [p["outcome"] for p in st["problems"]] == ["no_fix"]
    _free(prep, sid)


def test_a_key_taken_away_while_start_runs_says_no_ai_key_is_saved(prep, plain_grid_xlsx, monkeypatch):
    # the refuter's check of R2-2: no key when the question is sent is not "could not be reached"
    from app import llm

    keys = ["a-key"]
    real = server.answer_question

    def asked(*a, **k):
        keys.clear()  # the person deletes the key under Advanced > Assistant while Start runs
        return real(*a, **k)

    monkeypatch.setattr(llm, "find_api_key", lambda: ("a-key", "http://127.0.0.1:9") if keys else None)
    monkeypatch.setattr(server, "llm_available", lambda: bool(keys))
    monkeypatch.setattr(server, "answer_question", asked)
    world = World({"problems": {RULE: ["B4"], OTHER: ["A3"]}})
    _install(monkeypatch, world)
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    st = _run(prep, sid)
    assert st["state"] == "no_key" and st["ask"]["lines"][0] == "Something is left that needs the assistant, and no AI key is saved."
    assert [(p["tries"], p["outcome"]) for p in st["problems"]] == [(0, "not_tried"), (0, "not_tried")]
    assert server.HANDOFF_MEMO[sid] == {} and _incidents() == 0
    _free(prep, sid)


def test_nothing_of_the_fix_is_saved_by_the_check_before_it_is_judged(prep, plain_grid_xlsx, monkeypatch):
    # R2-7: the check of the workbook with the fix (step h) does not write the state file: the fix is
    # not judged yet, and a crash then would reopen the workbook on it
    world = World({"problems": {RULE: ["B4"]}}, {})
    _install(monkeypatch, world)
    monkeypatch.setattr(server, "answer_question", Assistant([_fix()]))
    monkeypatch.setattr(server, "compare_workbooks", _numbers())
    sid = _upload(prep, plain_grid_xlsx)["sessionId"]
    seen: list = []
    real = server.handoff.judge

    def judge(*a, **k):
        seen.append(_saved_versions(sid))  # what a crash right now would reopen
        return real(*a, **k)

    monkeypatch.setattr(server.handoff, "judge", judge)
    st = _run(prep, sid)
    assert seen == [["ver-001"]] and len(world.scans) == 2  # the check ran, and saved nothing
    assert st["state"] == "done" and _saved_versions(sid) == ["ver-001", "ver-002"]  # judged and kept: now it is saved
    _free(prep, sid)
