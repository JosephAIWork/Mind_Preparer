"""17 Nov 2026 (Palermo): the errors a Mind loop makes, told apart from the real ones.

Once the loop is built, Excel computes the loop's FIRST value (policy 1), which the original never
computed in those cells (the macro pasted results): 667 #N/A in Decrements stopped Start at the
recalculation, and pressing Start again stopped it the same way. Now the recalculation also
recalculates the version before the loop (once; a recalculation the session holds is reused) and
compares cell by cell, the cells followed through the rows Prep inserted since: an error that does
not show there shows only for the loop's first value and never blocks. Excel is replaced here."""
from __future__ import annotations

import time
from pathlib import Path

from app.readiness import compute_readiness
from app.web import server

READY = {"rule_id": "READY-001", "status": "NOT_SUPPORTED", "priority": "REQUIRED"}
NA = "#N/A"


def _session(tmp_path: Path, steps: list[tuple[str, list[dict]]]) -> server.Session:
    """A session with one version per step (source, change log), each a file of its own, and an analysis
    with nothing blocking: only the recalculation stands between it and Mind."""
    s = server.Session(id=f"lb{time.time_ns()}", work_dir=tmp_path / "work", mode="plan")
    for i, (source, log) in enumerate(steps, start=1):
        p = tmp_path / "versions" / f"v{i}" / "model.xlsm"
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"x")
        server._add_version(s, p, source, source, log, None, "0" * 64)
    s.result = {"validation_report": {"findings": [dict(READY)], "status": "NOT_SUPPORTED", "summary": {}}}
    s.plan = []
    for v in s.versions:
        s.ref_formulas[str(s.paths[v["id"]])] = set()
    return s


def _err(sheet: str, cell: str, error: str = NA) -> dict:
    return {"sheet": sheet, "cell": cell, "error": error, "formula": f"=VLOOKUP({cell},Decrements!$B$7:$D$125,2,0)"}


class FakeExcel:
    """app.recalc.recalculate without Excel: each file's error cells, by the version its copy was made
    from (v1, v2, ...); `refuse` names the versions whose recalculation does not run. `runs`: the
    versions recalculated, in order."""

    def __init__(self, s: server.Session, errors: dict[str, list[dict]], refuse: tuple[str, ...] = ()) -> None:
        self.s, self.errors, self.refuse, self.runs = s, errors, refuse, []
        self.copies: dict[str, str] = {}

    def recalculate(self, path, save=True):
        version = self.copies[str(Path(path))]
        self.runs.append(version)
        if version in self.refuse:
            return {"status": "ERROR", "message": "Excel COM recalculation failed: Excel closed or stopped answering", "ran": False, "formula_errors": [], "addin_gap_errors": []}
        return {"status": "ERROR" if self.errors.get(version) else "PASS", "message": "recalculated", "ran": True, "formula_errors": list(self.errors.get(version, [])), "addin_gap_errors": []}


def _wire(monkeypatch, s: server.Session, fake: FakeExcel) -> None:
    """Every disposable copy keeps the name of the version it was made from, so the fake knows the file."""
    real = server.make_immutable_copy

    def copy(source, work_dir):
        out, sha = real(source, work_dir)
        fake.copies[str(Path(out))] = Path(source).parent.name
        return out, sha

    monkeypatch.setattr(server, "make_immutable_copy", copy)
    monkeypatch.setattr(server, "recalculate", fake.recalculate)
    monkeypatch.setattr(server, "_save", lambda s: None)


def _recalc_current(s: server.Session, errors: list[dict]) -> dict:
    """The Recalculate step's own recalculation of the current version, handed in as the route does."""
    return server._recalc_out(s, {"status": "ERROR" if errors else "PASS", "message": "recalculated", "ran": True, "formula_errors": errors, "addin_gap_errors": []}, s.current_path)


# Palermo's line: the upload, Prep, the Mind loop, Prep again (a title row inserted on Decrements above the table)
PALERMO = [("upload", []), ("prep", []), ("loop", [{"op": "set_formula", "sheet": "Run", "cell": "C3", "after": '=MM_LOOP("Policy",1,6000)'}]),
           ("prep", [{"op": "insert_row", "sheet": "Decrements", "row": 2}])]


def test_errors_only_for_the_loops_first_value_never_stop_start(tmp_path, monkeypatch):
    s = _session(tmp_path, PALERMO)
    # the original and the version before the loop computed the stored policy: no error in Decrements
    fake = FakeExcel(s, {"v1": [], "v2": []})
    _wire(monkeypatch, s, fake)
    after = [_err("Decrements", f"M{r}") for r in range(704, 704 + 667)]
    out = _recalc_current(s, after)
    check = out["loop_check"]
    assert check["ran"] and check["from_loop"] == 667 and check["also_without"] == 0 and check["version_id"] == "ver-002" and check["label"] == "v2"
    assert check["example"] == {"sheet": "Decrements", "cell": "M704", "error": NA} and all(e.get("loop_only") for e in out["formula_errors"])
    assert fake.runs == ["v1", "v2"]  # the original's baseline, then the version before the loop
    r = out["readiness"]
    assert r["state"] == "ready" and r["recalc"]["clean"] and r["recalc"]["new_errors"] == 0 and r["recalc"]["loop_errors"] == 667
    assert r["recalc"]["loop_note"] == (
        "667 cells show an error only for the Mind loop's first value, for example Decrements!M704: #N/A means the value is not available "
        "(often: a lookup found nothing). The original never computed these cells, and Mind will show the same for that value. Check the results in Mind."
    )
    assert r["next_step"]["label"] == "Download or send to Mind" and r["next_step"]["detail"] == r["recalc"]["loop_note"]
    # READY-001 passes: the workbook can go to Mind
    finding = s.result["validation_report"]["findings"][0]
    assert finding["status"] == "PASS" and "only for the Mind loop's first value" in finding["message"]
    # the version before the loop is recalculated once: the next Recalculate reuses it
    runs = len(fake.runs)
    _recalc_current(s, after)
    assert len(fake.runs) == runs == 2 and s.recalc["loop_check"]["reused"] is True


def test_only_errors_that_also_show_without_the_loop_stop_start(tmp_path, monkeypatch):
    s = _session(tmp_path, PALERMO)
    # the version before the loop already errs at Results!B5 (not in the original): a real new error
    fake = FakeExcel(s, {"v1": [], "v2": [_err("Results", "B5", "#DIV/0!")]})
    _wire(monkeypatch, s, fake)
    after = [_err("Decrements", f"M{r}") for r in range(704, 714)] + [_err("Results", "B5", "#DIV/0!")]
    out = _recalc_current(s, after)
    assert out["loop_check"]["from_loop"] == 10 and out["loop_check"]["also_without"] == 1
    r = out["readiness"]
    assert r["state"] == "unverified" and not r["recalc"]["clean"] and r["recalc"]["new_errors"] == 1 and r["recalc"]["loop_errors"] == 10
    assert r["headline"] == "Acceptable by Mind, but the recalculation found 1 new formula error(s)"
    assert r["recalc"]["loop_note"] == "The other 10 cells show an error only for the Mind loop's first value, for example Decrements!M704: #N/A means the value is not available (often: a lookup found nothing)."
    assert r["next_step"]["detail"] == "1 cell(s) not in error in the original. " + r["recalc"]["loop_note"]
    assert s.result["validation_report"]["findings"][0]["status"] == "ERROR"


def test_the_rows_prep_inserted_after_the_loop_are_followed(tmp_path, monkeypatch):
    """Before the loop the error sat at Decrements!M704; Prep after the loop inserted a row above it, so
    the same cell is M705 now: it also shows without the loop. M704 now is another cell: from the loop."""
    s = _session(tmp_path, PALERMO)
    fake = FakeExcel(s, {"v1": [], "v2": [_err("Decrements", "M704")]})
    _wire(monkeypatch, s, fake)
    out = _recalc_current(s, [_err("Decrements", "M704"), _err("Decrements", "M705")])
    loop_only = {e["cell"] for e in out["formula_errors"] if e.get("loop_only")}
    assert loop_only == {"M704"} and out["loop_check"]["also_without"] == 1
    assert out["readiness"]["state"] == "unverified" and out["readiness"]["recalc"]["new_errors"] == 1


def test_a_cell_an_apply_after_the_loop_wrote_is_never_the_loops(tmp_path, monkeypatch):
    """Review LBr: Prep after the loop writes a formula that errs (Results!B5, #DIV/0!), a cell with no
    error before the loop. That error is Prep's, not the loop's first value's: it still stops Start.
    The written cells are followed like the rest: M703, written in the same Apply that inserts a row
    above it, is M704 now; M705 nobody wrote stays the loop's."""
    after_loop = [
        {"op": "insert_row", "sheet": "Decrements", "row": 2},
        {"op": "set_formula", "rule_id": "FRM-009", "sheet": "Results", "cell": "B5", "before": "=A5/A6", "after": "=N(A5/A6)"},
        {"op": "set_formula", "rule_id": "FRM-009", "sheet": "Decrements", "cell": "M703", "before": "=L703", "after": "=N(L703)"},
    ]
    s = _session(tmp_path, PALERMO[:3] + [("prep", after_loop)])
    fake = FakeExcel(s, {"v1": [], "v2": []})
    _wire(monkeypatch, s, fake)
    out = _recalc_current(s, [_err("Decrements", "M704"), _err("Decrements", "M705"), _err("Results", "B5", "#DIV/0!")])
    check = out["loop_check"]
    assert {e["cell"] for e in out["formula_errors"] if e.get("loop_only")} == {"M705"}
    assert check["from_loop"] == 1 and check["also_without"] == 2 and check["written_after"] == 2
    r = out["readiness"]
    assert r["state"] == "unverified" and r["recalc"]["new_errors"] == 2 and r["recalc"]["loop_errors"] == 1
    assert s.result["validation_report"]["findings"][0]["status"] == "ERROR"
    # the loop's own writes stay the loop's: a second loop after the Prep writes M705 again
    s2 = _session(tmp_path / "b", PALERMO[:3] + [("prep", after_loop), ("loop", [{"op": "set_formula", "sheet": "Decrements", "cell": "M705", "after": "=MM_LOOP(1)"}])])
    _wire(monkeypatch, s2, FakeExcel(s2, {"v1": [], "v2": []}))
    out = _recalc_current(s2, [_err("Decrements", "M705")])
    assert out["loop_check"]["from_loop"] == 1 and out["readiness"]["state"] == "ready"


def test_when_the_version_before_the_loop_cannot_be_recalculated_nothing_is_guessed(tmp_path, monkeypatch):
    s = _session(tmp_path, PALERMO)
    fake = FakeExcel(s, {"v1": []}, refuse=("v2",))
    _wire(monkeypatch, s, fake)
    out = _recalc_current(s, [_err("Decrements", "M704")])
    assert out["loop_check"]["ran"] is False and "Excel closed" in out["loop_check"]["message"]
    r = out["readiness"]
    assert r["state"] == "unverified" and r["recalc"]["new_errors"] == 1 and r["recalc"]["loop_errors"] == 0
    assert r["recalc"]["loop_note"] == "Whether these errors come only from the Mind loop's first value could not be checked: Excel could not recalculate the version before the loop."
    assert not any(e.get("loop_only") for e in out["formula_errors"])
    # nothing was kept for it: the next Recalculate tries again
    assert not any(k.startswith("before_loop:") for k in s.baselines)


def test_without_a_loop_nothing_changes_and_no_extra_recalculation_runs(tmp_path, monkeypatch):
    s = _session(tmp_path, [("upload", []), ("prep", [])])
    fake = FakeExcel(s, {"v1": []})
    _wire(monkeypatch, s, fake)
    out = _recalc_current(s, [_err("Decrements", "M704")])
    assert "loop_check" not in out and fake.runs == ["v1"]  # the original's baseline only
    r = out["readiness"]
    assert r["state"] == "unverified" and r["recalc"]["new_errors"] == 1 and "loop_note" not in r["recalc"] and "loop_errors" not in r["recalc"]


def test_a_clean_recalculation_after_the_loop_needs_no_extra_one(tmp_path, monkeypatch):
    s = _session(tmp_path, PALERMO)
    fake = FakeExcel(s, {"v1": [_err("Decrements", "M704")]})
    _wire(monkeypatch, s, fake)
    # an error the original already had (M704 moved to M705 by the title row) is no new error: nothing to check
    out = _recalc_current(s, [_err("Decrements", "M705")])
    assert out["new_errors"] == 0 and "loop_check" not in out and fake.runs == ["v1"]
    assert out["readiness"]["state"] == "ready"


def test_a_recalculation_the_session_holds_of_the_version_before_the_loop_is_reused(tmp_path, monkeypatch):
    s = _session(tmp_path, PALERMO)
    fake = FakeExcel(s, {"v1": []})
    _wire(monkeypatch, s, fake)
    # the last recalculation is of v2 itself (Start's own, before the loop was built)
    s.recalc = {"version_id": "ver-002", "ran": True, "formula_errors": [], "addin_gap_errors": []}
    out = _recalc_current(s, [_err("Decrements", "M704")])
    assert out["loop_check"]["reused"] is True and out["loop_check"]["from_loop"] == 1 and fake.runs == ["v1"]


def test_a_recalculation_saved_before_the_check_keeps_the_guess():
    """A recalculation stored before 17 Nov 2026 carries no loop check: the plain guess of 16 Nov stays."""
    dirty = {"version_id": "v", "ran": True, "formula_errors": [{"cell": "A1"}], "addin_gap_errors": []}
    r = compute_readiness({"findings": [dict(READY)]}, [], dirty, "v", loop_built=True)
    assert r["state"] == "unverified" and "may come from the Mind loop" in r["recalc"]["loop_note"]


def test_night_chain_goes_on_as_start_does(tmp_path, monkeypatch):
    """scripts/night_chain.py passes the Recalculate step on the verdict, as Start does: Palermo's 667
    errors for the loop's first value no longer fail the chain; a real new error still does."""
    import importlib.util

    spec = importlib.util.spec_from_file_location("night_chain_lb", Path(server.PACKAGE_ROOT) / "scripts" / "night_chain.py")
    chain = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(chain)
    s = _session(tmp_path, PALERMO)
    _wire(monkeypatch, s, FakeExcel(s, {"v1": [], "v2": []}))
    out = _recalc_current(s, [_err("Decrements", f"M{r}") for r in range(704, 714)])
    assert out["new_errors"] == 10 and chain.recalc_ok(out) is True
    # (review LBr: it said `recalc_ok({**out, "loop_check": None}) is recalc_ok(out)`, true whatever the
    # code does) a real error after the loop -- the version before the loop errs there too -- fails it
    s3 = _session(tmp_path / "c", PALERMO)
    _wire(monkeypatch, s3, FakeExcel(s3, {"v1": [], "v2": [_err("Results", "B5", "#DIV/0!")]}))
    out = _recalc_current(s3, [_err("Decrements", f"M{r}") for r in range(704, 714)] + [_err("Results", "B5", "#DIV/0!")])
    assert out["loop_check"]["from_loop"] == 10 and out["loop_check"]["also_without"] == 1 and chain.recalc_ok(out) is False
    s2 = _session(tmp_path / "b", [("upload", []), ("prep", [])])
    _wire(monkeypatch, s2, FakeExcel(s2, {"v1": []}))
    assert chain.recalc_ok(_recalc_current(s2, [_err("Results", "B5", "#DIV/0!")])) is False
