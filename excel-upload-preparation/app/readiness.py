"""The one answer the user needs (1.7.2): is this workbook acceptable by Mind,
and if not, what exactly stands in the way?

The engine already knows everything: each finding carries the rule's
priority (REQUIRED / RECOMMENDED / INFORMATIONAL), the prep plan says which
findings have an automatic fix, and the recalculation adapter says whether
Excel evaluated the file cleanly. This module only assembles those facts into
one verdict with three states -- it never judges a workbook on its own:

  blocked     one or more REQUIRED rules are in ERROR (or wait for user input):
              Mind's converter refuses the file or computes it wrong.
  unverified  nothing blocks, but the current version was not recalculated
              in Excel (or the recalculation found genuine formula errors).
  ready       nothing blocks and a real recalculation of this version is clean
              (add-in gaps -- #NAME? on MM_ functions when MMForExcel is
              missing on this machine -- are not workbook defects).

Everything else (WARNING findings, ERRORs from non-required rules) is
"optional": Mind reads the file as it is; the change improves it.
"""
from __future__ import annotations

from typing import Any

BLOCKING_STATUSES = ("ERROR", "REQUIRES_USER_INPUT")
RECALC_RULE = "READY-001"


def level_of(finding: dict[str, Any]) -> str:
    """'blocking' for a REQUIRED rule that fails, 'optional' for anything else
    that is not PASS, 'pass' otherwise. The recalculation rule is 'verify'."""
    if finding["rule_id"] == RECALC_RULE:
        return "verify"
    status = finding.get("status")
    if status == "PASS":
        return "pass"
    if status in BLOCKING_STATUSES and finding.get("priority") == "REQUIRED":
        return "blocking"
    return "optional"


def _fix_route(rule_id: str, plan: list[dict[str, Any]]) -> tuple[str, str | None]:
    """How a finding can be fixed: ('prep', action_id) when a plan action holds
    operations for the rule, ('prep_skipped', action_id) when the action
    exists but left every site for review, else ('assistant', None)."""
    for a in plan:
        if rule_id in a.get("rule_ids", []) and a.get("count", 0) > 0:
            return "prep", a["id"]
    for a in plan:
        if rule_id in a.get("rule_ids", []) and a.get("skipped"):
            return "prep_skipped", a["id"]
    return "assistant", None


def recalc_state(recalc: dict[str, Any] | None, version_id: str) -> dict[str, Any]:
    """What the last recalculation says about *this* version."""
    if not recalc or recalc.get("version_id") != version_id:
        return {"ran": False, "clean": False, "version_id": recalc.get("version_id") if recalc else None, "formula_errors": 0, "addin_gap_errors": 0, "message": "not recalculated for this version"}
    ran = bool(recalc.get("ran", recalc.get("status") != "NOT_SUPPORTED"))
    errors = recalc.get("formula_errors") or []
    gaps = recalc.get("addin_gap_errors") or []
    # An error cell that was already an error in the original upload is the
    # model's own (a #DIV/0! on an empty year, a #N/A lookup): it did not come
    # from the preparation and Mind takes the file with it. Only *new* errors
    # keep the verdict from turning green.
    preexisting = int(recalc.get("preexisting_errors") or 0) if "preexisting_errors" in recalc else 0
    new_errors = int(recalc["new_errors"]) if "new_errors" in recalc else len(errors)
    if ran:
        if not errors:
            message = "full Excel recalculation, no formula error"
        elif new_errors == 0:
            message = f"full Excel recalculation: {len(errors)} formula error(s), all already in the original workbook"
        else:
            message = f"{new_errors} new formula error(s) after a full Excel recalculation" + (f" ({preexisting} more were already in the original)" if preexisting else "")
    else:
        message = recalc.get("message")
    return {
        "ran": ran,
        "clean": ran and new_errors == 0,
        "version_id": version_id,
        "formula_errors": len(errors),
        "new_errors": new_errors,
        "preexisting_errors": preexisting,
        "addin_gap_errors": len(gaps),
        "message": message,
    }


def compute_readiness(report: dict[str, Any], plan: list[dict[str, Any]], recalc: dict[str, Any] | None, version_id: str) -> dict[str, Any]:
    findings = report.get("findings", [])
    blocking, optional = [], []
    for f in findings:
        lvl = level_of(f)
        if lvl == "blocking":
            route, action_id = _fix_route(f["rule_id"], plan)
            blocking.append({
                "rule_id": f["rule_id"],
                "status": f["status"],
                "message": f.get("message", ""),
                "location": f.get("location") or {},
                "fix": route,
                "action_id": action_id,
                "sites": _site_count(f),
            })
        elif lvl == "optional":
            optional.append({"rule_id": f["rule_id"], "status": f["status"]})
    rc = recalc_state(recalc, version_id)
    if blocking:
        state = "blocked"
        headline = f"Not acceptable by Mind yet: {len(blocking)} blocking problem(s)"
    elif not rc["ran"]:
        state = "unverified"
        headline = "Acceptable by Mind, not verified: run the Excel recalculation"
    elif not rc["clean"]:
        state = "unverified"
        headline = f"Acceptable by Mind, but the recalculation found {rc['new_errors']} new formula error(s)"
    else:
        state = "ready"
        headline = "Acceptable by Mind: no blocking problem and a clean Excel recalculation" + (f" ({rc['preexisting_errors']} error cell(s) already in the original)" if rc["preexisting_errors"] else "")
    prep_blocking = sum(a["count"] for a in plan if a.get("level") == "blocking")
    prep_optional = sum(a["count"] for a in plan if a.get("level") == "optional")
    manual = [b for b in blocking if b["fix"] != "prep"]
    next_step = _next_step(state, blocking, rc)
    return {
        "state": state,
        "headline": headline,
        "version_id": version_id,
        "blocking": blocking,
        "blocking_count": len(blocking),
        "manual_blocking_count": len(manual),
        "optional_count": len(optional),
        "prep": {"blocking_ops": prep_blocking, "optional_ops": prep_optional},
        "recalc": rc,
        "next_step": next_step,
    }


def _site_count(f: dict[str, Any]) -> int | None:
    obs = f.get("observed")
    if isinstance(obs, dict):
        for key in ("sites", "cells", "call_sites", "problems"):
            v = obs.get(key)
            if isinstance(v, list):
                return len(v)
    return None


def _next_step(state: str, blocking: list[dict[str, Any]], rc: dict[str, Any]) -> dict[str, Any]:
    if state == "blocked":
        # Only a Prep action that targets a *blocking finding* counts here -- a
        # blocking-level action with work left for an optional finding does not.
        via_prep = [b for b in blocking if b["fix"] == "prep"]
        if via_prep:
            return {"screen": "prep", "label": "Apply the blocking repairs in Prep", "detail": ", ".join(b["rule_id"] for b in via_prep)}
        first = next((b for b in blocking if b["fix"] != "prep"), blocking[0])
        return {"screen": "findings", "label": f"Fix {first['rule_id']} with the assistant", "detail": "no automatic repair for what is left", "rule_id": first["rule_id"]}
    if state == "unverified":
        if rc["ran"]:
            return {"screen": "recalculate", "label": "Review the new formula errors", "detail": f"{rc['new_errors']} cell(s) not in error in the original"}
        return {"screen": "recalculate", "label": "Recalculate in Excel", "detail": "the only way the app can call the file ready"}
    return {"screen": "reports", "label": "Download or send to Mind", "detail": None}


def prep_progress(history: list[dict[str, Any]]) -> dict[str, Any]:
    """How the prep work evolves over the session's versions: one entry per
    analysis with the operation counts, and whether the last rounds stalled
    (the same operations coming back -- Prep cannot resolve them)."""
    entries = [{k: h.get(k) for k in ("version_id", "blocking_ops", "optional_ops", "total_ops", "blocking_findings", "by_action")} for h in history]
    # Stalled: three analyses in a row with work planned and no net decrease.
    # (Exact repetition is not required -- a row insert that never resolves
    # shifts its row number every round, yet the count never goes down.)
    stalled = False
    if len(entries) >= 3:
        totals = [e["total_ops"] or 0 for e in entries[-3:]]
        stalled = all(t > 0 for t in totals) and totals[-1] >= totals[0]
    repeating = []
    if stalled:
        a, b = entries[-2].get("by_action") or {}, entries[-1].get("by_action") or {}
        repeating = sorted(k for k in b if b[k] > 0 and a.get(k) == b[k])
    return {
        "entries": entries,
        "stalled": stalled,
        "repeating_actions": repeating,
        "message": ("the planned repairs are not going down over the last three versions -- Prep cannot resolve what is left; read what each action left for review and fix it by hand or with the assistant" if stalled else None),
    }


def plan_counts(plan: list[dict[str, Any]]) -> dict[str, int]:
    return {a["id"]: int(a.get("count", 0)) for a in plan}
