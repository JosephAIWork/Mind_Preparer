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


def apply_record(outcome: dict[str, Any]) -> dict[str, Any]:
    """One line of the applied-changes gauge: what an Apply really wrote into
    the file, by level, and how many of its repairs resolved their problem."""
    by_level = {lvl: {"applied": 0, "failed": 0} for lvl in ("blocking", "optional")}
    verdicts: dict[str, int] = {}
    for a in outcome["actions"]:
        lvl = by_level.setdefault(a["level"], {"applied": 0, "failed": 0})
        lvl["applied"] += a["applied"]
        lvl["failed"] += a["failed"]
        verdicts[a["verdict"]] = verdicts.get(a["verdict"], 0) + 1
    return {
        "version_id": outcome["version_id"],
        "previous_version_id": outcome["previous_version_id"],
        "sent": outcome["sent"],
        "applied": outcome["applied"],
        "failed": outcome["failed"],
        "blocking_applied": by_level["blocking"]["applied"],
        "blocking_failed": by_level["blocking"]["failed"],
        "optional_applied": by_level["optional"]["applied"],
        "optional_failed": by_level["optional"]["failed"],
        "verdicts": verdicts,
        "resolved_blocking": list(outcome["resolved_blocking"]),
    }


def prep_progress(history: list[dict[str, Any]], applies: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """How the prep work evolves over the session's versions: one entry per
    analysis with the operation counts, whether the last rounds stalled (the
    same operations coming back -- Prep cannot resolve them), and one entry
    per Apply with the changes it really wrote (`applies`, `written`)."""
    applies = list(applies or [])
    written = {k: sum(a.get(k, 0) for a in applies) for k in ("sent", "applied", "failed", "blocking_applied", "blocking_failed", "optional_applied", "optional_failed")}
    entries = [{k: h.get(k) for k in ("version_id", "blocking_ops", "optional_ops", "total_ops", "blocking_findings", "open_findings", "status_counts", "by_action")} for h in history]
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
        "applies": applies,
        "written": written,
        "stalled": stalled,
        "repeating_actions": repeating,
        "message": ("the planned repairs are not going down over the last three versions -- Prep cannot resolve what is left; read what each action left for review and fix it by hand or with the assistant" if stalled else None),
    }


def plan_counts(plan: list[dict[str, Any]]) -> dict[str, int]:
    return {a["id"]: int(a.get("count", 0)) for a in plan}


# --- what an Apply did (1.7.3) --------------------------------------------------------------
STATUS_RANK = {"PASS": 0, "WARNING": 1, "REQUIRES_USER_INPUT": 2, "NOT_SUPPORTED": 3, "ERROR": 4}
MANUAL_SOURCES = {"assistant": "Assistant fix", "formula_replacement": "Formula fix"}


def _plural(n: int, word: str) -> str:
    return f"{n:,} {word}{'' if n == 1 else 's'}"


def _some(names: list[str], most: int = 6) -> str:
    return ", ".join(names[:most]) + (f" (+{len(names) - most} more)" if len(names) > most else "")


def _rule_move(rule_id: str, before: dict[str, dict[str, Any]], after: dict[str, dict[str, Any]] | None) -> dict[str, Any]:
    b = before.get(rule_id)
    a = after.get(rule_id) if after is not None else None
    return {
        "rule_id": rule_id,
        "from": b.get("status") if b else None,
        # a rule the new analysis no longer reports has nothing left to say: it passes
        "to": None if after is None else (a.get("status") if a else "PASS"),
        "sites_from": _site_count(b) if b else None,
        "sites_to": None if after is None else (_site_count(a) if a else 0),
        "blocking_before": bool(b) and level_of(b) == "blocking",
        "blocking_after": bool(a) and level_of(a) == "blocking",
    }


def _verdict(applied: int, rules: list[dict[str, Any]], planned_before: int | None, planned_after: int | None, reanalyzed: bool) -> str:
    """resolved: every rule the changes serve passes and nothing more is planned.
    partial: something moved (a better status, fewer sites, fewer repairs planned).
    unchanged: the changes were written and the new analysis says the same as before.
    failed: nothing could be written. written: written, no analysis to compare with."""
    if applied == 0:
        return "failed"
    if not reanalyzed:
        return "written"
    still_open = [r for r in rules if r["to"] not in (None, "PASS")]
    if not still_open and not planned_after:
        return "resolved" if rules or planned_before else "written"
    moved = any(
        STATUS_RANK.get(r["to"] or "PASS", 0) < STATUS_RANK.get(r["from"] or "PASS", 0)
        or (r["sites_from"] is not None and r["sites_to"] is not None and r["sites_to"] < r["sites_from"])
        for r in rules
    )
    fewer = planned_before is not None and planned_after is not None and planned_after < planned_before
    return "partial" if moved or fewer else "unchanged"


def _action_summary(a: dict[str, Any]) -> tuple[str, str]:
    """(the whole sentence, its conclusion alone): what was written and how the
    rules moved, then what is left and who acts."""
    written = f"{_plural(a['applied'], 'change')} written"
    if a["failed"]:
        written += f", {a['failed']:,} refused"
    moves = []
    for r in a["rules"]:
        if r["to"] is None:
            continue
        sites = f" ({r['sites_from']:,} -> {r['sites_to']:,} cells)" if r["sites_from"] is not None and r["sites_to"] is not None and r["sites_from"] != r["sites_to"] else ""
        worse = " (worse than before)" if STATUS_RANK.get(r["to"], 0) > STATUS_RANK.get(r["from"] or "PASS", 0) else ""
        moves.append(f"{r['rule_id']} {r['from'] or 'not reported'} -> {r['to']}{sites}{worse}" if r["from"] != r["to"] else f"{r['rule_id']} still {r['to']}{sites}")
    left = []
    if a["planned_after"]:
        left.append(f"{_plural(a['planned_after'], 'more repair')} planned by the new analysis")
    if a["skipped_after"]:
        left.append(f"{a['skipped_after']:,} left for review by hand")
    verdict = a["verdict"]
    if verdict == "failed":
        end = f"None of the {_plural(a['sent'], 'change')} could be written" + (f": {a['errors'][0]}" if a["errors"] else "") + "."
        return end, end
    text = written + ". " + ("; ".join(moves) + ". " if moves else "")
    if verdict == "written":
        end = "Not re-analyzed: what it resolved is not known yet."
    elif verdict == "resolved":
        end = "Resolved: nothing left for this repair."
    elif verdict == "unchanged":
        end = "The new analysis says the same as before" + (" and plans as many repairs again" if a["planned_after"] else "") + ": applying it again will not resolve it -- fix it by hand or with the assistant."
    else:
        end = "Still open: " + "; ".join(left) + "." if left else "Still open: nothing more Prep can do -- fix the rest by hand or with the assistant."
    return text + end, end


def apply_outcome(
    operations: list[dict[str, Any]],
    result: dict[str, Any],
    before_plan: list[dict[str, Any]],
    before_report: dict[str, Any] | None,
    after_plan: list[dict[str, Any]] | None,
    after_report: dict[str, Any] | None,
    version_id: str,
    previous_version_id: str | None,
) -> dict[str, Any]:
    """What an Apply did, repair by repair: how many changes were written, what
    the rule said before and says now, what is still planned or left by hand,
    and which repairs the new analysis plans that the previous one did not.
    Only facts from the two analyses and the executor's result -- nothing is
    judged here. Without a re-analysis (`after_*` None) only the writes are known."""
    reanalyzed = after_report is not None and after_plan is not None
    before = {f["rule_id"]: f for f in (before_report or {}).get("findings", [])}
    after = {f["rule_id"]: f for f in after_report.get("findings", [])} if reanalyzed else None
    plan_before = {a["id"]: a for a in before_plan}
    plan_after = {a["id"]: a for a in (after_plan or [])}

    def key(o: dict[str, Any]) -> tuple[str, str]:
        aid = str(o.get("action_id") or "")
        return (aid, "") if aid in plan_before else (aid, str(o.get("rule_id") or ""))

    groups: dict[tuple[str, str], dict[str, list[dict[str, Any]]]] = {}
    for name, ops in (("sent", operations), ("applied", result.get("applied", [])), ("failed", result.get("failed", []))):
        for o in ops:
            groups.setdefault(key(o), {"sent": [], "applied": [], "failed": []})[name].append(o)

    actions = []
    for (aid, rid), g in groups.items():
        known = plan_before.get(aid) if not rid else None
        rule_ids = list(dict.fromkeys(str(o.get("rule_id")) for o in g["sent"] if str(o.get("rule_id")) in before))
        if not rule_ids and known:
            rule_ids = [r for r in known["rule_ids"] if r in before]
        rules = [_rule_move(r, before, after) for r in rule_ids]
        now = plan_after.get(aid) if known and reanalyzed else None
        entry = {
            "id": aid if known else f"{aid}:{rid}",
            "title": known["title"] if known else f"{MANUAL_SOURCES.get(aid, 'Manual fix')}{' for ' + rid if rid in before else ''}",
            "level": known["level"] if known else ("blocking" if any(r["blocking_before"] for r in rules) else "optional"),
            "sent": len(g["sent"]),
            "applied": len(g["applied"]),
            "failed": len(g["failed"]),
            "planned_before": int(known["count"]) if known else None,
            "planned_after": int(now["count"]) if now else (0 if known and reanalyzed else None),
            "skipped_after": len(now["skipped"]) if now else (0 if known and reanalyzed else None),
            "rules": rules,
            "errors": list(dict.fromkeys(str(o.get("error")) for o in g["failed"] if o.get("error")))[:5],
        }
        entry["verdict"] = _verdict(entry["applied"], rules, entry["planned_before"], entry["planned_after"], reanalyzed)
        entry["summary"], entry["next"] = _action_summary(entry)
        actions.append(entry)
    actions.sort(key=lambda a: (a["level"] != "blocking", -a["sent"]))

    applied_ids = {aid for (aid, rid) in groups if not rid}
    appeared = [
        {"id": a["id"], "title": a["title"], "level": a["level"], "rule_ids": a["rule_ids"], "planned_before": int(plan_before.get(a["id"], {}).get("count", 0)), "planned_after": int(a["count"])}
        for a in (after_plan or [])
        if a["id"] not in applied_ids and int(a["count"]) > int(plan_before.get(a["id"], {}).get("count", 0))
    ]
    regressed = [
        {"rule_id": rid, "from": before[rid]["status"], "to": f["status"]}
        for rid, f in (after or {}).items()
        if rid in before and STATUS_RANK.get(f["status"], 0) > STATUS_RANK.get(before[rid]["status"], 0)
    ]
    touched = {r["rule_id"] for a in actions for r in a["rules"]}
    blocking_before = [rid for rid, f in before.items() if level_of(f) == "blocking"]
    remaining = []
    for rid, f in (after or {}).items():
        if level_of(f) == "blocking":
            route, action_id = _fix_route(rid, after_plan or [])
            remaining.append({"rule_id": rid, "fix": route, "action_id": action_id, "sites": _site_count(f), "touched": rid in touched, "new": rid not in blocking_before})
    resolved = [rid for rid in blocking_before if reanalyzed and rid not in {r["rule_id"] for r in remaining}]

    n_applied, n_failed = len(result.get("applied", [])), len(result.get("failed", []))
    headline = f"{n_applied:,} of {_plural(len(operations), 'change')} written"
    if not reanalyzed:
        headline += "; the new file was not re-analyzed."
    elif not blocking_before and not remaining:
        headline += ". No blocking problem before or after."
    else:
        headline += f". Blocking problems: {len(blocking_before)} before, {len(remaining)} now"
        if resolved:
            headline += f" -- resolved: {_some(resolved)}"
        appeared_blocking = [r["rule_id"] for r in remaining if r["new"]]
        if appeared_blocking:
            headline += f" -- new in this version: {_some(appeared_blocking)}"
        if remaining:
            by_hand = [r["rule_id"] for r in remaining if r["fix"] != "prep"]
            via_prep = [r["rule_id"] for r in remaining if r["fix"] == "prep"]
            parts = ([f"{_some(via_prep)} (Prep has a repair: apply it)"] if via_prep else []) + ([f"{_some(by_hand)} (no automatic repair: by hand or with the assistant)"] if by_hand else [])
            headline += f" -- left: {'; '.join(parts)}"
        headline += "."
    return {
        "version_id": version_id,
        "previous_version_id": previous_version_id,
        "reanalyzed": reanalyzed,
        "sent": len(operations),
        "applied": n_applied,
        "failed": n_failed,
        "headline": headline,
        "actions": actions,
        "appeared": appeared,
        "regressed": regressed,
        "blocking_before": len(blocking_before),
        "blocking_after": len(remaining) if reanalyzed else None,
        "resolved_blocking": resolved,
        "remaining_blocking": remaining,
    }
