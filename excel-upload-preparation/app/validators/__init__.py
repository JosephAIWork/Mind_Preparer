"""Validator orchestration.

Each `validators.<category>` module exposes plain functions named after a
rule's `validation.implementation` (e.g. `validators.loop.mm_loop_inventory`).
A validator function takes `(rule, analysis, config)` and returns a partial
Finding dict (at least `status` and `message`); `run_rule` fills in the rest
from the rule definition and returns a dict ready for `app.models.Finding`.

If a rule's implementation path doesn't resolve to a real function (not
built in this MVP pass), `run_rule` returns the rule's own `on_unevaluable`
status rather than silently skipping it -- SKILL.md non-negotiable rule #8.
"""
from __future__ import annotations

from typing import Any

from ..rules_engine import RulesEngine


def run_rule(rule: dict[str, Any], analysis: dict[str, Any], config: dict[str, Any]) -> dict[str, Any]:
    implementation = rule["validation"]["implementation"]
    fn = RulesEngine.resolve_validator(implementation)

    base = {
        "rule_id": rule["id"],
        "confidence": rule["confidence"],
        "source": rule.get("source", {}),
        "correction_available": bool(rule.get("correction", {}).get("automatic") is True),
        "priority": rule.get("priority"),
    }

    if fn is None:
        base.update(
            status=rule["on_unevaluable"],
            severity=rule.get("severity_source", rule["priority"]),
            evidence="DOCUMENTED_RULE",
            location={},
            observed=None,
            expected=None,
            message=f"Not implemented in this MVP: {implementation}. Rule text: {rule['description']}",
            readiness_impact=rule["on_unevaluable"],
        )
        return base

    result = fn(rule, analysis, config) or {}
    status = result.get("status", rule["on_unevaluable"])
    base.update(
        status=status,
        severity=result.get("severity", rule.get("severity_source", rule["priority"])),
        evidence=result.get("evidence", "DETERMINISTIC_FINDING"),
        location=result.get("location", {}),
        observed=result.get("observed"),
        expected=result.get("expected"),
        message=result.get("message", rule["description"]),
        readiness_impact=result.get("readiness_impact", status),
    )
    return base
