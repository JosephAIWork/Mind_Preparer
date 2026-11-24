"""Readiness  aggregation: turns  a list of Finding dicts into a ValidationReport-
shaped dict. Status is derived purely from the findings' statuses (worst wins),
per instructions/reporting.md: "the application derives status from enums and
structured fields, never prose. The  LLM never computes the final readiness state."
"""
from __future__ import annotations

import uuid
from typing import Any

from .models import Finding, Status, aggregate_status


def build_validation_report(findings: list[dict[str, Any]]) -> dict[str, Any]:
    findings = [Finding.model_validate(f).model_dump(mode="json") for f in findings]
    statuses = [Status(f["status"]) for f in findings]
    overall = aggregate_status(statuses)

    counts: dict[str, int] = {}
    for s in statuses:
        counts[s.value] = counts.get(s.value, 0) + 1

    return {
        "schema_version": "1.0.0",
        "report_id": str(uuid.uuid4()),
        "status": overall.value,
        "findings": findings,
        "summary": {
            "finding_count": len(findings),
            "status_counts": counts,
            "not_supported_rule_ids": [f["rule_id"] for f in findings if f["status"] == "NOT_SUPPORTED"],
            "blocking_rule_ids": [
                f["rule_id"] for f in findings if f["status"] in ("ERROR", "REQUIRES_USER_INPUT", "NOT_SUPPORTED")
            ],
        },
    }
