"""Shared ANALYZE -> VALIDATE -> ASK -> REPORT runner for all four modes.
TRANSFORM/VERIFY are intentionally absent (no change-set application or
recalculation adapter here: Prep, the fixer and app/recalc.py do those).
"""
from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any, Callable

from ..formula_utils import formula_memo
from ..inventory import build_analysis
from ..models import Mode, ProcessingResult, Status
from ..report import build_validation_report
from ..rules_engine import RulesEngine
from ..validators import run_rule


def _build_question(finding: dict[str, Any]) -> dict[str, Any]:
    return {
        "question_id": f"Q-{finding['rule_id']}",
        "status": "OPEN",
        "blocking": True,
        "rule_ids": [finding["rule_id"]],
        "location": finding.get("location") or {},
        "question": finding["message"],
        "options": [
            {"id": "acknowledge", "label": "Acknowledge and continue without a change"},
            {"id": "provide_target_version", "label": "Provide the missing information referenced above"},
        ],
        "default_if_unanswered": "REQUIRES_USER_INPUT",
    }


def run_mode(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """_run_mode with every formula parsed once for all the rules (formula_memo)."""
    with formula_memo():
        return _run_mode(*args, **kwargs)


def _run_mode(
    mode: Mode,
    source_path: Path,
    work_dir: Path,
    config: dict[str, Any],
    categories: list[str] | None,
    engine: RulesEngine | None = None,
    ignore_sheets: list[str] | None = None,
    progress: Callable[..., None] | None = None,
) -> dict[str, Any]:
    """categories=None means 'all active rules' (used by PLAN_MODE).
    `ignore_sheets`: sheets the user chose not to scan (large workbooks).
    `progress(stage, message, fraction, **facts)`: status callback -- the
    inventory stages first, then one call per rule, then 'report'."""
    engine = engine or RulesEngine()
    request_id = str(uuid.uuid4())

    analysis = build_analysis(source_path, work_dir, analysis_id=request_id, ignore_sheets=ignore_sheets, progress=progress)

    if categories is None:
        rules = engine.active_rules()
    else:
        rules = [r for r in engine.active_rules() if r["category"] in categories]

    findings = []
    for i, rule in enumerate(rules):
        if progress is not None:
            progress("rules", f"Checking rule {i + 1}/{len(rules)}: {rule['id']}", i / max(1, len(rules)), rule_id=rule["id"])
        findings.append(run_rule(rule, analysis, config))
    if progress is not None:
        progress("report", "Building the report", None)
    validation_report = build_validation_report(findings)

    questions = [
        _build_question(f) for f in findings if f["status"] == "REQUIRES_USER_INPUT"
    ]
    warnings = [f for f in findings if f["status"] == "WARNING"]
    unsupported = [f for f in findings if f["status"] == "NOT_SUPPORTED"]

    overall_status = Status(validation_report["status"])

    result = ProcessingResult(
        schema_version="1.0.0",
        request_id=request_id,
        mode=mode,
        status=overall_status,
        upload_readiness={
            "ready": overall_status == Status.PASS,
            "validation_status": validation_report["status"],
            "recalculation_performed": False,
            "reason": (
                "PASS needs a real recalculation in Excel: until one has run, "
                "READY-001 reports NOT_SUPPORTED."
                if overall_status != Status.PASS
                else None
            ),
        },
        source=analysis["source"],
        output=None,
        questions=questions,
        warnings=warnings,
        unsupported_features=unsupported,
    )
    return {
        "processing_result": result.model_dump(mode="json"),
        "workbook_analysis": analysis,
        "validation_report": validation_report,
        "rule_load_errors": [e.__dict__ for e in engine.errors],
    }
