"""Domain models mirroring schemas/*.json.

Every top-level model here corresponds 1:1 to one schema file and uses
``extra="forbid"`` because every schema sets ``additionalProperties: false``.
Fields typed loosely as ``dict``/``list`` mirror schema properties that are
themselves only declared as ``type: object`` / ``type: array`` with no
further nested schema -- tightening them further would be inventing
structure the package doesn't document (SKILL.md rule #4: never invent
requirements).
"""
from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class Status(str, Enum):
    PASS = "PASS"
    WARNING = "WARNING"
    ERROR = "ERROR"
    REQUIRES_USER_INPUT = "REQUIRES_USER_INPUT"
    NOT_SUPPORTED = "NOT_SUPPORTED"


# Worst-first precedence per validation/readiness-policy.md.
STATUS_PRECEDENCE: list[Status] = [
    Status.NOT_SUPPORTED,
    Status.ERROR,
    Status.REQUIRES_USER_INPUT,
    Status.WARNING,
    Status.PASS,
]


def aggregate_status(statuses: list[Status]) -> Status:
    """Worst-status-wins aggregation, per validation/readiness-policy.md."""
    if not statuses:
        return Status.PASS
    for candidate in STATUS_PRECEDENCE:
        if candidate in statuses:
            return candidate
    return Status.PASS


class Mode(str, Enum):
    PLAN_MODE = "PLAN_MODE"
    PREP_MIND_LOOPS = "PREP_MIND_LOOPS"
    FIX_INCOMPATIBLE_FORMULAS = "FIX_INCOMPATIBLE_FORMULAS"
    STRUCTURE_FIX = "STRUCTURE_FIX"


class EvidenceLabel(str, Enum):
    DOCUMENTED_RULE = "DOCUMENTED_RULE"
    DETERMINISTIC_FINDING = "DETERMINISTIC_FINDING"
    USER_PROVIDED = "USER_PROVIDED"
    INFERENCE = "INFERENCE"
    RECOMMENDATION = "RECOMMENDATION"


class ChangeSetState(str, Enum):
    PROPOSED = "PROPOSED"
    APPROVED = "APPROVED"
    APPLIED = "APPLIED"
    REJECTED = "REJECTED"
    FAILED = "FAILED"


class QuestionStatus(str, Enum):
    OPEN = "OPEN"
    ANSWERED = "ANSWERED"
    CANCELLED = "CANCELLED"


class PreservationClass(str, Enum):
    AUTHORIZED_CHANGE = "AUTHORIZED_CHANGE"
    EXPECTED_RECALCULATION_CHANGE = "EXPECTED_RECALCULATION_CHANGE"
    METADATA_CHANGE = "METADATA_CHANGE"
    UNAUTHORIZED_CHANGE = "UNAUTHORIZED_CHANGE"
    UNVERIFIABLE_CHANGE = "UNVERIFIABLE_CHANGE"


class Finding(BaseModel):
    """One validator result. Matches validation/validation-policy.md's required fields."""

    model_config = ConfigDict(extra="forbid")

    rule_id: str
    status: Status
    severity: str | None = None
    confidence: str
    evidence: EvidenceLabel
    location: dict[str, Any] = Field(default_factory=dict)
    observed: Any = None
    expected: Any = None
    message: str
    readiness_impact: Status
    correction_available: bool = False
    source: dict[str, Any] = Field(default_factory=dict)
    # 1.7.2: the rule's own priority (REQUIRED / RECOMMENDED / INFORMATIONAL),
    # so the UI can tell a change Mind *needs* from one it merely benefits from.
    priority: str | None = None


class WorkbookAnalysis(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: str
    analysis_id: str
    source: dict[str, Any]
    workbooks: list[dict[str, Any]]
    features: dict[str, Any]
    risks: list[dict[str, Any]]


class ValidationReport(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: str
    report_id: str
    status: Status
    findings: list[dict[str, Any]]
    summary: dict[str, Any]


class ChangeSet(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: str
    change_set_id: str
    source_sha256: str
    state: ChangeSetState
    changes: list[dict[str, Any]]


class UserQuestion(BaseModel):
    model_config = ConfigDict(extra="forbid")

    question_id: str
    status: QuestionStatus
    blocking: bool
    rule_ids: list[str]
    location: dict[str, Any] | None = None
    question: str
    options: list[Any]
    default_if_unanswered: str


class ProcessingResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: str
    request_id: str
    mode: Mode
    status: Status
    upload_readiness: dict[str, Any]
    source: dict[str, Any]
    output: dict[str, Any] | None = None
    questions: list[dict[str, Any]] = Field(default_factory=list)
    warnings: list[dict[str, Any]] = Field(default_factory=list)
    unsupported_features: list[dict[str, Any]] = Field(default_factory=list)
