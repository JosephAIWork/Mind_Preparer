"""FILE-*  validators: FILE-001  ensure_source_immutable, FILE-002 allowed_output_format."""
from __future__ import annotations

from typing import Any


def ensure_source_immutable(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    source = analysis["source"]
    ok = bool(source.get("sha256")) and source.get("copy_path") != source.get("path")
    return {
        "status": "PASS" if ok else "ERROR",
        "evidence": "DETERMINISTIC_FINDING",
        "observed": {"source_sha256": source.get("sha256"), "copy_path": source.get("copy_path")},
        "message": (
            "Analysis ran against an immutable copy; source hashed before copy."
            if ok
            else "Could not establish an immutable copy of the source file."
        ),
    }


def allowed_output_format(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    allowed = set(config.get("allowed_output_formats", ["xlsx", "xlsm"]))
    wb = analysis["workbooks"][0]
    file_type = wb["file_type"]
    ok = file_type in allowed
    return {
        "status": "PASS" if ok else "ERROR",
        "evidence": "DETERMINISTIC_FINDING",
        "observed": {"file_type": file_type},
        "expected": {"allowed_output_formats": sorted(allowed)},
        "message": (
            f"'{file_type}' is an allowed output format."
            if ok
            else f"'{file_type}' is not in the allowed output formats {sorted(allowed)}."
        ),
    }
