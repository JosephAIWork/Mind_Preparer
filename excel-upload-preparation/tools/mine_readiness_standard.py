#!/usr/bin/env python3
"""One-time  deterministic extraction  of the Mind Readiness Standard into rules/*.yaml.

Source: 01_Mind_Readiness_Standard.md (sibling doc to this package). That doc
uses a fully regular structure per rule:

    ## <ID> <Title>
    [intro text]
    PASS IF: / FAIL IF: / Allowed: /  Review ... / Inventory ... (criteria)
    SEVERITY:
    <LEVEL>
    ---

This script is a plain regex  parser (no LLM involved), so re-running  it
against an updated source doc reproduces the  same rules/*.yaml content.

Two existing hand-authored rule files predate this doc and  overlap with it:
  - loop-rules.yaml (LOOP-001..003) duplicates the  doc's own LOOP-001..004
    (section 7). The doc's IDs become canonical; the one rule with no doc
    equivalent (MM_LOOPLABELS size matching, old LOOP-003) is kept, renamed
    to LBL-001.
  - structure-rules.yaml (GRID-001/004), format-rules.yaml (FORMAT-001),
    and formula-rules.yaml (FORMULA-001/002) use different ID prefixes than
    the doc (GRID/FORMAT/FORMULA vs STR/FMT/FRM) and are complementary
    policy-level rules, not duplicates -- they are preserved as-is and the
    mined rules are appended alongside them.
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import yaml

SECTION_RE = re.compile(r"^## ([A-Z]+-\d{3}) (.+)$")

# rule  ID prefix  -> (target rules file, category name used in validator dotted paths)
CATEGORY_TARGETS: dict[str, tuple[str, str]] = {
    "STR": ("structure-rules.yaml", "structure"),
    "FMT": ("format-rules.yaml", "format"),
    "FRM": ("formula-rules.yaml", "formula"),
    "MMX": ("environment-rules.yaml", "environment"),
    "LOOP": ("loop-rules.yaml", "loop"),
    "RES": ("loop-rules.yaml", "loop"),
    "RZS": ("loop-rules.yaml", "loop"),
    "LKP": ("lookup-rules.yaml", "lookup"),
    "INP": ("io-rules.yaml", "io"),
    "EXP": ("io-rules.yaml", "io"),
    "PRJ": ("project-rules.yaml", "project"),
    "INS": ("multi-workbook-rules.yaml", "multi_workbook"),
    "LNK": ("multi-workbook-rules.yaml", "multi_workbook"),
    "PAR": ("parameter-rules.yaml", "parameter"),
    "CAL": ("iteration-rules.yaml", "iteration"),
    "DBG": ("performance-rules.yaml", "performance"),
    "RSK": ("risk-rules.yaml", "risk"),
}

# rule_id  -> validator  function name, for the MVP-implemented subset.
# Everything else in the mined set gets a registered but NOT_SUPPORTED validator.
IMPLEMENTED: dict[str, str] = {
    "STR-001": "grid_detection_heuristic",
    "STR-004": "grid_naming_heuristic",
    "STR-007": "hidden_sheets",
    "FMT-004": "merged_cells",
    "FMT-005": "locked_cells",
    "FMT-007": "comments",
    "FRM-002": "unsupported_functions",
    "FRM-004": "vba_presence",
    "LOOP-001": "mm_loop_inventory",
    "LOOP-002": "loop_naming_consistency",
    "LOOP-003": "repeated_definition_range_lengths",
    "RSK-003": "hyperlinks",
    "RSK-004": "let_function_usage",
    "RSK-005": "external_named_ranges",
}

SEVERITY_TO_PRIORITY = {
    "BLOCKER": "REQUIRED",
    "HIGH": "REQUIRED",
    "MEDIUM": "RECOMMENDED",
    "LOW": "INFORMATIONAL",
    "INFO": "INFORMATIONAL",
}

# Cross-workbook  categories: config/default.yaml  has multiple_workbook_context.enabled=false.
DRAFT_PREFIXES = {"INS", "LNK"}


def parse_sections(text: str) -> list[tuple[str, str, str]]:
    lines = text.splitlines()
    headers = [
        (i, m.group(1), m.group(2).strip())
        for i, line in enumerate(lines)
        if (m := SECTION_RE.match(line))
    ]
    sections = []
    for idx, (line_no, rule_id, title) in enumerate(headers):
        end = headers[idx + 1][0] if idx + 1 < len(headers) else len(lines)
        body = "\n".join(lines[line_no + 1 : end]).strip()
        body = re.sub(r"\n?-{3,}\s*$", "", body).strip()
        sections.append((rule_id, title, body))
    return sections


def split_body(body: str) -> tuple[str | None, str | None, str | None, str | None]:
    severity = None
    sev_m = re.search(r"SEVERITY:\s*\n+\s*([A-Z]+)", body)
    body_wo_sev = body[: sev_m.start()].strip() if sev_m else body.strip()
    if sev_m:
        severity = sev_m.group(1)

    pass_if = fail_if = None
    p_m = re.search(r"PASS IF:\s*\n+(.*?)(?=\n\s*FAIL IF:|\Z)", body_wo_sev, re.S)
    f_m = re.search(r"FAIL IF:\s*\n+(.*)\Z", body_wo_sev, re.S)
    if p_m:
        pass_if = p_m.group(1).strip()
    if f_m:
        fail_if = f_m.group(1).strip()

    if pass_if is None and fail_if is None:
        criteria_text = body_wo_sev.strip() or None
    else:
        criteria_text = None

    return severity, pass_if, fail_if, criteria_text


def bullets_or_text(s: str):
    raw_lines = [l for l in s.splitlines() if l.strip()]
    if raw_lines and all(l.strip().startswith("-") for l in raw_lines):
        return [l.strip("- ").strip() for l in raw_lines]
    return " ".join(l.strip() for l in raw_lines) or None


def build_rule(rule_id: str, title: str, body: str, source_doc: str) -> dict:
    prefix = rule_id.split("-")[0]
    severity, pass_if, fail_if, criteria_text = split_body(body)
    priority = SEVERITY_TO_PRIORITY.get(severity or "MEDIUM", "RECOMMENDED")
    status = "draft" if prefix in DRAFT_PREFIXES else "active"

    pass_criteria = bullets_or_text(pass_if) if pass_if else None
    fail_criteria = bullets_or_text(fail_if) if fail_if else None

    desc_parts = [title]
    if criteria_text:
        desc_parts.append(re.sub(r"\s+", " ", criteria_text))
    elif pass_criteria and not fail_criteria:
        flat = pass_criteria if isinstance(pass_criteria, str) else "; ".join(pass_criteria)
        desc_parts.append(f"Pass if: {flat}")
    description = " -- ".join(p for p in desc_parts if p)

    impl_fn = IMPLEMENTED.get(rule_id)
    _, category = CATEGORY_TARGETS[prefix]
    deterministic = impl_fn is not None
    implementation = (
        f"validators.{category}.{impl_fn}"
        if impl_fn
        else f"validators.{category}.{rule_id.lower().replace('-', '_')}"
    )
    on_unevaluable = "NOT_SUPPORTED"
    if deterministic:
        on_unevaluable = "REQUIRES_USER_INPUT" if priority == "REQUIRED" else "WARNING"
    if status == "draft":
        on_unevaluable = "NOT_SUPPORTED"

    rule = {
        "id": rule_id,
        "version": "1.0.0",
        "status": status,
        "category": category,
        "description": description,
        "confidence": "HIGH",
        "priority": priority,
        "source": {"document": source_doc, "article": rule_id},
        "validation": {"deterministic": deterministic, "implementation": implementation},
        "correction": {"automatic": False},
        "on_unevaluable": on_unevaluable,
    }
    if pass_criteria:
        rule["pass_if"] = pass_criteria
    if fail_criteria:
        rule["fail_if"] = fail_criteria
    if severity:
        rule["severity_source"] = severity
    return rule


def load_existing(path: Path) -> dict:
    if not path.exists():
        return {"version": "1.0.0", "rules": []}
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    data.setdefault("version", "1.0.0")
    data.setdefault("rules", [])
    return data


def write_yaml(path: Path, data: dict) -> None:
    path.write_text(
        yaml.dump(data, sort_keys=False, allow_unicode=True, width=100),
        encoding="utf-8",
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=str(Path(__file__).resolve().parents[2] / "01_Mind_Readiness_Standard.md"))
    ap.add_argument("--rules-dir", default=str(Path(__file__).resolve().parents[1] / "rules"))
    args = ap.parse_args()

    src_path = Path(args.source)
    rules_dir = Path(args.rules_dir)
    text = src_path.read_text(encoding="utf-8")
    sections = parse_sections(text)
    print(f"Parsed {len(sections)} rule sections from {src_path.name}")

    grouped: dict[str, list[dict]] = {}
    unknown = 0
    for rule_id, title, body in sections:
        prefix = rule_id.split("-")[0]
        if prefix not in CATEGORY_TARGETS:
            print(f"  WARNING: unknown prefix for {rule_id}, skipping")
            unknown += 1
            continue
        filename, _ = CATEGORY_TARGETS[prefix]
        grouped.setdefault(filename, []).append(build_rule(rule_id, title, body, src_path.name))

    # loop-rules.yaml:  drop the  old hand-authored LOOP-001..003 (superseded by
    # the doc's own LOOP-001..004), keep the one rule with no doc equivalent
    # (MM_LOOPLABELS size matching) renamed to LBL-001.
    loop_path = rules_dir / "loop-rules.yaml"
    loop_existing = load_existing(loop_path)
    labels_rule = next((r for r in loop_existing["rules"] if r["id"] == "LOOP-003"), None)
    kept_loop_rules = []
    if labels_rule:
        labels_rule = dict(labels_rule)
        labels_rule["id"] = "LBL-001"
        labels_rule["source"] = {"document": "CompleteMindDocn.docx", "article": "mm-looplabels-function", "section": "Remarks"}
        kept_loop_rules.append(labels_rule)
    grouped.setdefault("loop-rules.yaml", [])
    grouped["loop-rules.yaml"] = kept_loop_rules + grouped["loop-rules.yaml"]

    for filename, new_rules in grouped.items():
        path = rules_dir / filename
        if filename == "loop-rules.yaml":
            existing = {"version": "1.0.0", "rules": []}  # fully rebuilt above
        else:
            existing = load_existing(path)
        existing["rules"] = existing["rules"] + new_rules
        write_yaml(path, existing)
        print(f"  wrote {filename}: {len(new_rules)} mined rules ({len(existing['rules'])} total)")

    if unknown:
        print(f"WARNING: {unknown} sections had unrecognized ID prefixes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
