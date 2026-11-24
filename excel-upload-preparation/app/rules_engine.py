"""Loads  rules/*.yaml, validates  every rule against rules/rule-schema.json,
and resolves each rule's `validation.implementation` dotted path to an
actual Python validator function (falling back to NOT_SUPPORTED when the
implementation doesn't exist yet -- see validators/__init__.py).
"""
from __future__ import annotations

import importlib
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import jsonschema
import yaml

PACKAGE_ROOT = Path(__file__).resolve().parent.parent


@dataclass
class RuleLoadError:
    file: str
    rule_id: str | None
    message: str


class RulesEngine:
    def __init__(self, rules_dir: Path | None = None, schema_path: Path | None = None):
        self.rules_dir = rules_dir or (PACKAGE_ROOT / "rules")
        self.schema_path = schema_path or (self.rules_dir / "rule-schema.json")
        self._validator = jsonschema.Draft202012Validator(
            json.loads(self.schema_path.read_text(encoding="utf-8"))
        )
        self.rules: dict[str, dict[str, Any]] = {}
        self.errors: list[RuleLoadError] = []
        self._load()

    def _load(self) -> None:
        for path in sorted(self.rules_dir.glob("*-rules.yaml")):
            data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
            for rule in data.get("rules", []):
                errs = list(self._validator.iter_errors(rule))
                if errs:
                    self.errors.append(
                        RuleLoadError(path.name, rule.get("id"), "; ".join(e.message for e in errs))
                    )
                    continue
                rule_id = rule["id"]
                if rule_id in self.rules:
                    self.errors.append(
                        RuleLoadError(path.name, rule_id, f"duplicate rule id (also in {self.rules[rule_id].get('_file')})")
                    )
                    continue
                rule = dict(rule)
                rule["_file"] = path.name
                self.rules[rule_id] = rule

    def active_rules(self, category: str | None = None) -> list[dict[str, Any]]:
        rules = [r for r in self.rules.values() if r["status"] == "active"]
        if category:
            rules = [r for r in rules if r["category"] == category]
        return sorted(rules, key=lambda r: r["id"])

    def get(self, rule_id: str) -> dict[str, Any] | None:
        return self.rules.get(rule_id)

    @staticmethod
    def resolve_validator(implementation: str) -> Callable[..., Any] | None:
        """'validators.loop.mm_loop_inventory'  -> app.validators.loop.mm_loop_inventory,
        or  None if the module/function isn't implemented yet."""
        parts = implementation.split(".")
        if len(parts) != 3 or parts[0] != "validators":
            return None
        _, category, func_name = parts
        try:
            module = importlib.import_module(f"app.validators.{category}")
        except ImportError:
            return None
        return getattr(module, func_name, None)
