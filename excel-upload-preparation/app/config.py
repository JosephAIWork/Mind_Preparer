"""Loads  config/default.yaml +  config/target-version.yaml into one dict
passed to modes/validators. Kept as plain dicts (not pydantic) since these
files are host-application configuration, not part of the schema-bound
output contract."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

PACKAGE_ROOT = Path(__file__).resolve().parent.parent


def load_config(config_dir: Path | None = None) -> dict[str, Any]:
    config_dir = config_dir or (PACKAGE_ROOT / "config")
    config: dict[str, Any] = {}
    default_path = config_dir / "default.yaml"
    if default_path.exists():
        config.update(yaml.safe_load(default_path.read_text(encoding="utf-8")) or {})
    version_path = config_dir / "target-version.yaml"
    if version_path.exists():
        config.update(yaml.safe_load(version_path.read_text(encoding="utf-8")) or {})
    return config
