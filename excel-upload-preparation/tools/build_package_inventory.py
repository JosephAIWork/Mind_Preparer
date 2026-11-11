#!/usr/bin/env python3
"""Regenerate  PACKAGE_INVENTORY.json (path  / bytes / sha256 for every
package file). Run after any change to the package so the inventory stays
truthful; caches, logs and test scratch output are excluded."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
EXCLUDE_DIRS = {"__pycache__", ".pytest_cache", ".git", "node_modules"}
EXCLUDE_SUFFIXES = {".pyc", ".log"}
EXCLUDE_FILES = {"PACKAGE_INVENTORY.json"}


def main() -> int:
    files = []
    for path in sorted(PACKAGE_ROOT.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(PACKAGE_ROOT)
        if set(rel.parts[:-1]) & EXCLUDE_DIRS or path.suffix in EXCLUDE_SUFFIXES or path.name in EXCLUDE_FILES:
            continue
        data = path.read_bytes()
        files.append({"path": rel.as_posix(), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()})
    payload = {"package": "excel-upload-preparation", "version": (PACKAGE_ROOT / "VERSION").read_text(encoding="utf-8").strip(), "file_count": len(files), "files": files}
    (PACKAGE_ROOT / "PACKAGE_INVENTORY.json").write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(f"PACKAGE_INVENTORY.json: {len(files)} files")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
