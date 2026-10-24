"""Package-part comparison: a pure before/after diff of two OOXML zip
packages. Used by scripts/compare_packages.py. This is inspection only (no
mutation).

Per instructions/preservation.md, differences should be classified as
AUTHORIZED_CHANGE / EXPECTED_RECALCULATION_CHANGE / METADATA_CHANGE /
UNAUTHORIZED_CHANGE / UNVERIFIABLE_CHANGE. This comparison has no change-set
context to know what was *authorized*, so every changed/added/removed part
is reported as UNVERIFIABLE_CHANGE -- an honest default, not a guess.
"""
from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from typing import Any

# Parts that commonly change on save even with no logical edits (e.g. calc chain,
# last-saved metadata) -- flagged as METADATA_CHANGE rather than UNVERIFIABLE_CHANGE.
METADATA_PART_PREFIXES = ("docProps/", "xl/calcChain.xml")


def _part_hashes(path: Path) -> dict[str, str]:
    hashes = {}
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            hashes[name] = hashlib.sha256(zf.read(name)).hexdigest()
    return hashes


def compare_packages(before_path: Path, after_path: Path) -> dict[str, Any]:
    before = _part_hashes(Path(before_path))
    after = _part_hashes(Path(after_path))

    added = sorted(set(after) - set(before))
    removed = sorted(set(before) - set(after))
    changed = sorted(p for p in (set(before) & set(after)) if before[p] != after[p])

    def classify(part: str) -> str:
        return "METADATA_CHANGE" if part.startswith(METADATA_PART_PREFIXES) else "UNVERIFIABLE_CHANGE"

    differences = [
        {"part": p, "change": "added", "classification": classify(p)} for p in added
    ] + [
        {"part": p, "change": "removed", "classification": classify(p)} for p in removed
    ] + [
        {"part": p, "change": "modified", "classification": classify(p)} for p in changed
    ]

    return {
        "before": str(before_path),
        "after": str(after_path),
        "identical": not differences,
        "differences": differences,
    }
