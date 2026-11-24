"""manifest.json and SKILL.md carry the app's version as plain strings: they must follow VERSION
(they said 1.11.0 while the app was at 1.14.1). Bump all three together."""
import json
import re
from pathlib import Path

APP = Path(__file__).resolve().parents[2]


def test_manifest_and_skill_follow_the_version_file():
    version = (APP / "VERSION").read_text(encoding="utf-8").strip()
    manifest = json.loads((APP / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["version"] == version
    skill = (APP / "SKILL.md").read_text(encoding="utf-8")
    assert re.search(r'package-version:\s*"%s"' % re.escape(version), skill)
    assert "port" not in manifest["web"] and "settings.yaml" in manifest["web"]["portSetting"]
