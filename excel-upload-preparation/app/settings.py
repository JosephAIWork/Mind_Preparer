"""The settings a team may change, read once from `settings.yaml` next to the app.

A line left out of the file means the default below. A mistake in the file (bad YAML, an unknown
setting, a wrong type) stops the app with a message naming it, rather than running on a guess. The
environment variables that existed before this file still win: MIND_READY_PORT, MIND_READY_MODEL,
MIND_MCP_URL, MINDREADY_MM_XLL.
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import yaml

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = PACKAGE_ROOT / "settings.yaml"


@dataclass(frozen=True)
class Settings:
    port: int = 8600
    data_folder: Path = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "MindReady"
    api_pages: bool = False
    ai_endpoint: str = "https://apim-aiservices-prod-01.azure-api.net"
    ai_model: str = "claude-sonnet-5"
    mind_site: str = "https://shared.milliman-mind.com"
    mind_mcp_url: str = "https://vm-prod.milliman-mind.com/api/chatbot/mcp"
    mind_sandbox_folder: str = "MindReady_Sandbox"
    mind_project_prefix: str = "zzz_mindready"
    mm_for_excel_addin: str = ""

    @property
    def mind_main_url(self) -> str:
        return self.mind_site.rstrip("/") + "/main"

    @property
    def mind_host_suffix(self) -> str:
        """shared.milliman-mind.com -> milliman-mind.com: every page of the Mind app is under it."""
        host = urlparse(self.mind_site).hostname or ""
        return ".".join(host.split(".")[-2:])


# file key -> (Settings field, type)
_KEYS: dict[str, tuple[str, type]] = {
    "port": ("port", int),
    "data_folder": ("data_folder", str),
    "api_pages": ("api_pages", bool),
    "ai_gateway.endpoint": ("ai_endpoint", str),
    "ai_gateway.model": ("ai_model", str),
    "mind.site": ("mind_site", str),
    "mind.mcp_url": ("mind_mcp_url", str),
    "mind.sandbox_folder": ("mind_sandbox_folder", str),
    "mind.project_prefix": ("mind_project_prefix", str),
    "mm_for_excel_addin": ("mm_for_excel_addin", str),
}


def _flatten(data: dict[str, Any], prefix: str = "") -> dict[str, Any]:
    out: dict[str, Any] = {}
    for k, v in data.items():
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            out.update(_flatten(v, key + "."))
        else:
            out[key] = v
    return out


def _folder(text: str) -> Path:
    """%VARIABLES% expanded; a relative folder is under the app's folder; a variable this machine does
    not have falls back to the default."""
    expanded = os.path.expandvars(text.strip())
    if "%" in expanded or "$" in expanded:
        return Settings.data_folder
    path = Path(expanded)
    return path if path.is_absolute() else PACKAGE_ROOT / path


def load(path: Path = SETTINGS_FILE) -> Settings:
    if not path.is_file():
        return Settings()
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:
        raise ValueError(f"{path.name} is not valid YAML: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"{path.name} should hold 'name: value' lines")
    values: dict[str, Any] = {}
    for key, value in _flatten(data).items():
        if key not in _KEYS:
            raise ValueError(f"{path.name}: unknown setting '{key}' (known: {', '.join(_KEYS)})")
        field, kind = _KEYS[key]
        if value is None or value == "":
            continue  # empty = the default
        if kind is int and (isinstance(value, bool) or not isinstance(value, int)):
            raise ValueError(f"{path.name}: '{key}' should be a whole number, not {value!r}")
        if kind is bool and not isinstance(value, bool):
            raise ValueError(f"{path.name}: '{key}' should be true or false, not {value!r}")
        if kind is str:
            value = str(value).strip()
        values[field] = _folder(value) if field == "data_folder" else value
    prefix = values.get("mind_project_prefix", Settings.mind_project_prefix)
    if len(prefix) < 3:
        raise ValueError(f"{path.name}: 'mind.project_prefix' should be at least 3 characters: the app deletes only projects whose name starts with it")
    return Settings(**values)


SETTINGS = load()
