"""The key for the Milliman AI gateway, saved for the person who pasted it.

The web app asks for the key when none is saved (`POST /api/assistant-key`). It is stored as
`ai_key.bin` in the data folder (settings.yaml; `%LOCALAPPDATA%\\MindReady` by default), encrypted
with Windows' DPAPI for the current Windows user: only that user, on this machine, can decrypt it, so
a copy of the file is useless anywhere else. The key itself is never sent back to the page, only its
last 4 characters.
"""
from __future__ import annotations

import os
import re

from .settings import SETTINGS

KEY_FILE = SETTINGS.data_folder / "ai_key.bin"
_ENTROPY = b"MindReady.ai-key"  # an app-specific extra secret: another program of the same user needs it too
_DESCRIPTION = "Mind Ready AI gateway key"
_CRYPTPROTECT_UI_FORBIDDEN = 0x1
# Printable ASCII without spaces; a gateway subscription key is 32 hex characters.
_KEY_RE = re.compile(r"^[\x21-\x7e]{8,512}$")


def check_format(key: str) -> str | None:
    """None when `key` looks like a key, else what is wrong with it (shown to the user)."""
    if not key:
        return "Paste the key first."
    if not _KEY_RE.match(key):
        return "That does not look like a key: it should be one word of letters and digits, with no spaces."
    return None


def save(key: str) -> None:
    """Encrypt `key` for the current Windows user and write it (replacing any saved key)."""
    try:
        import win32crypt  # pywin32
    except ImportError as exc:
        raise RuntimeError("Saving the key needs Windows (pywin32).") from exc
    blob = win32crypt.CryptProtectData(key.encode("utf-8"), _DESCRIPTION, _ENTROPY, None, None, _CRYPTPROTECT_UI_FORBIDDEN)
    KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = KEY_FILE.with_suffix(".tmp")
    tmp.write_bytes(blob)
    os.replace(tmp, KEY_FILE)


def load() -> str | None:
    """The saved key, or None (no file, not Windows, or saved by another user or machine)."""
    try:
        blob = KEY_FILE.read_bytes()
        import win32crypt  # pywin32

        _, data = win32crypt.CryptUnprotectData(blob, _ENTROPY, None, None, _CRYPTPROTECT_UI_FORBIDDEN)
        return data.decode("utf-8") or None
    except Exception:
        return None


def last4(key: str | None) -> str | None:
    return key[-4:] if key else None
