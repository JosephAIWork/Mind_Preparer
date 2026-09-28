"""LLM abstraction (CLAUDE_CODE_PROMPT.md phase 7) over the shared Milliman
APIM Claude gateway.

Two uses, both advisory:

* `suggest_formula_fix` -- one-shot "suggest a fix" for a formula finding
  (UI button). Never auto-applied (SKILL.md non-negotiable rule #6). Output
  is always tagged evidence=RECOMMENDATION.
* `chat_completion` (1.4.0) -- multi-turn conversation used by the
  workbook chatbot (app/chat_context.py). The model only ever *answers*;
  every workbook change still goes through the reviewable prep plan
  (app/prep.py) and the user's explicit approval.

Credential discovery reuses the proven repo convention (see
reserve_narrator/utils/key_loader.py): a `secret.key` file holding a Fernet
key, with a sibling `config.enc` holding the Azure APIM subscription key,
searched for in nearby project directories (this package is one level
deeper than reserve_narrator, hence levels_up=4).
"""
from __future__ import annotations

import base64
import re
from pathlib import Path
from typing import Any

DEFAULT_AZURE_ENDPOINT = "https://apim-aiservices-prod-01.azure-api.net"
# 1.7.2: the gateway's Anthropic Messages passthrough. The former
# "/claude/invocations" operation answers HTTP 200 with an empty body since
# 2026-09 (any payload, even {}), which the app reported as "Request failed:
# Expecting value". This path speaks the native Messages API: `system` is a
# top-level field, never a leading {"role": "system"} message (400 otherwise).
CLAUDE_APIM_PATH = "/anthropic/v1/messages"
DEFAULT_MODEL_ID = "claude-sonnet-4-6"

_KV_LINE = re.compile(r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\S.*?)\s*$")


def _looks_like_fernet_key(text: str) -> bool:
    token = text.strip()
    if len(token) != 44:
        return False
    try:
        return len(base64.urlsafe_b64decode(token.encode())) == 32
    except (ValueError, TypeError):
        return False


def _candidate_secret_files(start_dir: Path, levels_up: int = 4) -> list[Path]:
    candidates: list[Path] = []
    seen: set[Path] = set()

    def _add(path: Path) -> None:
        try:
            path = path.resolve()
        except OSError:
            return
        if path not in seen:
            seen.add(path)
            candidates.append(path)

    level_dirs = [start_dir]
    current = start_dir
    for _ in range(levels_up):
        parent = current.parent
        if parent == current:
            break
        level_dirs.append(parent)
        current = parent

    for directory in level_dirs:
        _add(directory / "secret.key")
        try:
            children = sorted(
                (c for c in directory.iterdir() if c.is_dir() and not c.name.startswith(".")),
                key=lambda p: p.name.lower(),
            )
        except OSError:
            children = []
        for child in children:
            _add(child / "secret.key")
    return candidates


def find_api_key(start_dir: Path | None = None) -> tuple[str, str] | None:
    """Returns (api_key, endpoint) or None if no usable secret.key is found."""
    start_dir = start_dir or Path(__file__).resolve().parent.parent
    for candidate in _candidate_secret_files(start_dir):
        if not candidate.is_file():
            continue
        try:
            text = candidate.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue

        kv = {}
        for line in text.splitlines():
            if line.strip().startswith("#"):
                continue
            m = _KV_LINE.match(line)
            if m and m.group(2):
                kv[m.group(1)] = m.group(2)
        if kv.get("AZURE_OPENAI_API_KEY"):
            return kv["AZURE_OPENAI_API_KEY"], kv.get("AZURE_OPENAI_ENDPOINT", DEFAULT_AZURE_ENDPOINT)

        if _looks_like_fernet_key(text):
            enc_path = candidate.parent / "config.enc"
            if enc_path.is_file():
                try:
                    from cryptography.fernet import Fernet

                    api_key = Fernet(text.strip().encode()).decrypt(enc_path.read_bytes()).decode()
                    return api_key, DEFAULT_AZURE_ENDPOINT
                except Exception:
                    continue
    return None


def llm_available() -> bool:
    return find_api_key() is not None


def _text_block(text: str) -> list[dict[str, str]]:
    return [{"type": "text", "text": text}]


def chat_completion(
    messages: list[dict[str, str]],
    system_prompt: str,
    model_id: str = DEFAULT_MODEL_ID,
    max_tokens: int = 1200,
    temperature: float = 0.2,
    timeout: int = 120,
) -> dict[str, Any]:
    """Send a multi-turn conversation (`messages`: [{role: user|assistant,
    content: str}, ...]) through the gateway. Returns {available, text,
    message}. Native Anthropic Messages payload (top-level `system`, content
    blocks), which is what the gateway's /anthropic/v1/messages passthrough
    expects; `suggest_formula_fix` goes through the same function."""
    found = find_api_key()
    if found is None:
        return {"available": False, "text": None, "message": "No secret.key found nearby; the assistant isn't configured in this environment."}
    api_key, endpoint = found
    try:
        import requests

        url = endpoint.rstrip("/") + CLAUDE_APIM_PATH
        headers = {"Content-Type": "application/json", "api-key": api_key, "anthropic-version": "2023-06-01"}
        payload_messages: list[dict[str, Any]] = []
        for m in messages:
            role = "assistant" if m["role"] == "assistant" else "user"
            payload_messages.append({"role": role, "content": _text_block(m["content"])})
        payload = {"model": model_id, "max_tokens": max_tokens, "temperature": temperature, "system": system_prompt, "messages": payload_messages}
        resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
        resp.raise_for_status()
        body = resp.json()
        if "content" in body and isinstance(body["content"], list):
            text = "".join(part.get("text", "") for part in body["content"] if isinstance(part, dict))
        elif "choices" in body:
            text = body["choices"][0]["message"]["content"]
        else:
            return {"available": True, "text": None, "message": f"Unrecognized response shape: {list(body.keys())}"}
        return {"available": True, "text": text, "message": None}
    except Exception as exc:
        return {"available": True, "text": None, "message": f"Request failed: {exc}"}


def suggest_formula_fix(finding: dict[str, Any], formula_text: str, model_id: str = DEFAULT_MODEL_ID) -> dict[str, Any]:
    """One-shot suggestion for a formula-related finding. Never applied
    automatically -- the caller (UI) shows it as a RECOMMENDATION only."""
    system_prompt = (
        "You review Excel formulas for compatibility with Milliman Mind's MMForExcel add-in "
        "(MM_-prefixed functions) and Mind's supported native Excel function list. Given one finding and its "
        "formula, suggest a concrete, minimal fix. Only use documented MM_ functions if proposing one. If you "
        "are not confident, say so plainly rather than guessing. Keep the answer to 2-4 sentences, no preamble. "
        "If you propose a replacement formula, put the complete formula on its own line starting with '='."
    )
    user_prompt = f"Rule: {finding['rule_id']}\nFinding: {finding['message']}\nFormula: {formula_text}\n\nWhat's a concrete fix?"
    result = chat_completion([{"role": "user", "content": user_prompt}], system_prompt, model_id=model_id, max_tokens=512, temperature=0.1, timeout=60)
    return {"available": result["available"], "suggestion": result["text"], "message": result["message"]}


FORMULA_LINE_RE = re.compile(r"^\s*(=[^\n]+?)\s*$", re.MULTILINE)


def extract_formula(text: str | None) -> str | None:
    """The first line of an answer that is a complete '=...' formula, if any
    (used to pre-fill -- never to auto-apply -- the formula replacement box)."""
    if not text:
        return None
    for m in FORMULA_LINE_RE.finditer(text):
        candidate = m.group(1).strip().strip("`")
        if candidate.startswith("=") and len(candidate) > 1:
            return candidate
    return None
