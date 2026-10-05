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

1.7.2: the default model is Sonnet 5 (`MIND_READY_MODEL` overrides it), and
`chat_completion(..., on_delta=)` streams the reply (server-sent events) so
the web app can show the text as it is written (`MIND_READY_STREAM=0`
turns streaming off).
"""
from __future__ import annotations

import base64
import json
import os
import re
from pathlib import Path
from typing import Any, Callable, Iterator

DEFAULT_AZURE_ENDPOINT = "https://apim-aiservices-prod-01.azure-api.net"
# /claude/invocations stopped answering (200 + empty body, 2026-09-16); the
# native Anthropic Messages passthrough takes a top-level "system" param.
CLAUDE_APIM_PATH = "/anthropic/v1/messages"
# Measured through the gateway 2026-09-24 (same ~150-word answer): Sonnet 4.6
# 7.7 s, Sonnet 5 6.2 s, Opus 4.8 6.7 s; no Haiku on the gateway.
DEFAULT_MODEL_ID = os.environ.get("MIND_READY_MODEL", "claude-sonnet-5")

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


def _takes_temperature(model_id: str) -> bool:
    """Sonnet 5 answers 400 "`temperature` is deprecated for this model"
    (2026-09-24); the 4.x models still accept it."""
    return model_id.startswith(("claude-sonnet-4", "claude-opus-4"))


# Sonnet 5 thinks before it answers by default. Measured 2026-09-24 on a real
# chat turn (3k input tokens, max_tokens 1200): ~885 thinking tokens, ~9 s of
# silence (thinking is not streamed through the gateway), then the answer ran
# out of room -- stop_reason max_tokens, cut off or empty. With effort "low"
# it answered at once (no thinking block), complete, in ~4.5 s. Newer models
# get `output_config.effort` (env MIND_READY_EFFORT, "" = model default) and
# extra room so a turn that does think still has space for its answer.
MODEL_EFFORT = os.environ.get("MIND_READY_EFFORT", "low").strip()
THINKING_HEADROOM_TOKENS = 3000


def _takes_effort(model_id: str) -> bool:
    return not _takes_temperature(model_id)


def _streaming_on() -> bool:
    return os.environ.get("MIND_READY_STREAM", "1").strip().lower() not in ("0", "off", "false", "no")


class StreamInterrupted(Exception):
    """The gateway closed a streamed reply before `message_stop`, sent an
    `error` event, or sent data that is not JSON."""


def _sse_data(resp: Any) -> Iterator[dict[str, Any]]:
    """The decoded JSON of every `data:` line of an SSE body (event-name and
    blank lines skipped). Lines are read as BYTES and decoded as UTF-8 here:
    requests' text decoding follows the Content-Type charset and splits on
    Unicode separators such as U+2028 inside the model's text (ported from
    IFRS_DataScraper/app/apim.py)."""
    for raw in resp.iter_lines(decode_unicode=False):
        if not raw:
            continue
        line = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw
        if not line.startswith("data:"):
            continue
        data = line[5:].strip()
        try:
            ev = json.loads(data)
        except ValueError as exc:
            raise StreamInterrupted(f"undecodable SSE data: {data[:80]!r}") from exc
        if isinstance(ev, dict):
            yield ev


def _parse_claude_stream(resp: Any, on_delta: Callable[[str], None]) -> tuple[str, str | None]:
    """Consume an Anthropic Messages SSE stream: every `text_delta` goes to
    `on_delta` as it arrives (thinking deltas are skipped); returns the whole
    text and the stop_reason."""
    parts: list[str] = []
    stop_reason = None
    for ev in _sse_data(resp):
        kind = ev.get("type")
        if kind == "content_block_delta":
            d = ev.get("delta") or {}
            if d.get("type") == "text_delta" and d.get("text"):
                parts.append(d["text"])
                on_delta(d["text"])
        elif kind == "message_delta":
            stop_reason = (ev.get("delta") or {}).get("stop_reason") or stop_reason
        elif kind == "message_stop":
            return "".join(parts), stop_reason
        elif kind == "error":
            raise StreamInterrupted("error event: " + json.dumps(ev.get("error", ev), ensure_ascii=False)[:200])
    raise StreamInterrupted(f"stream ended without message_stop after {sum(map(len, parts))} chars")


def _body_text(body: dict[str, Any]) -> str | None:
    if "content" in body and isinstance(body["content"], list):
        return "".join(part.get("text", "") for part in body["content"] if isinstance(part, dict) and part.get("type", "text") == "text")
    if "choices" in body:
        return body["choices"][0]["message"]["content"]
    return None


TRUNCATED_NOTE = "\n\n_(The answer was cut off at the length limit.)_"


def _finish(text: str, stop_reason: str | None, on_delta: Callable[[str], None] | None) -> dict[str, Any]:
    """The result dict; says so when the model returned nothing or was cut off
    (never a silently truncated answer)."""
    if not text.strip():
        return {"available": True, "text": None, "message": f"The model returned no text (stop_reason {stop_reason or 'unknown'})."}
    if stop_reason == "max_tokens":
        text += TRUNCATED_NOTE
        if on_delta is not None:
            on_delta(TRUNCATED_NOTE)
    return {"available": True, "text": text, "message": None}


def chat_completion(
    messages: list[dict[str, str]],
    system_prompt: str,
    model_id: str = DEFAULT_MODEL_ID,
    max_tokens: int = 1200,
    temperature: float = 0.2,
    timeout: int = 120,
    on_delta: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """Send a multi-turn conversation (`messages`: [{role: user|assistant,
    content: str}, ...]) through the gateway's Anthropic Messages
    passthrough (system as the top-level `system` param). Returns
    {available, text, message}. With `on_delta` the reply is streamed and
    every text chunk is passed to it as it arrives; a gateway that answers a
    streamed request with plain JSON is read the old way. `temperature` is
    only sent to models that accept it (see `_takes_temperature`)."""
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
        payload: dict[str, Any] = {"model": model_id, "max_tokens": max_tokens, "system": system_prompt, "messages": payload_messages}
        if _takes_temperature(model_id):
            payload["temperature"] = temperature
        else:
            payload["max_tokens"] = max_tokens + THINKING_HEADROOM_TOKENS
            if MODEL_EFFORT:
                payload["output_config"] = {"effort": MODEL_EFFORT}
        stream = on_delta is not None and _streaming_on()
        if stream:
            payload["stream"] = True
        resp = requests.post(url, headers=headers, json=payload, timeout=timeout, **({"stream": True} if stream else {}))
        rejected = [k for k in ("temperature", "output_config") if k in payload and k in (resp.text or "")] if resp.status_code == 400 else []
        if rejected:
            # A model that refuses one of the tuning params (not yet known to
            # _takes_temperature / _takes_effort): drop it and ask once more.
            for k in rejected:
                payload.pop(k)
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout, **({"stream": True} if stream else {}))
        resp.raise_for_status()
        if stream and "text/event-stream" in (resp.headers.get("content-type") or ""):
            try:
                text, stop_reason = _parse_claude_stream(resp, on_delta)  # type: ignore[arg-type]
            finally:
                resp.close()
            return _finish(text, stop_reason, on_delta)
        body = resp.json()
        text = _body_text(body)
        if text is None:
            return {"available": True, "text": None, "message": f"Unrecognized response shape: {list(body.keys())}"}
        if on_delta is not None and text:
            on_delta(text)  # not streamed: hand over the whole reply at once
        return _finish(text, body.get("stop_reason"), on_delta)
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
