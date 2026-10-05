"""1.8.0 -- the assistant's part in the automatic fixer (app/autofix.py).

The fixer keeps an error formula and gives it a value to return when it
fails: =IFERROR(<formula>, <fallback>). Which fallback fits -- 0, empty text
or FALSE -- depends on what the cell *is* (an amount that is summed, a label,
a rate that is averaged, a check), and that is a judgement about meaning: the
assistant reads the cell's labels, its formula and what it reads, and says
which value fits and why, in one plain sentence shown to the user.

It only orders the ladder. Whatever it says, the numbers gate has the last
word: a fallback that changes a value that was good is rolled back and the
next one is tried. No answer (assistant not configured, a timeout, a reply
that is not the JSON asked for) means the built-in order -- never an error.
"""
from __future__ import annotations

import json
import re
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable

from .llm import DEFAULT_MODEL_ID, chat_completion, llm_available

ALLOWED_FALLBACKS = {"0": "0", '""': '""', "''": '""', "": '""', "EMPTY": '""', "FALSE": "FALSE"}
BATCH = 10
MAX_GROUPS = 80  # the biggest groups get an opinion; the long tail keeps the built-in order

SYSTEM_PROMPT = """You help repair an Excel model (usually an actuarial or financial model) so that every formula recalculates without an error.

For each error group below, the tool keeps the formula and wraps it: =IFERROR(<the formula>, <fallback>). You choose the fallback:
- 0      -- the cell is an amount, a count or a rate that other cells add up or multiply; a missing value should weigh nothing.
- ""     -- (empty text) the cell is a label or a piece of text, or it is a figure where a 0 would be wrong to show or would be counted / averaged as a real observation.
- FALSE  -- the cell is a logical test (TRUE/FALSE).

You see, per group: the sheet and cell, the labels left of / above it, the formula, the error, the values of the cells it reads, how many cells share the formula, and what the same formula returns where it works ("number", "text" or "bool").

Answer with JSON only, no prose around it:
[{"id": 1, "fallback": "0", "reason": "..."}, ...]
`reason` is ONE short plain sentence for a non-programmer: what goes wrong in this cell and why this value is the right thing to show instead. Do not mention IFERROR or JSON in it."""


def _prompt(batch: list[tuple[int, dict[str, Any]]]) -> str:
    lines = []
    for n, g in batch:
        ctx = g.get("context") or {}
        lines.append(
            f"### Group {n}\n"
            f"cell: {g['sheet']}!{g['cell']}  ({g['count']} cell(s) share this formula)\n"
            f"error: {g['error']}\n"
            f"formula: {str(g['formula'])[:600]}\n"
            f"labels left of the cell: {ctx.get('left_of_cell') or '-'}\n"
            f"labels above the cell: {ctx.get('above_cell') or '-'}\n"
            f"it reads: {'; '.join(ctx.get('reads') or []) or '-'}\n"
            f"the same formula elsewhere returns: {g.get('kind_of_value') or 'unknown'}"
        )
    return "\n\n".join(lines) + "\n\nGive the JSON list now, one entry per group."


def parse_answer(text: str | None) -> dict[int, dict[str, str]]:
    """{group id: {"fallback", "reason"}} out of the model's reply; anything
    that is not the JSON asked for is dropped, never guessed."""
    if not text:
        return {}
    m = re.search(r"\[.*\]", text, re.DOTALL)
    if not m:
        return {}
    try:
        items = json.loads(m.group(0))
    except ValueError:
        return {}
    out: dict[int, dict[str, str]] = {}
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        try:
            n = int(item.get("id"))
        except (TypeError, ValueError):
            continue
        raw = item.get("fallback")
        key = "FALSE" if raw is False else "0" if raw == 0 and raw is not False else str(raw if raw is not None else "").strip().upper()
        fallback = ALLOWED_FALLBACKS.get(key)
        if fallback is None:
            continue
        reason = re.sub(r"\s+", " ", str(item.get("reason") or "")).strip()[:300]
        out[n] = {"fallback": fallback, "reason": reason}
    return out


def make_advisor(model_id: str | None = None, complete: Callable[..., dict[str, Any]] | None = None) -> Callable[[list[dict[str, Any]]], dict[str, dict[str, Any]]] | None:
    """An advisor for AutoFixConfig, or None when no assistant is configured.
    `complete` replaces app.llm.chat_completion in tests."""
    if complete is None:
        if not llm_available():
            return None
        complete = chat_completion
    model = model_id or DEFAULT_MODEL_ID

    def advise(groups: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
        numbered = list(enumerate(groups[:MAX_GROUPS], start=1))
        batches = [numbered[i:i + BATCH] for i in range(0, len(numbered), BATCH)]

        def ask(batch: list[tuple[int, dict[str, Any]]]) -> dict[int, dict[str, str]]:
            try:
                res = complete([{"role": "user", "content": _prompt(batch)}], SYSTEM_PROMPT, model_id=model, max_tokens=1600, temperature=0.0, timeout=90)
            except Exception:
                return {}
            return parse_answer(res.get("text") if isinstance(res, dict) else None)

        answers: dict[int, dict[str, str]] = {}
        with ThreadPoolExecutor(max_workers=min(4, len(batches) or 1)) as pool:
            for got in pool.map(ask, batches):
                answers.update(got)
        return {g["key"]: answers[n] for n, g in numbered if n in answers}

    return advise
