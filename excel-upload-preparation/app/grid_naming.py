"""Context-aware grid names from the Milliman APIM Claude gateway.

Why this exists
---------------
`prep._context_name` reads a name off the cells around a grid. That is
deterministic and cheap, and it is right whenever the model actually wrote a
heading next to the block. When it isn't -- a block with no heading, a heading
that is a whole sentence, a header row of single letters -- it falls back to
`"<Sheet> <Anchor>"` ("Cashflows C4"), which is unique and useless: Mind shows
it in the template list and nobody can tell what the grid holds.

This module asks the model to read the same surroundings and propose a short
name, in ONE batched call for the whole workbook. It is advisory in exactly the
way every assistant feature of the app is:

  * it never writes anything -- it returns names, and the ordinary prep plan
    (`prep.plan_create_grid_titles`) turns them into reviewable operations the
    user approves;
  * every name is sanitised and made unique by the same helpers the
    deterministic path uses, so an odd answer cannot produce an invalid title;
  * when the gateway is unavailable the caller keeps the deterministic name --
    naming never becomes a hard dependency on the network.

`build_grid_context` is also useful on its own: it is what the UI shows when a
user asks "what is in this grid?".
"""
from __future__ import annotations

import json
import re
from typing import Any, Callable

from .formula_utils import ref_text
from .inventory import cell_value, values_cell
from .llm import DEFAULT_MODEL_ID, chat_completion

NAME_MAX = 60
CONTEXT_ROWS_ABOVE = 3
CONTEXT_COLS_LEFT = 3
CONTEXT_SAMPLE_ROWS = 4
CONTEXT_SAMPLE_COLS = 8
CELL_TEXT_MAX = 32

SYSTEM_PROMPT = (
    "You name tables ('grids') in an actuarial Excel model that is being prepared for "
    "Milliman Mind. For each grid you get the cells above it, the cells to its left, and "
    "its first rows (formulas are shown as 'formula => computed value'). Give each grid a "
    "short, specific name an actuary would recognise -- what the block holds, not where it "
    "sits. 2-5 words, Title Case, no sheet names, no cell references, no '#', no quotes, "
    "no trailing punctuation. Prefer the model's own wording when the surrounding cells "
    "supply it. If the surroundings genuinely say nothing, describe the contents instead "
    "(e.g. 'Mortality Rates By Age'). Never invent a business meaning the cells do not "
    "support. Answer with a JSON object mapping each grid's id to its name, nothing else."
)


def _clip(value: Any, limit: int = CELL_TEXT_MAX) -> str:
    return str(value).replace("\n", " ").replace("\r", " ")[:limit]


def _shown(analysis: dict[str, Any], sheet: str, row: int, col: int) -> str | None:
    """The cell as a human reads it: a formula with the value it computed."""
    raw = cell_value(analysis, sheet, row, col)
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return None
    if isinstance(raw, str) and raw.startswith("="):
        computed = values_cell(analysis, sheet, row, col)
        if computed is not None and str(computed).strip() != "":
            return f"{_clip(raw, 28)} => {_clip(computed, 20)}"
    return _clip(raw)


def build_grid_context(analysis: dict[str, Any], grid: dict[str, Any]) -> dict[str, Any]:
    """What surrounds a grid, in the form both the model and the UI can read."""
    sheet = grid["sheet"]
    r0, r1 = grid["first_row"], grid["last_row"]
    c0, c1 = grid["first_col"], grid["last_col"]

    above: list[str] = []
    for row in range(max(1, r0 - CONTEXT_ROWS_ABOVE), r0):
        cells = []
        for col in range(max(1, c0 - 2), min(c1, c0 + CONTEXT_SAMPLE_COLS) + 1):
            shown = _shown(analysis, sheet, row, col)
            if shown:
                cells.append(f"{ref_text(col, row)}={shown}")
        if cells:
            above.append(" | ".join(cells[:8]))

    left: list[str] = []
    for col in range(max(1, c0 - CONTEXT_COLS_LEFT), c0):
        cells = []
        for row in range(r0, min(r1, r0 + CONTEXT_SAMPLE_ROWS) + 1):
            shown = _shown(analysis, sheet, row, col)
            if shown:
                cells.append(f"{ref_text(col, row)}={shown}")
        if cells:
            left.append(" | ".join(cells[:6]))

    body: list[str] = []
    for row in range(r0, min(r1, r0 + CONTEXT_SAMPLE_ROWS - 1) + 1):
        cells = [_shown(analysis, sheet, row, col) or "" for col in range(c0, min(c1, c0 + CONTEXT_SAMPLE_COLS) + 1)]
        body.append(" | ".join(cells))

    return {
        "id": f"{sheet}!{grid['ref']}",
        "sheet": sheet,
        "ref": grid["ref"],
        "size": f"{grid['n_rows']} rows x {grid['n_cols']} cols",
        "flags": grid.get("flag_names") or [],
        "current_name": grid.get("name"),
        "above": above,
        "left": left,
        "sample": body,
    }


def _render(context: dict[str, Any]) -> str:
    parts = [f"id: {context['id']}  ({context['size']})"]
    if context["flags"]:
        parts.append(f"  flags: {', '.join(context['flags'])}")
    if context["above"]:
        parts.append("  above:")
        parts += [f"    {line}" for line in context["above"]]
    if context["left"]:
        parts.append("  left:")
        parts += [f"    {line}" for line in context["left"]]
    parts.append("  first rows:")
    parts += [f"    {line}" for line in context["sample"]]
    return "\n".join(parts)


_JSON_OBJECT = re.compile(r"\{.*\}", re.DOTALL)


def _parse_names(text: str | None) -> dict[str, str]:
    if not text:
        return {}
    match = _JSON_OBJECT.search(text)
    if not match:
        return {}
    try:
        data = json.loads(match.group(0))
    except (ValueError, TypeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {str(k): str(v) for k, v in data.items() if isinstance(v, (str, int, float)) and str(v).strip()}


def suggest_names(
    contexts: list[dict[str, Any]],
    model_id: str = DEFAULT_MODEL_ID,
    batch_size: int = 25,
    completion: Callable[..., dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """{available, names: {grid_id: name}, message}. Batched so a whole
    workbook costs a handful of calls; a batch that fails is skipped rather
    than losing the batches that worked."""
    if not contexts:
        return {"available": True, "names": {}, "message": None}
    call = completion or chat_completion
    names: dict[str, str] = {}
    messages: list[str] = []
    available = True
    for start in range(0, len(contexts), batch_size):
        batch = contexts[start : start + batch_size]
        prompt = (
            f"Name these {len(batch)} grids. Reply with JSON: "
            '{"<id>": "<name>", ...} using exactly these ids.\n\n'
            + "\n\n".join(_render(c) for c in batch)
        )
        result = call(
            [{"role": "user", "content": prompt}],
            SYSTEM_PROMPT,
            model_id=model_id,
            max_tokens=min(4000, 120 * len(batch) + 400),
            temperature=0.0,
            timeout=180,
        )
        if not result.get("available"):
            return {"available": False, "names": {}, "message": result.get("message")}
        if result.get("message"):
            messages.append(result["message"])
        parsed = _parse_names(result.get("text"))
        if not parsed:
            messages.append(f"no usable names for grids {start + 1}-{start + len(batch)}")
        names.update(parsed)
    return {"available": available, "names": names, "message": "; ".join(messages) or None}
