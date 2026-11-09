"""Stand-in input files, made with the assistant, for a model whose input files the user does not have.

The model's import macro opens outside files and pastes their data in (app/outside_inputs.py). When
the user has only the model, the assistant reads that macro (its VBA), the import table and the
shape of the sheets the files fill, and describes each file the macro expects as a small recipe:
fixed cells (the words the macro searches for) and blocks copied from the model's own cells. The app
builds the files in an Excel of its own from the numbers the model ALREADY holds, so bringing them in
changes nothing, and proves it before anything is kept: the model's macro runs on a copy with the
files (outside_inputs.run_import) and every cell of every area it writes must equal the model's own.
When not, the assistant is shown what happened -- the macro's message boxes, the cells that differ --
and tries again (MAX_ATTEMPTS).

The stand-ins only reproduce what the model holds: best results come with the actual input files
(the screens and INP-010 say so). Never the user's own Excel: DispatchEx only.
"""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any, Callable

from . import outside_inputs as core
from .formula_utils import num_to_col, parse_ref

ProgressFn = Callable[[str], None]
AskFn = Callable[[str, str], "str | None"]  # (system prompt, user prompt) -> the assistant's text (None: no answer)

MAX_ATTEMPTS = 4
MAX_CELLS_PER_FILE = 3_000_000
PREVIEW_ROWS = 6
MAX_AREAS_LISTED = 25
ASSISTANT_EFFORT = "low"  # measured 2026-11-09 on the Tokyo tool: "low" answers in ~1 min; "medium" thought past 8 min and the gateway returned nothing (each try is proven, and the next one is told what went wrong)
ASSISTANT_TIMEOUT_S = 600

# Excel's SaveAs formats by extension (the macro opens the file as the real one would be)
_FORMATS = {".xlsx": 51, ".xlsm": 52, ".xlsb": 50, ".xls": 56, ".csv": 6, ".txt": 42}
_BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_NUMBER_LIKE = re.compile(r"^\s*[-+]?(\d[\d,]*\.?\d*|\.\d+)([eE][-+]?\d+)?\s*%?\s*$")

SYSTEM_PROMPT = """You read Excel VBA import macros and describe the input files they expect.

A model workbook (the "model") has a macro that opens outside input files and pastes their data
into the model. The user does not have those files. For each import-table row you are given,
describe a stand-in file that the macro accepts and that, once the macro has pasted it, puts back
EXACTLY the numbers the model already holds in the area that file fills. The app builds each file
from your recipe by copying cells of the model (values only), runs the macro on a copy of the model
with the files, and checks that every cell of every area the macro writes equals the model's own.

Answer with ONE JSON object and nothing else:
{"files": [{"row": "<the row's cell, as listed, e.g. Ctrl_01!H35>",
            "name": "<file name: keep the expected name unless the macro needs otherwise>",
            "sheet": "<name of the sheet in the file>",
            "cells": {"A1": "Description"},
            "copies": [{"to": "A5", "from": "Data!AB2:DP2", "transpose": true}]}],
 "explain": "<one or two plain sentences: how the macro reads a file and where it pastes it>"}

Rules:
- "cells": fixed values -- the words the macro searches for (Find), headers, labels it checks.
- "copies": blocks copied from the model, values only. "from" is a range of the model with its
  sheet (Sheet!A1:B9, or 'My Sheet'!A1:B9); "to" is the top-left cell in the file; "transpose": true
  turns the model's rows into the file's columns (use it when the macro transposes on the way in).
- Follow the VBA exactly: Find(...), Offset/Cells(r, c) relative to a found cell (Cells(5, 1) of a
  cell is 4 rows below it), End(xlDown) / End(xlToRight), Resize, the loops that transpose, the sheet
  and the cell where the data lands.
- End(xlDown) / End(xlToRight) stop at the first EMPTY cell, and the macro often counts with them
  (the number of items, of periods). So every cell they walk must be filled: put the item labels in
  the column (or row) the macro walks -- copy them from the model's own labels for that area, in
  the same order -- and the values beside them, filled to the last period.
- A word the macro searches for (Find("...")) must be in the file exactly; when it is the first of
  the model's labels, the copied labels already hold it.
- The block must reproduce the model's area cell for cell: the same number of rows and columns the
  macro will write, in the orientation the macro reads, taken from the cells it will be pasted to.
- Copy labels (item names, headers) from the model when the macro reads or counts them.
- The macro may test the file name (Like "*BBA*"): keep the expected name. One file per row listed.
- Never invent numbers: every number in a file comes from the model.
- Copy the WHOLE block the macro fills, from the first cell it pastes to down to the last row (or
  column) of the model that holds numbers there -- even where the model is empty (its file was
  missing when the model last ran). The app writes 0 into the empty cells of a block of numbers,
  so the walks of End(...) always reach the end; it accepts 0 where the model is empty and then
  checks that no computed value changes. Copy labels and numbers as separate blocks."""


# --- 1. what the assistant reads -------------------------------------------------------------------
def chain_code(modules: list[tuple[str, str]], scan: dict[str, Any]) -> str:
    """The VBA of the import chain (the import macro down to the procedure that opens the files) and
    of every procedure those call, comments left out, as one text."""
    procs = core.vba_procedures(modules)
    info = scan.get("procedures") or {}
    chain = list(scan.get("chain") or [])
    wanted: list[str] = []
    for name in chain:
        for k in [name, *((info.get(name) or {}).get("calls") or [])]:
            if k in procs and k not in wanted:
                wanted.append(k)
    out = []
    for k in wanted:
        p = procs[k]
        kind = {"function": "Function", "property": "Property"}.get(p["kind"], "Sub")
        out.append(f"' module {p['module']}\n{kind} {p['name']}({p['params']})\n" + "\n".join(p["lines"]) + f"\nEnd {kind}")
    return "\n\n".join(out)


def _row_key(r: dict[str, Any]) -> str:
    return f"{r['sheet']}!{r['cell']}"


def _norm_key(text: Any) -> str:
    """'Ctrl_01'!$H$35 / ctrl_01!h35 / My Sheet!H35 -> ctrl_01!H35 (the rows' keys compare like this)."""
    s = str(text or "").strip()
    if "!" not in s:
        return s.casefold()
    sheet, cell = s.rsplit("!", 1)
    sheet = sheet.strip()
    if len(sheet) > 1 and sheet[0] == sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    ref = parse_ref(cell.strip())
    at = f"{num_to_col(ref['c1'])}{ref['r1']}" if ref and not ref.get("sheet") else cell.strip().replace("$", "").upper()
    return f"{sheet.casefold()}!{at}"


def _short(v: Any, n: int = 40) -> str:
    if isinstance(v, str):
        s = v.replace("\n", " ")
        return json.dumps(s[:n] + ("…" if len(s) > n else ""), ensure_ascii=False)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return repr(v) if not isinstance(v, (int, float)) else str(v)


def compress_row(row: int, first_col: int, values: list[Any]) -> str:
    """The non-empty cells of one row, runs of the same kind summed up:
    'A1="csv file"  AB1:DP1 = 93 texts ("Premium", "Claims", …, "Expense")  DQ1 = 4'."""
    parts: list[str] = []
    i, n = 0, len(values)
    while i < n:
        v = values[i]
        if v is None or v == "":
            i += 1
            continue
        kind = "text" if isinstance(v, str) else "number"
        j = i
        while j + 1 < n and values[j + 1] not in (None, "") and (isinstance(values[j + 1], str) == (kind == "text")):
            j += 1
        a, b = f"{num_to_col(first_col + i)}{row}", f"{num_to_col(first_col + j)}{row}"
        run = values[i:j + 1]
        if len(run) <= 4:
            parts.extend(f"{num_to_col(first_col + i + k)}{row}={_short(x)}" for k, x in enumerate(run))
        else:
            sample = ", ".join(_short(x, 24) for x in run[:3]) + f", …, {_short(run[-1], 24)}"
            parts.append(f"{a}:{b} = {len(run)} {kind}s ({sample})")
        i = j + 1
    text = "  ".join(parts) or "(empty)"
    return text if len(text) <= 1800 else text[:1800] + " …"


def _areas(rng: Any, kind: int, value: int | None) -> tuple[list[str], int]:
    try:
        found = rng.SpecialCells(kind, value) if value is not None else rng.SpecialCells(kind)
    except Exception:  # Excel raises when there is no such cell
        return [], 0
    count = int(found.Areas.Count)
    return [str(found.Areas(i).Address).replace("$", "") for i in range(1, min(count, MAX_AREAS_LISTED) + 1)], count


def _matrix(v: Any) -> list[list[Any]]:
    """Value2 of a range as a list of rows (one cell: [[v]])."""
    if isinstance(v, tuple):
        return [list(r) if isinstance(r, tuple) else [r] for r in v]
    return [[v]]


def sheet_shape(ws: Any) -> dict[str, Any]:
    """Where a sheet holds what: its used range, its first rows (the labels) and where its constant
    numbers, constant texts and formulas are -- a few lines however big the sheet."""
    ur = ws.UsedRange
    r0, c0 = int(ur.Row), int(ur.Column)
    nr, nc = int(ur.Rows.Count), int(ur.Columns.Count)
    used = core.rect_ref((r0, c0, r0 + nr - 1, c0 + nc - 1))
    rows = []
    for r in range(r0, min(r0 + PREVIEW_ROWS, r0 + nr)):
        vals = _matrix(ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + nc - 1)).Value2)[0]
        rows.append(f"row {r}: {compress_row(r, c0, vals)}")
    numbers, n_numbers = _areas(ur, 2, 1)   # xlCellTypeConstants, xlNumbers
    texts, n_texts = _areas(ur, 2, 2)       # xlCellTypeConstants, xlTextValues
    formulas, n_formulas = _areas(ur, -4123, None)  # xlCellTypeFormulas
    return {"name": str(ws.Name), "used": used, "rows": rows, "numbers": numbers, "numbers_areas": n_numbers,
            "texts": texts, "texts_areas": n_texts, "formulas": formulas, "formulas_areas": n_formulas}


def describe_model(model: Path, rows: list[dict[str, Any]], vba: str) -> dict[str, Any]:
    """What the assistant reads of the model, from an Excel of our own (read only, macros off): each
    import-table row (its cells and the path formula), the defined names the macro uses, and the
    shape of every sheet the files fill."""
    import pythoncom

    pythoncom.CoInitialize()
    xl = wb = None
    pid = None
    try:
        xl, pid = core._new_excel(macros=False)
        xl.EnableEvents = False
        wb = xl.Workbooks.Open(str(model), UpdateLinks=0, ReadOnly=True)
        sheet_names = [str(wb.Worksheets(i).Name) for i in range(1, int(wb.Worksheets.Count) + 1)]
        by_norm = {core.normalize_name(n): n for n in sheet_names}
        table = []
        targets: list[str] = []
        for r in rows:
            ws = wb.Worksheets(r["sheet"])
            ur = ws.UsedRange
            c0, nc = int(ur.Column), int(ur.Columns.Count)
            vals = _matrix(ws.Range(ws.Cells(r["row"], c0), ws.Cells(r["row"], c0 + nc - 1)).Value2)[0]
            target = r.get("target_sheet") or next((by_norm[core.normalize_name(k)] for k in r.get("keys") or [] if core.normalize_name(k) in by_norm), None)
            if target and target not in targets:
                targets.append(target)
            table.append({"key": _row_key(r), "expected": r.get("expected"), "keys": list(r.get("keys") or []), "target_sheet": target,
                          "formula": str(ws.Range(r["cell"]).Formula), "row": compress_row(int(r["row"]), c0, vals)})
        names = []
        for i in range(1, int(wb.Names.Count) + 1):
            nm = wb.Names(i)
            name = str(nm.Name).split("!")[-1]
            if re.search(rf'"{re.escape(name)}"', vba):
                names.append({"name": name, "refers_to": str(nm.RefersTo)})
        shapes = [sheet_shape(wb.Worksheets(t)) for t in targets]
        return {"sheets": sheet_names, "table": table, "names": names, "shapes": shapes}
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        wb = None
        core._close_excel(xl, None, pid)
        xl = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def build_prompt(vba: str, macro: str, model: dict[str, Any], need: list[str], history: list[dict[str, Any]]) -> str:
    """The user prompt: the macro, the names it uses, the import table, the shape of the sheets the
    files fill, the rows that need a file -- and what happened to the earlier attempts."""
    out = [f"## The import macro\nThe model runs `{macro}` to bring the files in (the app writes each file's full path into its "
           "import-table cell first).\n```vba\n" + vba + "\n```"]
    if model.get("names"):
        out.append("## Defined names the macro uses\n" + "\n".join(f"- {n['name']} -> {n['refers_to']}" for n in model["names"]))
    lines = []
    for t in model["table"]:
        lines.append(f"- {t['key']}: expected file {json.dumps(t['expected'], ensure_ascii=False) if t['expected'] else '(not known)'}"
                     + (f"; fills sheet {t['target_sheet']}" if t.get("target_sheet") else "")
                     + f"\n  its cell's formula: {t['formula'][:400]}\n  its row: {t['row']}")
    out.append("## The import table (one row per input file)\n" + "\n".join(lines))
    shapes = []
    for s in model["shapes"]:
        more = lambda k: f" (+{s[k + '_areas'] - len(s[k])} more areas)" if s[k + "_areas"] > len(s[k]) else ""  # noqa: E731
        shapes.append(f"### {s['name']} -- used range {s['used']}\n" + "\n".join(s["rows"])
                      + f"\nconstant numbers: {', '.join(s['numbers']) or '(none)'}{more('numbers')}"
                      + f"\nconstant texts: {', '.join(s['texts']) or '(none)'}{more('texts')}"
                      + f"\nformulas: {', '.join(s['formulas']) or '(none)'}{more('formulas')}")
    out.append("## The sheets the files fill (the model as it is now: its data came from the last import)\n" + "\n\n".join(shapes))
    out.append("## Your task\nOne stand-in file for each of these rows: " + ", ".join(need) + ". Copy the numbers from the model; never invent them.")
    for h in history:
        out.append(f"## Attempt {h['n']} did not work\nYour recipe:\n```json\n{json.dumps(h.get('recipe'), ensure_ascii=False)[:6000]}\n```\n"
                   f"What happened: {h['problem']}\n" + "\n".join(h.get("details") or [])
                   + "\nFix the recipe (answer with the whole recipe again, every file).")
    return "\n\n".join(out)


# --- 2. the recipe ---------------------------------------------------------------------------------
def _json_object(text: str) -> dict[str, Any]:
    m = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.S)
    raw = m.group(1) if m else text[text.find("{"): text.rfind("}") + 1]
    if not raw:
        raise ValueError("no JSON object in the answer")
    try:
        obj = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"the JSON does not parse ({exc.msg} at character {exc.pos})") from None
    if not isinstance(obj, dict):
        raise ValueError("the answer is not a JSON object")
    return obj


def _cell(text: Any) -> tuple[int, int]:
    ref = parse_ref(str(text or ""))
    if not ref or ref.get("sheet") or ref["whole_column"] or ref["whole_row"] or (ref["r1"], ref["c1"]) != (ref["r2"], ref["c2"]):
        raise ValueError(f"{text!r} is not one cell")
    return ref["r1"], ref["c1"]


def parse_recipe(text: str, rows: list[dict[str, Any]], need: list[str], sheets: list[str]) -> dict[str, Any]:
    """The assistant's answer as a checked recipe: {files: [{key, name, sheet, cells: [(r, c, value)],
    copies: [{to: (r, c), sheet, rect, transpose, rows, cols}]}], explain}. ValueError says what is wrong."""
    obj = _json_object(text)
    by_key = {_norm_key(_row_key(r)): r for r in rows}
    wanted = {_norm_key(k): k for k in need}
    sheet_by_fold = {s.casefold(): s for s in sheets}
    files, seen, names = [], set(), set()
    for f in obj.get("files") or []:
        if not isinstance(f, dict):
            raise ValueError("each file must be an object")
        k = _norm_key(f.get("row"))
        if k not in wanted:
            raise ValueError(f"row {f.get('row')!r} is not one of the rows asked for")
        if k in seen:
            raise ValueError(f"two files for row {f.get('row')!r}")
        seen.add(k)
        r = by_key[k]
        name = str(f.get("name") or r.get("expected") or "").strip() or f"{(r.get('keys') or ['input'])[0]}.xlsx"
        if Path(name).name != name or Path(name).suffix.lower() not in _FORMATS:
            raise ValueError(f"file name {name!r}: a plain name ending in {', '.join(sorted(_FORMATS))}")
        if core.normalize_name(name) in names:
            raise ValueError(f"two files named {name!r}")
        names.add(core.normalize_name(name))
        sheet = str(f.get("sheet") or "Sheet1").strip()[:31] or "Sheet1"
        if _BAD_SHEET_CHARS.search(sheet):
            raise ValueError(f"sheet name {sheet!r} has a character Excel refuses")
        cells = []
        for ref, value in (f.get("cells") or {}).items():
            if value is not None and not isinstance(value, (str, int, float, bool)):
                raise ValueError(f"cell {ref}: a text or a number")
            cells.append((*_cell(ref), value))
        copies, total = [], 0
        zeros = f.get("zeros") or []
        for z in ([zeros] if isinstance(zeros, dict) else zeros):
            if not isinstance(z, dict):
                raise ValueError("each zeros block must be an object {to, rows, cols}")
            try:
                zr, zc = int(z.get("rows")), int(z.get("cols"))
            except (TypeError, ValueError):
                raise ValueError(f"zeros at {z.get('to')!r}: rows and cols must be numbers") from None
            if zr < 1 or zc < 1:
                raise ValueError(f"zeros at {z.get('to')!r}: at least one row and one column")
            total += zr * zc
            copies.append({"to": _cell(z.get("to")), "zeros": True, "rows": zr, "cols": zc})
        for c in f.get("copies") or []:
            if not isinstance(c, dict):
                raise ValueError("each copy must be an object")
            src = parse_ref(str(c.get("from") or ""))
            if not src or not src.get("sheet") or src["whole_column"] or src["whole_row"]:
                raise ValueError(f"copy from {c.get('from')!r}: a range of the model with its sheet (Sheet!A1:B9)")
            if src["sheet"].casefold() not in sheet_by_fold:
                raise ValueError(f"copy from {c.get('from')!r}: the model has no sheet {src['sheet']!r}")
            nr, nc = src["r2"] - src["r1"] + 1, src["c2"] - src["c1"] + 1
            tr = bool(c.get("transpose"))
            total += nr * nc
            copies.append({"to": _cell(c.get("to")), "sheet": sheet_by_fold[src["sheet"].casefold()],
                           "rect": (src["r1"], src["c1"], src["r2"], src["c2"]), "transpose": tr,
                           "rows": nc if tr else nr, "cols": nr if tr else nc})
        if total > MAX_CELLS_PER_FILE:
            raise ValueError(f"{name}: {total:,} cells copied, more than {MAX_CELLS_PER_FILE:,}")
        if not copies:
            raise ValueError(f"{name}: no block copied from the model and no zeros (the file needs its data)")
        files.append({"key": _row_key(r), "name": name, "sheet": sheet, "cells": cells, "copies": copies})
    missing = [wanted[k] for k in wanted if k not in seen]
    if missing:
        raise ValueError(f"no file for {', '.join(missing)}")
    return {"files": files, "explain": str(obj.get("explain") or "").strip()[:600], "raw": obj}


# --- 3. the files ----------------------------------------------------------------------------------
def _writable(v: Any) -> Any:
    """A text Excel would read as a formula keeps its first character as text."""
    return "'" + v if isinstance(v, str) and v[:1] in ("=", "+", "-", "@") and not _NUMBER_LIKE.match(v) else v


def zero_fill(data: list[list[Any]]) -> list[list[Any]]:
    """A block of numbers (no text in it) gets 0 in its empty cells: a macro walks the data with
    End(xlDown) / End(xlToRight), which stop at the first empty cell (an area that was empty in the
    model -- its file was missing -- would otherwise make it read to the end of the sheet). A block
    with text (labels) is left as it is. The proof accepts 0 where the model is empty."""
    if any(isinstance(v, str) and v != "" for row in data for v in row):
        return data
    return [[0 if v is None or v == "" else v for v in row] for row in data]


def build_files(model: Path, recipe: dict[str, Any], folder: Path) -> dict[str, Path]:
    """The stand-in files of `recipe`, built in an Excel of our own from the model's own cells (Value2:
    the numbers exactly as the model holds them). {row key: path}."""
    import pythoncom

    folder = Path(folder)
    if folder.exists():
        shutil.rmtree(folder, ignore_errors=True)
    folder.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    xl = src = None
    pid = None
    out: dict[str, Path] = {}
    try:
        xl, pid = core._new_excel(macros=False)
        xl.EnableEvents = False
        src = xl.Workbooks.Open(str(model), UpdateLinks=0, ReadOnly=True)
        try:
            xl.SheetsInNewWorkbook = 1
        except Exception:
            pass
        for f in recipe["files"]:
            wb = xl.Workbooks.Add()
            try:
                ws = wb.Worksheets(1)
                if f["sheet"] != str(ws.Name):
                    ws.Name = f["sheet"]
                for r, c, v in f["cells"]:
                    cell = ws.Cells(r, c)
                    if isinstance(v, str) and _NUMBER_LIKE.match(v):
                        cell.NumberFormat = "@"  # a label that looks like a number stays the text it is
                    cell.Value2 = _writable(v)
                for cp in f["copies"]:
                    if cp.get("zeros"):
                        tr, tc = cp["to"]
                        ws.Range(ws.Cells(tr, tc), ws.Cells(tr + cp["rows"] - 1, tc + cp["cols"] - 1)).Value2 = 0
                        continue
                    r1, c1, r2, c2 = cp["rect"]
                    sws = src.Worksheets(cp["sheet"])
                    data = _matrix(sws.Range(sws.Cells(r1, c1), sws.Cells(r2, c2)).Value2)
                    if cp["transpose"]:
                        data = [list(x) for x in zip(*data)]
                    data = zero_fill(data)
                    if any(isinstance(v, str) and _NUMBER_LIKE.match(v) for row in data for v in row):
                        texts = True
                    else:
                        texts = False
                    tr, tc = cp["to"]
                    target = ws.Range(ws.Cells(tr, tc), ws.Cells(tr + len(data) - 1, tc + len(data[0]) - 1))
                    if texts:
                        target.NumberFormat = "@"  # numbers written below stay numbers; number-like texts stay texts
                    target.Value2 = tuple(tuple(_writable(v) for v in row) for row in data)
                path = folder / f["name"]
                fmt = _FORMATS[Path(f["name"]).suffix.lower()]
                wb.SaveAs(str(path), FileFormat=fmt)
                out[f["key"]] = path
            finally:
                wb.Close(SaveChanges=False)
        return out
    finally:
        try:
            if src is not None:
                src.Close(SaveChanges=False)
        except Exception:
            pass
        src = None
        core._close_excel(xl, None, pid)
        xl = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# --- 4. the proof ----------------------------------------------------------------------------------
def compare_matrices(sheet: str, rect: tuple[int, int, int, int], model: list[list[Any]], got: list[list[Any]],
                     samples: int = 5) -> dict[str, Any]:
    """Cell by cell, exactly: {sheet, ref, cells, differ, samples: [{cell, model, got}]}."""
    r1, c1, r2, c2 = rect
    differ, zero_for_empty, held, shown = 0, 0, 0, []
    for i in range(r2 - r1 + 1):
        a_row = model[i] if i < len(model) else []
        b_row = got[i] if i < len(got) else []
        for j in range(c2 - c1 + 1):
            a = a_row[j] if j < len(a_row) else None
            b = b_row[j] if j < len(b_row) else None
            if a == "":
                a = None
            if b == "":
                b = None
            if a is not None:
                held += 1
            if a is None and isinstance(b, (int, float)) and not isinstance(b, bool) and b == 0:
                zero_for_empty += 1  # a stand-in's 0 where the model is empty (its file was missing)
                continue
            if a != b or type(a) is not type(b) and not (isinstance(a, (int, float)) and isinstance(b, (int, float))):
                differ += 1
                if len(shown) < samples:
                    shown.append({"cell": f"{sheet}!{num_to_col(c1 + j)}{r1 + i}", "model": a, "got": b})
    return {"sheet": sheet, "ref": core.rect_ref(rect), "cells": (r2 - r1 + 1) * (c2 - c1 + 1), "differ": differ, "samples": shown,
            "zero_for_empty": zero_for_empty, "model_empty": held == 0}


def compare_areas(model: Path, imported: Path, areas: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every area the macro wrote, in the imported copy against the model, cell by cell."""
    import pythoncom

    pythoncom.CoInitialize()
    xl = a = b = None
    pid = None
    out = []
    try:
        xl, pid = core._new_excel(macros=False)
        xl.EnableEvents = False
        a = xl.Workbooks.Open(str(model), UpdateLinks=0, ReadOnly=True)
        copy = Path(imported).with_name("compare_" + Path(imported).name)  # two books of one name cannot be open together
        shutil.copy2(imported, copy)
        b = xl.Workbooks.Open(str(copy), UpdateLinks=0, ReadOnly=True)
        for area in areas:
            ref = parse_ref(area["ref"])
            rect = (ref["r1"], ref["c1"], ref["r2"], ref["c2"])
            wa, wb_ = a.Worksheets(area["sheet"]), b.Worksheets(area["sheet"])
            va = _matrix(wa.Range(wa.Cells(rect[0], rect[1]), wa.Cells(rect[2], rect[3])).Value2)
            vb = _matrix(wb_.Range(wb_.Cells(rect[0], rect[1]), wb_.Cells(rect[2], rect[3])).Value2)
            out.append({**compare_matrices(area["sheet"], rect, va, vb), "file": area.get("file"), "key": area.get("key")})
        return out
    finally:
        for book in (b, a):
            try:
                if book is not None:
                    book.Close(SaveChanges=False)
            except Exception:
                pass
        a = b = None
        core._close_excel(xl, None, pid)
        xl = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _fmt_value(v: Any) -> str:
    return "(empty)" if v is None else _short(v, 30)


def judge(report: dict[str, Any], need: list[str], compared: list[dict[str, Any]] | None) -> tuple[str | None, list[str]]:
    """(None, details) when every stand-in was opened, written and reproduced the model exactly;
    else (the problem in one line, the details for the assistant)."""
    by_key = {e["cell"]: e for e in report.get("files") or []}
    details: list[str] = []
    said = [str(d.get("message_repaired") or d.get("message") or "").strip() for d in report.get("dialogs") or []]
    for s in [x for x in said if x][:4]:
        details.append(f'- The macro showed a message box: "{s[:300]}"')
    if report.get("error"):
        details.append(f"- Excel: {str(report['error'])[:300]}")
    problems = []
    for k in need:
        e = by_key.get(k) or {}
        if not e.get("opened"):
            problems.append(f"{e.get('file') or k} was not opened")
            details.append(f"- {k} ({e.get('file')}): the macro did not open it" + (" (it stopped before)" if report.get("stopped_early") else ""))
        elif not e.get("areas"):
            problems.append(f"nothing was written from {e.get('file')}")
            details.append(f"- {k} ({e.get('file')}): opened, but the macro wrote nothing from it")
    for c in compared or []:
        line = f"- {c.get('key')} ({c.get('file')}): the macro wrote {c['sheet']}!{c['ref']}; "
        ref = parse_ref(c["ref"]) or {}
        if ref and (ref.get("r2", 0) >= 16000 or ref.get("c2", 0) >= 16000):
            details.append(f"- {c.get('key')}: the macro wrote down to {c['ref'].split(':')[-1]} -- an End(xlDown)/End(xlToRight) walk ran past "
                           "the data: the row or column it walks has an empty cell (or the labels and the values are not side by side)")
        if c["differ"]:
            problems.append(f"{c['differ']:,} of {c['cells']:,} cells of {c['sheet']}!{c['ref']} differ from the model")
            ex = "; ".join(f"{s['cell']}: model {_fmt_value(s['model'])}, file gave {_fmt_value(s['got'])}" for s in c["samples"])
            details.append(line + f"{c['differ']:,} of {c['cells']:,} cells differ from the model, e.g. {ex}")
        else:
            details.append(line + f"all {c['cells']:,} cells equal the model's")
    if report.get("stopped_early") and not problems:
        problems.append("the macro stopped early")
    if report.get("error") and not problems:
        problems.append(str(report["error"])[:200])
    return ("; ".join(problems[:4]) if problems else None), details


def numbers_changed(numbers: dict[str, Any]) -> tuple[int, str]:
    """How many computed values a numbers check found changed, and the first one in words."""
    c = numbers.get("counts") or {}
    n = int(c.get("good_to_error", 0)) + int(c.get("good_to_other", 0)) + int(c.get("missing_good", 0))
    samples = numbers.get("samples") or {}
    first = next((x for kind in ("good_to_error", "good_to_other", "missing") for x in samples.get(kind) or []), None)
    example = ""
    if first and first.get("original"):
        example = (f"{first['original']} would show {_fmt_value(first.get('is'))} instead of {_fmt_value(first.get('was'))}" if "is" in first
                   else f"{first['original']} would no longer be a formula")
    return n, example


def check_numbers(model: Path, imported: Path) -> dict[str, Any]:
    """Every computed value of the imported copy against the model's (app/numbers_check.py)."""
    from .numbers_check import Moves, compare_workbooks

    return compare_workbooks(model, imported, Moves([]))


# --- 5. the whole job ------------------------------------------------------------------------------
def assistant_ask(model_id: str | None = None, effort: str = ASSISTANT_EFFORT, timeout: int = ASSISTANT_TIMEOUT_S) -> AskFn | None:
    """The app's assistant (app/llm.py) as an AskFn; None when no AI key is saved."""
    from .llm import DEFAULT_MODEL_ID, chat_completion, llm_available

    if not llm_available():
        return None

    def ask(system: str, user: str) -> str | None:
        res = chat_completion([{"role": "user", "content": user}], system, model_id=model_id or DEFAULT_MODEL_ID,
                              max_tokens=12000, timeout=timeout, effort=effort)
        text = res.get("text") if isinstance(res, dict) else None
        if not text:  # why, in the gateway's words (a timeout, a refused key, no text)
            raise RuntimeError(str((res or {}).get("message") or "no text")[:200])
        return text

    return ask


def make_stand_ins(model: Path, rows: list[dict[str, Any]], need: list[str], macro: str, modules: list[tuple[str, str]],
                   work_dir: Path, ask: AskFn, real_files: dict[str, Path] | None = None, messages: list[dict[str, Any]] | None = None,
                   progress: ProgressFn | None = None, attempts: int = MAX_ATTEMPTS, run_import: Callable[..., dict[str, Any]] | None = None,
                   describe: Callable[..., dict[str, Any]] | None = None, build: Callable[..., dict[str, Path]] | None = None,
                   compare: Callable[..., list[dict[str, Any]]] | None = None,
                   numbers: Callable[..., dict[str, Any]] | None = None) -> dict[str, Any]:
    """Stand-in files for the import-table rows `need` ("sheet!cell" keys), proven before they are
    returned: the model's macro runs on a copy with them (and with `real_files`, the files the user
    did pick) and every area each stand-in fills must equal the model's own cell for cell (a 0 where
    the model is empty counts as equal: its file was missing when the model last ran); then no
    computed value of the model may change. When the zeros of empty areas are what changes a value,
    those rows get no stand-in (`left_out`: their blocks stay as the model holds them).

    {ok, files: {key: path}, left_out: [key], explain, problem, numbers, attempts: [{n, problem, details, recipe}]}.
    run_import / describe / build / compare / numbers replace the Excel steps in tests."""
    say = progress or (lambda _m: None)
    run_import = run_import or core.run_import
    describe = describe or describe_model
    build = build or build_files
    compare = compare or compare_areas
    numbers = numbers or check_numbers
    work_dir = Path(work_dir)
    scan = core.scan_vba(modules)
    vba = chain_code(modules, scan)
    say("Reading the model's sheets for the assistant")
    shape = describe(model, rows, vba)
    history: list[dict[str, Any]] = []
    for n in range(1, attempts + 1):
        say(f"Asking the assistant how the input files look (attempt {n} of {attempts})")
        try:
            text = ask(SYSTEM_PROMPT, build_prompt(vba, macro, shape, need, history))
        except Exception as exc:  # the gateway's own words
            history.append({"n": n, "problem": f"the assistant did not answer ({str(exc)[:160]})", "details": [], "recipe": None})
            continue
        if not text:
            history.append({"n": n, "problem": "the assistant did not answer", "details": [], "recipe": None})
            continue
        try:
            recipe = parse_recipe(text, rows, need, shape["sheets"])
        except ValueError as exc:
            history.append({"n": n, "problem": f"its answer could not be used: {exc}", "details": [], "recipe": None})
            continue
        folder = work_dir / f"attempt_{n}"
        say(f"Building {len(recipe['files'])} stand-in file{'' if len(recipe['files']) == 1 else 's'} from the model's own numbers")
        try:
            files = build(model, recipe, folder / "files")
        except Exception as exc:  # a recipe Excel cannot follow (a cell out of the sheet ...)
            history.append({"n": n, "problem": f"the files could not be built: {str(exc)[:200]}", "details": [], "recipe": recipe["raw"]})
            continue
        say(f"Running the model's import macro {macro} with the stand-in files")
        # the stand-ins alone (the rows of real files keep the model's data, as missing ones): what
        # they do is all that changes; their import-table rows keep the model's own status
        report = run_import(model, rows, files, macro, folder / "check" / Path(model).name, folder,
                            messages=messages or [], restore_rows=set(files))
        with_real = False
        if (report.get("stopped_early") or report.get("error")) and real_files:
            # a macro that wants every file: with the user's files too (their data changes values by
            # design, so the cell-by-cell check of the stand-ins' areas is the proof)
            say("Running the macro again with the files you picked as well")
            report = run_import(model, rows, {**real_files, **files}, macro, folder / "check_all" / Path(model).name, folder / "all",
                                messages=messages or [], restore_rows=set(files))
            with_real = True
        written = [{"sheet": a["sheet"], "ref": a["ref"], "file": e.get("file"), "key": e["cell"]}
                   for e in report.get("files") or [] if e["cell"] in files and e.get("opened") for a in e.get("areas") or []]
        compared = None
        if report.get("saved") and written:
            say("Checking every cell the macro wrote against the model")
            compared = compare(model, Path(report["saved"]), written)
        problem, details = judge(report, need, compared)
        left_out: list[str] = []
        checked: dict[str, Any] | None = None
        if problem is None and with_real:
            details.append("- The computed values were not compared: the macro needs every file, and the files you picked change values by design")
        elif problem is None:
            say("Checking that no computed value of the model changes")
            checked = numbers(model, Path(report["saved"]))
            changed, example = numbers_changed(checked) if checked.get("ran") and not checked.get("clean") else (0, "")
            empty = sorted({c["key"] for c in compared or [] if c.get("model_empty")})
            if changed and empty and len(empty) < len(files):
                # the zeros where the model is empty change a result: those rows keep no stand-in
                say("Checking again without the stand-ins of the empty areas")
                kept = {k: v for k, v in files.items() if k not in empty}
                report2 = run_import(model, rows, kept, macro, folder / "check_without_empty" / Path(model).name,
                                     folder / "without_empty", messages=messages or [], restore_rows=set(kept))
                problem2, _d2 = judge(report2, list(kept), None)
                checked2 = numbers(model, Path(report2["saved"])) if problem2 is None and report2.get("saved") else None
                if checked2 is not None and (not checked2.get("ran") or checked2.get("clean")):
                    files, left_out, checked, changed = kept, empty, checked2, 0
                    details.append(f"- Left out (the model holds no data there, and zeros would change a result): {', '.join(empty)}")
            if changed:
                problem = f"{changed:,} computed value{'' if changed == 1 else 's'} of the model would change" + (f", e.g. {example}" if example else "")
                details.append(f"- {problem}: the files must reproduce the model's data exactly")
        if problem is None:
            if checked is not None and not checked.get("ran"):
                details.append(f"- The computed values could not be compared: {checked.get('verdict')}")
            done = history + [{"n": n, "problem": None, "details": details, "recipe": recipe["raw"]}]
            _keep_record(work_dir, done)
            return {"ok": True, "files": files, "left_out": left_out, "explain": recipe["explain"], "problem": None, "numbers": checked,
                    "attempts": done}
        history.append({"n": n, "problem": problem, "details": details, "recipe": recipe["raw"]})
        _keep_record(work_dir, history)
    last = history[-1]["problem"] if history else "no attempt"
    return {"ok": False, "files": {}, "left_out": [], "explain": None, "problem": last, "numbers": None, "attempts": history}


def _keep_record(work_dir: Path, attempts: list[dict[str, Any]]) -> None:
    """Every try, its recipe and what happened, beside the files (attempts.json): to see later why."""
    try:
        work_dir.mkdir(parents=True, exist_ok=True)
        (work_dir / "attempts.json").write_text(json.dumps(attempts, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    except Exception:
        pass
