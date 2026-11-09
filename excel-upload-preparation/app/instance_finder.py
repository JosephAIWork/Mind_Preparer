"""Instance finder (1.12.5): where Mind instances or a Mind loop would help a model, read blind.

A model that repeats itself -- one block per product, one sheet per year, one pass of a macro per
group -- can run in Mind as one copy per value of a key (instances) or as one block in a loop. The
finder reads the workbook's analysis and points to where that would remove repetition or compute
every choice in one run. It reports and ranks; it writes nothing. No hint is taken from a name.

Four patterns, merged into dimensions by the input the copies step through (never by labels):
- P1 a column of the table the macro's loop steps through splits its rows into copies (the macro's
  own key chain, lookups keyed by the current row's value, formulas that branch on it, the counts);
- P2 the same formulas side by side in one sheet (bands of columns or rows);
- P3 sheets with the same layout and formulas;
- P4 near copies: what differs (typed-in numbers, numbers inside formulas, logic) and the edits
  that would unify them.
Cells that read the macro's paste area are sorted by what each needs in Mind (app.finder_grouping).
Chains (each copy reads the one before it), time axes and plain filled areas are set aside and said.

Deterministic: pure Python, no Excel, no numpy; the same workbook gives the same answer. An
optional assistant pass explains and ranks on top; any failure of it leaves the deterministic result.

    python -m app.instance_finder <workbook> [--json] [--assistant]
"""
from __future__ import annotations

import json
import math
import re
import sys
import time
from copy import deepcopy
from pathlib import Path
from typing import Any, Callable

from .finder_grouping import MAX_LIST_CELLS, MAX_SLOTS, MIN_BRANCHES, find_grouping, list_texts, paste_area_readers, reverse_names, same_values, share_covered
from .finder_patterns import BIG_FILL, CHAIN_SHARE, MAX_COPIES_INSTANCES, SMALL_GAIN, edits_for, sheet_copies, sheet_findings, step_target
from .finder_shapes import Budget, build_book, cell_name
from .formula_utils import formula_memo, num_to_col, parse_ref
from .instance_builder import _json_value, _resolve_range
from .inventory import open_cached
from .llm import DEFAULT_MODEL_ID, chat_completion, llm_available
from .loop_builder import _qs, _split, build_loop_plan
from .model_map import _Names

SCHEMA = "instance-finder/1"
MAX_DIMENSIONS = 12
MAX_EVIDENCE = 8
MAX_CELLS = 6
MAX_CHAINS = 10
MAX_FINDINGS = 60  # bands + sheet groups kept before the merge
MAX_FILLED_DETAIL = 12
MAX_VBA = 6
MAX_SWITCHES = 8
MANY_COPIES = 50  # above this, instances fit "partly": a run slows down
LIST_COVER = 0.90  # share of a key column's values a list must hold to be "the same values"
RELOAD_NOTE_S = 2.0
DIMENSION_KINDS = ("col", "band", "sheets", "fill", "stepping", "keyed")
JOINED_KINDS = ("stepping", "keyed", "list", "bands", "sheets")
KIND_ORDER = ("macro", "lookups", "branches", "stepping", "keyed", "list", "bands", "sheets", "counts")
ADDERS_REASON = "{n} adding cells move to the main workbook (there a plain link to the copied model is the sum over the copies)"


def _version() -> str:
    try:
        return (Path(__file__).resolve().parents[1] / "VERSION").read_text(encoding="utf-8").strip() or "dev"
    except OSError:
        return "dev"


# --- small helpers ------------------------------------------------------------------------------------
def _jsonable(x: Any) -> Any:
    """Only str keys, lists, ints, floats, bools, None and str: json.dumps(result) needs no default."""
    if isinstance(x, dict):
        return {str(k): _jsonable(v) for k, v in x.items()}
    if isinstance(x, (list, tuple, set, frozenset)):
        return [_jsonable(v) for v in (sorted(x, key=str) if isinstance(x, (set, frozenset)) else x)]
    if isinstance(x, float):
        return x if math.isfinite(x) else None
    if x is None or isinstance(x, (bool, int, str)):
        return x
    return str(_json_value(x))


def _rect_a1(r1: int, r2: int, c1: int, c2: int) -> str:
    a = f"{num_to_col(c1)}{r1}"
    return a if (r1, c1) == (r2, c2) else f"{a}:{num_to_col(c2)}{r2}"


def _intersects(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return a[0] <= b[1] and a[1] >= b[0] and a[2] <= b[3] and a[3] >= b[2]


def _range_id(text: str | None, own_sheet: str, names: Any) -> str | None:
    """'Sheet'!A1:B2 for a range written any way (a defined name, with or without '$'), else None."""
    if not text:
        return None
    got = _resolve_range(str(text), own_sheet, names)
    if not got or got[1]["whole_column"] or got[1]["whole_row"]:
        return None
    sheet, r, _name = got
    return f"{_qs(sheet)}!{r['ref']}"


def _log2(n: float) -> float:
    return math.log2(1 + max(0.0, n))


class _Union:
    """A union-find over anchor ids; the smallest id of a set is its root (deterministic)."""

    def __init__(self) -> None:
        self.parent: dict[str, str] = {}

    def add(self, x: str) -> None:
        self.parent.setdefault(x, x)

    def find(self, x: str) -> str:
        self.add(x)
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: str, b: str) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            lo, hi = sorted((ra, rb))
            self.parent[hi] = lo

    def groups(self) -> list[list[str]]:
        out: dict[str, list[str]] = {}
        for x in self.parent:
            out.setdefault(self.find(x), []).append(x)
        return [sorted(v) for _k, v in sorted(out.items())]


class _Found:
    """Everything the detectors found, as anchors joined by edges; built by find_instances."""

    def __init__(self, names: Any) -> None:
        self.names = names
        self.uf = _Union()
        self.findings: dict[str, dict[str, Any]] = {}  # anchor id -> {"kind", ...}
        self.lists: dict[str, dict[str, Any]] = {}  # list anchor id -> {"ref", "name", "texts"}
        self.slots: dict[str, tuple[str, tuple[int, int, int, int], int, str]] = {}  # slots id -> (sheet, rect, w, axis)
        self.cleared: list[dict[str, Any]] = []  # the macro's cleared areas, each with the slot width when known
        self.key_cols: dict[str, str] = {}  # upper token of a key cell (address or name) -> the col: anchor id
        self.chains: list[dict[str, Any]] = []
        self.ignored = {"time_axes": 0, "filled_areas": 0, "small_fills": 0, "label_columns": 0}
        self.filled_detail: list[dict[str, Any]] = []
        self.readers: dict[str, Any] | None = None
        self.slots_area: str | None = None  # the slots: id of the cleared area the readers were classified on
        self.truncated: dict[str, bool] = {}
        self.table: dict[str, Any] | None = None
        self.notes: list[str] = []  # plain notes for `skipped`, written while the anchors and edges are built

    def slots_id(self, sheet: str, rect: tuple[int, int, int, int], w: int, axis: str) -> str:
        """The anchor of an area cut into slots: the macro's cleared area it lies in (same width) when
        there is one, else its own address."""
        for a in self.cleared:
            if a["sheet"] == sheet and a.get("w") in (None, w) and _intersects(rect, (a["r1"], a["r2"], a["c1"], a["c2"])):
                sid = f"slots:{_qs(sheet)}!{_rect_a1(a['r1'], a['r2'], a['c1'], a['c2'])}/{w}/{axis}"
                self.slots.setdefault(sid, (sheet, (a["r1"], a["r2"], a["c1"], a["c2"]), w, axis))
                return sid
        sid = f"slots:{_qs(sheet)}!{_rect_a1(*rect)}/{w}/{axis}"
        self.slots.setdefault(sid, (sheet, rect, w, axis))
        return sid

    def list_id(self, text: str | None, own_sheet: str, name: str | None = None) -> str | None:
        rid = _range_id(text, own_sheet, self.names)
        if rid is None:
            return None
        lid = f"list:{rid}"
        entry = self.lists.setdefault(lid, {"ref": rid, "name": None, "texts": None})
        if not entry["name"]:
            r = parse_ref(rid)
            entry["name"] = name or (self.names.name_of(r["sheet"], r["ref"]) if r and r["sheet"] else None)
        return lid

    def col_of(self, key_text: str | None, own_sheet: str) -> str | None:
        """The col: anchor a key text (a key cell's address or defined name) belongs to."""
        if not key_text:
            return None
        s = str(key_text).strip().replace("$", "").lstrip("@").upper()
        if "!" in s:
            sh, ref = _split(s)
            s = f"{(sh or own_sheet).upper()}!{ref}"
        elif parse_ref(s):
            s = f"{own_sheet.upper()}!{s}"
        return self.key_cols.get(s)


# --- the detectors, as anchors joined by edges ----------------------------------------------------------
def _p1_anchors(found: _Found, grouping: dict[str, Any], names: Any) -> None:
    """Every P1 candidate is a col: anchor; the macro chain joins it to the list it matches into, the
    slots it pastes into and the tables keyed by the value. Every key cell (current row, alias, mapped
    key) is remembered by address and name so a MATCH on it joins its column."""
    table = grouping.get("table") or {}
    sheet = table.get("sheet") or ""
    rev = reverse_names(names)
    for col, cells in (grouping.get("keys") or {}).items():
        cid = f"col:{_qs(sheet)}!{num_to_col(col)}"
        for s, r, c in cells:
            ref = f"{num_to_col(c)}{r}"
            found.key_cols[f"{s.upper()}!{ref}"] = cid
            for n in rev.get((s.upper(), ref), []):
                found.key_cols[n.upper()] = cid
    for cand in grouping.get("candidates") or []:
        cid = f"col:{_qs(sheet)}!{cand['column']}"
        found.findings[cid] = {"kind": "col", "cand": cand}
        found.uf.add(cid)
        m = cand.get("macro")
        if not m:
            continue
        lid = found.list_id(m.get("list"), sheet, m.get("list_name")) if m.get("list") else None
        if lid:
            found.uf.union(cid, lid)
        slots = m.get("slots") or {}
        r = parse_ref(slots["anchor"]) if slots.get("anchor") and slots.get("w") else None
        if r:
            sid = found.slots_id(r["sheet"] or sheet, (r["r1"], r["r2"], r["c1"], r["c2"]), int(slots["w"]), slots.get("axis") or "columns")
            found.uf.union(lid or cid, sid)
        for t in m.get("keyed_tables") or []:
            aid = f"area:{t['ref']}"
            found.findings.setdefault(aid, {"kind": "area", "name": t.get("name"), "ref": t["ref"]})
            found.uf.union(cid, aid)


def _cleared_areas(found: _Found, loop_plan: dict[str, Any], chain: dict[str, Any] | None) -> None:
    """The areas the macro empties before a run, with the slot width of the one its paste anchor lies in."""
    for c in loop_plan.get("clears") or []:
        if c.get("status") == "ready" and c.get("sheet") and c.get("r1") and c.get("c1"):
            found.cleared.append({"sheet": c["sheet"], "r1": c["r1"], "r2": c["r2"], "c1": c["c1"], "c2": c["c2"], "w": None, "axis": None})
    r = parse_ref(chain["anchor"]) if chain and chain.get("anchor") and chain.get("w") else None
    if r:
        for a in found.cleared:
            if a["sheet"] == r["sheet"] and _intersects((r["r1"], r["r2"], r["c1"], r["c2"]), (a["r1"], a["r2"], a["c1"], a["c2"])):
                a["w"], a["axis"] = int(chain["w"]), chain.get("axis") or "columns"


def _step_edges(found: _Found, fid: str, st: dict[str, Any], own_sheet: str) -> None:
    """Join a band / fill to what it steps through: slots of width w (the macro's cleared area when it lies
    there), an area read row by row, or -- for a MATCH on a key -- the list and the key's column."""
    rect = (st["r1"], st["r2"], st["c1"], st["c2"])
    c = st.get("computed") or {}
    if c.get("form") == "match":
        lid = found.list_id(c.get("list_ref") or c.get("list"), own_sheet)
        if lid:
            found.uf.union(fid, lid)
        cid = found.col_of(c.get("key"), own_sheet)
        if cid and cid in found.findings:
            found.uf.union(fid, cid)
    if st.get("stride", 1) > 1:
        found.uf.union(fid, found.slots_id(st["sheet"], rect, int(st["stride"]), st["axis"]))
    elif c.get("form") != "match":
        aid = f"area:{_qs(st['sheet'])}!{_rect_a1(*rect)}"
        found.findings.setdefault(aid, {"kind": "area", "name": None, "ref": aid[5:]})
        found.uf.union(fid, aid)


def _example_cell(finding: dict[str, Any], copy: int = 1) -> list[str]:
    cells = finding.get("cells") or {}
    lst = cells.get(copy) or cells.get(0) or []
    return [cell_name(finding["sheet"], *min(lst))] if lst else []


def _covering_groups(found: _Found, finding: dict[str, Any]) -> tuple[list[str], bool]:
    """(the reader-group anchors whose cells the finding overlaps, does it overlap the adders' block?): a
    fill or band at the place of cells the readers pass already classified belongs to them."""
    sheet = finding["sheet"]
    sample = [rc for lst in (finding.get("cells") or {}).values() for rc in lst[:40]]
    hits: list[str] = []
    for fid, f in found.findings.items():
        if f["kind"] in ("stepping", "keyed") and f["g"]["sheet"] == sheet:
            own = f["g"].get("own")
            if own and any(own[0] <= r <= own[1] and own[2] <= c <= own[3] for r, c in sample):
                hits.append(fid)
    adders = False
    for key in ("adders_hull", "followers_hull"):  # the adding block, and the cells that read it: both already in the reader lines
        hull = ((found.readers or {}).get(key) or {}).get(sheet)
        adders = adders or (bool(hull) and any(hull[0] <= r <= hull[1] and hull[2] <= c <= hull[3] for r, c in sample))
    return hits, adders


def _repetitions(found: _Found, book: Any, names: Any, budget: Budget) -> None:
    """P2 bands and fills, P3 sheet copies: chains are set aside, the rest become anchors with their step edges."""
    bands_all: list[dict[str, Any]] = []
    fills: list[dict[str, Any]] = []
    order = list(book.sheet_order) + [n for n in book.sheets if n not in book.sheet_order]
    for name in order:
        sh = book.sheets.get(name)
        if sh is None:
            continue
        if budget.out():
            budget.skip(f"repeated blocks on sheet {name} and the sheets after it not checked: the time budget ran out")
            break
        for f in sheet_findings(book, sh, budget):
            (fills if f.get("fill") else bands_all).append(f)
    groups: list[dict[str, Any]] = []
    if budget.out():
        budget.skip("repeated sheets not checked: the time budget ran out")
    else:
        groups = sheet_copies(book, budget)
    if len(bands_all) + len(groups) > MAX_FINDINGS:
        found.truncated["findings"] = True
        bands_all = sorted(bands_all, key=lambda f: (-f.get("n_cells", 0), f["sheet"], f["where"]))[:max(0, MAX_FINDINGS - len(groups))]
    for band in bands_all:
        label = f"{_qs(band['sheet'])} {band['where']}"
        if band.get("chain_share", 0) >= CHAIN_SHARE:
            found.chains.append({"label": label, "copies": band["copies"], "per_copy": band["per_copy"], "chain_share": band["chain_share"],
                                 "why": "each block reads the block before it", "cells": _example_cell(band)})
            continue
        fid = f"band:{label}"
        a1 = book.a1_text(band["sheet"], *band["first_cell"]) if band.get("first_cell") else None
        st = step_target(book, band, a1, names)
        hits, adders = _covering_groups(found, band)
        found.findings[fid] = {"kind": "band", "f": band, "step": st, "covered": bool(hits), "adders": adders}
        found.uf.add(fid)
        for h in hits:
            found.uf.union(fid, h)
        if adders and found.slots_area:
            found.uf.union(fid, found.slots_area)
        if st:
            _step_edges(found, fid, st, band["sheet"])
    for g in groups:
        members = list(g["members"])
        label = ", ".join(members)
        if g.get("chain"):
            cells = g.get("cells") or {}
            second = cells.get(members[1]) if len(members) > 1 else None
            found.chains.append({"label": label, "copies": g["copies"], "per_copy": g["per_copy"], "chain_share": 1.0,
                                 "why": "each sheet reads the sheet before it", "cells": [cell_name(members[1], *min(second))] if second else []})
            continue
        fid = f"sheets:{'+'.join(members)}"
        found.findings[fid] = {"kind": "sheets", "g": g}
        found.uf.add(fid)
    for fill in fills:
        n_cells = fill.get("n_cells") or 0
        time_axes = fill.get("time_axes") or {fill["axis"]: bool(fill.get("time_axis"))}
        hits, adders = _covering_groups(found, fill)
        if hits or adders:
            # the cells the readers pass classified: the fill belongs to that dimension, said by the reader lines
            fid = f"fill:{_qs(fill['sheet'])} {fill['where']}"
            found.findings[fid] = {"kind": "fill", "f": fill, "step": None, "covered": bool(hits), "adders": adders}
            found.uf.add(fid)
            for h in hits:
                found.uf.union(fid, h)
            if adders and found.slots_area:
                found.uf.union(fid, found.slots_area)
            continue
        if n_cells < BIG_FILL:
            if fill.get("time_axis"):
                found.ignored["time_axes"] += 1
            else:
                found.ignored["small_fills"] += 1
            continue
        st = None
        if not budget.out():
            a1 = book.a1_text(fill["sheet"], *fill["first_cell"]) if fill.get("first_cell") else None
            st = step_target(book, fill, a1, names)
        c = (st or {}).get("computed")
        if c:
            fid = f"fill:{_qs(fill['sheet'])} {fill['where']}"
            found.findings[fid] = {"kind": "fill", "f": fill, "step": st}
            found.uf.add(fid)
            _step_edges(found, fid, st, fill["sheet"])
            found.ignored["time_axes"] += sum(1 for ax, is_time in time_axes.items() if is_time and ax != c.get("step_axis"))
            continue
        if fill.get("time_axis"):
            found.ignored["time_axes"] += sum(1 for v in time_axes.values() if v) or 1
            continue
        found.ignored["filled_areas"] += 1
        if len(found.filled_detail) < MAX_FILLED_DETAIL:
            found.filled_detail.append({"sheet": fill["sheet"], "where": fill["where"], "cells": n_cells})
        else:
            found.truncated["filled_areas"] = True


def _reader_anchors(found: _Found, analysis: dict[str, Any], book: Any, names: Any, grouping: dict[str, Any], budget: Budget) -> None:
    """The cells that read the macro's paste area, classified; the stepping and keyed groups become
    anchors joined to the area's slots, to their list and key column, and to a fill at the same place."""
    area = next((a for a in found.cleared if a.get("w")), None)
    if area is None:
        if found.cleared:
            budget.skip("the macro's paste area is known but not how wide one copy's slot is: the cells that read it were not sorted by class")
        return
    if budget.out():
        budget.skip("the cells that read the macro's paste area were not sorted by class: the time budget ran out")
        return
    rd = paste_area_readers(analysis, book, names, area, area["w"], grouping.get("key_set") or set(), area["axis"])
    if rd is None:  # the width is known, so the only other way out: the area is cut into more slots than the pass reads
        budget.skip(f"the macro's paste area is cut into more than {MAX_SLOTS:,} slots: the cells that read it were not sorted by class")
        return
    found.readers = rd
    sid = found.slots_id(area["sheet"], (area["r1"], area["r2"], area["c1"], area["c2"]), area["w"], area["axis"])
    found.slots_area = sid
    for g in rd.get("groups") or []:
        own = tuple(g.get("own") or ())
        if g["kind"] == "computed":
            pos = g.get("positions") or []
            if g.get("axis") == "rows" and pos:
                where = f"rows {pos[0]}-{pos[-1]}"
            elif g.get("axis") == "columns" and pos:
                where = f"{num_to_col(pos[0])}:{num_to_col(pos[-1])}"
            else:
                where = _rect_a1(*own) if own else g["example"]
            fid = f"stepping:{_qs(g['sheet'])} {where}"
        else:
            fid = f"keyed:{_qs(g['sheet'])}!{_rect_a1(*own) if own else g['example']}"
        found.findings[fid] = {"kind": "stepping" if g["kind"] == "computed" else "keyed", "g": g}
        found.uf.union(fid, sid)
        if g["kind"] == "keyed":
            lid = found.list_id(g.get("list"), g["sheet"])
            if lid:
                found.uf.union(fid, lid)
            cid = found.col_of(g.get("key"), g["sheet"])
            if cid and cid in found.findings:
                found.uf.union(fid, cid)


def _value_edges(found: _Found, analysis: dict[str, Any]) -> None:
    """Equality edges: two lists with the same set of texts; a list holding a key column's values (or
    90 % of them); two slot anchors of the same width on one sheet whose areas overlap."""
    for lid in sorted(found.lists):
        entry = found.lists[lid]
        if entry["texts"] is None:
            r = parse_ref(entry["ref"])
            if r and r["sheet"] and (r.get("cells") or 0) > MAX_LIST_CELLS:
                found.notes.append(f"the list {entry['name'] + ' ' if entry.get('name') else ''}({entry['ref']}) has more than {MAX_LIST_CELLS:,} cells: its values were not compared")
            entry["texts"] = list_texts(analysis, r["sheet"], r) if r and r["sheet"] else []
    ids = sorted(found.lists)
    for i, a in enumerate(ids):
        for b in ids[i + 1:]:
            if same_values(found.lists[a]["texts"], found.lists[b]["texts"]):
                found.uf.union(a, b)
    for cid, f in found.findings.items():
        if f["kind"] != "col":
            continue
        values = f["cand"].get("values") or []
        for lid in ids:
            texts = found.lists[lid]["texts"]
            if texts and values and (same_values(texts, values) or share_covered(texts, values) >= LIST_COVER):
                found.uf.union(cid, lid)
    sids = sorted(found.slots)
    for i, a in enumerate(sids):
        sa = found.slots[a]
        for b in sids[i + 1:]:
            sb = found.slots[b]
            if sa[0] == sb[0] and sa[2] == sb[2] and sa[3] == sb[3] and _intersects(sa[1], sb[1]):
                found.uf.union(a, b)


# --- one dimension out of one component ---------------------------------------------------------------
def _fill_copies(f: dict[str, Any]) -> tuple[int, int, str]:
    """(copies, per copy, the stepping axis) of a fill, along the axis its computed index steps through."""
    fill, st = f["f"], f.get("step") or {}
    c = st.get("computed") or {}
    axis = c.get("step_axis") or fill["axis"]
    axes = fill.get("axes") or {}
    copies = int(axes.get(axis) or fill["copies"] or 1)
    n_cells = int(fill.get("n_cells") or 0)
    return copies, (n_cells // copies if copies else n_cells), axis


_EDIT_FIELDS = {"numbers": ("cell", "values_per_copy", "action"), "formula_numbers": ("cell", "literals_per_copy", "action"), "logic": ("cell", "formulas_per_copy", "action"), "size": ("cell", "action")}


def _merge_edits(parts: list[dict[str, Any]]) -> dict[str, Any]:
    """The edits of several blocks as one list per kind (each capped at 40, the counts complete)."""
    out: dict[str, Any] = {"numbers": [], "formula_numbers": [], "logic": [], "size": [], "numbers_count": 0, "formula_numbers_count": 0, "logic_count": 0, "size_count": 0, "branches": None}
    for e in parts:
        for k, fields in _EDIT_FIELDS.items():
            out[k].extend({f: x.get(f) for f in fields} for x in e.get(k) or [])
            out[f"{k}_count"] += int(e.get(f"{k}_count") or 0)
    for k in _EDIT_FIELDS:
        out[k] = out[k][:40]
    return out


def _score(key: dict[str, Any] | None, kinds: list[str], blocks: list[dict[str, Any]], edits: dict[str, Any]) -> int:
    s = 0.0
    if key:
        if "macro" in kinds:
            s += 50
        if key.get("keyed_formulas"):
            s += 10 + min(20.0, 4 * _log2(key["keyed_formulas"]))
        if key.get("branch_formulas", 0) >= MIN_BRANCHES:
            s += 5 + min(10.0, 2 * _log2(key["branch_formulas"]))
        d = key.get("distinct") or 0
        if 2 <= d <= MAX_COPIES_INSTANCES:
            s += 5 + 10 * float(key.get("balance") or 0) * (1 if 3 <= d <= 100 else 0.5)
        if d == 2:
            s -= 10
    for b in blocks:
        saved = max(0, (b["copies"] - 1) * b["per_copy"])
        s += min(30.0, 10 * _log2(saved / 50)) + 15
        if not (b.get("diff") or {}).get("different logic"):
            s += 10
    if edits.get("numbers_count") or edits.get("formula_numbers_count"):
        s += 5
    joined = [k for k in kinds if k in JOINED_KINDS]
    extra = len(joined) - (0 if key else 1)
    s += min(40, 20 * max(0, extra))
    return int(round(s))


def _fit(key: dict[str, Any] | None, lists: list[dict[str, Any]], rd: dict[str, Any] | None, copies: int, biggest: int | None, saved: int,
         edits: dict[str, Any], cross_reads: list[dict[str, Any]], has_blocks: bool) -> dict[str, Any]:
    """How each option fits, in plain words. Mind facts: a plain link from another workbook to the copied
    model is the sum over the copies; MM_LOOPINSTANCE runs a loop of its own size inside each copy."""
    inst: list[str] = []
    loop: list[str] = []
    leave: list[str] = []
    inst_fit = loop_fit = "partly"
    if key or lists:
        inst_fit = "yes"
        if key and key.get("distinct") == 2:
            inst_fit = "partly"
            inst.append("two copies: an IF or a loop of 2 does the same")
        if key and key.get("type") != "text":
            inst_fit = "partly"
            inst.append(f"column {key['column']} holds numbers: check that its values are groups, not amounts")
        if copies > MAX_COPIES_INSTANCES:
            inst_fit = "no"
            inst.append(f"{copies} copies are more than instances are meant for ({MAX_COPIES_INSTANCES})")
        elif copies > MANY_COPIES:
            inst_fit = "partly"
            inst.append("many copies make a run slower; Mind's docs suggest /InstanceSelect groups for models with lots of instances "
                        "(the team's own note: instances use more memory and processor than loops)")
        if cross_reads:
            inst_fit = "partly" if inst_fit != "no" else inst_fit
            inst.append("these cells need MM_INSTANCE(cell, key) in the main workbook, one call per formula")
    else:
        inst.append("no list or column holds the copies' names yet: make a table with one row per copy (its first column is the key)")
        if edits.get("numbers_count") or edits.get("formula_numbers_count"):
            inst.append("the copies' own numbers go into that table, one row per copy")
    if rd:
        if rd["all_slots"]["count"]:
            inst.append(ADDERS_REASON.format(n=rd["all_slots"]["count"]))
        if rd["keyed_slot"]["count"]:
            inst.append(f"{rd['keyed_slot']['count']} cells read their own key's slot: in a copy they are the copy's own block, nothing to write")
        if rd["fixed_slot"]["count"]:
            inst.append(f"{rd['fixed_slot']['count']} cells read one slot: MM_INSTANCE(cell, key) in the main workbook")
    if key:
        grouped = bool(rd and (rd["all_slots"]["count"] or rd["keyed_slot"]["count"]))
        loop_fit = "partly" if grouped else "yes"
        loop.append(f"one loop of {key.get('rows') or copies} passes (the Mind loop panel); instances cut it to {biggest or '?'} passes per copy")
        if grouped:
            loop.append("one loop over every row adds every group into one block: the groups are lost")
    elif has_blocks:
        if not edits.get("logic_count") and not edits.get("size_count"):
            loop_fit = "yes"
            loop.append("one block and MM_LOOP over the copies; the cells that read a different input read the loop index instead")
        else:
            if edits.get("logic_count"):
                loop.append(f"{edits['logic_count']} cells of different logic each need IF(MM_DIMINDEX(...))")
            if edits.get("size_count"):
                loop.append("copies of different size: a plain loop would have to run the biggest size and pad the shorter copies; instances with MM_LOOPINSTANCE inside fit better")
                if key or lists:
                    inst_fit = "yes" if inst_fit == "partly" and copies <= MANY_COPIES else inst_fit
    else:
        loop.append("the rows that read one slot each become one row per copy; a loop needs the copies as a list")
    if saved < SMALL_GAIN:
        leave_fit = "yes"
        leave.append("little to gain")
    else:
        leave_fit = "partly"
        leave.append(f"keeps {saved} formula cells to maintain")
    return {"instances": {"fits": inst_fit, "reasons": inst}, "loop": {"fits": loop_fit, "reasons": loop}, "leave": {"fits": leave_fit, "reasons": leave}}


def _headline(key: dict[str, Any] | None, kinds: list[str], bands: list[dict[str, Any]], sheets_g: list[dict[str, Any]], steps: list[dict[str, Any]],
              fills: list[dict[str, Any]], keyeds: list[dict[str, Any]], edits: dict[str, Any], copies: int) -> str:
    if key:
        what = key.get("header") or f"column {key['column']}"
        if key.get("distinct") == 2 and "macro" not in kinds:
            return f"Column {key['column']} ({what}) has 2 values: little to gain from 2 copies."
        text = f"Column {key['column']} ({what}) splits the {key.get('rows', 0):,} rows into {key.get('distinct', 0)} groups"
        if "macro" in kinds:
            return text + "; the macro already groups its output by it."
        if "lookups" in kinds:
            return text + f"; {key['keyed_formulas']} formulas look values up by it."
        if "branches" in kinds:
            return text + f"; {key['branch_formulas']} formulas branch on it."
        return text + "."
    if bands:
        b = bands[0]["f"]
        unit = "columns" if b["axis"] == "columns" else "rows"
        return f"Sheet {b['sheet']}: {b['copies']} blocks of {b['width']} {unit} compute the same, each from a different input: one block and a loop, or {b['copies']} instances."
    if sheets_g:
        g = sheets_g[0]["g"]
        kept = edits.get("logic_count") or 0
        return (f"Sheets {', '.join(g['members'])} are copies ({g['min_similarity']:.0%} the same): one sheet and {g['copies']} instances"
                + (f", {kept} cell{'s' if kept != 1 else ''} kept per copy." if kept else "."))
    if steps:
        g = steps[0]["g"]
        unit = g.get("axis") or "cells"
        return f"{_qs(g['sheet'])}: {len(g.get('positions') or [])} {unit} each read one slot of the macro's paste area: one row per copy in the main workbook."
    if fills:
        c, per, axis = _fill_copies(fills[0])
        return f"{_qs(fills[0]['f']['sheet'])} {fills[0]['f']['where']}: {c} {axis} each read one slot of an area: one row per copy, or {c} instances."
    if keyeds:
        g = keyeds[0]["g"]
        return f"{_qs(g['sheet'])}: {g['cells']} cells read the slot of the current key: in a copy they are the copy's own block."
    return f"{copies} copies."


def _dimension(found: _Found, members: list[str], book: Any, wb: Any) -> dict[str, Any] | None:
    f_of = found.findings
    kinds_found = {f_of[m]["kind"] for m in members if m in f_of}
    if not kinds_found & set(DIMENSION_KINDS):
        return None
    by_kind = lambda k: [f_of[m] for m in members if m in f_of and f_of[m]["kind"] == k]  # noqa: E731
    cands = [x["cand"] for x in by_kind("col")]
    bands, sheets_g, steps, keyeds = by_kind("band"), by_kind("sheets"), by_kind("stepping"), by_kind("keyed")
    fills = [f for f in by_kind("fill") if not f.get("covered") and not f.get("adders")]
    bands = [b for b in bands if not b.get("covered") and not b.get("adders")]
    lists = sorted((found.lists[m] for m in members if m in found.lists), key=lambda x: (not x.get("name"), x["ref"]))
    if not (cands or bands or sheets_g or steps or keyeds or fills):
        return None  # only the slots area and fills the reader lines already cover: nothing of its own (find_instances notes the readers)
    key = max(cands, key=lambda c: (bool(c.get("macro")), c.get("keyed_formulas", 0), c.get("branch_formulas", 0), -c["col"])) if cands else None
    rd = found.readers if (found.slots_area in members and found.readers) else None
    computed_fills = [f for f in fills if ((f.get("step") or {}).get("computed") or {}).get("form") != "match"]
    keyed_fills = [f for f in fills + bands if ((f.get("step") or {}).get("computed") or {}).get("form") == "match"]
    kinds: list[str] = []
    if key and key.get("macro"):
        kinds.append("macro")
    if key and key.get("keyed_formulas"):
        kinds.append("lookups")
    if key and key.get("branch_formulas", 0) >= MIN_BRANCHES:
        kinds.append("branches")
    if steps or computed_fills:
        kinds.append("stepping")
    if keyeds or keyed_fills:
        kinds.append("keyed")
    if lists:
        kinds.append("list")
    if bands:
        kinds.append("bands")
    if sheets_g:
        kinds.append("sheets")
    if key and not {"macro", "lookups", "branches"} & set(kinds):
        kinds.append("counts")
    kinds.sort(key=KIND_ORDER.index)
    # blocks and edits
    blocks: list[dict[str, Any]] = []
    edit_parts: list[dict[str, Any]] = []
    for b in bands:
        f = b["f"]
        st = b.get("step") or {}
        blocks.append({"sheet": f["sheet"], "axis": f["axis"], "where": f["where"], "copies": f["copies"], "per_copy": f["per_copy"],
                       "computed": ((st.get("computed") or {}).get("text")), "diff": f.get("diff"), "chain_share": f.get("chain_share", 0.0)})
        edit_parts.append(edits_for(book, f, wb))
    for g in sheets_g:
        gg = g["g"]
        blocks.append({"sheet": gg["members"][0], "axis": "sheets", "where": ", ".join(gg["members"]), "copies": gg["copies"], "per_copy": gg["per_copy"],
                       "computed": None, "diff": gg.get("diff"), "chain_share": 0.0})
        edit_parts.append(edits_for(book, gg, wb))
    scored = list(blocks)  # bands and sheet groups: what a loop or instances would save
    for f in fills:
        copies_f, per_f, axis_f = _fill_copies(f)
        blocks.append({"sheet": f["f"]["sheet"], "axis": axis_f, "where": f["f"]["where"], "copies": copies_f, "per_copy": per_f,
                       "computed": ((f.get("step") or {}).get("computed") or {}).get("text"), "diff": f["f"].get("diff"), "chain_share": f["f"].get("chain_share", 0.0)})
    edits = _merge_edits(edit_parts)
    if key and key.get("branch_formulas", 0) >= MIN_BRANCHES:
        edits["branches"] = {"count": key["branch_formulas"], "action": "move the values chosen by these IFs into a table with one row per copy"}
    # copies, biggest copy, what is saved and what changes
    if key:
        copies = int(key.get("distinct") or 0)
        biggest = int(key["biggest"]["rows"]) if key.get("biggest") else None
    else:
        options = [b["copies"] for b in blocks] + [len(x["texts"] or []) for x in lists] + [len(s["g"].get("positions") or []) for s in steps] + ([rd["n_slots"]] if rd else [])
        copies = max(options) if options else 0
        pers = [b["per_copy"] for b in blocks] + [s["g"].get("per_copy") or 0 for s in steps]
        biggest = max(pers) if pers else None
    saved = sum(max(0, (b["copies"] - 1) * b["per_copy"]) for b in scored)
    if rd:
        saved += int(rd["computed_slot"]["count"])
    else:
        saved += sum(int(f["f"].get("n_cells") or 0) for f in computed_fills)
    changed = edits["numbers_count"] + edits["formula_numbers_count"] + edits["logic_count"]
    if rd:
        changed += int(rd["all_slots"]["count"]) + int(rd["fixed_slot"]["count"])
    if edits.get("branches"):
        changed += int(edits["branches"]["count"])
    cross: list[dict[str, Any]] = [{"from": c["from"], "to": ", ".join(c["to"]) if isinstance(c.get("to"), list) else c.get("to"), "cells": c["cells"]} for c in rd["cross_reads"]] if rd else []
    for g in sheets_g:
        cross.extend({"from": c.get("from"), "to": c.get("to"), "cells": c.get("cells")} for c in g["g"].get("cross_reads") or [])
    # evidence, in a fixed order
    evidence: list[dict[str, Any]] = []
    counts_lines: list[dict[str, Any]] = []  # the weakest lines go last, so the cap cuts them first
    if key:
        cid = f"col:{_qs(key['sheet'])}!{key['column']}"
        # the block that reads its own key's slot (OFFSET + MATCH on this key) is a keyed lookup by the rule and already a reader line: said, not counted twice
        own_block = sum(int(g.get("cells") or 0) for g in (rd or {}).get("groups") or [] if g.get("kind") == "keyed" and found.col_of(g.get("key"), g["sheet"]) == cid)
        for e in key.get("evidence") or []:
            if e["kind"] == "header":
                continue  # a label is never evidence (weight 0)
            line = {"kind": e["kind"], "text": e["text"], "cells": list(e.get("cells") or [])[:MAX_CELLS]}
            if e.get("count") is not None:
                line["count"] = e["count"]
            if e.get("vba"):
                line["vba"] = [str(v)[:160] for v in e["vba"][:MAX_VBA]]
            if e["kind"] == "lookups" and 0 < own_block <= int(e.get("count") or 0):
                line["text"] += f"; {own_block} of them are the block that reads its own key's slot"
            (counts_lines if e["kind"] == "counts" else evidence).append(line)
        for other in cands:
            if other is not key:
                counts_lines.append({"kind": "counts", "text": f"column {other['column']} ({other.get('header') or 'no header'}) joins the same dimension with {other['distinct']} values",
                                     "cells": [f"{_qs(other['sheet'])}!{other['column']}{(found.table or {}).get('first_row') or ''}"], "count": other["distinct"]})
    area_text = f"{_qs(found.readers['area']['sheet'])}!{_rect_a1(found.readers['area']['r1'], found.readers['area']['r2'], found.readers['area']['c1'], found.readers['area']['c2'])}" if rd else ""
    for s in steps:
        g = s["g"]
        pos = g.get("positions") or []
        unit = g.get("axis") or "cells"
        where = f"rows {pos[0]}-{pos[-1]}" if unit == "rows" and pos else (f"{num_to_col(pos[0])}:{num_to_col(pos[-1])}" if unit == "columns" and pos else "")
        how = "COUNTA" if g.get("form") == "counta" else g.get("text") or g.get("form")
        evidence.append({"kind": "stepping", "text": f"{_qs(g['sheet'])} {where}: {len(pos)} {unit} each read one {g['w']}-{'column' if g['slot_axis'] == 'columns' else 'row'} slot of {area_text} ({g.get('fn') or 'OFFSET'} with {how})",
                         "cells": [g["example"]], "count": len(pos), "per_copy": g.get("per_copy")})
    covered = {(s["g"]["sheet"], tuple(s["g"].get("positions") or [])) for s in steps}
    for f in computed_fills:
        c, per, axis = _fill_copies(f)
        st = f["step"]
        if (f["f"]["sheet"], tuple(sorted({(r if axis == "rows" else cc) for lst in (f["f"].get("cells") or {}).values() for r, cc in lst}))) in covered:
            continue
        target = f"{_qs(st['sheet'])}!{_rect_a1(st['r1'], st['r2'], st['c1'], st['c2'])}"
        evidence.append({"kind": "stepping", "text": f"{_qs(f['f']['sheet'])} {f['f']['where']}: {c} {axis} each read one slot of {target} ({st['computed'].get('call') or 'OFFSET'} with {st['computed'].get('text')})",
                         "cells": _example_cell(f["f"], 0), "count": c, "per_copy": per})
    for k in keyeds:
        g = k["g"]
        own = g.get("own")
        evidence.append({"kind": "keyed", "text": f"{_qs(g['sheet'])}!{_rect_a1(*own) if own else g['example']} reads the current key's slot and adds the pass to it: in a copy it is the copy's own block",
                         "cells": [g["example"]], "count": g["cells"]})
    for f in keyed_fills:
        st = f["step"]
        evidence.append({"kind": "keyed", "text": f"{_qs(f['f']['sheet'])} {f['f']['where']} reads the slot of a key ({st['computed'].get('text')}): in a copy it is the copy's own block",
                         "cells": _example_cell(f["f"], 0), "count": int(f["f"].get("n_cells") or 0)})
    with_values = [x for x in lists if x["texts"]]  # a list too big to read, or empty, gets no line (the note says so)
    if with_values:
        first = with_values[0]
        n = len(first["texts"])
        text = f"the list {first['name'] + ' ' if first.get('name') else ''}({first['ref']}) holds the {n} values the macro matches into" if key and key.get("macro") else f"the list {first['name'] + ' ' if first.get('name') else ''}({first['ref']}) holds the {n} values of the copies"
        others = [x["ref"] for x in with_values[1:]]
        if others:
            text += f"; {', '.join(others[:3])} hold{'s' if len(others) == 1 else ''} the same values"
        r = parse_ref(first["ref"])
        evidence.append({"kind": "list", "text": text, "cells": [f"{_qs(r['sheet'])}!{num_to_col(r['c1'])}{r['r1']}"] if r and r["sheet"] else [], "count": n})
    for b in bands:
        f = b["f"]
        unit = "columns" if f["axis"] == "columns" else "rows"
        st = b.get("step") or {}
        reads = f" (each from a different input: {_qs(st['sheet'])}!{_rect_a1(st['r1'], st['r2'], st['c1'], st['c2'])})" if st else ""
        evidence.append({"kind": "bands", "text": f"{_qs(f['sheet'])} {f['where']}: {f['copies']} blocks of {f['width']} {unit} compute the same{reads}", "cells": _example_cell(f, 0), "count": f["copies"], "per_copy": f["per_copy"]})
    for g in sheets_g:
        gg = g["g"]
        evidence.append({"kind": "sheets", "text": f"sheets {', '.join(gg['members'])} are copies ({gg['min_similarity']:.0%} the same, {gg['per_copy']} formulas each)"
                         + (f"; {gg['adding_cells']} cells on other sheets add them up ({gg['adding_example']})" if gg.get("adding_cells") else ""), "cells": [], "count": gg["copies"], "per_copy": gg["per_copy"]})
    evidence += counts_lines
    truncated_evidence = len(evidence) > MAX_EVIDENCE
    evidence = evidence[:MAX_EVIDENCE]
    # name, score, fit
    if key:
        label, label_from = (key.get("header") or f"column {key['column']}"), ("header" if key.get("header") else "address")
    elif lists and lists[0].get("name"):
        label, label_from = lists[0]["name"], "defined name"
    elif bands and bands[0]["f"].get("header") == "text":
        label, label_from = f"{bands[0]['f']['sheet']} {bands[0]['f']['where']}", "header hint"
    elif sheets_g:
        label, label_from = ", ".join(sheets_g[0]["g"]["members"]), "address"
    else:
        src = (bands or fills or steps or keyeds)[0]
        src_f = src.get("f") or src.get("g")
        label, label_from = f"{src_f['sheet']} {src_f.get('where') or src_f.get('example') or ''}".strip(), "address"
    score = _score(key, kinds, scored, edits)
    fit = _fit(key, lists, rd, copies, biggest, saved, edits, cross, bool(bands or sheets_g))
    headline = _headline(key, kinds, bands, sheets_g, steps, fills, keyeds, edits, copies)
    key_block = None
    if key:
        table = found.table or {}
        key_block = {"sheet": key["sheet"], "column": key["column"], "header": key.get("header"), "table": table.get("ref"), "distinct": key["distinct"], "rows": key["rows"],
                     "biggest": key.get("biggest"), "smallest": key.get("smallest"), "balance": key.get("balance"), "counts": key.get("counts") or {}, "type": key.get("type"),
                     "keyed_formulas": key.get("keyed_formulas", 0), "keyed_tables": list(key.get("keyed_tables") or [])[:6], "mapped_through": list(key.get("mapped_through") or []),
                     "branch_formulas": key.get("branch_formulas", 0), "branch_literals_count": key.get("branch_literals_count", 0), "literals_are_values": key.get("literals_are_values", 0)}
    readers_block = {k: {"count": rd[k]["count"], "example": rd[k]["example"], "text": rd[k]["text"]} for k in ("all_slots", "fixed_slot", "computed_slot", "keyed_slot", "follow_adders")} if rd else None
    return {"id": None, "rank": None, "score": score, "label": label, "label_from": label_from, "headline": headline, "kinds": kinds, "key": key_block,
            "copies": copies, "biggest_copy": biggest, "formulas_saved": int(saved), "cells_changed": int(changed), "evidence": evidence, "blocks": blocks,
            "readers": readers_block, "cross_reads": cross[:10], "edits": edits, "fit": fit, "assistant_rank": None, "assistant_text": None, "assistant_edits": [],
            "_truncated_evidence": truncated_evidence, "_members": members}


# --- the entry point -----------------------------------------------------------------------------------
def _switches(model_map: dict[str, Any], wb0: dict[str, Any]) -> list[dict[str, Any]]:
    """The run switches the model map found, with how many formulas read each: by its defined name, or
    by its address written whole (B3, $B$3, B$3 -- never B30 or AB3) on the switch's own sheet, or with
    the sheet in front from any sheet."""
    out: list[dict[str, Any]] = []
    sws = ((model_map.get("run_settings") or {}).get("switches") or [])[:MAX_SWITCHES]
    if not sws:
        return out
    formulas = [(str(f.get("sheet") or "").upper(), str(f.get("formula") or "").upper()) for f in wb0.get("formulas") or []]
    for sw in sws:
        name = str(sw.get("name") or "").strip()
        sheet, ref = _split(str(sw.get("cell") or ""))
        r = parse_ref(ref) if ref else None
        quick: list[str] = []  # cheap substring tests first; the exact word tests only on the formulas that pass
        exact: list[tuple[re.Pattern, bool]] = []  # (pattern, own sheet only)
        if name:
            quick.append(name.upper())
            exact.append((re.compile(rf"(?<![A-Za-z0-9_.])(?:{re.escape(name)})(?![A-Za-z0-9_(])", re.I), False))
        if r:
            letters, row = num_to_col(r["c1"]), r["r1"]
            quick += [f"{letters}{row}", f"${letters}${row}", f"{letters}${row}", f"${letters}{row}"]
            cell = rf"\$?{letters}\$?{row}(?![0-9A-Za-z_(])"
            exact.append((re.compile(r"(?<![A-Za-z0-9_.$!'])" + cell, re.I), True))
            if sheet:
                quoted = sheet.replace("'", "''")
                exact.append((re.compile(rf"(?:'{re.escape(quoted)}'|{re.escape(sheet)})!" + cell, re.I), False))
        own = str(sheet or "").upper()
        readers = 0
        for f_sheet, text in formulas:
            if not any(q in text for q in quick):
                continue
            if any(p.search(text) for p, own_only in exact if not own_only or f_sheet == own):
                readers += 1
        out.append({"cell": sw.get("cell"), "name": sw.get("name") or None, "readers": readers})
    return out


def _file_name(wb0: dict[str, Any]) -> str:
    name = str(wb0.get("file_name") or Path(str(wb0.get("copy_path") or "model")).stem)
    ext = str(wb0.get("file_type") or "").lower().lstrip(".")
    return name if "." in name or not ext else f"{name}.{ext}"


def find_instances(analysis: dict[str, Any], model_map: dict[str, Any] | None, *, loop_plan: dict[str, Any] | None = None, assistant: bool = False,
                   complete: Callable[..., Any] | None = None, budget_s: float = 120) -> dict[str, Any]:
    """Where Mind instances or a loop would help, ranked, with evidence; nothing written (section 7 of
    the design is the result's shape). The time budget starts after the workbook is open: a reload of a
    dropped workbook is reported in `counts.reload_seconds`, not charged. `assistant=True` adds the
    assistant's explanation and order on top; `complete` replaces the gateway call in tests."""
    t_load = time.monotonic()
    wb_note = None
    try:
        wb = open_cached(analysis)
    except Exception as exc:  # no copy, a deleted copy, a file openpyxl cannot read: the detectors that need only the analysis still run
        wb, wb_note = None, f"the workbook copy could not be opened ({type(exc).__name__}): typed-in numbers and the key column's values were not read"
    reload_s = round(time.monotonic() - t_load, 1)
    budget = Budget(budget_s)
    budget.restart()
    wb0 = analysis["workbooks"][0]
    names = _Names(wb0.get("defined_names") or [])
    model_map = model_map or {}
    if loop_plan is None:
        loop_plan = build_loop_plan(analysis, model_map)
    found = _Found(names)
    skipped: list[str] = []
    if reload_s > RELOAD_NOTE_S:
        skipped.append(f"the workbook had to be read again ({reload_s:.0f} s): the app keeps two workbooks open and this one had been dropped")
    if wb_note:
        skipped.append(wb_note)
    with formula_memo():
        book = build_book(analysis, budget)
        grouping = find_grouping(analysis, model_map, loop_plan, names, wb, budget)
        found.table = grouping.get("table")
        found.ignored["label_columns"] = int(grouping.get("identifier_columns") or 0)
        _p1_anchors(found, grouping, names)
        _cleared_areas(found, loop_plan, grouping.get("macro"))
        _reader_anchors(found, analysis, book, names, grouping, budget)
        _repetitions(found, book, names, budget)
        _value_edges(found, analysis)
        dims: list[dict[str, Any]] = []
        for members in found.uf.groups():
            d = _dimension(found, members, book, wb)
            if d is not None:
                dims.append(d)
    dims.sort(key=lambda d: (-d["score"], -d["formulas_saved"], d["label"]))
    rd = found.readers
    if rd and not any(d.get("readers") for d in dims):
        parts = [f"{rd[k]['count']} {what}" for k, what in (("all_slots", "add every slot up"), ("fixed_slot", "read one slot"), ("computed_slot", "read a slot by row or column"), ("keyed_slot", "read a key's slot")) if rd[k]["count"]]
        if parts:
            total = sum(rd[k]["count"] for k in ("all_slots", "fixed_slot", "computed_slot", "keyed_slot"))
            skipped.append(f"{total} cells read the macro's paste area but no key column, list or stepping block joins them ({'; '.join(parts)}): they are in this note only")
    truncated = dict(found.truncated)
    for d in dims:
        if d.pop("_truncated_evidence", False):
            truncated["evidence"] = True
        d.pop("_members", None)
    n_dims = len(dims)
    if n_dims > MAX_DIMENSIONS:
        truncated["dimensions"] = True
        dims = dims[:MAX_DIMENSIONS]
    for i, d in enumerate(dims, 1):
        d["id"], d["rank"] = f"dim-{i}", i
    chains = sorted(found.chains, key=lambda c: (-(c["copies"] * c["per_copy"]), c["label"]))
    if len(chains) > MAX_CHAINS:
        truncated["chains"] = True
        chains = chains[:MAX_CHAINS]
    macro_loops = [{"name": le.get("name"), "driver": le.get("driver"), "kind": le.get("kind"), "status": le.get("status"), "first": le.get("first"), "last": le.get("last"),
                    "handled_by": "the Mind loop step"} for le in loop_plan.get("loops") or []]
    builder = dict(grouping.get("builder") or {"column": None, "header": None, "chosen_by": None, "candidates": [], "matches_top": None, "source": "rule"})
    top = next((d for d in dims if d.get("key")), None)
    builder["matches_top"] = (top["key"]["column"] == builder["column"]) if (top and builder.get("column")) else None  # None: nothing to compare
    skipped += list(grouping.get("notes") or [])
    skipped += list(budget.skipped)
    skipped += list(found.notes)
    counts = {"sheets": len(wb0.get("sheets") or []), "formulas": len(wb0.get("formulas") or []), "dimensions": n_dims, "chains": len(found.chains),
              "time_axes": found.ignored["time_axes"], "seconds": round(budget.elapsed(), 1), "reload_seconds": reload_s}
    result = {"schema": SCHEMA, "version": _version(), "file": _file_name(wb0), "counts": counts, "dimensions": dims, "chains": chains, "macro_loops": macro_loops,
              "switches": _switches(model_map, wb0), "builder": builder, "ignored": dict(found.ignored), "ignored_detail": {"filled_areas": found.filled_detail},
              "skipped": skipped, "truncated": truncated, "assistant": dict(EMPTY_ASSISTANT)}
    result = _jsonable(result)
    if assistant:
        advise = make_finder_advisor(complete=complete)
        block = advise(result) if advise is not None else {**EMPTY_ASSISTANT, "asked": True, "available": False, "message": NO_KEY_MESSAGE}
        result = merge_assistant(result, block)
    return result


# --- words for a person --------------------------------------------------------------------------------
def describe(result: dict[str, Any]) -> str:
    lines: list[str] = []
    dims = result.get("dimensions") or []
    for d in dims:
        lines.append(f"#{d['rank']} {d['label']}: {d['headline']}")
        for e in d.get("evidence") or []:
            lines.append(f"   {e['text']}" + (f"  ({', '.join(e['cells'])})" if e.get("cells") else ""))
        rd = d.get("readers") or {}
        for k in ("all_slots", "computed_slot", "keyed_slot", "fixed_slot", "follow_adders"):
            if (rd.get(k) or {}).get("count"):
                lines.append(f"   {rd[k]['text']}")
        for opt, word in (("instances", "Instances"), ("loop", "A loop"), ("leave", "Leave as is")):
            f = (d.get("fit") or {}).get(opt) or {}
            lines.append(f"   {word}: {f.get('fits')}" + (f" -- {'; '.join(f['reasons'])}" if f.get("reasons") else ""))
    if not dims:
        lines.append("Nothing repeats enough to split: leave the model as it is.")
    if result.get("chains"):
        lines.append("Set aside (chains): " + "; ".join(f"{c['label']} ({c['copies']} copies: {c['why']})" for c in result["chains"]))
    if result.get("macro_loops"):
        lines.append("Macro loops the Mind loop step handles: " + ", ".join(f"{l['name']} ({l['kind']}, {l['status']})" for l in result["macro_loops"]))
    detail = (result.get("ignored_detail") or {}).get("filled_areas") or []
    if detail:
        lines.append("Filled areas that join nothing: " + ", ".join(f"{x['sheet']}!{x['where']} ({x['cells']} cells)" for x in detail))
    b = result.get("builder") or {}
    if b.get("column"):
        tag = " (the builder's rule)" if b.get("source") == "rule" else ""
        top = next((d for d in dims if d.get("key")), None)
        if b.get("matches_top"):
            verdict = "the same as the top dimension"
        elif top is not None:
            verdict = f"DIFFERENT: the top dimension is column {top['key']['column']} ({top['key'].get('header') or 'no header'})"
        else:
            verdict = "no column dimension to compare with"
        lines.append(f"Builder's key: {b['column']} {b.get('header') or ''} ({b.get('chosen_by')}){tag} -- {verdict}")
    for s in result.get("skipped") or []:
        lines.append(f"Note: {s}")
    return "\n".join(lines)


# --- the assistant pass --------------------------------------------------------------------------------
EMPTY_ASSISTANT: dict[str, Any] = {"asked": False, "available": False, "text": None, "ranking": None, "explanations": {}, "edits": {}, "extra": [], "message": None}
NO_KEY_MESSAGE = "No AI key is saved yet: paste one under Advanced > Assistant."
NOT_JSON_MESSAGE = "The assistant's answer was not the JSON asked for"
NO_ANSWER_MESSAGE = "The assistant gave no usable answer"
EXTRA_FLAG = "suggested by the assistant, not checked"
PROMPT_LIMIT = 12_000

SYSTEM_PROMPT = """You help prepare an Excel model (usually actuarial) for Milliman Mind. Mind can run one copy of a model per value of a key (instances) or one block in a loop. A tool found where the model repeats itself; its facts are checked and must not be contradicted. For each dimension explain in two plain sentences for a non-programmer what repeats, what would change, and which option fits best. Rank the dimensions by how much work they save and how safe the change is. For near copies propose the edits that unify them. You may add candidates the tool did not find; they are shown as 'suggested by the assistant, not checked'. Answer with JSON only, no prose around it:
{"ranking": [1, 3, 2], "explanations": {"1": "…"}, "edits": {"2": ["move the three rates in rows 1-3 into a table with one row per product"]},
 "extra": [{"headline": "…", "why": "…", "cells": ["Sheet!A1"]}]}"""


def _clip(text: Any, n: int = 600) -> str:
    s = str(text or "")
    return s if len(s) <= n else s[: n - 1] + "…"


def _prompt(result: dict[str, Any]) -> str:
    """What the assistant sees: the dimensions (at most 8) with their evidence, the reader classes, the key's
    top counts and branch literals, the VBA lines; then the chains and the builder's choice. Clipped."""
    parts: list[str] = []
    for d in (result.get("dimensions") or [])[:8]:
        lines = [f"### Dimension {d['rank']}", f"headline: {_clip(d['headline'])}", f"kinds: {', '.join(d.get('kinds') or [])}",
                 f"copies: {d.get('copies')}; biggest copy: {d.get('biggest_copy')}; formulas saved: {d.get('formulas_saved')}; cells changed: {d.get('cells_changed')}"]
        for e in d.get("evidence") or []:
            lines.append(f"- {_clip(e.get('text'))}" + (f" [{', '.join(e.get('cells') or [])}]" if e.get("cells") else ""))
            for v in e.get("vba") or []:
                lines.append(f"  VBA: {_clip(v, 160)}")
        rd = d.get("readers") or {}
        for k in ("all_slots", "fixed_slot", "computed_slot", "keyed_slot", "follow_adders"):
            if (rd.get(k) or {}).get("count"):
                lines.append(f"- {_clip(rd[k]['text'])}")
        key = d.get("key") or {}
        if key.get("counts"):
            lines.append("key values and rows: " + _clip(", ".join(f"{k}: {v}" for k, v in list(key["counts"].items())[:12])))
        for b in (d.get("blocks") or [])[:3]:
            lines.append(f"block: {b.get('sheet')} {b.get('where')} ({b.get('copies')} copies, {b.get('per_copy')} formulas each" + (f", index {b['computed']}" if b.get("computed") else "") + ")")
        ed = d.get("edits") or {}
        if any(ed.get(f"{k}_count") for k in ("numbers", "formula_numbers", "logic", "size")):
            lines.append(f"edits: {ed.get('numbers_count', 0)} numbers, {ed.get('formula_numbers_count', 0)} numbers inside formulas, {ed.get('logic_count', 0)} logic cells, {ed.get('size_count', 0)} size differences")
            for x in (ed.get("logic") or [])[:3]:
                lines.append(f"  logic {x.get('cell')}: " + _clip(" | ".join(str(f) for f in x.get("formulas_per_copy") or []), 300))
        if ed.get("branches"):
            lines.append(f"branches: {ed['branches'].get('count')} formulas choose a value by the key")
        for opt in ("instances", "loop", "leave"):
            f = (d.get("fit") or {}).get(opt) or {}
            lines.append(f"{opt}: {f.get('fits')}" + (f" ({'; '.join(_clip(r, 200) for r in f.get('reasons') or [])})" if f.get("reasons") else ""))
        parts.append("\n".join(lines))
    if result.get("chains"):
        parts.append("### Chains (set aside)\n" + "\n".join(f"- {c['label']}: {c['copies']} copies, {c['why']}" for c in result["chains"]))
    b = result.get("builder") or {}
    if b.get("column"):
        parts.append(f"### Builder's choice\ncolumn {b['column']} ({b.get('header')}), chosen by {b.get('chosen_by')}; the same as the top dimension: {b.get('matches_top')}")
    text = "\n\n".join(parts) + "\n\nGive the JSON now."
    return text if len(text) <= PROMPT_LIMIT else text[: PROMPT_LIMIT - 20] + "\n\nGive the JSON now."


def _first_json_object(text: str) -> str | None:
    """The first balanced {...} of a text, braces inside string literals ignored."""
    start = text.find("{")
    while start != -1:
        depth, in_str, esc = 0, False, False
        for i in range(start, len(text)):
            ch = text[i]
            if in_str:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return text[start:i + 1]
        start = text.find("{", start + 1)
    return None


def parse_assistant(text: str | None, result: dict[str, Any], available: bool = True, message: str | None = None) -> dict[str, Any]:
    """The assistant block out of the model's reply: the first JSON object, validated against the result's
    dimensions; anything else is dropped, never guessed."""
    block = {**EMPTY_ASSISTANT, "asked": True, "available": available, "message": message}
    if not available or not text:
        return block
    raw = _first_json_object(str(text))
    try:
        data = json.loads(raw) if raw else None
    except ValueError:
        data = None
    if not isinstance(data, dict):
        block["message"] = NOT_JSON_MESSAGE
        return block
    ids = {d["rank"] for d in result.get("dimensions") or []}
    ranking = data.get("ranking")
    if isinstance(ranking, list) and ranking:
        ok = all(isinstance(x, (int, float)) and not isinstance(x, bool) and int(x) == x and int(x) in ids for x in ranking)
        clean = [int(x) for x in ranking] if ok else []
        block["ranking"] = clean if clean and len(clean) == len(set(clean)) else None  # every entry an existing dimension, each once, else nothing
    expl = data.get("explanations")
    if isinstance(expl, dict):
        block["explanations"] = {str(k): _clip(v, 600) for k, v in expl.items() if str(k).isdigit() and int(k) in ids and isinstance(v, str) and v.strip()}
    edits = data.get("edits")
    if isinstance(edits, dict):
        block["edits"] = {str(k): [_clip(e, 300) for e in v if isinstance(e, str) and e.strip()][:8] for k, v in edits.items() if str(k).isdigit() and int(k) in ids and isinstance(v, list)}
        block["edits"] = {k: v for k, v in block["edits"].items() if v}
    extra = data.get("extra")
    if isinstance(extra, list):
        kept = []
        for item in extra[:5]:
            if not isinstance(item, dict) or not isinstance(item.get("headline"), str) or not isinstance(item.get("why"), str):
                continue
            cells = []
            for c in (item.get("cells") or [])[:6]:
                if isinstance(c, str) and parse_ref(_split(c)[1]) and _split(c)[0]:
                    cells.append(c)
            kept.append({"headline": _clip(item["headline"], 300), "why": _clip(item["why"], 600), "cells": cells, "flag": EXTRA_FLAG})
        block["extra"] = kept
    if block["explanations"]:
        block["text"] = " ".join(block["explanations"][k] for k in sorted(block["explanations"], key=int))
    if not (block["ranking"] or block["explanations"] or block["edits"] or block["extra"]):
        block["message"] = NO_ANSWER_MESSAGE  # JSON, but none of the four fields holds anything: a failed answer, asked again next time
    return block


def make_finder_advisor(model_id: str | None = None, complete: Callable[..., Any] | None = None) -> Callable[[dict[str, Any]], dict[str, Any]] | None:
    """The assistant's pass over a finder result, or None when no assistant is configured. `complete`
    replaces app.llm.chat_completion in tests. The returned function never raises."""
    if complete is None:
        if not llm_available():
            return None
        complete = chat_completion
    model = model_id or DEFAULT_MODEL_ID

    def explain(result: dict[str, Any]) -> dict[str, Any]:
        try:
            res = complete([{"role": "user", "content": _prompt(result)}], SYSTEM_PROMPT, model_id=model, max_tokens=2000, temperature=0.0, timeout=90)
            if not isinstance(res, dict):
                return {**EMPTY_ASSISTANT, "asked": True, "available": True, "message": NO_ANSWER_MESSAGE}
            return parse_assistant(res.get("text"), result, available=bool(res.get("available", True)), message=res.get("message"))
        except Exception as exc:  # the whole body: the call, the prompt, the parsing
            return {**EMPTY_ASSISTANT, "asked": True, "available": True, "message": f"Request failed: {type(exc).__name__}"}

    return explain


def merge_assistant(result: dict[str, Any], block: dict[str, Any]) -> dict[str, Any]:
    """A deep copy of the result with the assistant block and the per-dimension assistant_* fields set.
    The deterministic dimensions, ranks, scores, evidence and fits are never changed."""
    out = deepcopy(result)
    out["assistant"] = {**EMPTY_ASSISTANT, **(block or {})}
    ranking = out["assistant"].get("ranking") or []
    for d in out.get("dimensions") or []:
        d["assistant_rank"] = ranking.index(d["rank"]) + 1 if d["rank"] in ranking else None
        d["assistant_text"] = (out["assistant"].get("explanations") or {}).get(str(d["rank"]))
        d["assistant_edits"] = list((out["assistant"].get("edits") or {}).get(str(d["rank"]), []))
    return out


# --- command line --------------------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    import argparse
    import tempfile

    from .inventory import build_analysis
    from .model_map import build_model_map

    ap = argparse.ArgumentParser(description="Where Mind instances or a loop would help a model: ranked, with evidence; nothing written.")
    ap.add_argument("workbook")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--assistant", action="store_true", help="add the assistant's explanation and order (needs a saved AI key)")
    ap.add_argument("--work", default=None)
    a = ap.parse_args(argv)
    work = Path(a.work) if a.work else Path(tempfile.mkdtemp(prefix="mind_finder_"))
    analysis = build_analysis(Path(a.workbook), work, Path(a.workbook).stem)
    mp = build_model_map(analysis)
    result = find_instances(analysis, mp, loop_plan=build_loop_plan(analysis, mp), assistant=a.assistant)
    out = sys.stdout
    try:
        out.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    text = json.dumps(result, indent=1, ensure_ascii=False) if a.json else describe(result)
    try:
        out.write(text + "\n")
    except UnicodeEncodeError:  # a console that cannot show every character of the sheet names
        out.write(text.encode("ascii", "replace").decode("ascii") + "\n")
    out.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
