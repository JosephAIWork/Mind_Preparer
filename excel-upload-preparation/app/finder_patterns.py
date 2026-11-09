"""Where a workbook repeats itself, read on formula shapes (the instance finder, 1.12.5).

On a `ShapeBook` (app/finder_shapes.py) this module finds:
  * P2 bands: the same formulas side by side in one sheet, as blocks of columns or rows
    (`bands`, `sheet_findings`), and fills: one column or row repeated many times (a time axis, a
    list of items, or a uniform area with both axes);
  * P3 sheet copies: sheets with the same layout and formulas (`sheet_copies`);
  * P4 how copies differ: identical / reads a different input / differs only in a typed-in number /
    different logic (`diff_kinds`), with the edits that would unify them (`edits_for`: typed-in
    numbers, numbers inside formulas, logic cells, cells present in some copies only);
  * chains: copies that read the copy before them (`chain_share`, `consecutive_reads`);
  * what a repetition steps through (`step_target`): differing relative references, or a computed
    index (ROW/COLUMN/COUNTA/MATCH inside OFFSET/INDEX) with the anchor resolved through the
    defined names.

A pure-Python port of the prototype in tools/structure_check/analyze.py (no numpy, no Excel), with
the fixes agreed in the design: both axes deduplicated, a uniform area is one fill with two axes,
the chain test counts relative reads only, computed indices are read. Nothing here names a client
workbook; the texts are plain words for the card.
"""
from __future__ import annotations

import random
import re
from collections import Counter
from datetime import date, datetime
from typing import Any

from .autofix import parse_r1c1, ref_rect
from .finder_shapes import Budget, ShapeBook, SheetShape, cell_name, number_literals
from .formula_utils import STORAGE_PREFIX_RE, find_calls, num_to_col, parse_ref, split_top_level_args

SIM_SHEETS = 0.75  # share of formula cells that must match for two sheets to be copies
SIM_COLS = 0.90  # share of a column's (row's) formula cells that must match its copy
MIN_FORMULAS = 10  # sheets / blocks smaller than this are not worth a finding
PMAX = 40  # widest block looked for
SAMPLE = 3000  # cells sampled per reference test
CHAIN_SHARE = 0.10  # share of sampled cells reading an earlier copy that makes a chain
MIN_FILL = 6  # positions a fill needs
BIG_FILL = 500  # cells: a fill this big is handed to the merge step
SMALL_GAIN = 200  # formula cells: below this, "little to gain"
MAX_COPIES_INSTANCES = 200
MAX_IF_FORMS = 3  # distinct formulas a logic cell may have for "IF on the copy's key"
MAX_EDITS = 40  # entries per edits list (counts stay complete)
MAX_VALUES = 12  # values / literals listed per edit
MAX_FORMS = 6  # formulas listed per logic cell
MAX_BAND_FORMULAS = 2_000_000  # a sheet with more formula cells skips the band search (announced)
PREFILTER = 0.5  # multiset overlap two sheets need before the aligned compare

NUMBERS_ACTION = "move into a table row per copy"
FORMULA_NUMBERS_ACTION = "replace the typed number by a reference to a per-copy table cell"
IF_ACTION = "IF on the copy's key"
OVERRIDE_ACTION = "override kept per copy"
SIZE_DIFFERS = "size differs"
SIZE_ACTION = ("copies of different size: instances, with MM_LOOPINSTANCE for the loop inside each copy "
               "(a plain loop would have to run the biggest size and pad the shorter copies)")
DIFF_LABELS = ("identical", "reads a different input", "differs only in a typed-in number", "different logic")


# --- stacks: the copies of a finding, aligned ---------------------------------------------------
class Stacks(list):
    """One dict per copy, {(row offset, col offset): shape id}, the same key = the same place in
    every copy. `sheets[i]` and `origins[i]` turn a key back into a cell of copy i."""

    sheets: list[str]
    origins: list[tuple[int, int]]

    def address(self, i: int, key: tuple[int, int]) -> tuple[str, int, int]:
        r0, c0 = self.origins[i]
        return self.sheets[i], r0 + key[0], c0 + key[1]


def stacks_of(book: ShapeBook, finding: dict[str, Any]) -> Stacks:
    """The aligned copies of a band / fill (`positions` per copy) or of a sheet group (`members`)."""
    if "members" in finding:
        members = list(finding["members"])
        st = Stacks(dict(book.sheets[m].cells) if m in book.sheets else {} for m in members)
        st.sheets = members
        st.origins = [(0, 0)] * len(members)
        return st
    sheet = book.sheets.get(finding["sheet"])
    axis = finding["axis"]
    st = Stacks()
    st.sheets, st.origins = [], []
    for positions in finding["positions"]:
        start = positions[0]
        origin = (0, start) if axis == "columns" else (start, 0)
        d: dict[tuple[int, int], int] = {}
        if sheet is not None:
            for p in positions:
                line = (sheet.cols if axis == "columns" else sheet.rows).get(p)
                if not line:
                    continue
                if axis == "columns":
                    for row, sid in line.items():
                        d[(row, p - start)] = sid
                else:
                    for col, sid in line.items():
                        d[(p - start, col)] = sid
        st.append(d)
        st.sheets.append(finding["sheet"])
        st.origins.append(origin)
    return st


def diff_kinds(book: ShapeBook, stacks: list[dict[tuple[int, int], int]]) -> dict[str, int]:
    """Cells of the copies sorted by how the copies differ there: identical (same text everywhere),
    reads a different input (same but for the references), differs only in a typed-in number,
    different logic (anything else, including a formula in some copies only)."""
    out = {label: 0 for label in DIFF_LABELS}
    if not stacks:
        return out
    L1, L2 = book.loose[1], book.loose[2]
    keys: set[tuple[int, int]] = set()
    for s in stacks:
        keys.update(s)
    for k in keys:
        ids = [s.get(k) for s in stacks]
        first = ids[0]
        if first is None or any(i is None for i in ids):
            out["different logic"] += 1
        elif all(i == first for i in ids):
            out["identical"] += 1
        elif all(L1[i] == L1[first] for i in ids):
            out["reads a different input"] += 1
        elif all(L2[i] == L2[first] for i in ids):
            out["differs only in a typed-in number"] += 1
        else:
            out["different logic"] += 1
    return out


# --- chains -------------------------------------------------------------------------------------
def _relative_span(ref: tuple[Any, Any, Any], axis: str, row: int, col: int) -> tuple[int, int] | None:
    """The rows (columns) a reference covers along `axis` when its component on that axis is
    RELATIVE at both ends; None for an absolute component or a whole row/column (they read a fixed
    place, not "the copy before")."""
    spec, a, b = ref
    i = 0 if axis == "rows" else 1
    ea = a[i]
    if ea is None or ea[0]:
        return None
    if b is not None:
        eb = b[i]
        if eb is None or eb[0]:
            return None
    r1, r2, c1, c2 = ref_rect(ref, row, col)
    return (r1, r2) if axis == "rows" else (c1, c2)


def chain_share(book: ShapeBook, finding: dict[str, Any]) -> float:
    """Share of sampled formula cells of copies 2..k that read an EARLIER copy of the same finding
    through a relative reference on its own sheet (an absolute reference or a whole row/column reads
    a fixed cell -- a parameter inside copy 1's columns is not a chain)."""
    sheet = book.sheets.get(finding["sheet"])
    if sheet is None:
        return 0.0
    axis = finding["axis"]
    positions = finding["positions"]
    if len(positions) < 2:
        return 0.0
    lo = positions[0][0]
    cells: list[tuple[int, int, int]] = []
    all_cells = finding.get("cells") or {}
    for j in range(1, len(positions)):
        hi = positions[j][0] - 1
        for (r, c) in all_cells.get(j, ()):
            cells.append((r, c, hi))
    if len(cells) > SAMPLE:
        cells = random.Random(0).sample(cells, SAMPLE)
    if not cells:
        return 0.0
    hits = 0
    for r, c, hi in cells:
        sid = sheet.cells.get((r, c))
        if sid is None:
            continue
        for ref in book.parsed(sid).refs:
            if ref[0] is not None:
                continue
            span = _relative_span(ref, axis, r, c)
            if span is not None and span[1] >= lo and span[0] <= hi:
                hits += 1
                break
    return round(hits / len(cells), 3)


def consecutive_reads(book: ShapeBook, group: dict[str, Any]) -> tuple[bool, list[dict[str, Any]]]:
    """Do the sheet copies read each other? A CHAIN when every member after the first reads the
    member before it (any number of cells); every other link is a cross read
    [{from, to, cells}] (a value one copy needs from another: an override need, not a chain)."""
    members = list(group["members"])
    mset = set(members)
    reads: dict[str, Counter] = {m: Counter() for m in members}
    for m in members:
        sh = book.sheets.get(m)
        if sh is None:
            continue
        for sid, n in Counter(sh.cells.values()).items():
            for t in book.sheets_read(sid, m) & mset:
                reads[m][t] += n
    chain = len(members) >= 2 and all(reads[members[i]].get(members[i - 1], 0) > 0 for i in range(1, len(members)))
    idx = {m: i for i, m in enumerate(members)}
    links = [{"from": m, "to": t, "cells": int(n)} for m in members for t, n in sorted(reads[m].items(), key=lambda kv: idx[kv[0]])]
    if chain:
        links = [l for l in links if idx[l["from"]] != idx[l["to"]] + 1]
    return chain, links


def adding_cells(book: ShapeBook, members: list[str]) -> tuple[int, str | None]:
    """Formula cells on other sheets that read two or more of the members (who adds the copies
    up), and one example cell."""
    mset = set(members)
    count, example = 0, None
    for name in book.sheet_order:
        sh = book.sheets.get(name)
        if sh is None or name in mset:
            continue
        hits = {sid for sid in set(sh.cells.values()) if len(book.sheets_read(sid, name) & mset) >= 2}
        if not hits:
            continue
        for (r, c), sid in sh.cells.items():
            if sid in hits:
                count += 1
                if example is None:
                    example = cell_name(name, r, c)
    return count, example


# --- P2: repetition inside a sheet ---------------------------------------------------------------
def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _increasing(values: list[Any]) -> bool:
    nums = [v.toordinal() if isinstance(v, (date, datetime)) else v for v in values]
    return len(nums) >= 2 and all(_is_number(n) for n in nums) and all(nums[i] < nums[i + 1] for i in range(len(nums) - 1))


def _header_hint(book: ShapeBook, sheet: SheetShape, axis: str, lo: int, hi: int) -> str:
    """What the first filled row above (column left of) the repetition holds, reading from the top
    (left) of the sheet: "formulas", "numbers" (rising: a time axis of years, months, durations),
    "text" (names of items; also numbers that do not rise), or "none". Constants come from the
    workbook when it is open; without it only formula cells are seen."""
    ws = book.worksheet(sheet.name)
    const = getattr(ws, "_cells", None) if ws is not None else None
    width = hi - lo + 1
    step = max(1, width // 2000)
    positions = list(range(lo, hi + 1, step))
    need = max(2, len(positions) // 2)
    first, last = (sheet.r_lo, sheet.r_hi) if axis == "columns" else (sheet.c_lo, sheet.c_hi)
    cells = sheet.cells
    for line in range(max(1, first - 50), min(last, first + 50) + 1):  # a header sits near the block
        n_f = n_num = n_txt = 0
        nums: list[Any] = []
        for pos in positions:
            key = (line, pos) if axis == "columns" else (pos, line)
            if key in cells:
                n_f += 1
                continue
            if const is None:
                continue
            c = const.get(key)
            v = None if c is None else c.value
            if v is None or v == "":
                continue
            if isinstance(v, bool):
                n_txt += 1
            elif _is_number(v) or isinstance(v, (date, datetime)):
                n_num += 1
                nums.append(v)
            elif isinstance(v, str):
                if v.startswith("="):
                    n_f += 1
                else:
                    n_txt += 1
            else:  # ArrayFormula / DataTableFormula objects are formulas
                n_f += 1
        if n_f + n_num + n_txt >= need:
            if n_f:
                return "formulas"
            if n_num >= n_txt:
                return "numbers" if _increasing(nums) else "text"
            return "text"
    return "none"


def _match_profile(prof: dict[int, dict[int, int]], npos: int, p: int) -> list[float | None]:
    """For each position i: share of the formula cells of positions i and i+p that are equal
    (loose-1), None when both hold no formula. Positions past the end count as empty."""
    out: list[float | None] = [None] * npos
    for i in range(npos):
        a = prof.get(i)
        b = prof.get(i + p)
        if not a and not b:
            continue
        if not a or not b:
            out[i] = 0.0
            continue
        eq = common = 0
        for k, v in a.items():
            w = b.get(k)
            if w is not None:
                common += 1
                if w == v:
                    eq += 1
        out[i] = eq / (len(a) + len(b) - common)
    return out


def _runs(mask: list[bool]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    start = None
    for i, v in enumerate(list(mask) + [False]):
        if v and start is None:
            start = i
        elif not v and start is not None:
            out.append((start, i - 1))
            start = None
    return out


def _where(axis: str, first: int, last: int) -> str:
    return f"{num_to_col(first)}:{num_to_col(last)}" if axis == "columns" else f"rows {first}-{last}"


def _cells_of_positions(sheet: SheetShape, axis: str, positions: list[int]) -> list[tuple[int, int]]:
    lines = sheet.cols if axis == "columns" else sheet.rows
    out: list[tuple[int, int]] = []
    for p in positions:
        line = lines.get(p)
        if line:
            out.extend(((o, p) for o in line) if axis == "columns" else ((p, o) for o in line))
    return out


def _uniform_axes(book: ShapeBook, cells: list[tuple[int, int]], sheet: SheetShape) -> dict[str, int] | None:
    """A uniform area: one loose-1 shape in >= 90 % of its cells and >= 90 % of its rectangle
    filled -> both axes, else None."""
    if not cells:
        return None
    L1 = book.loose[1]
    top = Counter(L1[sheet.cells[rc]] for rc in cells).most_common(1)[0][1]
    rows = {r for r, _ in cells}
    cols = {c for _, c in cells}
    if len(rows) < 2 or len(cols) < 2 or top < SIM_COLS * len(cells) or len(cells) < SIM_COLS * len(rows) * len(cols):
        return None
    return {"rows": len(rows), "columns": len(cols)}


def _finding(book: ShapeBook, sheet: SheetShape, axis: str, positions: list[list[int]], width: int, fill: bool) -> dict[str, Any]:
    cells = {j: _cells_of_positions(sheet, axis, pos) for j, pos in enumerate(positions)}
    first_cell = min(cells[0]) if cells[0] else None
    all_cells = [rc for lst in cells.values() for rc in lst]
    f: dict[str, Any] = {
        "sheet": sheet.name, "axis": axis, "where": _where(axis, positions[0][0], positions[-1][-1]),
        "start": positions[0][0], "width": width, "copies": len(positions), "positions": positions,
        "per_copy": len(cells[0]), "n_cells": len(all_cells), "cells": cells, "fill": fill,
        "axes": _uniform_axes(book, all_cells, sheet) if fill else None,
        "header": _header_hint(book, sheet, axis, positions[0][0], positions[-1][-1]),
        "first_cell": first_cell,
        "first_formula_index": book.index_of(sheet.name, *first_cell) if first_cell else None,
    }
    f["chain_share"] = chain_share(book, f)
    f["diff"] = diff_kinds(book, stacks_of(book, f))
    if fill:
        f["time_axis"] = f["chain_share"] >= CHAIN_SHARE or f["header"] == "numbers"
        f["chain_shares"] = {axis: f["chain_share"]}
        f["time_axes"] = {axis: f["time_axis"]}
        f["headers"] = {axis: f["header"]}
    return f


def bands(book: ShapeBook, sheet: SheetShape, axis: str, budget: Budget | None = None) -> list[dict[str, Any]]:
    """Blocks of `width` positions (columns or rows) repeated `copies` times side by side, each
    block >= MIN_FORMULAS formula cells and >= 90 % equal (loose-1) to the next; then fills: one
    position repeated >= MIN_FILL times (`fill: True`; a uniform area carries both `axes`). Each
    finding: sheet, axis, where, start, width, copies, positions (per copy), per_copy, n_cells,
    cells {copy index: [(row, col)..]}, fill, axes, header, first_cell, first_formula_index,
    chain_share, diff, and for fills time_axis."""
    if sheet.n < MIN_FORMULAS:
        return []
    if sheet.n > MAX_BAND_FORMULAS:
        if budget is not None:
            budget.skip(f"repeated blocks on sheet {sheet.name} not checked: {sheet.n:,} formula cells is more than the finder reads for blocks")
        return []
    if budget is not None and budget.out():
        budget.skip(f"repeated blocks on sheet {sheet.name} not checked: the time budget ran out")
        return []
    lines = sheet.cols if axis == "columns" else sheet.rows
    lo, hi = (sheet.c_lo, sheet.c_hi) if axis == "columns" else (sheet.r_lo, sheet.r_hi)
    npos = hi - lo + 1
    if npos < 4:
        return []
    L1 = book.loose[1]
    prof: dict[int, dict[int, int]] = {pos - lo: {o: L1[sid] for o, sid in line.items()} for pos, line in lines.items()}
    found: list[dict[str, Any]] = []
    taken = [False] * npos
    m1 = _match_profile(prof, npos, 1)
    # bands of width p >= 2 (positions empty in both count as matching); a range that is one pure
    # fill (every position equal to the next) holds no band: every block would fail the fill test
    if not all(v is None or v >= SIM_COLS for v in m1[:npos - 1]):
        for p in range(2, min(PMAX, npos // 2) + 1):
            if budget is not None and budget.out():
                budget.skip(f"repeated blocks on sheet {sheet.name}: blocks wider than {p - 1} not checked, the time budget ran out")
                break
            mp = _match_profile(prof, npos, p)
            ok = [v is None or v >= SIM_COLS for v in mp]
            for a, b in _runs(ok):
                k = (b - a + 1 + p) // p
                if k < 2 or any(taken[a:a + k * p]):
                    continue
                if sum(len(prof.get(a + j, ())) for j in range(p)) < MIN_FORMULAS:
                    continue
                inner = m1[a:a + p - 1]  # a pure fill (every position equals the next) is an axis, not a band
                if inner and min(0.0 if v is None else v for v in inner) >= SIM_COLS:
                    continue
                positions = [[lo + a + j * p + t for t in range(p)] for j in range(k)]
                found.append(_finding(book, sheet, axis, positions, p, False))
                for i in range(a, min(a + k * p, npos)):
                    taken[i] = True
    # fills: one position repeated
    ok1 = [v is not None and v >= SIM_COLS for v in m1]
    for a, b in _runs(ok1):
        n = b - a + 2
        if n < MIN_FILL or any(taken[a:a + n]):
            continue
        positions = [[lo + a + j] for j in range(n)]
        found.append(_finding(book, sheet, axis, positions, 1, True))
        for i in range(a, a + n):
            taken[i] = True
    return found


def _cell_set(f: dict[str, Any]) -> frozenset:
    return frozenset(rc for lst in f["cells"].values() for rc in lst)


def sheet_findings(book: ShapeBook, sheet: SheetShape, budget: Budget | None = None) -> list[dict[str, Any]]:
    """Bands and fills of one sheet along both axes, deduplicated. The same cells found twice are
    kept once: two blocks -> the version with more copies; a block and a fill -> the block, which
    keeps the fill as its `inner_axis` (the rows of a band of column blocks repeat too: that is the
    band's own time or item axis, not a finding of its own); two fills of the same area -> ONE fill
    with both axes (`axes`, `chain_shares`, `time_axes`, `headers` per axis). A coarse fill (whole
    columns repeated, several formulas down each) that holds a UNIFORM area found along the other
    axis gives way to that area: the area keeps the coarse fill's axis facts and names it in
    `within`; the coarse fill is dropped (its computed index would be read from the wrong row)."""
    out = list(bands(book, sheet, "columns", budget))
    sets = [_cell_set(f) for f in out]
    dropped: set[int] = set()
    for f in bands(book, sheet, "rows", budget):
        fs = _cell_set(f)
        overlaps = []  # dropped coarse fills stay in: every uniform area inside them inherits their axis facts
        for i, gs in enumerate(sets):
            inter = len(fs & gs)
            if inter and inter >= SIM_COLS * min(len(fs), len(gs)):
                overlaps.append((i, inter))
        if not overlaps:
            out.append(f)
            sets.append(fs)
            continue
        band_ov = [(i, n) for i, n in overlaps if not out[i]["fill"]]
        fill_ov = [(i, n) for i, n in overlaps if out[i]["fill"]]
        if not f["fill"]:  # a band of row blocks
            if band_ov:
                i = max(band_ov, key=lambda t: t[1])[0]
                if f["copies"] > out[i]["copies"]:
                    out[i], sets[i] = f, fs
            else:  # it covers fills of columns: the band wins, the biggest fill is its inner axis
                i = max(fill_ov, key=lambda t: t[1])[0]
                f["inner_axis"] = _inner_axis(out[i])
                dropped.update(j for j, _ in fill_ov)
                out.append(f)
                sets.append(fs)
            continue
        if band_ov:  # a fill of rows inside a band of column blocks: the band's inner axis
            i = max(band_ov, key=lambda t: t[1])[0]
            out[i]["inner_axis"] = _inner_axis(f)
            continue
        same = [i for i, n in fill_ov if n >= SIM_COLS * max(len(fs), len(sets[i])) and i not in dropped]
        if same:
            _merge_fill_axis(out[same[0]], f)
            continue
        if f["axes"]:  # a uniform area inside coarse fills of the other axis
            for i, _ in fill_ov:
                if not out[i]["axes"]:
                    _merge_fill_axis(f, out[i])
                    f["within"] = _summary(out[i])
                    dropped.add(i)
            out.append(f)
            sets.append(fs)
            continue
        holds = [i for i, _ in fill_ov if out[i]["axes"]]  # f is the coarse one
        for i in holds:
            _merge_fill_axis(out[i], f)
            out[i]["within"] = _summary(f)
        if not holds:
            out.append(f)
            sets.append(fs)
    return [g for i, g in enumerate(out) if i not in dropped]


def _merge_fill_axis(target: dict[str, Any], source: dict[str, Any]) -> None:
    """The other axis's facts of the same (or the enclosing) fill."""
    target["chain_shares"][source["axis"]] = source["chain_share"]
    target["time_axes"][source["axis"]] = source["time_axis"]
    target["headers"][source["axis"]] = source["header"]
    target["time_axis"] = any(target["time_axes"].values())
    if target["axes"] is None and source["axes"] is not None:
        target["axes"] = source["axes"]


def _summary(fill: dict[str, Any]) -> dict[str, Any]:
    return {"axis": fill["axis"], "where": fill["where"], "copies": fill["copies"], "n_cells": fill["n_cells"]}


def _inner_axis(fill: dict[str, Any]) -> dict[str, Any]:
    return {"axis": fill["axis"], "copies": fill["copies"], "chain_share": fill["chain_share"], "header": fill["header"],
            "time_axis": fill["time_axis"]}


# --- P3: sheet copies ----------------------------------------------------------------------------
def _overlap(a: Counter, b: Counter) -> float:
    keys = a.keys() | b.keys()
    mn = sum(min(a.get(k, 0), b.get(k, 0)) for k in keys)
    mx = sum(max(a.get(k, 0), b.get(k, 0)) for k in keys)
    return mn / mx if mx else 0.0


def _aligned_similarity(a: dict[tuple[int, int], int], b: dict[tuple[int, int], int], L1: list[int]) -> float:
    """Share of the formula cells of two sheets (by address) that hold the same loose-1 shape."""
    same = common = 0
    for k, v in a.items():
        w = b.get(k)
        if w is not None:
            common += 1
            if L1[w] == L1[v]:
                same += 1
    union = len(a) + len(b) - common
    return same / union if union else 0.0


def sheet_copies(book: ShapeBook, budget: Budget | None = None) -> list[dict[str, Any]]:
    """Groups of sheets with the same formulas at the same addresses (>= SIM_SHEETS of the union
    equal in loose-1), in workbook order. Per group: members, per_copy (formulas of the first),
    per_copy_formulas, copies, min_similarity (the weakest pair), cells {member: {(row, col): id}},
    diff, chain, cross_reads, adding_cells, adding_example."""
    order = list(book.sheet_order) + [n for n in book.sheets if n not in book.sheet_order]
    cands = [n for n in order if n in book.sheets and book.sheets[n].n >= MIN_FORMULAS]
    L1 = book.loose[1]
    counters = {n: Counter(L1[s] for s in book.sheets[n].cells.values()) for n in cands}
    parent = {n: n for n in cands}

    def find(x: str) -> str:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    sims: dict[tuple[str, str], float] = {}
    for i, a in enumerate(cands):
        if budget is not None and budget.out():
            budget.skip("repeated sheets: not every pair of sheets was compared, the time budget ran out")
            break
        for b in cands[i + 1:]:
            if _overlap(counters[a], counters[b]) < PREFILTER:
                continue
            sim = _aligned_similarity(book.sheets[a].cells, book.sheets[b].cells, L1)
            sims[(a, b)] = sim
            if sim >= SIM_SHEETS:
                parent[find(a)] = find(b)
    groups: dict[str, list[str]] = {}
    for n in cands:
        groups.setdefault(find(n), []).append(n)
    out: list[dict[str, Any]] = []
    for members in groups.values():
        if len(members) < 2:
            continue
        members.sort(key=cands.index)
        pair_sims = []
        for i, a in enumerate(members):
            for b in members[i + 1:]:
                s = sims.get((a, b))
                if s is None:
                    s = sims[(a, b)] = _aligned_similarity(book.sheets[a].cells, book.sheets[b].cells, L1)
                pair_sims.append(s)
        group = {"members": members}
        chain, cross = consecutive_reads(book, group)
        adds, example = adding_cells(book, members)
        per = [book.sheets[m].n for m in members]
        out.append({
            "members": members, "per_copy": per[0], "per_copy_formulas": per, "copies": len(members),
            "min_similarity": round(min(pair_sims), 3), "cells": {m: book.sheets[m].cells for m in members},
            "diff": diff_kinds(book, stacks_of(book, group)), "chain": chain, "cross_reads": cross,
            "adding_cells": adds, "adding_example": example, "where": ", ".join(members),
        })
    return out


# --- P4: how the copies differ, and the edits that would unify them ------------------------------
def _constant(cells: Any, row: int, col: int) -> Any:
    """The typed-in value of a cell (None for an empty cell, a formula, or anything unusual)."""
    c = cells.get((row, col)) if cells is not None else None
    if c is None:
        return None
    v = c.value
    if v is None or v == "" or (isinstance(v, str) and v.startswith("=")):
        return None
    if not isinstance(v, (int, float, str, bool, date, datetime)):
        return None  # ArrayFormula / DataTableFormula objects
    return v


def _json_value(v: Any) -> Any:
    return v.isoformat() if isinstance(v, (date, datetime)) else v


def _cells_lookup(wb: Any):
    cache: dict[str, Any] = {}

    def cells_of(sheet: str) -> Any:
        if sheet not in cache:
            try:
                cache[sheet] = getattr(wb[sheet], "_cells", None) if sheet in wb.sheetnames else None
            except Exception:
                cache[sheet] = None
        return cache[sheet]

    return cells_of


def _constants_differ(wb: Any, positions_per_copy: list[list[tuple[str, int, int]]], want) -> list[dict[str, Any]]:
    if wb is None or not positions_per_copy or any(not p for p in positions_per_copy):
        return []
    cells_of = _cells_lookup(wb)
    out: list[dict[str, Any]] = []
    for i in range(min(len(p) for p in positions_per_copy)):
        vals = [_constant(cells_of(sheet), r, c) for sheet, r, c in (copy[i] for copy in positions_per_copy)]
        if all(want(v) for v in vals) and len(set(vals)) > 1:
            sheet, r, c = positions_per_copy[0][i]
            out.append({"cell": cell_name(sheet, r, c), "values": [_json_value(v) for v in vals]})
    return out


def numbers_differ(wb: Any, positions_per_copy: list[list[tuple[str, int, int]]]) -> list[dict[str, Any]]:
    """Positions (aligned per copy, as (sheet, row, col)) where EVERY copy holds a typed-in number
    and the numbers differ: [{cell (copy 1), values (per copy)}]."""
    return _constants_differ(wb, positions_per_copy, _is_number)


def texts_differ(wb: Any, positions_per_copy: list[list[tuple[str, int, int]]]) -> list[dict[str, Any]]:
    """Like numbers_differ for typed-in texts."""
    return _constants_differ(wb, positions_per_copy, lambda v: isinstance(v, str))


def constant_positions(wb: Any, sheet: str, rect: tuple[int, int, int, int] | None = None, want=_is_number,
                       limit: int = 200_000) -> list[tuple[int, int]]:
    """(row, col) of the typed-in values of a sheet (numbers by default), inside `rect`
    (r1, c1, r2, c2) when given; row-major. Reads the sheet's own cells, nothing is created."""
    try:
        cells = getattr(wb[sheet], "_cells", None) if wb is not None and sheet in wb.sheetnames else None
    except Exception:
        cells = None
    if cells is None:
        return []
    out: list[tuple[int, int]] = []
    if rect is not None and (rect[2] - rect[0] + 1) * (rect[3] - rect[1] + 1) <= min(limit, len(cells)):
        for r in range(rect[0], rect[2] + 1):
            for c in range(rect[1], rect[3] + 1):
                if want(_constant(cells, r, c)):
                    out.append((r, c))
        return out
    n = 0
    for (r, c), cell in cells.items():
        n += 1
        if n > limit:
            break
        if rect is not None and not (rect[0] <= r <= rect[2] and rect[1] <= c <= rect[3]):
            continue
        if want(_constant(cells, r, c)):
            out.append((r, c))
    out.sort()
    return out


def copy_positions(book: ShapeBook, finding: dict[str, Any], wb: Any) -> list[list[tuple[str, int, int]]]:
    """The typed-in numbers of copy 1 (a sheet group: the whole first member; a band: its own
    columns or rows) at the same place in every copy, for numbers_differ."""
    if wb is None:
        return []
    if "members" in finding:
        members = list(finding["members"])
        pos = constant_positions(wb, members[0])
        return [[(m, r, c) for r, c in pos] for m in members]
    sheet = book.sheets.get(finding["sheet"])
    if sheet is None or not finding["positions"]:
        return []
    axis = finding["axis"]
    first = finding["positions"][0]
    rect = (1, first[0], sheet.r_hi, first[-1]) if axis == "columns" else (first[0], 1, first[-1], sheet.c_hi)
    pos = constant_positions(wb, sheet.name, rect)
    out = []
    for positions in finding["positions"]:
        d = positions[0] - first[0]
        out.append([(sheet.name, r, c + d) if axis == "columns" else (sheet.name, r + d, c) for r, c in pos])
    return out


def _address(stacks: Any, i: int, key: tuple[int, int]) -> tuple[str, int, int]:
    if isinstance(stacks, Stacks):
        return stacks.address(i, key)
    return "?", key[0], key[1]


def _all_keys(stacks: list[dict[tuple[int, int], int]]) -> list[tuple[int, int]]:
    keys: set[tuple[int, int]] = set()
    for s in stacks:
        keys.update(s)
    return sorted(keys)


def formula_numbers_differ(book: ShapeBook, stacks: list[dict[tuple[int, int], int]]) -> list[dict[str, Any]]:
    """Cells whose copies differ only in a typed-in number inside the formula (=C7*0.1 vs =C7*0.12):
    [{cell (copy 1), literals (per copy, in order)}]."""
    L1, L2 = book.loose[1], book.loose[2]
    out: list[dict[str, Any]] = []
    for k in _all_keys(stacks):
        ids = [s.get(k) for s in stacks]
        first = ids[0]
        if first is None or any(i is None for i in ids) or all(i == first for i in ids):
            continue
        if all(L1[i] == L1[first] for i in ids) or not all(L2[i] == L2[first] for i in ids):
            continue
        sheet, r, c = _address(stacks, 0, k)
        out.append({"cell": cell_name(sheet, r, c), "literals": [number_literals(book.texts[i] or "") for i in ids]})
    return out


def logic_cells(book: ShapeBook, stacks: list[dict[tuple[int, int], int]]) -> list[dict[str, Any]]:
    """Cells of different logic: [{cell, formulas (A1 per copy, None where the copy has none),
    action}]. The action: "IF on the copy's key" when the cell exists in every copy with at most
    MAX_IF_FORMS distinct formulas, "override kept per copy" for more forms, "size differs" for a
    cell present in some copies only."""
    L1, L2 = book.loose[1], book.loose[2]
    out: list[dict[str, Any]] = []
    for k in _all_keys(stacks):
        ids = [s.get(k) for s in stacks]
        present = [i for i in ids if i is not None]
        if len(present) < len(ids):
            action = SIZE_DIFFERS
        else:
            first = ids[0]
            if all(i == first for i in ids) or all(L1[i] == L1[first] for i in ids) or all(L2[i] == L2[first] for i in ids):
                continue
            action = IF_ACTION if len(set(ids)) <= MAX_IF_FORMS else OVERRIDE_ACTION
        formulas = [book.a1_text(*_address(stacks, i, k)) if sid is not None else None for i, sid in enumerate(ids)]
        owner = next(i for i, sid in enumerate(ids) if sid is not None)
        sheet, r, c = _address(stacks, owner, k)
        out.append({"cell": cell_name(sheet, r, c), "formulas": formulas, "action": action, "copies_with_cell": len(present)})
    return out


def edits_for(book: ShapeBook, finding: dict[str, Any], wb: Any = None) -> dict[str, Any]:
    """The edits that would unify the copies of a band / fill / sheet group, each list capped at
    MAX_EDITS with complete counts: numbers (typed-in numbers that differ; needs the workbook),
    formula_numbers, logic (with the IF-or-override action), size (cells in some copies only)."""
    st = stacks_of(book, finding)
    fnums = formula_numbers_differ(book, st)
    logic_all = logic_cells(book, st)
    logic = [x for x in logic_all if x["action"] != SIZE_DIFFERS]
    size = [x for x in logic_all if x["action"] == SIZE_DIFFERS]
    numbers = numbers_differ(wb, copy_positions(book, finding, wb)) if wb is not None else []
    return {
        "numbers": [{"cell": x["cell"], "values_per_copy": x["values"][:MAX_VALUES], "action": NUMBERS_ACTION} for x in numbers[:MAX_EDITS]],
        "formula_numbers": [{"cell": x["cell"], "literals_per_copy": x["literals"][:MAX_VALUES], "action": FORMULA_NUMBERS_ACTION} for x in fnums[:MAX_EDITS]],
        "logic": [{"cell": x["cell"], "formulas_per_copy": x["formulas"][:MAX_FORMS], "action": x["action"]} for x in logic[:MAX_EDITS]],
        "size": [{"cell": x["cell"], "copies_with_cell": x["copies_with_cell"], "action": SIZE_ACTION} for x in size[:MAX_EDITS]],
        "numbers_count": len(numbers), "formula_numbers_count": len(fnums), "logic_count": len(logic), "size_count": len(size),
    }


# --- what a repetition steps through ----------------------------------------------------------------
def _relative_step(book: ShapeBook, finding: dict[str, Any]) -> dict[str, Any] | None:
    """Copies whose cells read a DIFFERENT input at the same place (same loose-1 shape, different
    references): when the targets of copy j sit (j-1)*s rows or (j-1)*w columns after copy 1's,
    the copies step through one area. The most common target wins."""
    stacks = stacks_of(book, finding)
    if len(stacks) < 2:
        return None
    L1 = book.loose[1]
    cands: Counter = Counter()  # (target sheet, axis, stride) -> cells
    hulls: dict[tuple, list[int]] = {}  # -> [r1, r2, c1, c2] over every target seen
    last_starts: dict[tuple, int] = {}  # -> the first row/column of the last copy's targets
    for k in stacks[0]:
        ids = [s.get(k) for s in stacks]
        first = ids[0]
        if any(i is None for i in ids) or all(i == first for i in ids) or not all(L1[i] == L1[first] for i in ids):
            continue
        parsed = [book.parsed(i) for i in ids]
        nrefs = len(parsed[0].refs)
        if any(len(p.refs) != nrefs for p in parsed):
            continue
        for ri in range(nrefs):
            refs = [p.refs[ri] for p in parsed]
            if all(r == refs[0] for r in refs):
                continue
            spec = refs[0][0]
            if any(r[0] != spec for r in refs) or (spec is not None and spec[0] != "sheet"):
                continue
            target = finding["sheet"] if spec is None else spec[1]
            rects = [ref_rect(refs[j], *stacks.address(j, k)[1:]) for j in range(len(ids))]
            if len({(r2 - r1, c2 - c1) for r1, r2, c1, c2 in rects}) != 1:
                continue
            r1s, c1s = [r[0] for r in rects], [r[2] for r in rects]
            axis = stride = None
            if len({(r[2], r[3]) for r in rects}) == 1:
                d = r1s[1] - r1s[0]
                if d > 0 and all(r1s[j] - r1s[j - 1] == d for j in range(1, len(r1s))):
                    axis, stride = "rows", d
            elif len({(r[0], r[1]) for r in rects}) == 1:
                d = c1s[1] - c1s[0]
                if d > 0 and all(c1s[j] - c1s[j - 1] == d for j in range(1, len(c1s))):
                    axis, stride = "columns", d
            if axis is None:
                continue
            key = (target, axis, stride)
            cands[key] += 1
            h = hulls.get(key)
            seen = [min(r[0] for r in rects), max(r[1] for r in rects), min(r[2] for r in rects), max(r[3] for r in rects)]
            hulls[key] = seen if h is None else [min(h[0], seen[0]), max(h[1], seen[1]), min(h[2], seen[2]), max(h[3], seen[3])]
            last = r1s[-1] if axis == "rows" else c1s[-1]
            last_starts[key] = min(last_starts.get(key, last), last)
    if not cands:
        return None
    (target, axis, stride), n = cands.most_common(1)[0]
    hull = hulls[(target, axis, stride)]
    if axis == "rows":  # the last copy's target reaches to the end of its slot
        hull[1] = max(hull[1], last_starts[(target, axis, stride)] + stride - 1)
    else:
        hull[3] = max(hull[3], last_starts[(target, axis, stride)] + stride - 1)
    return {"sheet": target, "r1": hull[0], "r2": hull[1], "c1": hull[2], "c2": hull[3], "stride": stride, "axis": axis,
            "computed": None, "cells": n}


_ROWCOL_FORMS = (
    re.compile(r"^\(\s*(?P<fn>ROW|COLUMN)\(\)\s*-\s*(?P<n>\d+)\s*\)\s*\*\s*(?P<w>\d+)$", re.IGNORECASE),
    re.compile(r"^(?P<w>\d+)\s*\*\s*\(\s*(?P<fn>ROW|COLUMN)\(\)\s*-\s*(?P<n>\d+)\s*\)$", re.IGNORECASE),
    re.compile(r"^(?P<fn>ROW|COLUMN)\(\)\s*-\s*(?P<n>\d+)$", re.IGNORECASE),
)
_CALL_FORMS = (
    re.compile(r"^(?P<w>\d+)\s*\*\s*\(\s*(?P<fn>COUNTA|MATCH)\((?P<args>.*)\)\s*-\s*1\s*\)$", re.IGNORECASE | re.DOTALL),
    re.compile(r"^\(\s*(?P<fn>COUNTA|MATCH)\((?P<args>.*)\)\s*-\s*1\s*\)\s*\*\s*(?P<w>\d+)$", re.IGNORECASE | re.DOTALL),
    re.compile(r"^(?P<fn>COUNTA|MATCH)\((?P<args>.*)\)\s*-\s*1$", re.IGNORECASE | re.DOTALL),
)


def _unwrap(text: str) -> str:
    """The text without a pair of parentheses that wraps all of it: "(COLUMN()-14)" -> "COLUMN()-14".
    Parentheses inside string literals do not count; "(a)*(b)" is not wrapped and stays."""
    t = text.strip()
    while len(t) > 1 and t.startswith("(") and t.endswith(")"):
        depth, in_q = 0, False
        for i, ch in enumerate(t):
            if ch == '"':
                in_q = not in_q
            elif in_q:
                continue
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0 and i < len(t) - 1:
                    return t  # the first parenthesis closes before the end: no wrap
        t = t[1:-1].strip()
    return t


def _computed_arg(text: str) -> dict[str, Any] | None:
    """One OFFSET/INDEX displacement argument read for the stepping idioms: (ROW()-n)*w, w*(ROW()-n),
    ROW()-n (w 1), the same with COLUMN(); w*(COUNTA(expanding range)-1); w*(MATCH(key, list, 0)-1).
    The parentheses must balance -- ROW()-21*10 is ROW()-210 in Excel, not one of the idioms -- and a
    pair wrapping the whole text is dropped first. The finder's one parser of these idioms:
    app.finder_grouping.displacement_form reads through it, so both passes see the same forms."""
    t = _unwrap(text.replace("@", ""))
    for rx in _ROWCOL_FORMS:
        m = rx.match(t)
        if m:
            fn = m.group("fn").upper()
            return {"form": "row" if fn == "ROW" else "column", "w": int(m.groupdict().get("w") or 1), "n": int(m.group("n")),
                    "step_axis": "rows" if fn == "ROW" else "columns", "text": t}
    for rx in _CALL_FORMS:
        m = rx.match(t)
        if not m:
            continue
        fn = m.group("fn").upper()
        args = [a.strip().replace("@", "") for a in split_top_level_args(m.group("args"))]
        w = int(m.groupdict().get("w") or 1)
        if fn == "COUNTA" and args:
            rng = parse_ref(args[0])
            if not rng or rng["whole_column"] or rng["whole_row"]:
                return None  # a whole column is no expanding range: its count is the same in every cell
            # an expanding range down one column counts rows; along one row it counts columns; a block is no idiom
            if rng["c1"] == rng["c2"]:
                step_axis = "rows"
            elif rng["r1"] == rng["r2"]:
                step_axis = "columns"
            else:
                return None
            return {"form": "counta", "w": w, "n": None, "step_axis": step_axis, "text": t, "range": args[0]}
        if fn == "MATCH" and len(args) >= 2:
            lst = parse_ref(args[1])
            return {"form": "match", "w": w, "n": None, "step_axis": None, "text": t, "key": args[0], "list": args[1],
                    "list_ref": f"{lst['sheet'] + '!' if lst and lst['sheet'] else ''}{lst['ref']}" if lst else None,
                    "list_len": lst["cells"] if lst else None}
    return None


def _resolve_anchor(text: str, own_sheet: str, names: Any) -> tuple[str, int, int] | None:
    """A reference or a defined name -> (sheet, row, col) of its top-left cell."""
    t = text.strip().replace("@", "")
    r = parse_ref(t)
    if r is None and names is not None:
        ref = names.ref(t) if hasattr(names, "ref") else None
        r = parse_ref(ref) if ref else None
    if r is None or r["whole_column"] or r["whole_row"]:
        return None
    return r["sheet"] or own_sheet, r["r1"], r["c1"]


def step_target(book: ShapeBook, finding: dict[str, Any], a1_text: str | None, names: Any = None) -> dict[str, Any] | None:
    """What the copies of a finding step through. Two ways: differing relative references (a band
    whose copy j reads row/column j of an input area) or a computed index in the A1 text of the
    first cell (OFFSET/INDEX with ROW/COLUMN/COUNTA/MATCH arithmetic, the use case's way). Returns
    {sheet, r1, c1, r2, c2 (the area stepped through), stride, axis (slots along rows|columns),
    computed: None | {form, anchor, w, slot_axis, step_axis, key, list, ...}} or None."""
    out = None
    if not finding.get("fill") and "members" not in finding:
        out = _relative_step(book, finding)
    if out is None and a1_text:
        out = _computed_step(book, finding, a1_text, names)
    return out


def _computed_step(book: ShapeBook, finding: dict[str, Any], a1_text: str, names: Any) -> dict[str, Any] | None:
    text = STORAGE_PREFIX_RE.sub("", a1_text)
    own = finding.get("sheet") or (finding.get("members") or [""])[0]
    cands: list[dict[str, Any]] = []
    for fn, is_index in (("OFFSET", False), ("INDEX", True)):
        for args in find_calls(text, fn):
            if len(args) < 2:
                continue
            for pos, slot_axis in ((1, "rows"), (2, "columns")):
                if pos >= len(args):
                    continue
                c = _computed_arg(args[pos])
                if c is None:
                    continue
                c.update({"slot_axis": slot_axis, "anchor_text": args[0].strip(), "index": is_index, "call": fn})
                cands.append(c)
    if not cands:
        return None
    cells = [rc for lst in (finding.get("cells") or {}).values() for rc in lst]
    rows = {r for r, _ in cells}
    cols = {c for _, c in cells}
    axis = finding.get("axis")

    def rank(c: dict[str, Any]) -> tuple:
        return (c["w"] > 1, c["form"] != "match", c.get("step_axis") == axis)

    cands.sort(key=rank, reverse=True)
    best = cands[0]
    anchor = _resolve_anchor(best["anchor_text"], own, names)
    if anchor is None:
        return None
    a_sheet, a_row, a_col = anchor
    w = best["w"]
    form = best["form"]
    if form == "match":
        npos = best.get("list_len") or (len(rows) if axis == "rows" else len(cols)) or 1
        v0 = 0
        step_axis = None
    else:
        step_axis = best["step_axis"]
        along = rows if step_axis == "rows" else cols
        npos = len(along) or finding.get("copies") or 1
        v0 = (min(along) - best["n"]) * w if form in ("row", "column") and along else 0
    if best["index"]:
        v0 -= 1
    # the slots along the slot axis; the other axis follows a second computed argument (w 1) or is one line.
    # The hull runs from the first read cell to the last one (anchor .. anchor + w*(positions-1)): the
    # index may land anywhere inside its slot, so a whole last slot would reach w-1 cells past the area.
    other = next((c for c in cands if c is not best and c["slot_axis"] != best["slot_axis"] and c["form"] in ("row", "column")), None)
    if other is not None:
        o_along = rows if other["step_axis"] == "rows" else cols
        o_v0 = (min(o_along) - other["n"]) * other["w"] if o_along else 0
        o_len = other["w"] * ((len(o_along) or 1) - 1) + 1
        if best["index"]:
            o_v0 -= 1
    else:
        o_v0, o_len = 0, 1
    if best["slot_axis"] == "columns":
        c1 = a_col + v0
        c2 = c1 + w * (npos - 1)
        r1 = a_row + o_v0
        r2 = r1 + o_len - 1
    else:
        r1 = a_row + v0
        r2 = r1 + w * (npos - 1)
        c1 = a_col + o_v0
        c2 = c1 + o_len - 1
    computed = {
        "form": form, "anchor": f"{cell_name(a_sheet, a_row, a_col).rsplit('!', 1)[0]}!${num_to_col(a_col)}${a_row}",
        "w": w, "slot_axis": best["slot_axis"], "step_axis": step_axis, "key": best.get("key"), "list": best.get("list"),
        "list_ref": best.get("list_ref"), "list_len": best.get("list_len"), "n": best.get("n"), "text": best["text"],
        "call": best["call"], "positions": npos,
        "others": [{k: v for k, v in c.items() if k not in ("index",)} for c in cands[1:]],
    }
    return {"sheet": a_sheet, "r1": max(1, r1), "c1": max(1, c1), "r2": max(1, r2), "c2": max(1, c2), "stride": w,
            "axis": best["slot_axis"], "computed": computed}
