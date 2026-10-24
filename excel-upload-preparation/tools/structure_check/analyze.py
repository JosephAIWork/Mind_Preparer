"""Stage 2: find repeated structure in a workbook and say, with evidence, whether a Mind loop,
Mind instances, or nothing would fit it. It checks; it does not choose.

Repetition is read on R1C1 formula text: cells filled down/across or copied share one text.
  A. sheet copies   - sheets with the same formulas at the same addresses
  B. block bands    - a block of columns (or rows) repeated side by side inside one sheet
  C. fill axes      - one column (row) repeated many times: a time axis or a list of items
For every repetition: are the copies PARALLEL (no copy reads another) or a CHAIN (each reads the
previous)? Only parallel copies are a dimension; a chain is a sequence and fits neither method.
"""
from __future__ import annotations

import pickle
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np

APP = Path(__file__).resolve().parents[2]   # excel-upload-preparation
sys.path.insert(0, str(APP))
from app.autofix import _REF_RE, _STRING_RE, parse_r1c1, ref_rect  # noqa: E402  (read-only use of the app's parser)

_NUM_RE = re.compile(r"(?<![A-Za-z_\]\[\d.])\d+(?:\.\d+)?(?:E[+-]?\d+)?(?![\d\]])", re.I)


def _outside_strings(text, fn):
    out, pos = [], 0
    for m in _STRING_RE.finditer(text):
        out.append(fn(text[pos:m.start()])); out.append(m.group(0)); pos = m.end()
    out.append(fn(text[pos:]))
    return "".join(out)


def loose_keys(text):
    """(refs masked but sheets kept, refs and sheets masked, + typed-in numbers masked)."""
    k1 = _outside_strings(text, lambda s: _REF_RE.sub(lambda m: (m.group("sheet") or "") + "@", s))
    k2 = _outside_strings(text, lambda s: _REF_RE.sub("@", s))
    k3 = _outside_strings(k2, lambda s: _NUM_RE.sub("#", s))
    return k1, k2, k3

SIM_SHEETS = 0.75      # share of formula cells that must match for two sheets to be copies
SIM_COLS = 0.90        # share of a column's formula cells that must match its copy
MIN_FORMULAS = 10      # sheets / blocks smaller than this are not worth a finding
PMAX = 40              # widest block looked for
SAMPLE = 3000          # cells sampled per reference test
RANDOM_FUNCS = {"RAND", "RANDBETWEEN", "NORMINV", "NORM.INV", "NORMSINV", "NORM.S.INV", "LOGINV", "LOGNORM.INV", "BETAINV", "GAMMAINV"}


class Book:
    def __init__(self, path: Path):
        d = pickle.load(open(path, "rb"))
        self.file = d["file"]
        self.texts = d["texts"]
        self.sheets = d["sheets"]
        self.by_name = {s["name"]: s for s in self.sheets}
        self._parsed: dict[int, object] = {}
        self.names = defined_cells(self.file)
        n = len(self.texts)
        self.L = [np.zeros(n, np.int32) for _ in range(3)]
        ids = [{}, {}, {}]
        for i, tx in enumerate(self.texts):
            if tx is None:
                continue
            for lvl, key in enumerate(loose_keys(tx)):
                self.L[lvl][i] = ids[lvl].setdefault(key, len(ids[lvl]) + 10)

    def loose(self, F, lvl=1):
        """A formula-id grid mapped to loose ids (0 stays 0). lvl 0: refs masked, 1: + sheets, 2: + numbers."""
        return np.where(F > 0, self.L[lvl][F], 0)

    def parsed(self, sid: int):
        p = self._parsed.get(sid)
        if p is None:
            p = self._parsed[sid] = parse_r1c1(self.texts[sid])
        return p

    def abs_grid(self, s, which="shape"):
        """The sheet's grid placed at absolute row/col 1 (so copies line up by address)."""
        g = s[which]
        out = np.zeros((s["r0"] - 1 + g.shape[0], s["c0"] - 1 + g.shape[1]), dtype=g.dtype)
        out[s["r0"] - 1:, s["c0"] - 1:] = g
        return out


def formulas(g):
    return np.where(g >= 10, g, 0)


def sheets_read(book: Book, sid: int, own: str) -> set[str]:
    out = set()
    for spec, _, _ in book.parsed(sid).refs:
        if spec and spec[0] == "sheet" and spec[1] != own:
            out.add(spec[1])
        elif spec and spec[0] == "3d":
            names = [s["name"] for s in book.sheets]
            if spec[1] in names and spec[2] in names:
                i, j = sorted((names.index(spec[1]), names.index(spec[2])))
                out.update(names[i:j + 1])
    return out


def same_sheet_cols(book: Book, sid: int, row: int, col: int, axis: int):
    """For a formula at (row, col): the spans (lo, hi) along `axis` (1 = columns, 0 = rows) of
    every reference to its own sheet."""
    spans = []
    for ref in book.parsed(sid).refs:
        if ref[0] is not None:
            continue
        r1, r2, c1, c2 = ref_rect(ref, row, col)
        spans.append((c1, c2) if axis == 1 else (r1, r2))
    return spans


# --- A. sheet copies --------------------------------------------------------------------------
def sheet_copies(book: Book):
    cands = [s for s in book.sheets if int((s["shape"] >= 10).sum()) >= MIN_FORMULAS]
    grids = {s["name"]: formulas(book.abs_grid(s)) for s in cands}
    loose = {k: book.loose(v, 1) for k, v in grids.items()}
    parent = {s["name"]: s["name"] for s in cands}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    sims = {}
    for i, a in enumerate(cands):
        for b in cands[i + 1:]:
            ga, gb = loose[a["name"]], loose[b["name"]]
            h, w = max(ga.shape[0], gb.shape[0]), max(ga.shape[1], gb.shape[1])
            A = np.zeros((h, w), np.int32); A[:ga.shape[0], :ga.shape[1]] = ga
            B = np.zeros((h, w), np.int32); B[:gb.shape[0], :gb.shape[1]] = gb
            union = int(((A > 0) | (B > 0)).sum())
            same = int(((A == B) & (A > 0)).sum())
            sim = same / union if union else 0
            sims[(a["name"], b["name"])] = sim
            if sim >= SIM_SHEETS:
                parent[find(a["name"])] = find(b["name"])
    groups = defaultdict(list)
    for n in parent:
        groups[find(n)].append(n)
    out = []
    order = [s["name"] for s in book.sheets]
    for members in groups.values():
        if len(members) < 2:
            continue
        members.sort(key=order.index)
        out.append(describe_sheet_group(book, members, grids, sims))
    return out


def describe_sheet_group(book: Book, members, grids, sims):
    h = max(grids[m].shape[0] for m in members)
    w = max(grids[m].shape[1] for m in members)
    stack = np.zeros((len(members), h, w), np.int32)
    cst = np.zeros((len(members), h, w), np.int64)
    kind = np.zeros((len(members), h, w), np.int32)
    for i, m in enumerate(members):
        g = grids[m]; stack[i, :g.shape[0], :g.shape[1]] = g
        c = book.abs_grid(book.by_name[m], "const"); cst[i, :c.shape[0], :c.shape[1]] = c
        k = book.abs_grid(book.by_name[m]); kind[i, :k.shape[0], :k.shape[1]] = k
    has_f = stack > 0
    kinds = diff_kinds(book, stack)
    identical = kinds["identical"]
    overrides = kinds["different logic"]
    per_copy = [int(has_f[i].sum()) for i in range(len(members))]
    # constants present in every copy at the same address
    num_all = (kind == 1).all(0)
    txt_all = (kind == 2).all(0)
    num_diff = int((num_all & ~(cst == cst[0]).all(0)).sum())
    num_same = int((num_all & (cst == cst[0]).all(0)).sum())
    txt_diff = int((txt_all & ~(cst == cst[0]).all(0)).sum())
    only_some = int(((kind > 0).any(0) & ~(kind > 0).all(0)).sum())   # cells some copies have and others not (size)
    # where the differing numbers sit (a compact block = an assumption table candidate)
    diff_pos = np.argwhere(num_all & ~(cst == cst[0]).all(0))
    diff_area = None
    if len(diff_pos):
        (r1, c1), (r2, c2) = diff_pos.min(0), diff_pos.max(0)
        diff_area = (int(r1) + 1, int(c1) + 1, int(r2) + 1, int(c2) + 1)
    # chain or parallel: do copies read each other?
    reads = {m: Counter() for m in members}
    for m in members:
        ids, counts = np.unique(stack[members.index(m)][stack[members.index(m)] > 0], return_counts=True)
        for sid, n in zip(ids, counts):
            for t in sheets_read(book, int(sid), m) & set(members):
                reads[m][t] += int(n)
    links = {m: dict(reads[m]) for m in members if reads[m]}
    # who adds the copies up
    agg_cells, agg_example = 0, None
    member_set = set(members)
    for s in book.sheets:
        if s["name"] in member_set:
            continue
        ids, counts = np.unique(s["shape"][s["shape"] >= 10], return_counts=True)
        for sid, n in zip(ids, counts):
            if len(sheets_read(book, int(sid), s["name"]) & member_set) >= 2:
                agg_cells += int(n)
                agg_example = agg_example or (s["name"], book.texts[int(sid)][:160])
    return {
        "kind": "sheet copies", "members": members, "per_copy_formulas": per_copy,
        "identical_formula_cells": identical, "override_cells": overrides, "diff_kinds": kinds,
        "numbers_differ": num_diff, "numbers_same": num_same, "texts_differ": txt_diff,
        "cells_in_some_copies_only": only_some, "differing_numbers_area": diff_area,
        "copies_read_each_other": links, "adding_cells": agg_cells, "adding_example": agg_example,
        "min_similarity": round(min(sims.get((a, b), sims.get((b, a), 0)) for i, a in enumerate(members) for b in members[i + 1:]), 3),
    }


# --- B/C. repetition inside a sheet ----------------------------------------------------------
def match_profile(F, p):
    """For each column c: share of formula cells of columns c and c+p that are equal."""
    a, b = F[:, :-p], F[:, p:]
    any_ = (a > 0) | (b > 0)
    eq = (a == b) & (a > 0)
    n = any_.sum(0)
    with np.errstate(invalid="ignore", divide="ignore"):
        m = np.where(n > 0, eq.sum(0) / np.maximum(n, 1), np.nan)
    return m


def runs(mask):
    out, start = [], None
    for i, v in enumerate(list(mask) + [False]):
        if v and start is None:
            start = i
        elif not v and start is not None:
            out.append((start, i - 1))
            start = None
    return out


def reads_earlier(book, F, s, cells, lo_of, axis):
    """Share of sampled formula cells that read an earlier position of the same repetition
    (lo_of(pos) gives the first position that still belongs to 'earlier copies')."""
    hits = tot = 0
    for (r, c) in cells:
        sid = int(F[r, c])
        R, C = r + s["r0"], c + s["c0"]
        pos = C if axis == 1 else R
        lo, hi = lo_of(pos)
        if hi < lo:
            continue
        tot += 1
        for a, b in same_sheet_cols(book, sid, R, C, axis):
            if b >= lo and a <= hi:
                hits += 1
                break
    return hits / tot if tot else 0.0


def sample_cells(F, rows, cols, n=SAMPLE):
    sub = F[np.ix_(rows, cols)] if not isinstance(rows, slice) else F[rows, cols]
    pos = np.argwhere(sub > 0)
    if len(pos) > n:
        pos = pos[np.random.default_rng(0).choice(len(pos), n, replace=False)]
    r_off = rows.start if isinstance(rows, slice) else 0
    c_off = cols.start if isinstance(cols, slice) else 0
    return [(int(r) + r_off, int(c) + c_off) for r, c in pos]


def header_hint(s, axis, lo, hi):
    """What the first constant row (column) above (left of) the repetition holds: numbers rising
    (a time axis), or text (names of items)."""
    shape = s["shape"]
    G = shape if axis == 1 else shape.T
    for r in range(G.shape[0]):
        seg = G[r, lo:hi + 1]
        if (seg > 0).sum() >= max(2, (hi - lo + 1) // 2):
            if (seg >= 10).any():
                return "formulas"
            return "numbers" if (seg == 1).sum() >= (seg == 2).sum() else "text"
    return "none"


def diff_kinds(book: Book, stack):
    """Cells of copies stacked on axis 0, sorted by how the copies differ there."""
    has = stack > 0
    anyf = has.any(0)
    exact = (stack == stack[0]).all(0) & has.all(0)
    out = {"identical": int(exact.sum())}
    prev = exact
    for lvl, label in ((1, "reads a different input"), (2, "differs only in a typed-in number")):
        L = book.loose(stack, lvl)
        same = (L == L[0]).all(0) & has.all(0) & ~prev
        out[label] = int(same.sum())
        prev = prev | same
    out["different logic"] = int((anyf & ~prev).sum())
    return out


def inside_sheet(book: Book, s, axis: int):
    F0 = formulas(s["shape"])
    Fx = F0 if axis == 1 else F0.T
    F = book.loose(Fx, 1)
    if F.shape[1] < 4 or int((F > 0).sum()) < MIN_FORMULAS:
        return []
    found, taken = [], np.zeros(F.shape[1], bool)
    m1 = match_profile(F, 1)
    # bands of width p >= 2 (columns empty in both count as matching)
    for p in range(2, min(PMAX, F.shape[1] // 2) + 1):
        mp = match_profile(F, p)
        ok = np.where(np.isnan(mp), True, mp >= SIM_COLS)
        for a, b in runs(ok):
            n_cols = b - a + 1 + p
            k = n_cols // p
            if k < 2 or taken[a:a + k * p].any():
                continue
            block = F[:, a:a + p]
            if int((block > 0).sum()) < MIN_FORMULAS:
                continue
            # a pure fill (every column equals the next) is an axis, not a band
            inner = m1[a:a + p - 1]
            if len(inner) and np.nanmin(np.where(np.isnan(inner), 0, inner)) >= SIM_COLS:
                continue
            band = {"axis": "columns" if axis == 1 else "rows", "sheet": s["name"], "start": a, "width": p, "copies": k}
            starts = [a + j * p for j in range(k)]
            cells = sample_cells(F, slice(0, F.shape[0]), slice(a + p, a + k * p))
            off = s["c0"] if axis == 1 else s["r0"]

            def lo_of(pos, starts=starts, off=off, a=a):
                j = max(i for i, st in enumerate(starts) if st + off <= pos)
                return a + off, starts[j] + off - 1

            share = reads_earlier(book, Fx, s, cells, lo_of, 1) if axis == 1 else reads_earlier_t(book, Fx, s, cells, lo_of)
            band["chain_share"] = round(share, 2)
            band["formulas_per_copy"] = int((block > 0).sum())
            blocks = np.stack([Fx[:, st:st + p] for st in starts])
            band["diff_kinds"] = diff_kinds(book, blocks)
            band["override_cells"] = band["diff_kinds"]["different logic"]
            cb = np.stack([(s["const"] if axis == 1 else s["const"].T)[:, st:st + p] for st in starts])
            kb = np.stack([(s["shape"] if axis == 1 else s["shape"].T)[:, st:st + p] for st in starts])
            num_all = (kb == 1).all(0)
            band["numbers_differ"] = int((num_all & ~(cb == cb[0]).all(0)).sum())
            band["adding_cells"] = adders(book, s, Fx, axis, starts, p)
            band["header"] = header_hint(s, axis, a, a + k * p - 1)
            found.append(band)
            taken[a:a + k * p] = True
    # fill axes: one column (row) repeated
    ok1 = np.where(np.isnan(m1), False, m1 >= SIM_COLS)
    for a, b in runs(ok1):
        n = b - a + 2
        if n < 6 or taken[a:a + n].any():
            continue
        cells = sample_cells(F, slice(0, F.shape[0]), slice(a + 1, a + n))
        off = s["c0"] if axis == 1 else s["r0"]
        if axis == 1:
            share = reads_earlier(book, Fx, s, cells, lambda pos, a=a, off=off: (a + off, pos - 1), 1)
        else:
            share = reads_earlier_t(book, Fx, s, cells, lambda pos, a=a, off=off: (a + off, pos - 1))
        found.append({"axis": "columns" if axis == 1 else "rows", "sheet": s["name"], "start": a, "width": 1, "copies": n,
                      "chain_share": round(share, 2), "formulas_per_copy": int((F[:, a] > 0).sum()),
                      "override_cells": 0, "adding_cells": 0, "header": header_hint(s, axis, a, a + n - 1), "fill": True})
        taken[a:a + n] = True
    return found


def reads_earlier_t(book, Ft, s, cells, lo_of):
    """reads_earlier for the transposed grid: cells are (col, row) of the original sheet."""
    hits = tot = 0
    for (c, r) in cells:
        sid = int(Ft[c, r])
        R, C = r + s["r0"], c + s["c0"]
        lo, hi = lo_of(R)
        if hi < lo:
            continue
        tot += 1
        for a, b in same_sheet_cols(book, sid, R, C, 0):
            if b >= lo and a <= hi:
                hits += 1
                break
    return hits / tot if tot else 0.0


def adders(book, s, F, axis, starts, p):
    """Formula cells outside the band that read two or more of its copies."""
    k = len(starts)
    lo, hi = starts[0], starts[-1] + p - 1
    off = s["c0"] if axis == 1 else s["r0"]
    outside = np.ones(F.shape[1], bool); outside[lo:hi + 1] = False
    pos = np.argwhere((F > 0) & outside[None, :])
    if len(pos) > SAMPLE:
        pos = pos[np.random.default_rng(1).choice(len(pos), SAMPLE, replace=False)]
    n = 0
    for r, c in pos:
        sid = int(F[r, c])
        R, C = (r + s["r0"], c + s["c0"]) if axis == 1 else (c + s["r0"], r + s["c0"])
        touched = set()
        for a, b in same_sheet_cols(book, sid, R, C, axis):
            for j, st in enumerate(starts):
                if b >= st + off and a <= st + off + p - 1:
                    touched.add(j)
        if len(touched) >= 2:
            n += 1
    return n


STEER_FUNCS = {"CHOOSE", "INDEX", "MATCH", "OFFSET", "VLOOKUP", "HLOOKUP", "XLOOKUP", "SWITCH", "INDIRECT", "SUMIFS", "SUMIF", "COUNTIFS", "LOOKUP"}


KEY_ARGS = {"VLOOKUP": (0,), "HLOOKUP": (0,), "XLOOKUP": (0,), "MATCH": (0,), "LOOKUP": (0,), "CHOOSE": (0,),
            "SWITCH": (0,), "INDIRECT": (0,), "INDEX": (1, 2), "OFFSET": (1, 2), "SUMIF": (1,),
            "SUMIFS": (2, 4, 6, 8, 10), "COUNTIFS": (1, 3, 5, 7, 9)}
_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


def key_args(text):
    """The argument texts in 'key' positions (what a lookup looks for, which choice, which row)."""
    masked = _STRING_RE.sub(lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', text)
    out = []
    for m in _CALL_RE.finditer(masked):
        name = m.group(1).upper().split(".")[-1] if m.group(1).upper().startswith("_XL") else m.group(1).upper()
        want = KEY_ARGS.get(name)
        if not want:
            continue
        depth, start, args = 0, m.end(), []
        for i in range(m.end(), len(masked)):
            ch = masked[i]
            if ch in "({":
                depth += 1
            elif ch in ")}":
                if depth == 0:
                    args.append(text[start:i]); break
                depth -= 1
            elif ch == "," and depth == 0:
                args.append(text[start:i]); start = i + 1
        out += [args[k] for k in want if k < len(args)]
    return out


def defined_cells(path):
    """Defined names that point at one cell: NAME -> (sheet, row, col). Read from the file's
    workbook.xml (xlsx/xlsm only); an .xlsb gives none."""
    import zipfile
    import xml.etree.ElementTree as ET
    out = {}
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("xl/workbook.xml"))
    except Exception:
        return out
    ns = {"m": root.tag.split("}")[0][1:]} if root.tag.startswith("{") else {}
    for dn in root.iter((("{%s}" % ns["m"]) if ns else "") + "definedName"):
        m = re.fullmatch(r"'?(.+?)'?!\$?([A-Z]{1,3})\$?(\d+)", (dn.text or "").strip())
        if m and dn.get("localSheetId") is None:
            col = 0
            for ch in m.group(2):
                col = col * 26 + ord(ch) - 64
            out[dn.get("name").upper()] = (m.group(1).replace("''", "'"), int(m.group(3)), col)
    return out


def switch_cells(book: Book, top=8, min_cells=10):
    """Typed-in cells (no formula) that many lookup/choice formulas read through an absolute
    single-cell reference: the model computes one choice of that cell at a time."""
    tally = Counter()
    funcs_seen = defaultdict(Counter)
    named = {}
    for s in book.sheets:
        ids, counts = np.unique(s["shape"][s["shape"] >= 10], return_counts=True)
        for sid, n in zip(ids, counts):
            pz = book.parsed(int(sid))
            steer = pz.funcs & STEER_FUNCS
            if not steer:
                continue
            args = key_args(book.texts[int(sid)])
            key_refs = [r for arg in args for r in parse_r1c1("=" + arg).refs]
            for arg in args:
                for nm in parse_r1c1("=" + arg).names:
                    if nm in book.names:
                        sh, rr, cc = book.names[nm]
                        key_refs.append((("sheet", sh), ((True, rr), (True, cc)), None))
                        named[(sh, rr, cc)] = nm
            positions = None
            for spec, a, b in key_refs:
                if b is not None or a[0] is None or a[1] is None:
                    continue
                target = s["name"] if spec is None else (spec[1] if spec[0] == "sheet" else None)
                if target is None:
                    continue
                if a[0][0] and a[1][0]:
                    keys = [((target, a[0][1], a[1][1]), int(n))]
                elif n <= 50:                       # a relative ref in a few cells still names fixed cells
                    if positions is None:
                        positions = np.argwhere(s["shape"] == sid)
                    keys = []
                    for pr, pc in positions:
                        R, C = int(pr) + s["r0"], int(pc) + s["c0"]
                        keys.append(((target, a[0][1] if a[0][0] else R + a[0][1], a[1][1] if a[1][0] else C + a[1][1]), 1))
                else:
                    continue
                for key, m in keys:
                    tally[key] += m
                    for f in steer:
                        funcs_seen[key][f] += m
    # a run of neighbouring cells each read about as often is a column (row) of labels, not a switch
    strong = {k: v for k, v in tally.items() if v >= min_cells}
    labels = set()
    for (sh, r, c), v in strong.items():
        for nb in ((sh, r + 1, c), (sh, r - 1, c), (sh, r, c + 1), (sh, r, c - 1)):
            w = strong.get(nb)
            if w and 0.5 <= v / w <= 2:
                labels.add((sh, r, c))
    out = []
    for (sh, r, c), n in tally.most_common():
        if (sh, r, c) in labels:
            continue
        if n < min_cells or len(out) >= top:
            break
        s = book.by_name.get(sh)
        if s is None:
            continue
        rr, cc = r - s["r0"], c - s["c0"]
        if not (0 <= rr < s["shape"].shape[0] and 0 <= cc < s["shape"].shape[1]):
            continue
        k = int(s["shape"][rr, cc])
        if k not in (1, 2):
            continue                      # a formula or an empty cell steers nothing the user chooses
        label = None
        for dc in range(1, 4):            # a text label to the left, if any
            if cc - dc >= 0 and s["shape"][rr, cc - dc] == 2:
                label = (r, c - dc)
                break
        out.append({"cell": f"{sh}!{letter(c)}{r}", "name": named.get((sh, r, c)), "formula_cells": n, "kind": "number" if k == 1 else "text",
                    "via": dict(funcs_seen[(sh, r, c)].most_common(4)), "label_cell": f"{letter(label[1])}{label[0]}" if label else None})
    return out


def mind_constructs(book: Book):
    c = Counter()
    for t in book.texts[10:]:
        if t is None:
            continue
        u = t.upper()
        for f in ("MM_LOOP(", "MM_LOOPINSTANCE(", "MM_RESULT(", "MM_INSTANCE(", "MM_INSTANCEKEY("):
            if f in u:
                c[f[:-1]] += 1
        if parse_r1c1(t).funcs & RANDOM_FUNCS:
            c["random functions"] += 1
    return dict(c)


def letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main(tag):
    book = Book(Path(__file__).parent / "data" / f"{tag}.pkl")
    res = {"file": book.file, "sheets": len(book.sheets),
           "formulas": int(sum((s["shape"] >= 10).sum() for s in book.sheets)),
           "mind": mind_constructs(book), "sheet_copies": sheet_copies(book), "inside": [],
           "switches": switch_cells(book)}
    for s in book.sheets:
        for axis in (1, 0):
            for f in inside_sheet(book, s, axis):
                off = s["c0"] if axis == 1 else s["r0"]
                a = f["start"] + off
                b = a + f["copies"] * f["width"] - 1
                f["where"] = f"{letter(a)}:{letter(b)}" if axis == 1 else f"rows {a}-{b}"
                res["inside"].append(f)
    pickle.dump(res, open(Path(__file__).parent / "data" / f"{tag}.result.pkl", "wb"))
    return res


if __name__ == "__main__":
    import json
    for tag in sys.argv[1:]:
        r = main(tag)
        print(json.dumps(r, indent=1, default=str)[:20000])
