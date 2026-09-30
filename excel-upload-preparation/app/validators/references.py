"""Reference checks Mind's converter or runner enforces (1.7.3), each learnt
from a real refusal:

  * FRM-005 circular references -- Mind's runner stops with "Run error:
    Circular reference found" (model RAROK, 2026-09-30). Excel tolerates a
    cycle when iterative calculation is on; Mind never does. The dependency
    graph of every formula cell (references, ranges, defined names) is
    searched for cycles and each cycle is reported as its chain of cells.
  * FRM-006 3-D references -- 'First:Last'!A1 spans every sheet between two
    tabs. Mind reads 'First:Last' as one sheet name and fails with "Sheet
    '...' not found in workbook ... on compiling formula" (model PVFP,
    2026-09-30). The sheets the reference spans are listed in tab order and
    the equivalent explicit formula is proposed.

Neither has an automatic repair: a cycle is broken by a modelling decision
(the Mind documentation's mechanism for a value feeding the next round is
MM_ITERATIONS with /iterationinput and /iterationoutput), and a 3-D
reference is rewritten by hand or with the assistant from the proposal.
"""
from __future__ import annotations

import re
from bisect import bisect_left, bisect_right
from typing import Any, Iterator

from ..formula_utils import called_functions, cell_refs_in_formula, mask_strings, name_tokens, ref_text
from ._common import finding

DYNAMIC_REFERENCE_FUNCTIONS = {"INDIRECT", "OFFSET"}
# ROW($A10), COLUMN(B:B), ROWS(...), COLUMNS(...) use a reference as a coordinate and never read its
# value: Excel does not count it as a dependency (=ROW(A1) in A1 is not circular), so neither do we.
# Arguments a function evaluates only on some path: IF's two branches, IFERROR/IFNA's fallback,
# CHOOSE's and SWITCH's values, IFS' pairs after the first condition. A cycle that only closes
# through such an argument is one Excel computes (the branches never evaluate together).
CONDITIONAL_ARGS: dict[str, Any] = {"IF": {1, 2}, "IFERROR": {1}, "IFNA": {1}, "CHOOSE": "rest", "SWITCH": "rest", "IFS": "rest"}
COORDINATE_ARGS_RE = re.compile(r"(?<![A-Za-z0-9_.])(?:ROW|COLUMN|ROWS|COLUMNS)\(([^()]*)\)", re.IGNORECASE)
MAX_CYCLES_REPORTED = 12
MAX_CHAIN = 12

QUOTED_3D_RE = re.compile(r"'((?:[^']|'')*:(?:[^']|'')*)'!")
BARE_3D_RE = re.compile(r"(?<![A-Za-z0-9_.!'\[])([A-Za-z0-9_.]+):([A-Za-z0-9_.]+)!")
SHEET_NEEDS_QUOTES_RE = re.compile(r"[^A-Za-z0-9_.]")


# --- FRM-005 -----------------------------------------------------------------------------
def unconditional_text(formula: str) -> str:
    """The formula with every conditionally evaluated argument blanked (same
    length, so positions hold): what is left is read on every calculation."""
    masked = mask_strings(formula)
    out = list(formula)
    for name, positions in CONDITIONAL_ARGS.items():
        pattern = re.compile(rf"(?<![A-Za-z0-9_.])(?:_xl[a-z]+\.)?{name}\s*\(", re.IGNORECASE)
        for m in pattern.finditer(masked):
            i = m.end()
            depth = 1
            arg_start = i
            k = 0
            spans = []
            while i < len(masked) and depth > 0:
                ch = masked[i]
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        spans.append((k, arg_start, i))
                elif ch in ",;" and depth == 1:
                    spans.append((k, arg_start, i))
                    k += 1
                    arg_start = i + 1
                i += 1
            for k, a, b in spans:
                if (positions == "rest" and k >= 1) or (positions != "rest" and k in positions):
                    for j in range(a, b):
                        out[j] = " "
    return "".join(out)


class _Graph:
    """Formula cells as nodes; the edges of a node -- every formula cell it
    reads directly, through a range or through a defined name -- are produced
    on demand and never stored: a model can hold 100k formulas each summing a
    column of formulas, and materialising that is what runs out of memory."""

    def __init__(self, wb0: dict[str, Any]) -> None:
        self.formulas = wb0.get("formulas", [])
        self.sheets = {s["name"] for s in wb0.get("sheets", [])}
        # sheet -> row -> sorted (col, node), for the rectangle lookups; a multi-cell
        # array formula owns every cell of its range
        self.rows: dict[str, dict[int, list[tuple[int, int]]]] = {}
        self.row_keys: dict[str, list[int]] = {}
        self.label: list[str] = []
        for i, f in enumerate(self.formulas):
            sheet = f["sheet"]
            span = cell_refs_in_formula("=" + str(f.get("array_ref") or f["cell"]))
            self.label.append(f"{sheet}!{f['cell']}")
            if not span:
                continue
            ref = span[0]
            rows = self.rows.setdefault(sheet, {})
            for r in range(ref["r1"], ref["r2"] + 1):
                row = rows.setdefault(r, [])
                for c in range(ref["c1"], ref["c2"] + 1):
                    row.append((c, i))
        for sheet, rows in self.rows.items():
            for r in rows:
                rows[r].sort()
            self.row_keys[sheet] = sorted(rows)
        self.names: dict[str, list[dict[str, Any]]] = {}
        for d in wb0.get("defined_names", []):
            value = str(d.get("value") or "")
            if "#REF!" in value or not d.get("name"):
                continue
            refs = [r for r in cell_refs_in_formula("=" + value) if r.get("sheet") in self.sheets]
            if refs:
                self.names.setdefault(str(d["name"]).upper(), []).extend(refs)
        self.visited = 0  # member cells looked at, against the budget
        self.range_ids: dict[tuple[str, int, int, int, int], int] = {}
        self.ranges: list[dict[str, Any]] = []

    def refs_of(self, i: int, unconditional: bool = False) -> tuple[list[dict[str, Any]], bool]:
        f = self.formulas[i]
        formula = f.get("formula") or ""
        dynamic = bool(DYNAMIC_REFERENCE_FUNCTIONS & {fn.upper() for fn in called_functions(formula)})
        if unconditional:
            formula = unconditional_text(formula)
        refs = []
        for r in cell_refs_in_formula(COORDINATE_ARGS_RE.sub(lambda m: m.group(0)[: m.start(1) - m.start(0)] + " " * len(m.group(1)) + ")", formula)):
            sheet = r.get("sheet") or f["sheet"]
            if sheet.startswith("["):  # another workbook
                continue
            refs.append({**r, "sheet": sheet})
        for tok in name_tokens(formula):
            refs.extend(self.names.get(tok.upper(), []))
        return refs, dynamic

    def members(self, ref: dict[str, Any]) -> Iterator[int]:
        """The formula cells inside a rectangle."""
        rows = self.rows.get(ref["sheet"])
        if not rows:
            return
        keys = self.row_keys[ref["sheet"]]
        for k in range(bisect_left(keys, ref["r1"]), bisect_right(keys, ref["r2"])):
            cols = rows[keys[k]]
            lo = bisect_left(cols, (ref["c1"], -1))
            hi = bisect_right(cols, (ref["c2"], 1 << 30))
            self.visited += hi - lo
            for _, node in cols[lo:hi]:
                yield node

    # A multi-cell range is a node of its own, shared by every formula that reads it,
    # so a column of 20k formulas summed by 2k formulas is walked once, not 2k times.
    def range_node(self, ref: dict[str, Any]) -> int:
        key = (ref["sheet"], ref["r1"], ref["r2"], ref["c1"], ref["c2"])
        node = self.range_ids.get(key)
        if node is None:
            node = len(self.formulas) + len(self.ranges)
            self.range_ids[key] = node
            self.ranges.append(ref)
        return node

    def neighbours(self, node: int) -> Iterator[int]:
        """What a node reads: a formula yields single cells and range nodes; a
        range node yields the formula cells inside it (repeats possible)."""
        n = len(self.formulas)
        if node >= n:
            yield from self.members(self.ranges[node - n])
            return
        refs, _ = self.refs_of(node)
        for ref in refs:
            if ref["r1"] == ref["r2"] and ref["c1"] == ref["c2"]:
                yield from self.members(ref)
            else:
                yield self.range_node(ref)

    def formula_neighbours(self, i: int, unconditional: bool = False) -> set[int]:
        """The formula cells formula i reads, ranges expanded (small use only)."""
        refs, _ = self.refs_of(i, unconditional)
        out: set[int] = set()
        for ref in refs:
            out |= set(self.members(ref))
        return out


WORK_BUDGET = 200_000_000  # member cells looked at before the search gives up (memory and time)


def _cycles(graph: _Graph) -> tuple[list[list[int]], bool]:
    """Tarjan's strongly connected components, iteratively and with the edges
    produced on demand. Range nodes are created as they are met, so the
    bookkeeping lists grow with them. A component is a cycle when it holds
    more than one node (a formula and the range that contains it is one), or
    when a formula reads its own cell. Returns (components of formula nodes,
    gave_up)."""
    n = len(graph.formulas)
    index = [-1] * n
    low = [0] * n
    on_stack = [False] * n

    def ensure(w: int) -> None:
        while len(index) <= w:
            index.append(-1)
            low.append(0)
            on_stack.append(False)

    stack: list[int] = []
    comps: list[list[int]] = []
    self_loops: set[int] = set()  # a formula reading its own cell directly
    counter = 0
    for root in range(n):
        if index[root] != -1:
            continue
        work: list[tuple[int, Iterator[int]]] = [(root, graph.neighbours(root))]
        index[root] = low[root] = counter
        counter += 1
        stack.append(root)
        on_stack[root] = True
        while work:
            if graph.visited > WORK_BUDGET:
                return comps, True
            v, todo = work[-1]
            w = next(todo, None)
            if w is not None:
                if w == v:
                    self_loops.add(v)
                    continue
                ensure(w)
                if index[w] == -1:
                    index[w] = low[w] = counter
                    counter += 1
                    stack.append(w)
                    on_stack[w] = True
                    work.append((w, graph.neighbours(w)))
                elif on_stack[w]:
                    low[v] = min(low[v], index[w])
                continue
            work.pop()
            if work:
                u = work[-1][0]
                low[u] = min(low[u], low[v])
            if low[v] == index[v]:
                comp = []
                while True:
                    x = stack.pop()
                    on_stack[x] = False
                    comp.append(x)
                    if x == v:
                        break
                cells = sorted(x for x in comp if x < n)
                if cells and (len(comp) > 1 or v in self_loops):
                    comps.append(cells)
    return comps, False


def _has_cycle(nodes: list[int], adj: dict[int, set[int]]) -> bool:
    """A cycle inside a small subgraph (a node reading itself counts)."""
    inside = set(nodes)
    colour: dict[int, int] = {}
    for start in nodes:
        if start in colour:
            continue
        stack = [(start, iter(sorted(adj[start] & inside)))]
        colour[start] = 1
        while stack:
            v, it = stack[-1]
            w = next(it, None)
            if w is None:
                colour[v] = 2
                stack.pop()
                continue
            c = colour.get(w)
            if c == 1:
                return True
            if c is None:
                colour[w] = 1
                stack.append((w, iter(sorted(adj[w] & inside))))
    return False


def _chain(graph: _Graph, comp: list[int], adj: dict[int, set[int]]) -> list[str]:
    """One closed walk inside the component, from its first cell back to itself."""
    start = comp[0]
    inside = set(comp)
    path = [start]
    seen = {start}
    node = start
    while len(path) < MAX_CHAIN:
        nxt = sorted(w for w in adj[node] if w in inside)
        if not nxt:
            break
        if start in nxt:
            path.append(start)
            break
        step = next((w for w in nxt if w not in seen), nxt[0])
        if step in seen:
            path.append(step)
            break
        seen.add(step)
        path.append(step)
        node = step
    return [graph.label[i] for i in path]


def circular_references(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FRM-005: cycles among the workbook's formulas."""
    wb0 = analysis["workbooks"][0]
    graph = _Graph(wb0)
    comps, gave_up = _cycles(graph)
    iterative = bool(analysis["features"].get("iterative_calculation"))
    dynamic = sum(1 for f in graph.formulas if DYNAMIC_REFERENCE_FUNCTIONS & {fn.upper() for fn in called_functions(f.get("formula") or "")})
    adj = {i: graph.formula_neighbours(i) for comp in comps for i in comp}
    always = {i: graph.formula_neighbours(i, unconditional=True) for comp in comps for i in comp}
    cycles = []
    for comp in comps:
        chain = _chain(graph, comp, adj)
        first = graph.formulas[comp[0]]
        cycles.append({
            "sheet": first["sheet"], "cell": first["cell"], "cells": len(comp), "chain": chain, "formula": str(first.get("formula") or "")[:160],
            # closes on every calculation, or only through an IF/IFERROR/CHOOSE branch
            "unconditional": _has_cycle(comp, always),
        })
    cycles.sort(key=lambda c: (not c["unconditional"], -c["cells"], c["sheet"], c["cell"]))
    hard = [c for c in cycles if c["unconditional"]]
    soft = [c for c in cycles if not c["unconditional"]]
    observed = {
        "cycles": cycles[:MAX_CYCLES_REPORTED],
        "cycle_count": len(comps),
        "unconditional_cycles": len(hard),
        "conditional_cycles": len(soft),
        "cells_in_cycles": sum(len(c) for c in comps),
        "iterative_calculation": iterative,
        "dynamic_reference_formulas_not_followed": dynamic,
        "formulas": len(graph.formulas),
        "member_cells_followed": graph.visited,
        "gave_up": gave_up,
        # the cells the Fix panel lists
        "sites": [{"sheet": c["sheet"], "cell": c["cell"], "chain": " -> ".join(c["chain"])} for c in cycles[:MAX_CYCLES_REPORTED]],
    }
    fix = ("No automatic repair: break each cycle by hand (a value that feeds the next round is what MM_ITERATIONS with /iterationinput and /iterationoutput is for; "
           "otherwise reference the previous period or a fixed starting value).")
    dyn = f" {dynamic} formula(s) use INDIRECT/OFFSET, whose targets cannot be followed here." if dynamic else ""
    if hard:
        first = hard[0]
        return finding(
            "ERROR",
            f"{len(hard)} circular reference(s) among the formulas, read on every calculation: Mind's runner stops on the first one "
            f"('Run error: Circular reference found'){' -- Excel tolerates them because iterative calculation is on in this workbook' if iterative else ''}. "
            f"First: {' -> '.join(first['chain'])}. "
            + (f"{len(soft)} more cycle(s) close only through an IF/IFERROR/CHOOSE branch (see the warning text). " if soft else "")
            + fix + dyn,
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if soft:
        first = soft[0]
        return finding(
            "WARNING",
            f"{len(soft)} circular reference(s) that close only through an IF/IFERROR/CHOOSE branch, e.g. {' -> '.join(first['chain'])}: Excel computes them "
            "because the two branches never evaluate together; whether Mind's calculation order accepts them is not documented -- a run in Mind tells "
            "('Run error: Circular reference found' if not). " + fix + dyn,
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
            evidence="INFERENCE",
        )
    if gave_up:
        return finding(
            "NOT_SUPPORTED",
            f"The dependency search stopped after {WORK_BUDGET:,} cell lookups without finding a cycle ({len(graph.formulas):,} formulas): this workbook is too large to "
            "follow completely here. Mind refuses circular references: check Excel's circular-reference warning and turn iterative calculation off"
            + (" (it is on in this workbook)" if iterative else "") + ".",
            observed,
            evidence="INFERENCE",
        )
    if iterative:
        return finding(
            "WARNING",
            "Excel's iterative calculation is on in this workbook (the setting that lets circular references compute), but no cycle was found among the static references"
            + (f"; {dynamic} formula(s) use INDIRECT/OFFSET, whose targets cannot be followed here" if dynamic else "")
            + ". Mind refuses circular references: turn the setting off in Excel and check that no cell shows a circular-reference warning.",
            observed,
            evidence="INFERENCE",
        )
    return finding("PASS", f"No circular reference among {len(graph.formulas):,} formula(s) ({graph.visited:,} dependencies followed).", observed)


# --- FRM-006 -----------------------------------------------------------------------------
def _quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'" if SHEET_NEEDS_QUOTES_RE.search(name) or name[:1].isdigit() else name


def _split_3d(text: str) -> tuple[str, str] | None:
    """'First:Last' (quotes removed) -> (First, Last); sheet names never contain ':'."""
    if ":" not in text:
        return None
    first, last = text.split(":", 1)
    return first.replace("''", "'"), last.replace("''", "'")


def three_d_references(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FRM-006: 'First:Last'!ref references spanning several sheets."""
    wb0 = analysis["workbooks"][0]
    order = [s["name"] for s in wb0.get("sheets", [])] + [s["name"] for s in wb0.get("ignored_sheets", [])]
    known = set(order)
    sites = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        masked = mask_strings(formula, mask_sheet_quotes=False)
        found: list[tuple[int, int, str, str]] = []  # start, end, first, last
        for m in QUOTED_3D_RE.finditer(masked):
            pair = _split_3d(formula[m.start(1) : m.end(1)])
            if pair:
                found.append((m.start(), m.end(), *pair))
        for m in BARE_3D_RE.finditer(masked):
            if m.group(1) in known and m.group(2) in known:
                found.append((m.start(), m.end(), m.group(1), m.group(2)))
        if not found:
            continue
        found.sort()
        suggested = formula
        spans = []
        for start, end, first, last in reversed(found):
            if first in order and last in order:
                i, j = sorted((order.index(first), order.index(last)))
                sheets = order[i : j + 1]
            else:
                sheets = []
            spans.append({"reference": formula[start:end], "first": first, "last": last, "sheets": sheets})
            if sheets:
                # the reference token is followed by the cell/range it applies to: repeat it per sheet
                m = re.match(r"(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})", suggested[end:])
                if m:
                    cell = m.group(1)
                    replacement = ",".join(f"{_quote_sheet(s)}!{cell}" for s in sheets)
                    suggested = suggested[:start] + replacement + suggested[end + len(cell) :]
        spans.reverse()
        unresolved = [s for s in spans if not s["sheets"]]
        sites.append({
            "sheet": f["sheet"],
            "cell": f["cell"],
            "formula": formula[:200],
            "references": spans,
            "suggested_formula": None if unresolved else suggested[:600],
            "issue": (f"sheet(s) not in the workbook: {', '.join(x for s in unresolved for x in (s['first'], s['last']) if x not in known)}" if unresolved else None),
        })
    if sites:
        first = sites[0]
        ref0 = first["references"][0]
        n_sheets = len(ref0["sheets"])
        return finding(
            "ERROR",
            f"{len(sites)} formula(s) use a 3-D reference ('First:Last'!cell, every sheet between two tabs): Mind reads 'First:Last' as one sheet name and refuses the model "
            f"(\"Sheet '{ref0['first']}:{ref0['last']}' not found in workbook ... on compiling formula\"). First: {first['sheet']}!{first['cell']} = {first['formula'][:80]}"
            + (f" spans {n_sheets} sheet(s) in tab order: {', '.join(ref0['sheets'][:8])}{' ...' if n_sheets > 8 else ''}." if n_sheets else f" ({first['issue']}).")
            + " Fix by hand or with the assistant: list every sheet explicitly, as in the proposed formula"
            + (f" {first['suggested_formula'][:160]}" if first["suggested_formula"] else "")
            + ". No automatic repair.",
            {"sites": sites[:80]},
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    return finding("PASS", "No 3-D reference ('First:Last'!cell) in the workbook's formulas.", {"sites": []})
