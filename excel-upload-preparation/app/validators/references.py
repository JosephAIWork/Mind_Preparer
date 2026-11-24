"""Reference  checks Mind's  converter or runner enforces, each learnt from a
real refusal:

  * FRM-005 circular references -- Mind's runner stops with "Run error:
    Circular reference found" (observed 2026-09-30). Excel tolerates a
    cycle when iterative calculation is on; Mind  never does. The dependency
    graph of every formula cell (references, ranges, defined  names) is
    searched for cycles and each  cycle is reported as its chain of  cells.
  * FRM-006 3-D references -- 'First:Last'!A1 spans every sheet between two
    tabs. Mind  reads 'First:Last' as one sheet name and  fails with "Sheet
    '...' not found in workbook ... on compiling formula" (observed
    2026-09-30). The sheets the reference spans are listed in tab order and
    the equivalent explicit formula is proposed -- without the sheets that
    hold no grid: Mind creates no spreadsheet for those (divider tabs such
    as '>>>' that hold only a heading: "Spreadsheet not found error",
    2026-10-01), and an empty sheet adds nothing to  SUM/AVERAGE/COUNT/MIN/
    MAX. The same rule flags any other formula that reads a grid-less sheet.

Neither has an automatic repair: a cycle is broken by a modelling decision
(the Mind documentation's mechanism for a value feeding the next round is
MM_ITERATIONS with /iterationinput and /iterationoutput), and a 3-D
reference is rewritten by hand or with the assistant from the proposal.
"""
from __future__ import annotations

import re
from bisect import bisect_left, bisect_right
from typing import Any, Iterator

from ..formula_utils import REF_RE, called_functions, cell_refs_in_formula, mask_strings, name_tokens, parse_ref, ref_text
from ._common import finding

DYNAMIC_REFERENCE_FUNCTIONS = {"INDIRECT", "OFFSET"}
# ROW($A10),  COLUMN(B:B), ROWS(...),  COLUMNS(...) and OFFSET's first argument use a reference as a
# coordinate and never read its value: Excel does not count it as a dependency (=ROW(A1) in A1 and
# =OFFSET(A1,0,1) in A1 are not  circular), so neither do we. "all" = every argument, else the  positions.
# Arguments a function evaluates only on  some path: IF's two branches, IFERROR/IFNA's fallback,
# CHOOSE's  and SWITCH's values, IFS' pairs after the first condition. A cycle  that only closes
# through such an argument is  one Excel computes (the branches never evaluate together).
CONDITIONAL_ARGS: dict[str, Any] = {"IF": {1, 2}, "IFERROR": {1}, "IFNA": {1}, "CHOOSE": "rest", "SWITCH": "rest", "IFS": "rest"}
# A  lookup reads  ONE cell of its table. A cycle that closes only through the table of an
# INDEX / VLOOKUP / HLOOKUP / MATCH / LOOKUP / XLOOKUP is not a cycle for Excel  (once INDIRECT
# became direct references, =INDEX(<a large input table>, ...) reached a  cell that reads the sheet
# the lookup  sits on; Excel reported no circular reference)  -- it is reported as a conditional one.
LOOKUP_TABLE_ARGS: dict[str, Any] = {"INDEX": {0}, "VLOOKUP": {1}, "HLOOKUP": {1}, "MATCH": {1}, "LOOKUP": "rest", "XLOOKUP": {1, 2}}
# CELL("filename",  A1) in  A1 asks where the cell is, not what it holds: no cycle for Excel either
# (Model D (an IFRS 17 tool), 2026-11-07). Only CELL("contents", A1) reads the value; it is rare and  FRM-012 freezes CELL anyway.
COORDINATE_ARGS: dict[str, Any] = {"ROW": "all", "COLUMN": "all", "ROWS": "all", "COLUMNS": "all", "OFFSET": {0}, "CELL": {1}}
MAX_CYCLES_REPORTED = 12
# INDIRECT("'"&sheet&"'!"&ADDRESS(ROW(),COLUMN())):  the cell  reads its own address on a sheet chosen
# at run time. Mind resolved that to the cell itself and stopped ("Run error: Circular reference found",
# 70k such cells on one sheet of a  real model, 2026-09-30).
SELF_ADDRESS_RE = re.compile(r"ADDRESS\(\s*ROW\(\s*\)\s*,\s*COLUMN\(\s*\)", re.IGNORECASE)
MAX_CHAIN = 12

QUOTED_3D_RE = re.compile(r"'((?:[^']|'')*:(?:[^']|'')*)'!")
BARE_3D_RE = re.compile(r"(?<![\w.!'\[])([\w.]+):([\w.]+)!")
SHEET_NEEDS_QUOTES_RE = re.compile(r"[^\w.]")


# ---  FRM-005 -----------------------------------------------------------------------------
def blank_arguments(formula: str, spec: dict[str, Any]) -> str:
    """The  formula with  the chosen arguments of the chosen functions blanked
    (same length, so positions hold). `spec`: function -> set of argument
    positions, "rest" (every argument but the first) or "all"."""
    masked = mask_strings(formula)
    out = list(formula)
    for name, positions in spec.items():
        pattern = re.compile(rf"(?<![\w.])(?:_xl[a-z]+\.)?{name}\s*\(", re.IGNORECASE)
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
                if positions == "all" or (positions == "rest" and k >= 1) or (positions not in ("all", "rest") and k in positions):
                    for j in range(a, b):
                        out[j] = " "
    return "".join(out)


def unconditional_text(formula: str) -> str:
    """What  a formula  reads on every calculation: its conditionally evaluated arguments blanked."""
    return blank_arguments(blank_arguments(formula, CONDITIONAL_ARGS), LOOKUP_TABLE_ARGS)


class _Graph:
    """Formula  cells as  nodes; the edges of a node -- every formula cell it
    reads directly, through a range or through a defined name -- are produced
    on demand and never stored: a model can hold  100k formulas each summing a
    column of formulas, and materialising that is  what runs out of memory."""

    def __init__(self, wb0: dict[str, Any]) -> None:
        self.formulas = wb0.get("formulas", [])
        self.sheets = {s["name"] for s in wb0.get("sheets", [])}
        # sheet ->  row -> sorted (col, node), for the  rectangle lookups; a multi-cell
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
        self.visited = 0  # member  cells looked at, against the budget
        self.range_ids: dict[tuple[str, int, int, int, int], int] = {}
        self.ranges: list[dict[str, Any]] = []

    def refs_of(self, i: int, unconditional: bool = False) -> tuple[list[dict[str, Any]], bool]:
        f = self.formulas[i]
        formula = f.get("formula") or ""
        dynamic = bool(DYNAMIC_REFERENCE_FUNCTIONS & {fn.upper() for fn in called_functions(formula)})
        if unconditional:
            formula = unconditional_text(formula)
        refs = []
        for r in cell_refs_in_formula(blank_arguments(formula, COORDINATE_ARGS)):
            sheet = r.get("sheet") or f["sheet"]
            if sheet.startswith("["):  # another workbook
                continue
            refs.append({**r, "sheet": sheet})
        for tok in name_tokens(formula):
            refs.extend(self.names.get(tok.upper(), []))
        return refs, dynamic

    def members(self, ref: dict[str, Any]) -> Iterator[int]:
        """The  formula cells inside a rectangle."""
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
        """The formula cells formula i reads,  ranges expanded (small use only)."""
        refs, _ = self.refs_of(i, unconditional)
        out: set[int] = set()
        for ref in refs:
            out |= set(self.members(ref))
        return out


WORK_BUDGET = 200_000_000  # member  cells looked  at before the search gives up (memory and time)


def _cycles(graph: _Graph) -> tuple[list[list[int]], bool]:
    """Tarjan's  strongly connected  components, iteratively and with the edges
    produced on demand. Range nodes are created as they are met, so the
    bookkeeping lists grow with them. A component is a cycle when it holds
    more than  one node (a formula and the range that contains it is  one), or
    when a formula reads its  own cell. Returns (components of formula nodes,
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
    self_loops: set[int] = set()  # a  formula reading its own cell directly
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
    """A  cycle inside  a small subgraph (a node reading itself counts)."""
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
    """One  closed walk  inside the component, from its first cell back to itself."""
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
    """FRM-005:  cycles among  the workbook's formulas."""
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
    self_address = [
        {"sheet": f["sheet"], "cell": f["cell"], "formula": str(f.get("formula") or "")[:160]}
        for f in graph.formulas
        if "INDIRECT" in str(f.get("formula") or "").upper() and SELF_ADDRESS_RE.search(mask_strings(str(f.get("formula") or "")))
    ]
    by_sheet: dict[str, int] = {}
    for x in self_address:
        by_sheet[x["sheet"]] = by_sheet.get(x["sheet"], 0) + 1
    observed = {
        "cycles": cycles[:MAX_CYCLES_REPORTED],
        "self_addressing_indirect": {"count": len(self_address), "by_sheet": by_sheet, "first": self_address[:12]},
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
    if self_address:
        first = self_address[0]
        sheets = ", ".join(f"{k} ({v:,})" for k, v in sorted(by_sheet.items(), key=lambda kv: -kv[1])[:6])
        return finding(
            "ERROR",
            f"{len(self_address):,} cell(s) read their own address on a sheet chosen at run time -- INDIRECT(...ADDRESS(ROW(),COLUMN())) -- on {sheets}. "
            f"Mind resolves that reference to the cell itself and stops ('Run error: Circular reference found' at the first such cell, as it did on a real model built this way). First: {first['sheet']}!{first['cell']}. "
            "No automatic repair: replace the run-time sheet choice by direct references to the sheet (or by a lookup Mind can follow) by hand or with the assistant."
            + (f" Also {len(hard)} unconditional and {len(soft)} conditional cycle(s) among the other formulas (see observed.cycles)." if comps else "")
            + dyn,
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if hard:
        first = hard[0]
        return finding(
            "ERROR",
            f"{len(hard)} circular reference(s) among the formulas, read on every calculation: Mind's runner stops on the first one "
            f"('Run error: Circular reference found'){' -- Excel tolerates them because iterative calculation is on in this workbook' if iterative else ''}. "
            f"First: {' -> '.join(first['chain'])}. "
            + (f"{len(soft)} more cycle(s) close only through an IF/IFERROR/CHOOSE branch or through the table of a lookup (INDEX, VLOOKUP, MATCH ...) (see the warning text). " if soft else "")
            + fix + dyn,
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if soft:
        first = soft[0]
        return finding(
            "WARNING",
            f"{len(soft)} circular reference(s) that close only through an IF/IFERROR/CHOOSE branch or through the table of a lookup (INDEX, VLOOKUP, MATCH ...), e.g. {' -> '.join(first['chain'])}: Excel computes them "
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
    """'First:Last'  (quotes removed)  -> (First, Last); sheet names never contain ':'."""
    if ":" not in text:
        return None
    first, last = text.split(":", 1)
    return first.replace("''", "'"), last.replace("''", "'")


def three_d_references(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = 80) -> dict:
    """FRM-006:  'First:Last'!ref references  spanning several sheets, and
    references to sheets Mind does not import. `limit` caps the sites kept
    in the finding (None: every site, for a caller that applies the proposals)."""
    from ..inventory import ignored_sheet_names, sheet_order

    wb0 = analysis["workbooks"][0]
    # tab order, the sheets the  scan skipped at their own index: a span 'First:Last' runs over
    # them  like any other sheet (a skipped  sheet exists; its cells were just never  read)
    order = sheet_order(wb0)
    known = set(order)
    unscanned = set(ignored_sheet_names(analysis))
    # sheets Mind creates no spreadsheet for: nothing on them makes a  grid. A skipped sheet's grids
    # are unknown and  not guessed: it is kept in the proposal, as without the skip
    gridless = {s["name"] for s in wb0.get("sheets", []) if not s.get("grids")}
    sites = []
    gridless_sites = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        masked = mask_strings(formula, mask_sheet_quotes=False)
        if gridless:
            targets = sorted({r["sheet"] for r in cell_refs_in_formula(formula) if r.get("sheet") in gridless})
            if targets:
                gridless_sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:200], "sheets": targets})
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
            kept = [x for x in sheets if x not in gridless]
            spans.append({"reference": formula[start:end], "first": first, "last": last, "sheets": sheets, "omitted_empty": [x for x in sheets if x in gridless], "unscanned": [x for x in sheets if x in unscanned]})
            if sheets and kept:
                # the reference token is followed by the cell/range it applies to: repeat it per sheet
                m = re.match(r"(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})", suggested[end:])
                if m:
                    cell = m.group(1)
                    replacement = ",".join(f"{_quote_sheet(s)}!{cell}" for s in kept)
                    suggested = suggested[:start] + replacement + suggested[end + len(cell) :]
            elif sheets:
                sheets = []  # every sheet of the span is empty for Mind: nothing to propose
        spans.reverse()
        unresolved = [s for s in spans if not s["sheets"]]
        sites.append({
            "sheet": f["sheet"],
            "cell": f["cell"],
            "formula": formula[:200],
            "references": spans,
            "suggested_formula": None if unresolved else suggested,  # whole, never cut: it is meant to be applied
            "issue": (
                f"sheet(s) not in the workbook: {', '.join(x for s in unresolved for x in (s['first'], s['last']) if x not in known)}" if any(x not in known for s in unresolved for x in (s["first"], s["last"]))
                else "every sheet of the span holds no grid: Mind imports none of them, so the reference has nothing to read" if unresolved
                else None
            ),
        })
    cap = slice(None) if limit is None else slice(0, limit)
    observed = {"sites": sites[cap], "sites_count": len(sites), "gridless_sheets": sorted(gridless), "formulas_reading_gridless_sheets": gridless_sites[cap], "formulas_reading_gridless_sheets_count": len(gridless_sites)}
    if sites or gridless_sites:
        parts = []
        if sites:
            first = sites[0]
            ref0 = first["references"][0]
            n_sheets = len(ref0["sheets"])
            parts.append(
                f"{len(sites)} formula(s) use a 3-D reference ('First:Last'!cell, every sheet between two tabs): Mind reads 'First:Last' as one sheet name and refuses the model "
                f"(\"Sheet '{ref0['first']}:{ref0['last']}' not found in workbook ... on compiling formula\"). First: {first['sheet']}!{first['cell']} = {first['formula'][:80]}"
                + (f" spans {n_sheets} sheet(s) in tab order: {', '.join(ref0['sheets'][:8])}{' ...' if n_sheets > 8 else ''}" if n_sheets else f" ({first['issue']})")
                + (f"; {', '.join(ref0['omitted_empty'][:6])} hold no grid and are left out of the proposal (Mind imports no empty sheet)" if ref0.get("omitted_empty") else "")
                + ". The Prep action 'Write 3-D references out sheet by sheet' lists every sheet explicitly, as in the proposed formula"
                + (f" {first['suggested_formula'][:160]}" if first["suggested_formula"] else "")
                + " (a sheet left out must hold nothing at those cells; otherwise by hand)."
            )
        if gridless_sites:
            g0 = gridless_sites[0]
            parts.append(
                f"{len(gridless_sites)} formula(s) read a sheet Mind does not import because nothing on it makes a grid ({', '.join(sorted({x for g in gridless_sites for x in g['sheets']})[:6])}): "
                f"Mind stops with 'Spreadsheet not found'. First: {g0['sheet']}!{g0['cell']} = {g0['formula'][:80]}. Point the formula at a sheet with a grid, or give that sheet a grid, by hand."
            )
        first_site = sites[0] if sites else gridless_sites[0]
        return finding(
            "ERROR",
            " ".join(parts) + (" No automatic repair for the formulas reading a sheet without a grid." if gridless_sites else ""),
            observed,
            location={"sheet": first_site["sheet"], "cell": first_site["cell"]},
        )
    return finding("PASS", "No 3-D reference ('First:Last'!cell) and no reference to a sheet without a grid in the workbook's formulas.", observed)


# --- FRM-007 -----------------------------------------------------------------------------
# INDIRECT's result used as a value -- compared, multiplied, divided, negated, concatenated:
# Mind  returns a reference object for INDIRECT and cannot turn it into a number ("Run error:
# Unable to cast object of type 'AM.Models.AMReference' to type 'System.IConvertible'",
# observed 2026-10-01). Where the text INDIRECT builds can be worked out from constants (string
# literals, cells holding text, ROW()/COLUMN(), simple arithmetic and text functions), the
# equivalent direct reference is proposed; a target that does not exist in the workbook is an
# error in Excel already (#REF!) and that call becomes NA() in its place -- never the whole
# formula: the call may  sit in an IF branch  that is never taken.
VALUE_OPERATORS = set("=<>+-*/^&")
# functions  whose argument  is a value, never a range: NOT(INDIRECT(...)) converts the result too
VALUE_ONLY_FUNCTIONS = {"NOT", "AND", "OR", "ABS", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "SQRT", "EXP", "LN", "LOG", "LOG10", "POWER", "MOD", "LEN", "LEFT", "RIGHT", "MID", "UPPER", "LOWER", "TRIM", "TEXT", "VALUE", "YEAR", "MONTH", "DAY", "ISNUMBER", "ISTEXT", "ISLOGICAL"}
INDIRECT_RE = re.compile(r"(?<![\w.])(?:_xlfn\.)?INDIRECT\s*\(", re.IGNORECASE)


class Unresolvable(Exception):
    pass


class _Expr:
    """A  small evaluator  for the text argument of INDIRECT, from the workbook's
    own constants: enough for "'"&$G710&"'!$Z$16", "LoB"&COLUMN()-6&"_Boolean",
    RIGHT(E2,IF(LEN(E2)=5,1,2)). Anything else raises Unresolvable."""

    TOKEN = re.compile(r"\s*(?:(\"(?:[^\"]|\"\")*\")|(\d+(?:\.\d+)?)|('[^']+'![$A-Za-z0-9:]+|[^\W\d][\w.]*![$A-Za-z0-9:]+|\$?[A-Z]{1,3}\$?\d{1,7})|([^\W\d][\w.]*)|(<>|<=|>=|[=<>&+\-*/()^,;]))")
    FUNCS = {"ROW", "COLUMN", "LEN", "LEFT", "RIGHT", "MID", "IF", "UPPER", "LOWER", "TRIM", "VALUE", "INT", "TRUE", "FALSE", "N", "T"}

    def __init__(self, text: str, lookup, sheet: str, row: int, col: int, names: dict[str, str]):
        self.toks = []
        pos = 0
        text = text.strip()
        while pos < len(text):
            m = self.TOKEN.match(text, pos)
            if not m or m.end() == pos:
                raise Unresolvable(f"cannot read {text[pos:pos + 12]!r}")
            pos = m.end()
            if m.group(1) is not None:
                self.toks.append(("str", m.group(1)[1:-1].replace('""', '"')))
            elif m.group(2) is not None:
                self.toks.append(("num", float(m.group(2))))
            elif m.group(3) is not None:
                self.toks.append(("ref", m.group(3)))
            elif m.group(4) is not None:
                self.toks.append(("name", m.group(4)))
            else:
                self.toks.append(("op", m.group(5)))
        self.i = 0
        self.lookup, self.sheet, self.row, self.col, self.names = lookup, sheet, row, col, names

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self, kind=None, value=None):
        tok = self.peek()
        if tok[0] is None or (kind and tok[0] != kind) or (value is not None and tok[1] != value):
            raise Unresolvable(f"unexpected {tok}")
        self.i += 1
        return tok

    def eval(self):
        v = self.cmp()
        if self.peek()[0] is not None:
            raise Unresolvable("trailing text")
        return v

    def cmp(self):
        left = self.concat()
        tok = self.peek()
        if tok == ("op", "=") or tok == ("op", "<>") or tok == ("op", "<") or tok == ("op", ">") or tok == ("op", "<=") or tok == ("op", ">="):
            self.take()
            right = self.concat()
            op = tok[1]
            a, b = _num_or_text(left), _num_or_text(right)
            return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]
        return left

    def concat(self):
        v = self.add()
        while self.peek() == ("op", "&"):
            self.take()
            v = _text(v) + _text(self.add())
        return v

    def add(self):
        v = self.mul()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            w = self.mul()
            v = _num(v) + _num(w) if op == "+" else _num(v) - _num(w)
        return v

    def mul(self):
        v = self.unary()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            w = self.unary()
            v = _num(v) * _num(w) if op == "*" else _num(v) / _num(w)
        return v

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return -_num(self.unary())
        return self.atom()

    def atom(self):
        kind, val = self.peek()
        if kind == "str" or kind == "num":
            self.take()
            return val
        if kind == "op" and val == "(":
            self.take()
            v = self.cmp()
            self.take("op", ")")
            return v
        if kind == "ref":
            self.take()
            return self.cell(val)
        if kind == "name":
            self.take()
            up = val.upper()
            if self.peek() == ("op", "("):
                self.take()
                args = []
                if self.peek() != ("op", ")"):
                    args.append(self.cmp())
                    while self.peek() in (("op", ","), ("op", ";")):
                        self.take()
                        args.append(self.cmp())
                self.take("op", ")")
                return self.call(up, args)
            if up == "TRUE":
                return True
            if up == "FALSE":
                return False
            ref = self.names.get(up)
            if ref is None:
                raise Unresolvable(f"name {val}")
            return self.cell(ref)
        raise Unresolvable(f"unexpected {kind} {val}")

    def call(self, fn, args):
        if fn == "ROW" and not args:
            return float(self.row)
        if fn == "COLUMN" and not args:
            return float(self.col)
        if fn == "LEN" and len(args) == 1:
            return float(len(_text(args[0])))
        if fn == "LEFT" and 1 <= len(args) <= 2:
            return _text(args[0])[: int(_num(args[1])) if len(args) == 2 else 1]
        if fn == "RIGHT" and 1 <= len(args) <= 2:
            n = int(_num(args[1])) if len(args) == 2 else 1
            return _text(args[0])[-n:] if n else ""
        if fn == "MID" and len(args) == 3:
            s = int(_num(args[1])) - 1
            return _text(args[0])[s : s + int(_num(args[2]))]
        if fn == "IF" and 2 <= len(args) <= 3:
            return args[1] if _truth(args[0]) else (args[2] if len(args) == 3 else False)
        if fn == "UPPER" and len(args) == 1:
            return _text(args[0]).upper()
        if fn == "LOWER" and len(args) == 1:
            return _text(args[0]).lower()
        if fn == "TRIM" and len(args) == 1:
            return " ".join(_text(args[0]).split())
        if fn in ("VALUE", "N") and len(args) == 1:
            return _num(args[0])
        if fn == "INT" and len(args) == 1:
            return float(int(_num(args[0])))
        if fn == "T" and len(args) == 1:
            return _text(args[0]) if isinstance(args[0], str) else ""
        raise Unresolvable(f"function {fn}")

    def cell(self, ref_text_: str):
        refs = cell_refs_in_formula("=" + ref_text_)
        if len(refs) != 1 or refs[0]["r1"] != refs[0]["r2"] or refs[0]["c1"] != refs[0]["c2"]:
            raise Unresolvable(f"reference {ref_text_}")
        r = refs[0]
        v = self.lookup(r.get("sheet") or self.sheet, r["r1"], r["c1"])
        if v is None or (isinstance(v, str) and v.startswith("=")):
            raise Unresolvable(f"{ref_text_} has no stored value")
        return v


class _Blank(str):
    """An  empty cell  read by a formula: "" as text, 0 as a number (Excel's own reading)."""


BLANK = _Blank("")


def _text(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    return str(v)


def _num(v) -> float:
    if isinstance(v, _Blank):
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError as exc:
        raise Unresolvable(f"not a number: {v!r}") from exc


def _num_or_text(v):
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        return str(v).lower()


def _truth(v) -> bool:
    if isinstance(v, str):
        return v.upper() == "TRUE"
    return bool(v)


def _indirect_calls(formula: str) -> list[dict[str, Any]]:
    """Every  INDIRECT call:  its span, its argument text (first argument only)
    and whether its result is used as a value (an operator right before the
    call or right after it)."""
    masked = mask_strings(formula)
    out = []
    for m in INDIRECT_RE.finditer(masked):
        i = m.end()
        depth = 1
        arg_end = None
        while i < len(masked) and depth:
            ch = masked[i]
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch in ",;" and depth == 1 and arg_end is None:
                arg_end = i
            i += 1
        if depth:
            continue
        end = i
        before = masked[1 : m.start()].rstrip()[-1:]  # the formula's leading '=' is not an  operator: =INDIRECT(...) alone is a reference
        after = masked[end:].lstrip()[:1]
        as_value = before in VALUE_OPERATORS or after in VALUE_OPERATORS
        if not as_value and before in "(,;":
            as_value = _enclosing_function(masked, m.start()) in VALUE_ONLY_FUNCTIONS or (before == "(" and _enclosing_function(masked, m.start()) == "IF")
        out.append({
            "start": m.start(), "end": end,
            "argument": formula[m.end() : arg_end if arg_end is not None else end - 1],
            "as_value": as_value,
        })
    return out


def _enclosing_function(masked: str, pos: int) -> str | None:
    """The  function whose  argument list the text at `pos` sits in, if any."""
    depth = 0
    i = pos - 1
    while i >= 0:
        ch = masked[i]
        if ch == ")":
            depth += 1
        elif ch == "(":
            if depth == 0:
                m = re.search(r"([^\W\d][\w.]*)\s*$", masked[:i])
                return m.group(1).upper() if m else None
            depth -= 1
        i -= 1
    return None


# INDIRECT("Cluster"&n) where the names Cluster1..ClusterN are N neighbouring single cells of one column
# (or one row), in order: the same cell as INDEX(<those  N cells>, n) for every whole n from 1 to N,  and
# #REF! in both for a whole  n above N. Model B, 2026-11-19: 2,320  such formulas, n built from ROW, COLUMN,
# ROUNDUP and COUNTA, so the  text cannot be worked out from constants.  Only a whole n that is never below 1
# is accepted: INDEX truncates 2.5 where INDIRECT finds no name "Cluster2.5", INDEX(..., 0) is the whole list
# and INDEX(..., -1) #VALUE! where INDIRECT gives #REF! (Haiku 5.5 refuter, 20 Nov 2026).
_NAMED_LIST_ARG_RE = re.compile(r'^\s*"([A-Za-z_\\][\w.]*)"\s*&\s*(.+?)\s*$', re.S)
_COUNTS = frozenset({"COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS"})
_ROUNDINGS = frozenset({"ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"})
_INF = float("inf")
_BIG = 10**11  # more  than the  cells of a sheet (1,048,576 x 16,384): no count, row or column number reaches it
_MAX_ROW, _MAX_COL = 1_048_576, 16_384


def _top_level_args(text: str) -> list[str]:
    """The  arguments of  a call's inside, split at the commas of its own level (strings already masked)."""
    out, depth, start = [], 0, 0
    for i, ch in enumerate(text):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch in ",;" and depth == 0:
            out.append(text[start:i])
            start = i + 1
    out.append(text[start:])
    return out


def _to_whole(name: str, x):
    """ROUND  / ROUNDUP  / ROUNDDOWN / TRUNC to 0 digits (away from or towards zero, as Excel), INT (down)."""
    import math
    from fractions import Fraction

    if x in (_INF, -_INF):
        return x
    x = Fraction(x)
    if name == "INT":
        return Fraction(math.floor(x))
    sign, a = (-1 if x < 0 else 1), abs(x)
    whole = math.ceil(a) if name == "ROUNDUP" else math.floor(a + Fraction(1, 2)) if name == "ROUND" else math.floor(a)
    return Fraction(sign * whole)


def _times(a, b):
    return 0 if a == 0 or b == 0 else a * b


def _bounds(expr: str):
    """(low,  high, whole)  of the number `expr` can give, or None when it is not plain arithmetic (text, &, ^,
    a comparison). A cell's value or an unknown function is any number (-inf, inf, not  whole); ROW and COLUMN
    of a written cell are that number; ROWS  / COLUMNS of a written range  its size; the COUNT functions 0 or
    more;  INT and ROUND / ROUNDUP / ROUNDDOWN / TRUNC to 0  digits are whole. A division by a  whole number
    skips 0 (that is #DIV/0! in INDIRECT's text and in INDEX's number alike)."""
    from fractions import Fraction

    s = mask_strings(expr)
    pos = 0
    anything = (-_INF, _INF, False)

    def skip() -> None:
        nonlocal pos
        while pos < len(s) and s[pos] in " \t\r\n":
            pos += 1

    def closing(start: int) -> int:
        depth = 0
        for i in range(start, len(s)):
            if s[i] == "(":
                depth += 1
            elif s[i] == ")":
                depth -= 1
                if depth == 0:
                    return i
        return -1

    def call(name: str, raw_args: list[str]):
        if name in ("ROW", "COLUMN"):
            big = _MAX_ROW if name == "ROW" else _MAX_COL
            r = parse_ref(raw_args[0]) if len(raw_args) == 1 and raw_args[0].strip() else None
            if r and not r.get("whole_column") and not r.get("whole_row"):
                v = Fraction(r["r1"] if name == "ROW" else r["c1"])
                return (v, v, True)
            return (Fraction(1), Fraction(big), True)
        if name in ("ROWS", "COLUMNS"):
            r = parse_ref(raw_args[0]) if len(raw_args) == 1 else None
            if r and not r.get("whole_column") and not r.get("whole_row"):
                v = Fraction(r["r2"] - r["r1"] + 1 if name == "ROWS" else r["c2"] - r["c1"] + 1)
                return (v, v, True)
            return (Fraction(1), Fraction(_BIG), True)
        if name in _COUNTS:
            return (Fraction(0), Fraction(_BIG), True)
        if name == "INT" and len(raw_args) == 1:
            inner = _bounds(raw_args[0])
            return None if inner is None else (_to_whole("INT", inner[0]), _to_whole("INT", inner[1]), True)
        if name in _ROUNDINGS and (len(raw_args) == 2 and raw_args[1].strip() == "0" or name == "TRUNC" and len(raw_args) == 1):
            inner = _bounds(raw_args[0])
            return None if inner is None else (_to_whole(name, inner[0]), _to_whole(name, inner[1]), True)
        return anything

    def factor():
        nonlocal pos
        skip()
        if pos >= len(s):
            return None
        if s[pos] in "+-":
            negative = s[pos] == "-"
            pos += 1
            v = factor()
            if v is None or not negative:
                return v
            return (-v[1], -v[0], v[2])
        m = re.match(r"\d+(?![\d.eE])", s[pos:])
        if m:
            pos += m.end()
            v = Fraction(int(m.group()))
            return (v, v, True)
        m = re.match(r"\d*\.\d+(?![\d.eE])", s[pos:])
        if m:
            pos += m.end()
            v = Fraction(m.group())
            return (v, v, v.denominator == 1)
        if s[pos] == "(":
            pos += 1
            v = total()
            skip()
            if v is None or pos >= len(s) or s[pos] != ")":
                return None
            pos += 1
            return v
        m = re.match(r"(?:_xlfn\.)?([A-Za-z][A-Za-z0-9.]*)\s*\(", s[pos:])
        if m:
            open_at = pos + m.end() - 1
            close_at = closing(open_at)
            if close_at < 0:
                return None
            masked_args = _top_level_args(s[open_at + 1 : close_at])
            raw_args, at = [], open_at + 1
            for a in masked_args:  # masking keeps every position: the same slices of the written text
                raw_args.append(expr[at : at + len(a)])
                at += len(a) + 1
            pos = close_at + 1
            return call(m.group(1).upper(), raw_args)
        m = re.match(r"(?:'[^']*'!|[\w.]+!)?[$\w.:]+", s[pos:])
        if m and not m.group().isdigit():
            pos += m.end()
            return anything  # a cell, a range or a name: any value
        return None

    def term():
        nonlocal pos
        a = factor()
        while a is not None:
            skip()
            if pos >= len(s) or s[pos] not in "*/":
                return a
            op = s[pos]
            pos += 1
            b = factor()
            if b is None:
                return None
            if op == "*":
                c = [_times(x, y) for x in (a[0], a[1]) for y in (b[0], b[1])]
                a = (min(c), max(c), a[2] and b[2])
                continue
            lo, hi = b[0], b[1]
            if b[2] and lo >= 0:
                lo = max(lo, 1)
            elif b[2] and hi <= 0:
                hi = min(hi, -1)
            if lo <= 0 <= hi or _INF in (abs(a[0]), abs(a[1]), abs(lo), abs(hi)):
                a = anything
                continue
            c = [x / y for x in (a[0], a[1]) for y in (lo, hi)]
            a = (min(c), max(c), False)
        return None

    def total():
        nonlocal pos
        a = term()
        while a is not None:
            skip()
            if pos >= len(s) or s[pos] not in "+-":
                return a
            op = s[pos]
            pos += 1
            b = term()
            if b is None:
                return None
            a = (a[0] + b[0], a[1] + b[1], a[2] and b[2]) if op == "+" else (a[0] - b[1], a[1] - b[0], a[2] and b[2])
        return None

    try:
        v = total()
        skip()
    except (RecursionError, ValueError, OverflowError, ZeroDivisionError):
        return None
    return v if v is not None and pos == len(s) else None


def _whole_bounds(expr: str):
    """(low,  high) of  a number `expr` can only give as a whole number (or an error), else None."""
    v = _bounds(expr)
    return (v[0], v[1]) if v is not None and v[2] else None


def _whole_number(expr: str) -> bool:
    """True  when `expr`  can only give a whole number (or an error)."""
    return _whole_bounds(expr) is not None


_POSITION_CALL_RE = re.compile(r"\b(ROW|COLUMN)\(\s*([^()]+?)\s*\)", re.IGNORECASE)


def _position_value(expr: str):
    """The  number `expr`  gives when it is arithmetic over whole constants and ROW/COLUMN of written
    references only (Model B's (COLUMN(V$1)-COLUMN($B$1)-3)/17), as a Fraction: it never depends on a
    cell's value, so it is the same number  whatever the inputs. None for anything else."""
    import ast
    from fractions import Fraction

    def position(m: re.Match) -> str:
        r = parse_ref(m.group(2))
        if not r or r.get("whole_column") or r.get("whole_row"):
            raise ValueError("not a written cell")
        return str(r["r1"] if m.group(1).upper() == "ROW" else r["c1"])

    try:
        text = _POSITION_CALL_RE.sub(position, expr)
    except ValueError:
        return None
    if not re.fullmatch(r"[\d\s+\-*/()]+", text):
        return None

    def value(node) -> Fraction:
        if isinstance(node, ast.Expression):
            return value(node.body)
        if isinstance(node, ast.Constant) and type(node.value) is int:
            return Fraction(node.value)
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            v = value(node.operand)
            return -v if isinstance(node.op, ast.USub) else v
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
            a, b = value(node.left), value(node.right)
            if isinstance(node.op, ast.Add):
                return a + b
            if isinstance(node.op, ast.Sub):
                return a - b
            if isinstance(node.op, ast.Mult):
                return a * b
            return a / b  # ZeroDivisionError: not whole
        raise ValueError("not arithmetic")

    try:
        return value(ast.parse(text.strip(), mode="eval"))
    except (SyntaxError, ValueError, ZeroDivisionError, RecursionError):
        return None


def _named_list_index(argument: str, wb0: dict[str, Any], sheets: dict[str, Any], unscanned: set[str]) -> dict[str, Any] | None:
    """INDIRECT's  argument "Prefix"&n,  with n a whole number and the names Prefix1..PrefixN N neighbouring
    single cells of one column or row, in order, each defined once for the whole workbook, and no other
    PrefixK name: {"formula":  "INDEX('Sheet'!$B$7:$B$26,n)", "range", "names", "sheet"}. Else None."""
    from ..formula_utils import num_to_col

    m = _NAMED_LIST_ARG_RE.match(argument)
    if not m:
        return None
    prefix, number = m.group(1), m.group(2)
    numbered: dict[int, list[dict]] = {}
    pattern = re.compile(re.escape(prefix) + r"(0|[1-9]\d*)", re.IGNORECASE)
    for d in wb0.get("defined_names", []):
        hit = pattern.fullmatch(str(d.get("name") or ""))
        if hit:
            numbered.setdefault(int(hit.group(1)), []).append(d)
    count = len(numbered)
    if count < 2 or sorted(numbered) != list(range(1, count + 1)):
        return None
    fixed = _position_value(number)
    if fixed is not None:
        if fixed.denominator != 1 or not 1 <= fixed <= count:
            return None  # INDIRECT finds no name for it  today, and INDEX(..., 0) or INDEX(...,  2.5) would not be the same error
    else:
        bounds = _whole_bounds(number)
        if bounds is None or bounds[0] < 1:
            return None  # a  fraction, or a number that may be 0 or less: INDEX  would not give INDIRECT's #REF! there
    if re.fullmatch(r"[A-Za-z]{1,3}", prefix):
        return None  # "CN"&21 is  the cell CN21, not a name: past the list INDIRECT would read it
    cells = []
    for k in range(1, count + 1):
        defs = numbered[k]
        if len(defs) != 1 or defs[0].get("local_sheet_id") is not None:
            return None  # a name defined for one sheet as well: which one INDIRECT finds depends on the sheet
        r = parse_ref(str(defs[0].get("value") or ""))
        if not r or not r.get("sheet") or r["r1"] != r["r2"] or r["c1"] != r["c2"] or r.get("whole_column") or r.get("whole_row"):
            return None
        cells.append(r)
    sheet = cells[0]["sheet"]
    if any(c["sheet"] != sheet for c in cells):
        return None
    down = all(c["c1"] == cells[0]["c1"] and c["r1"] == cells[0]["r1"] + i for i, c in enumerate(cells))
    across = all(c["r1"] == cells[0]["r1"] and c["c1"] == cells[0]["c1"] + i for i, c in enumerate(cells))
    if not (down or across):
        return None
    if sheet not in unscanned and not (sheet in sheets and sheets[sheet].get("grids")):
        return None  # Mind imports no sheet without a grid
    first, last = cells[0], cells[-1]
    rng = f"{_quote_sheet(sheet)}!${num_to_col(first['c1'])}${first['r1']}:${num_to_col(last['c1'])}${last['r1']}"
    return {"formula": f"INDEX({rng},{number})", "range": rng, "names": f"{prefix}1..{prefix}{count}", "sheet": sheet}


def indirect_as_value(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = 80) -> dict:
    """FRM-007:  INDIRECT, with  the direct reference proposed where the text it builds can be
    worked out from the workbook's constants. Used as a value Mind stops the run; used as a range
    (VLOOKUP(x, INDIRECT($A5&"!A:B"), 2, FALSE))  Mind does not resolve it either and the formula
    falls to its  IFERROR -- confirmed in Mind on  Model D (an IFRS 17 tool), 2026-11-08:  480 cells showed
    "-" where Excel shows the results; the same lookup  on BBA_04!$A:$B gave Excel's values.

    A sheet the  scan skipped at the user's request still exists: a target on it is written as the
    plain reference, exactly as without the skip (its cells and grids were never read, so nothing
    about them is assumed, and the site lists it under `unscanned_sheets`). A text that needs a cell
    of a skipped sheet cannot be worked out and is left for a person --  never NA(). Only a sheet
    that truly is not in the workbook makes NA()."""
    from ..inventory import cell_value, ignored_sheet_names, values_cell

    wb0 = analysis["workbooks"][0]
    sheets = {s["name"]: s for s in wb0.get("sheets", [])}
    unscanned = set(ignored_sheet_names(analysis))
    names = {}
    for d in wb0.get("defined_names", []):
        if d.get("name") and d.get("value") and "#REF!" not in str(d["value"]):
            names[str(d["name"]).upper()] = str(d["value"])
    canonical = {str(d["name"]).upper(): str(d["name"]) for d in wb0.get("defined_names", []) if d.get("name")}

    def lookup(sheet: str, row: int, col: int):
        if sheet in unscanned:
            raise Unresolvable(f"{sheet} was left out of the scan at your request, so its cells were not read")
        v = cell_value(analysis, sheet, row, col)
        if v is None and sheet in sheets:
            return BLANK  # an empty cell: "" (a sheet chooser left blank, Model D (an IFRS 17 tool)'s ロールフォワード!T19)
        if isinstance(v, str) and v.startswith("="):
            v = values_cell(analysis, sheet, row, col)  # a formula cell: the value Excel last stored
        return v

    sites = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        if "INDIRECT" not in formula.upper():
            continue
        calls = _indirect_calls(formula)
        if not calls:
            continue
        origin = parse_ref(f["cell"])
        resolved = []
        suggested = formula
        missing = False
        unresolved = None
        unscanned_targets: list[str] = []
        for call in reversed(calls):
            try:
                text = _text(_Expr(call["argument"], lookup, f["sheet"], origin["r1"], origin["c1"], names).eval())
            except (Unresolvable, ZeroDivisionError, ValueError) as exc:
                listed = _named_list_index(call["argument"], wb0, sheets, unscanned)
                if listed:
                    if listed["sheet"] in unscanned and listed["sheet"] not in unscanned_targets:
                        unscanned_targets.append(listed["sheet"])
                    resolved.append({"argument": call["argument"][:120], "resolved": None, "target": listed["range"], "exists": True,
                                     "as_value": call["as_value"], "by_position": listed["names"]})
                    suggested = suggested[: call["start"]] + listed["formula"] + suggested[call["end"] :]
                    continue
                unresolved = f"{call['argument'][:60]}: {exc}"
                resolved.append({"argument": call["argument"][:120], "resolved": None, "as_value": call["as_value"]})
                continue
            target = None
            exists = False
            refs = cell_refs_in_formula("=" + text)
            if len(refs) == 1 and REF_RE.fullmatch(text.strip()):
                r = refs[0]
                sheet = r.get("sheet") or f["sheet"]
                if sheet in unscanned:
                    # the sheet exists; whether it has a grid was not looked at and is not guessed
                    exists = True
                    if sheet not in unscanned_targets:
                        unscanned_targets.append(sheet)
                else:
                    exists = sheet in sheets and bool(sheets[sheet].get("grids"))
                target = (f"{_quote_sheet(sheet)}!" if r.get("sheet") else "") + text.strip().rsplit("!", 1)[-1]
            elif text.upper() in canonical:
                target = canonical[text.upper()]
                exists = True
            resolved.append({"argument": call["argument"][:120], "resolved": text[:120], "target": target, "exists": exists, "as_value": call["as_value"]})
            if target and exists:
                suggested = suggested[: call["start"]] + target + suggested[call["end"] :]
            else:
                missing = True
                # where a reference must stand (OFFSET(INDIRECT(...), ...)) Excel refuses NA(): INDEX($A$1,NA()) is the
                # same #N/A as a reference (Model D (an IFRS 17 tool), 2026-11-08)
                suggested = suggested[: call["start"]] + ("NA()" if call["as_value"] else "INDEX($A$1,NA())") + suggested[call["end"] :]
        resolved.reverse()
        if unresolved:
            proposal, issue = None, f"the text INDIRECT builds cannot be worked out from the workbook's constants ({unresolved})"
        elif missing:
            gone = sorted({c["resolved"] for c in resolved if not c.get("exists")})
            proposal, issue = suggested, f"target(s) not in the workbook (or without a grid, so Mind does not import them): {', '.join(gone[:4])} -- that call is an error in Excel already and becomes NA() in its place"
        else:
            proposal, issue = suggested, None
        sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:300], "calls": resolved, "suggested_formula": proposal, "issue": issue, "unscanned_sheets": sorted(unscanned_targets)})
    cap = slice(None) if limit is None else slice(0, limit)
    observed = {
        "sites": sites[cap], "sites_count": len(sites),
        "with_proposal": sum(1 for s in sites if s["suggested_formula"] and not s["issue"]),
        "with_missing_targets": sum(1 for s in sites if s["suggested_formula"] and s["issue"]),
        "without_proposal": sum(1 for s in sites if s["suggested_formula"] is None),
        "reading_skipped_sheets": sum(1 for s in sites if s["unscanned_sheets"]),
        "by_position": sum(1 for s in sites if s["suggested_formula"] and any(c.get("by_position") for c in s["calls"])),
    }
    if sites:
        first = next((s for s in sites if s["suggested_formula"] and not s["issue"]), sites[0])
        return finding(
            "ERROR",
            f"{len(sites)} formula(s) use INDIRECT. Mind does not follow it: used as a value (compared, multiplied, concatenated) Mind stops the run "
            "('Run error: Unable to cast object of type AM.Models.AMReference to type System.IConvertible'); used as a range (VLOOKUP(x, INDIRECT($A5&\"!A:B\"), 2, FALSE)) "
            "Mind computes the formula without it and an IFERROR around it shows its fallback (confirmed in Mind, 2026-11-08: \"-\" in place of the results). "
            f"First: {first['sheet']}!{first['cell']} = {first['formula'][:90]}"
            + (f" -> proposed {first['suggested_formula'][:120]}" if first["suggested_formula"] else "")
            + f". {observed['with_proposal']} can be rewritten as the direct reference the text designates (or, for a numbered list of names such as Cluster1..Cluster20, as INDEX over the cells the names stand for); {observed['with_missing_targets']} also point at a sheet or name that does not exist "
            "(an error in Excel already: that call becomes NA() in its place); "
            f"{observed['without_proposal']} build their text from values the tool cannot work out (by hand). The Prep action 'Replace INDIRECT by the reference its text builds' writes the proposals.",
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    return finding("PASS", "No formula uses INDIRECT.", observed)



# --- REF-002 -----------------------------------------------------------------------------
def external_links_in_formulas(rule, analysis: dict[str, Any], config: dict[str, Any], limit: int | None = 80) -> dict:
    """REF-002:  formulas that  read another workbook. Mind cannot follow an external link; Excel
    shows the copy it kept of each external cell (the link's cache), or #REF! where it kept none."""
    from ..external_links import external_references, freeze_formula, read_external_links

    wb0 = analysis["workbooks"][0]
    try:
        links = read_external_links(wb0["copy_path"])
    except Exception as exc:
        return finding("NOT_SUPPORTED", f"The workbook's external links could not be read ({type(exc).__name__}: {str(exc)[:120]}).", {"sites": [], "links": []})
    sites = []
    for f in wb0.get("formulas", []):
        formula = f.get("formula") or ""
        if "[" not in formula or not external_references(formula):
            continue
        fixed, notes, issue = freeze_formula(formula, links)
        sites.append({"sheet": f["sheet"], "cell": f["cell"], "formula": formula[:200], "suggested_formula": fixed, "values": notes[:4], "issue": issue})
    link_info = [{"index": x["index"], "file_name": x["file_name"], "target": x["target"], "reachable": x["reachable"], "sheets": x["sheets"][:8], "cached_cells": x["cached_cells"]} for x in links]
    observed = {"sites": sites if limit is None else sites[:limit], "sites_count": len(sites), "links": link_info, "with_proposal": sum(1 for x in sites if x["suggested_formula"]), "without_proposal": sum(1 for x in sites if not x["suggested_formula"])}
    if not sites and not links:
        return finding("PASS", "No formula reads another workbook, and the package holds no external link.", observed)
    if not sites:
        return finding("WARNING", f"{len(links)} external link part(s) in the package but no formula reads them: the Prep action 'Freeze external links' removes them.", observed)
    first = sites[0]
    files = ", ".join(f"{x['file_name'] or '?'} ({'reachable' if x['reachable'] else 'not reachable from here'}, {x['cached_cells']} cell(s) kept)" for x in link_info[:3])
    return finding(
        "ERROR",
        f"{len(sites)} formula(s) read another workbook ({files}). Mind cannot follow an external link: each of these is an error there, and the converter reports the link "
        f"(seen in a real Mind conversion: 20,934 such formulas, '20,000 deterministic errors'). Excel shows the copy it kept of each external cell, or #REF! where it kept none. "
        f"First: {first['sheet']}!{first['cell']} = {first['formula'][:90]}"
        + (f" -> proposed {first['suggested_formula'][:120]}" if first["suggested_formula"] else "")
        + f". The Prep action 'Freeze external links' writes that copy into {observed['with_proposal']} formula(s) (NA() where Excel kept none) and removes the link; Excel's values stay the same"
        + (f"; {observed['without_proposal']} read a whole range of the other workbook and are left for a person." if observed["without_proposal"] else "."),
        observed,
        location={"sheet": first["sheet"], "cell": first["cell"]},
    )
