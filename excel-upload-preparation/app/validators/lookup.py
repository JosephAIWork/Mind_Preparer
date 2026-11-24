"""LKP-*  validators: MM_READTABLE  / MM_READTABLENAN (KB 'MM_READTABLE
function': Range must include headers; Header names exactly one column; up
to 6 comparisons test the first columns in order; first matching row wins;
not-found returns a text message vs  NaN for the NAN variant)."""
from __future__ import annotations

from typing import Any

from ..formula_utils import find_calls, parse_ref, unquote
from ..grids import all_grids, grid_containing
from ..inventory import cell_value, mm_calls
from ._common import finding, fmt_sites, is_blank

READTABLE_FUNCTIONS = ("MM_READTABLE", "MM_READTABLENAN")
MAX_COMPARISONS = 6


def _calls(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for fn in READTABLE_FUNCTIONS:
        for c in mm_calls(analysis, fn):
            out.append({**c, "function": fn})
    return out


def _resolve_range(analysis: dict[str, Any], call: dict[str, Any]) -> dict[str, Any] | None:
    """The  table range  of a call as {sheet, r1, c1, r2, c2, via}: a literal
    range, or MM_TABLE(cell) resolved through the detected grid."""
    if not call["args"]:
        return None
    arg = call["args"][0]
    ref = parse_ref(arg)
    if ref and not ref["whole_column"] and not ref["whole_row"]:
        return {"sheet": ref["sheet"] or call["sheet"], "r1": ref["r1"], "c1": ref["c1"], "r2": ref["r2"], "c2": ref["c2"], "via": "literal"}
    tables = find_calls(arg if arg.startswith("=") else "=" + arg, "MM_TABLE")
    if tables and tables[0]:
        cell_ref = parse_ref(tables[0][0])
        if cell_ref:
            sheet = cell_ref["sheet"] or call["sheet"]
            g = grid_containing([g for g in all_grids(analysis) if g["sheet"] == sheet], cell_ref["r1"], cell_ref["c1"])
            if g:
                return {"sheet": sheet, "r1": cell_ref["r1"], "c1": cell_ref["c1"], "r2": g["last_row"], "c2": g["last_col"], "via": "MM_TABLE", "grid_first_row": g["first_row"]}
    return None


def _header_row(analysis: dict[str, Any], rng: dict[str, Any]) -> list[Any]:
    return [cell_value(analysis, rng["sheet"], rng["r1"], c) for c in range(rng["c1"], rng["c2"] + 1)]


def lkp_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    calls = _calls(analysis)
    counts = {fn: sum(1 for c in calls if c["function"] == fn) for fn in READTABLE_FUNCTIONS}
    arrays = sum(1 for c in calls if c.get("array_ref"))
    return finding("PASS", f"Inventoried {len(calls)} MM_READTABLE/MM_READTABLENAN call(s) {counts}, {arrays} as array formulas.", {"counts": counts, "array_formula_calls": arrays, "sites": [{"sheet": c["sheet"], "cell": c["cell"]} for c in calls[:50]]})


def lkp_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """LKP-002:  the Range  must include the header row."""
    calls = _calls(analysis)
    if not calls:
        return finding("PASS", "No MM_READTABLE calls.", {"checked": 0})
    bad, unresolvable, too_many = [], [], []
    checked = 0
    for c in calls:
        if len(c["args"]) - 2 > MAX_COMPARISONS:
            too_many.append({"sheet": c["sheet"], "cell": c["cell"], "comparisons": len(c["args"]) - 2})
        rng = _resolve_range(analysis, c)
        if rng is None:
            unresolvable.append({"sheet": c["sheet"], "cell": c["cell"], "range_arg": c["args"][0] if c["args"] else None})
            continue
        checked += 1
        header = _header_row(analysis, rng)
        header_ok = header and all(isinstance(h, str) and not h.startswith("=") and not is_blank(h) for h in header)
        if rng["via"] == "MM_TABLE" and rng["r1"] != rng.get("grid_first_row"):
            header_ok = False
        if not header_ok:
            bad.append({"sheet": c["sheet"], "cell": c["cell"], "range": f"{rng['sheet']}!{rng['r1']}:{rng['r2']}", "first_row_values": [str(h)[:20] for h in header[:6]]})
    observed = {"checked": checked, "missing_headers": bad[:50], "unresolvable_ranges": unresolvable[:50], "too_many_comparisons": too_many}
    if bad or too_many:
        first = (bad or too_many)[0]
        return finding(
            "ERROR",
            (f"{len(bad)} MM_READTABLE range(s) do not start on a text-only header row: {fmt_sites(bad)}. " if bad else "")
            + (f"{len(too_many)} call(s) pass more than {MAX_COMPARISONS} comparisons: {fmt_sites(too_many)}" if too_many else ""),
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if unresolvable:
        return finding("REQUIRES_USER_INPUT", f"{checked} range(s) include headers; {len(unresolvable)} range argument(s) are expressions that could not be resolved: {fmt_sites(unresolvable)}", observed)
    return finding("PASS", f"All {checked} MM_READTABLE range(s) start on a header row.", observed)


def lkp_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """LKP-003:  the Header  argument must match exactly one header of the range."""
    calls = _calls(analysis)
    if not calls:
        return finding("PASS", "No MM_READTABLE calls.", {"checked": 0})
    missing, duplicated, unresolvable = [], [], []
    checked = 0
    for c in calls:
        if len(c["args"]) < 2:
            missing.append({"sheet": c["sheet"], "cell": c["cell"], "header": None, "reason": "no Header argument"})
            continue
        rng = _resolve_range(analysis, c)
        header_name = unquote(c["args"][1])
        if header_name is None:
            ref = parse_ref(c["args"][1])
            if ref and ref["cells"] == 1:
                v = cell_value(analysis, ref["sheet"] or c["sheet"], ref["r1"], ref["c1"])
                header_name = v if isinstance(v, str) and not v.startswith("=") else None
        if rng is None or header_name is None:
            unresolvable.append({"sheet": c["sheet"], "cell": c["cell"], "header_arg": c["args"][1]})
            continue
        checked += 1
        headers = [str(h).strip() if h is not None else "" for h in _header_row(analysis, rng)]
        n = headers.count(header_name.strip())
        if n == 0:
            missing.append({"sheet": c["sheet"], "cell": c["cell"], "header": header_name, "available": headers[:12]})
        elif n > 1:
            duplicated.append({"sheet": c["sheet"], "cell": c["cell"], "header": header_name, "occurrences": n})
    observed = {"checked": checked, "header_not_found": missing[:50], "header_duplicated": duplicated[:50], "unresolvable": unresolvable[:50]}
    if missing or duplicated:
        first = (missing or duplicated)[0]
        return finding(
            "ERROR",
            (f"{len(missing)} MM_READTABLE Header argument(s) match no header of the range: " + "; ".join(f"{m['sheet']}!{m['cell']} '{m.get('header')}'" for m in missing[:5]) + ". " if missing else "")
            + (f"{len(duplicated)} Header argument(s) match several columns: {fmt_sites(duplicated)}" if duplicated else ""),
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if unresolvable:
        return finding("REQUIRES_USER_INPUT", f"{checked} header(s) verified unique; {len(unresolvable)} could not be resolved (expression range or header): {fmt_sites(unresolvable)}", observed)
    return finding("PASS", f"All {checked} MM_READTABLE Header argument(s) match exactly one column.", observed)


def lkp_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """LKP-004:  duplicate key  combinations in the tested columns -> first match wins."""
    calls = _calls(analysis)
    if not calls:
        return finding("PASS", "No MM_READTABLE calls.", {"checked": 0})
    dup_tables = []
    checked = 0
    seen_ranges = set()
    for c in calls:
        rng = _resolve_range(analysis, c)
        n_keys = max(len(c["args"]) - 2, 0)
        if rng is None or n_keys == 0:
            continue
        key = (rng["sheet"], rng["r1"], rng["c1"], rng["r2"], rng["c2"], n_keys)
        if key in seen_ranges:
            continue
        seen_ranges.add(key)
        checked += 1
        combos: dict[tuple, int] = {}
        has_formula_keys = False
        for r in range(rng["r1"] + 1, rng["r2"] + 1):
            vals = []
            for col in range(rng["c1"], min(rng["c1"] + n_keys, rng["c2"] + 1)):
                v = cell_value(analysis, rng["sheet"], r, col)
                if isinstance(v, str) and v.startswith("="):
                    has_formula_keys = True
                vals.append(str(v).strip().lower() if v is not None else "")
            combo = tuple(vals)
            combos[combo] = combos.get(combo, 0) + 1
        dups = {k: v for k, v in combos.items() if v > 1}
        if dups and not has_formula_keys:
            dup_tables.append({"sheet": c["sheet"], "cell": c["cell"], "table": f"{rng['sheet']}!R{rng['r1']}C{rng['c1']}:R{rng['r2']}C{rng['c2']}", "keys_tested": n_keys, "duplicate_combinations": len(dups), "example": list(next(iter(dups)))})
    if dup_tables:
        return finding("WARNING", f"{len(dup_tables)} MM_READTABLE table(s) have duplicate key combinations in the tested column(s) -- only the first matching row is returned: {fmt_sites(dup_tables)}", {"checked": checked, "duplicates": dup_tables[:50]}, location={"sheet": dup_tables[0]["sheet"], "cell": dup_tables[0]["cell"]})
    return finding("PASS", f"{checked} distinct MM_READTABLE table(s) checked; no duplicate key combinations among constant keys.", {"checked": checked, "duplicates": []})


def lkp_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """LKP-005:  not-found behaviour  -- text message (MM_READTABLE) vs NaN (MM_READTABLENAN)."""
    calls = _calls(analysis)
    plain = [c for c in calls if c["function"] == "MM_READTABLE"]
    nan = [c for c in calls if c["function"] == "MM_READTABLENAN"]
    if plain:
        return finding(
            "WARNING",
            f"{len(plain)} MM_READTABLE call(s) return the text 'Combination of criteria ... not found' when no row matches, which breaks downstream arithmetic; use MM_READTABLENAN where a NaN is preferable ({len(nan)} already do): {fmt_sites(plain)}",
            {"mm_readtable": len(plain), "mm_readtablenan": len(nan)},
            location={"sheet": plain[0]["sheet"], "cell": plain[0]["cell"]},
        )
    return finding("PASS", f"{len(nan)} MM_READTABLENAN call(s), no plain MM_READTABLE (not-found yields NaN)." if nan else "No MM_READTABLE calls.", {"mm_readtable": 0, "mm_readtablenan": len(nan)})
