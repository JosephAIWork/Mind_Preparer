"""LOOP-* / LBL-* / RES-* / RZS-* validators: MM_LOOP, MM_LOOPLABELS,
MM_LOOPINSTANCE, MM_RESULT and the resize family (MM_SETSIZE, /Resize
flags, MM_LASTROWCELL / MM_LASTCOLUMNCELL / MM_GETRANGE), via formula
tokenization plus the detected grids.

KB sources: 'MM_LOOP function', 'MM_LOOPLABELS function', 'MM_LOOPINSTANCE
function', 'MM_RESULT function', 'MM_SETSIZE function', 'Dynamic resizing
of grids', 'MM_LASTROWCELL function', 'MM_GETRANGE function', and the FAQ
'How to solve a run error' (loop referenced but never created; SetSize of a
SetSize)."""
from __future__ import annotations

import re
from typing import Any

from ..formula_utils import as_number, cell_refs_in_formula, find_calls_detailed, mask_strings, parse_ref, ref_text, unquote
from ..grids import all_grids, grid_containing, grid_flags
from ..inventory import cell_value, known_loop_names, loop_definitions, loopinstance_definitions, looplabels_definitions, mm_calls
from ._common import finding, fmt_grids, fmt_sites, grid_location, is_blank

RESIZE_ROW_FLAGS = ("resize", "resizerow")
RESIZE_COL_FLAGS = ("resizecolumn",)
DYNAMIC_FUNCTIONS = {"MM_FILTER", "MM_REMOVEDUPLICATES", "MM_HUNION", "MM_VUNION", "MM_SORT", "MM_SORTRANGE", "MM_SORTRANGES", "MM_CONCATENATE", "MM_GROUPBY", "MM_MIX",
                     "MM_LEFTOUTERJOIN", "MM_RIGHTOUTERJOIN", "MM_LEFTINNERJOIN", "MM_RIGHTINNERJOIN", "MM_FLATTENDIM", "MM_FILTERDIM"}
TEXT_SETSIZE_RE = re.compile(r"&\s*IF\s*\(\s*MM_SETSIZE\s*\([^()]*\)\s*=\s*0\s*,\s*\"\"\s*,\s*0\s*\)", re.IGNORECASE)


# --- LOOP-001..003, LBL-001 (unchanged behaviour, shared helpers) -------------
def mm_loop_inventory(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    defs = loop_definitions(analysis)
    inst = loopinstance_definitions(analysis)
    total_sites = sum(len(sites) for sites in defs.values())
    return finding(
        "PASS",
        f"Inventoried {len(defs)} MM_LOOP name(s) across {total_sites} definition site(s) and {len(inst)} MM_LOOPINSTANCE name(s).",
        {"loop_names": sorted(defs), "definition_site_count": total_sites, "loopinstance_names": sorted(inst)},
    )


def loop_naming_consistency(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """Loop names are case-sensitive in Mind; flag names that only differ by case."""
    defs = loop_definitions(analysis)
    by_lower: dict[str, set[str]] = {}
    for name in defs:
        by_lower.setdefault(name.lower(), set()).add(name)
    collisions = {k: sorted(v) for k, v in by_lower.items() if len(v) > 1}
    return finding(
        "ERROR" if collisions else "PASS",
        f"Loop name(s) used with inconsistent capitalization (names are case-sensitive): {collisions}" if collisions else "No loop-name capitalization inconsistencies found.",
        {"case_variants": collisions},
    )


def repeated_definition_range_lengths(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    defs = loop_definitions(analysis)
    mismatched = {}
    unresolvable = []
    dynamic_names = []
    for name, sites in defs.items():
        if len(sites) < 2:
            continue
        dyn = [s for s in sites if s.get("dynamic")]
        if dyn and len(dyn) == len(sites):
            # every definition reads its range from the table (MM_COLUMN / MM_ROW / MM_TABLE): Mind sizes them
            # from the same grid, so they stay equal, and they grow together when rows are added
            dynamic_names.append(name)
            continue
        if dyn:
            mismatched[name] = [{"cell": f"{s['sheet']}!{s['cell']}", "range_length": s["range_length"], "dynamic": bool(s.get("dynamic"))} for s in sites]
            continue
        lengths = {s["range_length"] for s in sites}
        if None in lengths:
            unresolvable.append(name)
        elif len(lengths) > 1:
            mismatched[name] = [{"cell": f"{s['sheet']}!{s['cell']}", "range_length": s["range_length"]} for s in sites]
    status = "ERROR" if mismatched else ("REQUIRES_USER_INPUT" if unresolvable else "PASS")
    mixed = [n for n, sites in mismatched.items() if any(x.get("dynamic") for x in sites)]
    return finding(
        status,
        (f"Repeated MM_LOOP definitions that mix table-following ranges (MM_COLUMN / MM_ROW) with fixed ranges: {sorted(mixed)}; make them all follow the table or all fixed." if mixed else f"Repeated MM_LOOP definitions with unequal range lengths: {sorted(mismatched)}.")
        if mismatched
        else (
            f"Could not resolve range length for repeated definitions of: {unresolvable} (likely a defined name or expression rather than a literal range)."
            if unresolvable
            else ("All repeated MM_LOOP definitions use equal range lengths." + (f" {sorted(dynamic_names)} follow their table (MM_COLUMN / MM_ROW): Mind sizes every definition from the same grid." if dynamic_names else ""))
        ),
        {"mismatched": mismatched, "unresolvable_range_refs": unresolvable, "table_following": sorted(dynamic_names)},
    )


def label_size_matches(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    loop_defs = loop_definitions(analysis)
    label_defs = looplabels_definitions(analysis)
    mismatched = []
    no_matching_loop = []
    unresolvable = []
    for label in label_defs:
        matching = loop_defs.get(label["name"])
        if not matching:
            no_matching_loop.append(label["name"])
            continue
        loop_len = matching[0]["range_length"]
        if label["range_length"] is None or loop_len is None:
            unresolvable.append(label["name"])
        elif label["range_length"] != loop_len:
            mismatched.append({"name": label["name"], "label_length": label["range_length"], "loop_length": loop_len})
    status = "ERROR" if mismatched else ("REQUIRES_USER_INPUT" if (no_matching_loop or unresolvable) else "PASS")
    return finding(
        status,
        f"MM_LOOPLABELS size mismatch vs. referenced loop: {mismatched}." if mismatched else "All MM_LOOPLABELS sizes match their referenced loop.",
        {"mismatched": mismatched, "no_matching_loop": no_matching_loop, "unresolvable": unresolvable},
    )


def loop_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """LOOP-004: MM_LOOPINSTANCE(Name, MaxSizeRange, InstanceSizeRange, [Index], [IndexStart])."""
    calls = mm_calls(analysis, "MM_LOOPINSTANCE")
    if not calls:
        return finding("PASS", "No MM_LOOPINSTANCE usage (loops do not vary by instance).", {"sites": []})
    bad = [c for c in calls if len(c["args"]) < 3 or len(c["args"]) > 5]
    names = sorted({unquote(c["args"][0]) or c["args"][0] for c in calls if c["args"]})
    loop_names_lower = {n.lower(): n for n in loop_definitions(analysis)}
    clashes = [n for n in names if n.lower() in loop_names_lower and loop_names_lower[n.lower()] != n]
    if bad:
        return finding("ERROR", f"{len(bad)} MM_LOOPINSTANCE call(s) with the wrong argument count (3 to 5 expected): {fmt_sites(bad)}", {"bad_sites": bad[:20], "names": names}, location={"sheet": bad[0]["sheet"], "cell": bad[0]["cell"]})
    if clashes:
        return finding("ERROR", f"MM_LOOPINSTANCE name(s) differ only by case from an MM_LOOP name (case-sensitive): {clashes}", {"names": names, "case_clashes": clashes})
    return finding("WARNING", f"{len(calls)} MM_LOOPINSTANCE call(s) defining {names}: review that the InstanceSizeRange varies per instance and MaxSizeRange is the true maximum (elements beyond it are dropped).", {"names": names, "sites": [{"sheet": c["sheet"], "cell": c["cell"]} for c in calls[:20]]})


# --- RES-* MM_RESULT ------------------------------------------------------------
def _result_calls(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for c in mm_calls(analysis, "MM_RESULT"):
        args = c["args"]
        specs = []
        for i in range(1, len(args) - 1, 2):
            specs.append({"name_arg": args[i], "name": unquote(args[i]), "index_arg": args[i + 1]})
        dangling = (len(args) - 1) % 2 == 1
        out.append({**c, "cell_arg": args[0] if args else None, "loop_specs": specs, "dangling_name": dangling})
    return out


def res_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    calls = _result_calls(analysis)
    names = sorted({s["name"] for c in calls for s in c["loop_specs"] if s["name"]})
    return finding("PASS", f"Inventoried {len(calls)} MM_RESULT call(s) referencing loop names {names}.", {"call_count": len(calls), "loop_names_referenced": names, "sites": [{"sheet": c["sheet"], "cell": c["cell"]} for c in calls[:50]]})


def res_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RES-002: every literal loop name in MM_RESULT / MM_DIMSIZE / MM_DIMINDEX /
    MM_LOOPLABELS must match a created loop exactly (case-sensitive)."""
    known = known_loop_names(analysis)
    known_lower = {n.lower(): n for n in known}
    missing, case_only, unresolvable = [], [], []
    for fn, name_positions in (("MM_RESULT", None), ("MM_DIMSIZE", [0]), ("MM_DIMINDEX", [0]), ("MM_LOOPLABELS", [0])):
        for c in mm_calls(analysis, fn):
            args = c["args"]
            positions = name_positions if name_positions is not None else list(range(1, len(args) - 1, 2))
            for pos in positions:
                if pos >= len(args):
                    continue
                name = unquote(args[pos])
                site = {"sheet": c["sheet"], "cell": c["cell"], "function": fn, "name": name or args[pos]}
                if name is None:
                    unresolvable.append(site)
                elif name in known:
                    continue
                elif name.lower() in known_lower:
                    case_only.append({**site, "defined_as": known_lower[name.lower()]})
                else:
                    missing.append(site)
    observed = {"known_loop_names": sorted(known), "missing": missing[:50], "case_mismatch": case_only[:50], "unresolvable_name_args": unresolvable[:50]}
    if missing or case_only:
        first = (missing or case_only)[0]
        return finding(
            "ERROR",
            f"{len(missing)} reference(s) to a loop that is never created and {len(case_only)} with different capitalization -- Mind's 'Limit of 100 calculation iteration' error. Missing: {fmt_sites(missing)}; case: {fmt_sites(case_only)}",
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    if unresolvable:
        return finding("WARNING", f"All literal loop names match; {len(unresolvable)} loop-name argument(s) are expressions that could not be checked: {fmt_sites(unresolvable)}", observed, evidence="INFERENCE")
    return finding("PASS", "Every loop name referenced by MM_RESULT/MM_DIMSIZE/MM_DIMINDEX/MM_LOOPLABELS matches a created loop exactly.", observed)


def res_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RES-003: MM_RESULT sums every loop it does not name."""
    calls = _result_calls(analysis)
    all_loops = set(loop_definitions(analysis)) | set(loopinstance_definitions(analysis))
    no_spec = [c for c in calls if not c["loop_specs"]]
    partial = []
    for c in calls:
        named = {s["name"] for s in c["loop_specs"] if s["name"]}
        omitted = sorted(all_loops - named)
        if c["loop_specs"] and omitted and len(all_loops) > 1:
            partial.append({"sheet": c["sheet"], "cell": c["cell"], "named": sorted(named), "not_named": omitted})
    if not calls:
        return finding("PASS", "No MM_RESULT calls.", {"summed_all_dimensions": [], "partially_specified": []})
    observed = {"summed_all_dimensions": [{"sheet": c["sheet"], "cell": c["cell"]} for c in no_spec[:50]], "partially_specified": partial[:50], "loops_in_workbook": sorted(all_loops)}
    if no_spec or partial:
        first = (no_spec or partial)[0]
        return finding(
            "WARNING",
            f"{len(no_spec)} MM_RESULT call(s) name no loop (sum over every dimension) and {len(partial)} name only some of the {len(all_loops)} loop(s) in the workbook -- any loop the target cell depends on but is not named is summed. Confirm intended: {fmt_sites(no_spec + partial)}",
            observed,
            location={"sheet": first["sheet"], "cell": first["cell"]},
            evidence="INFERENCE",
        )
    return finding("PASS", "Every MM_RESULT call names every loop defined in the workbook.", observed)


def res_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RES-004: the same loop named twice in one MM_RESULT call."""
    dups = []
    malformed = []
    for c in _result_calls(analysis):
        names = [s["name"] for s in c["loop_specs"] if s["name"]]
        repeated = sorted({n for n in names if names.count(n) > 1})
        if repeated:
            dups.append({"sheet": c["sheet"], "cell": c["cell"], "repeated": repeated})
        if c["dangling_name"]:
            malformed.append({"sheet": c["sheet"], "cell": c["cell"]})
    if dups or malformed:
        first = (dups or malformed)[0]
        return finding(
            "ERROR",
            (f"{len(dups)} MM_RESULT call(s) name the same loop more than once (use two MM_RESULT calls to sum two dimensions): {fmt_sites(dups)}. " if dups else "")
            + (f"{len(malformed)} call(s) have a loop name without an index: {fmt_sites(malformed)}" if malformed else ""),
            {"duplicates": dups, "name_without_index": malformed},
            location={"sheet": first["sheet"], "cell": first["cell"]},
        )
    return finding("PASS", "No MM_RESULT call names a loop twice.", {"duplicates": []})


def res_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RES-005: 'SIM' dimension usage in MM_RESULT / MM_DIMSIZE."""
    sim_sites = []
    for fn in ("MM_RESULT", "MM_DIMSIZE", "MM_DIMINDEX"):
        for c in mm_calls(analysis, fn):
            if any((unquote(a) or "").upper() == "SIM" for a in c["args"]):
                sim_sites.append({"sheet": c["sheet"], "cell": c["cell"], "function": fn})
    stochastic = sorted(fn for fn in analysis["features"].get("function_usage", {}) if fn in ("RAND", "RANDBETWEEN", "RANDARRAY") or fn.startswith(("MM_SIMULATE", "MM_RANDOM", "MM_SIMCOPULA")))
    if sim_sites:
        return finding("WARNING", f"{len(sim_sites)} reference(s) to the SIM dimension (a specific simulation index / the simulation count): {fmt_sites(sim_sites)}. Confirm this is intended rather than a risk measure.", {"sim_sites": sim_sites[:50], "stochastic_functions": stochastic}, location={"sheet": sim_sites[0]["sheet"], "cell": sim_sites[0]["cell"]})
    return finding("PASS", "No explicit SIM dimension references." + (f" Stochastic functions present: {stochastic}." if stochastic else ""), {"sim_sites": [], "stochastic_functions": stochastic})


# --- RZS-* resize ---------------------------------------------------------------
def _setsize_calls(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            for call in find_calls_detailed(f["formula"], "MM_SETSIZE"):
                out.append({"sheet": f["sheet"], "cell": f["cell"], "formula": f["formula"], "array_ref": f.get("array_ref"), **call})
    return out


def _resolve_int(analysis: dict[str, Any], sheet: str, arg: str) -> int | None:
    """Literal integer, or the constant in a same-sheet single-cell reference."""
    n = as_number(arg)
    if n is not None:
        return int(n)
    ref = parse_ref(arg)
    if ref and ref["cells"] == 1:
        v = cell_value(analysis, ref["sheet"] or sheet, ref["r1"], ref["c1"])
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return int(v)
    return None


def rzs_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    calls = _setsize_calls(analysis)
    return finding("PASS", f"Inventoried {len(calls)} MM_SETSIZE call(s).", {"sites": [{"sheet": c["sheet"], "cell": c["cell"], "args": c["args"]} for c in calls[:50]]})


def rzs_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RZS-002: MM_SETSIZE(NbRows, NbCols) added to a formula with '+' (or the
    documented text form), never nested in another function, never a
    SetSize of a SetSize."""
    problems = []
    for c in _setsize_calls(analysis):
        issues = []
        if len(c["args"]) != 2:
            issues.append(f"{len(c['args'])} argument(s), 2 expected (NbRows, NbCols)")
        masked = mask_strings(c["formula"])
        before = masked[: c["start"]].rstrip()
        after = masked[c["end"] :].lstrip()
        text_form = TEXT_SETSIZE_RE.search(masked) is not None
        added_with_plus = before.endswith("+") or after.startswith("+")
        bare = before.strip() in ("=", "") and after.strip() == ""
        if not (added_with_plus or text_form or bare):
            # nested inside another function's parentheses?
            depth = before.count("(") - before.count(")")
            if depth > 0:
                issues.append("nested inside another function (must be added with '+')")
            else:
                issues.append("not combined with '+' (documented form: =Formula + MM_SETSIZE(r, c))")
        if bare:
            issues.append("MM_SETSIZE alone in the cell: there is no formula to copy")
        if masked.upper().count("MM_SETSIZE") > 1:
            issues.append("more than one MM_SETSIZE in one formula (SetSize of a SetSize)")
        if any("MM_SETSIZE" in (a or "").upper() for a in c["args"]):
            issues.append("MM_SETSIZE inside MM_SETSIZE arguments")
        if issues:
            problems.append({"sheet": c["sheet"], "cell": c["cell"], "issues": issues, "formula": c["formula"][:160]})
    if problems:
        return finding("ERROR", f"{len(problems)} MM_SETSIZE call(s) with syntax problems: " + "; ".join(f"{p['sheet']}!{p['cell']}: {', '.join(p['issues'])}" for p in problems[:6]), {"problems": problems[:50]}, location={"sheet": problems[0]["sheet"], "cell": problems[0]["cell"]})
    n = len(_setsize_calls(analysis))
    return finding("PASS", f"All {n} MM_SETSIZE call(s) use the documented '+ MM_SETSIZE(NbRows, NbCols)' form." if n else "No MM_SETSIZE calls.", {"problems": []})


def rzs_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RZS-003: the NbRows x NbCols block starting at the MM_SETSIZE cell must be
    empty or hold the copies MM_SETSIZE itself made ('+0' where it was
    copied) -- it never overwrites existing content."""
    violations, unresolvable = [], []
    checked = 0
    for c in _setsize_calls(analysis):
        origin = parse_ref(c["cell"])
        if len(c["args"]) != 2 or origin is None:
            continue
        rows = _resolve_int(analysis, c["sheet"], c["args"][0])
        cols = _resolve_int(analysis, c["sheet"], c["args"][1])
        if rows is None or cols is None:
            unresolvable.append({"sheet": c["sheet"], "cell": c["cell"], "args": c["args"]})
            continue
        checked += 1
        bad = []
        for r in range(origin["r1"], origin["r1"] + max(rows, 1)):
            for col in range(origin["c1"], origin["c1"] + max(cols, 1)):
                if (r, col) == (origin["r1"], origin["c1"]):
                    continue
                v = cell_value(analysis, c["sheet"], r, col)
                if is_blank(v):
                    continue
                if isinstance(v, str) and v.startswith("=") and ("+0" in v.replace(" ", "") or "MM_SETSIZE" in v.upper() or v.replace(" ", "").endswith("&IF(0=0,\"\",0)")):
                    continue  # a copy MM_SETSIZE made itself
                bad.append(ref_text(col, r))
                if len(bad) >= 10:
                    break
            if len(bad) >= 10:
                break
        if bad:
            violations.append({"sheet": c["sheet"], "cell": c["cell"], "size": [rows, cols], "occupied_destination_cells": bad})
    if violations:
        return finding(
            "ERROR",
            f"{len(violations)} MM_SETSIZE destination block(s) contain content that is not a SetSize copy -- it will not be overwritten and the resize silently fails: "
            + "; ".join(f"{v['sheet']}!{v['cell']} ({v['size'][0]}x{v['size'][1]}): {v['occupied_destination_cells'][:5]}" for v in violations[:5]),
            {"violations": violations, "unresolvable": unresolvable, "checked": checked},
            location={"sheet": violations[0]["sheet"], "cell": violations[0]["cell"]},
        )
    if unresolvable:
        return finding("REQUIRES_USER_INPUT", f"{checked} MM_SETSIZE block(s) verified empty/copies; {len(unresolvable)} could not be sized (NbRows/NbCols are formulas, not constants): {fmt_sites(unresolvable)}", {"violations": [], "unresolvable": unresolvable, "checked": checked})
    return finding("PASS", f"{checked} MM_SETSIZE destination block(s) hold only copies or empty cells." if checked else "No MM_SETSIZE calls.", {"violations": [], "checked": checked})


def _resize_grids(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return [g for g in all_grids(analysis) if set(g["flag_names"]) & set(RESIZE_ROW_FLAGS + RESIZE_COL_FLAGS)]


def rzs_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _resize_grids(analysis)
    inventory = [{"grid": f"{g['sheet']}!{g['anchor']}", "name": g["display_name"], "flags": [f["raw"] for f in g["flags"] if f["name"] in RESIZE_ROW_FLAGS + RESIZE_COL_FLAGS]} for g in grids]
    has_setsize = bool(_setsize_calls(analysis))
    has_input = any("input" in g["flag_names"] or "inputlink" in g["flag_names"] for g in all_grids(analysis))
    if grids and not (has_setsize or has_input):
        return finding("WARNING", f"{len(grids)} grid(s) carry resize flags but nothing drives a resize (no MM_SETSIZE reference grid, no /Input or /InputLink grid resized by import): {fmt_grids(grids)}", {"resize_grids": inventory, "drivers": {"mm_setsize": has_setsize, "input_grids": has_input}}, location=grid_location(grids[0]))
    return finding("PASS", f"{len(grids)} grid(s) with /Resize, /ResizeRow or /ResizeColumn flags." + (" Drivers present." if grids else ""), {"resize_grids": inventory, "drivers": {"mm_setsize": has_setsize, "input_grids": has_input}})


def rzs_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RZS-005: /Resize.NAME groups need at least two members; a flag may
    carry at most one name."""
    groups: dict[str, list[dict[str, Any]]] = {}
    malformed = []
    for g in _resize_grids(analysis):
        for f in g["flags"]:
            if f["name"] not in RESIZE_ROW_FLAGS + RESIZE_COL_FLAGS:
                continue
            if len(f["args"]) > 1:
                malformed.append({"sheet": g["sheet"], "cell": g["anchor"], "flag": f["raw"]})
            key = (("col" if f["name"] in RESIZE_COL_FLAGS else "row"), f["args"][0] if f["args"] else "")
            groups.setdefault(f"{key[0]}:{key[1] or '(unnamed)'}", []).append(g)
    singles = {k: v for k, v in groups.items() if len(v) == 1}
    if malformed:
        return finding("ERROR", f"Resize flag(s) with more than one name segment: {[m['flag'] for m in malformed]}", {"groups": {k: len(v) for k, v in groups.items()}, "malformed": malformed}, location={"sheet": malformed[0]["sheet"], "cell": malformed[0]["cell"]})
    if singles:
        first = next(iter(singles.values()))[0]
        return finding("WARNING", f"Resize group(s) with a single member have nothing to synchronise with: {sorted(singles)} ({fmt_grids([v[0] for v in singles.values()])})", {"groups": {k: len(v) for k, v in groups.items()}, "single_member_groups": sorted(singles)}, location=grid_location(first))
    return finding("PASS", f"{len(groups)} resize group(s), each with at least two grids." if groups else "No resize flags.", {"groups": {k: len(v) for k, v in groups.items()}})


def _resizable_grids(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    setsize_cells = {(c["sheet"], c["cell"]) for c in _setsize_calls(analysis)}
    out = []
    for g in all_grids(analysis):
        flags = set(g["flag_names"])
        resizable = bool(flags & set(RESIZE_ROW_FLAGS + RESIZE_COL_FLAGS)) or "input" in flags or "inputlink" in flags or "instancesplit" in flags
        if not resizable:
            for (s, cell) in setsize_cells:
                if s == g["sheet"]:
                    ref = parse_ref(cell)
                    if ref and g["first_row"] <= ref["r1"] <= g["last_row"] and g["first_col"] <= ref["c1"] <= g["last_col"]:
                        resizable = True
                        break
        if resizable:
            out.append(g)
    return out


def rzs_006(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RZS-006: formulas that read a resizable grid with a fixed range ending
    exactly on its current last row/column will not follow the resize --
    they must use MM_LASTROWCELL / MM_LASTCOLUMNCELL (or MM_TABLE/MM_COLUMN/MM_ROW)."""
    resizable = _resizable_grids(analysis)
    if not resizable:
        return finding("PASS", "No resizable grids (no resize flags, /Input grids or MM_SETSIZE), so fixed ranges cannot go stale.", {"resizable_grids": 0})
    by_sheet: dict[str, list[dict[str, Any]]] = {}
    for g in resizable:
        by_sheet.setdefault(g["sheet"], []).append(g)
    dynamic_refs = sum(analysis["features"].get("function_usage", {}).get(fn, 0) for fn in ("MM_LASTROWCELL", "MM_LASTCOLUMNCELL", "MM_TABLE", "MM_COLUMN", "MM_ROW", "MM_LASTROW", "MM_LASTCOLUMN"))
    stale = []
    for wb in analysis["workbooks"]:
        for f in wb["formulas"]:
            for ref in cell_refs_in_formula(f["formula"]):
                if ref["whole_column"] or ref["whole_row"] or ref["cells"] == 1:
                    continue
                sheet = ref["sheet"] or f["sheet"]
                for g in by_sheet.get(sheet, []):
                    inside = g["first_col"] <= ref["c1"] and ref["c2"] <= g["last_col"] and g["first_row"] <= ref["r1"] and ref["r2"] <= g["last_row"]
                    ends_on_edge = (ref["r2"] == g["last_row"] and ref["r1"] > g["first_row"]) or (ref["c2"] == g["last_col"] and ref["c1"] > g["first_col"])
                    own_cell = parse_ref(f["cell"])
                    in_same_grid = own_cell and f["sheet"] == g["sheet"] and g["first_row"] <= own_cell["r1"] <= g["last_row"] and g["first_col"] <= own_cell["c1"] <= g["last_col"]
                    if inside and ends_on_edge and not in_same_grid:
                        stale.append({"sheet": f["sheet"], "cell": f["cell"], "reference": (ref["sheet"] + "!" if ref["sheet"] else "") + ref["ref"], "grid": g["display_name"]})
                        break
                else:
                    continue
                break
    observed = {"resizable_grids": [f"{g['sheet']}!{g['anchor']} ({g['display_name']})" for g in resizable[:50]], "dynamic_reference_calls": dynamic_refs, "fixed_ranges_to_resizable_grids": stale[:50]}
    if stale:
        return finding(
            "ERROR",
            f"{len(stale)} formula(s) read a resizable grid through a fixed range that stops at its current last row/column and will not follow a resize (use MM_LASTROWCELL/MM_LASTCOLUMNCELL or MM_TABLE/MM_COLUMN): {fmt_sites(stale)}",
            observed,
            location={"sheet": stale[0]["sheet"], "cell": stale[0]["cell"]},
            evidence="INFERENCE",
        )
    return finding("PASS", f"{len(resizable)} resizable grid(s); no fixed range from outside a grid stops on its current edge; {dynamic_refs} dynamic reference call(s) in use.", observed, evidence="INFERENCE")


def rzs_007(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RZS-007: MM_GETRANGE is system-generated next to dynamic functions."""
    getrange = mm_calls(analysis, "MM_GETRANGE")
    usage = analysis["features"].get("function_usage", {})
    dynamic = {fn: usage[fn] for fn in sorted(DYNAMIC_FUNCTIONS & set(usage))}
    if dynamic and not getrange:
        return finding("WARNING", f"Dynamic function(s) {dynamic} are used but no MM_GETRANGE cells exist next to them -- the add-in normally generates those when the formula is entered; the results may not spill in Excel (they will in Mind).", {"mm_getrange_calls": 0, "dynamic_functions": dynamic})
    return finding("PASS", f"{len(getrange)} MM_GETRANGE cell(s) (system-generated) for dynamic functions {dynamic or 'none'}.", {"mm_getrange_calls": len(getrange), "dynamic_functions": dynamic})
