"""PRJ-*  validators: the  /ProjectSettings grid (KB '/ProjectSettings grid':
unique grid with Name | Value | Locked | Hidden; 'Projects settings edition'
for the Runs/Debug settings; 'General structure and guidelines' for the
#NbSimulations grid). The KB's list of  setting *names* lives in an attached
sample file that is not part  of the scrape, so KB names  are not validated -- but a subset
confirmed  rejected by Mind's real converter is flagged by PRJ-006."""
from __future__ import annotations

from typing import Any

from ..grids import all_grids
from ._common import as_bool, finding, flagged_grids, fmt_grids, grid_location, grid_rows, header_names, is_blank

STOCHASTIC_NATIVE = {"RAND", "RANDBETWEEN", "RANDARRAY"}
STOCHASTIC_MM_PREFIXES = ("MM_SIMULATE", "MM_RANDOM", "MM_SIMCOPULA", "MM_RANDOMSAMPLE", "MM_SIMQUANTILE", "MM_SIMINDEXQUANTILE")
DEBUG_KEYWORDS = ("profiler", "debug", "nan")
SIM_KEYWORDS = ("simulation", "stochastic", "seed", "multicore", "multithread", "batch")
# /ProjectSettings  setting names  Mind's converter rejects as non-existent (proven by real
# upload testing 2026-08-30: declaring them fails the Convert step). Compared lower-cased.
MIND_INVALID_SETTINGS = {"enablecheckformatinputmanager", "enabledebugmode"}


def _settings(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return flagged_grids(analysis, "projectsettings")


def _setting_rows(analysis: dict[str, Any], g: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for r, row in enumerate(grid_rows(analysis, g)[1:], start=g["first_row"] + 1):
        if not row or is_blank(row[0]):
            continue
        out.append({"row": r, "name": str(row[0]).strip(), "value": row[1] if len(row) > 1 else None, "locked": row[2] if len(row) > 2 else None, "hidden": row[3] if len(row) > 3 else None})
    return out


def prj_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _settings(analysis)
    if len(grids) > 1:
        return finding("ERROR", f"{len(grids)} /ProjectSettings grids -- it must be unique per project: {fmt_grids(grids)}", {"projectsettings_grids": [f"{g['sheet']}!{g['ref']}" for g in grids]}, location=grid_location(grids[0]))
    if not grids:
        return finding("PASS", "No /ProjectSettings grid: Mind's default settings apply (stochastic mode auto-enables if stochastic formulas exist).", {"projectsettings_grids": [], "settings": []})
    g = grids[0]
    rows = _setting_rows(analysis, g)
    return finding("PASS", f"/ProjectSettings grid at {g['sheet']}!{g['ref']} pre-sets {len(rows)} setting(s).", {"projectsettings_grids": [f"{g['sheet']}!{g['ref']}"], "settings": [r["name"] for r in rows]})


def prj_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _settings(analysis)
    if not grids:
        return finding("PASS", "No /ProjectSettings grid.", {"problems": []})
    g = grids[0]
    headers = [h.lower() for h in header_names(g)]
    expected = ["name", "value", "locked", "hidden"]
    problems = []
    if headers[:4] != expected:
        problems.append({"issue": f"headers {header_names(g)[:4]} must be Name | Value | Locked | Hidden"})
    else:
        for r in _setting_rows(analysis, g):
            for key in ("locked", "hidden"):
                v = r[key]
                if not is_blank(v) and not (isinstance(v, str) and v.startswith("=")) and as_bool(v) is None:
                    problems.append({"row": r["row"], "name": r["name"], "issue": f"{key.capitalize()}='{v}' is not TRUE/FALSE"})
    if problems:
        return finding("ERROR", f"/ProjectSettings structure problem(s): {problems[:6]}", {"problems": problems, "headers": header_names(g)}, location=grid_location(g))
    return finding("PASS", "/ProjectSettings grid has Name | Value | Locked | Hidden with boolean Locked/Hidden values.", {"problems": [], "headers": header_names(g)})


def prj_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _settings(analysis)
    if not grids:
        return finding("PASS", "No /ProjectSettings grid, so no hidden settings.", {"hidden_settings": []})
    hidden = [{"name": r["name"], "value": r["value"], "locked": r["locked"]} for r in _setting_rows(analysis, grids[0]) if as_bool(r["hidden"]) is True]
    locked = [r["name"] for r in _setting_rows(analysis, grids[0]) if as_bool(r["locked"]) is True]
    if hidden:
        return finding("WARNING", f"{len(hidden)} setting(s) are hidden from the settings panel (users cannot see or change them): {[h['name'] for h in hidden]}. Confirm intended.", {"hidden_settings": hidden, "locked_settings": locked}, location=grid_location(grids[0]))
    return finding("PASS", f"No hidden settings ({len(locked)} locked).", {"hidden_settings": [], "locked_settings": locked})


def _stochastic_functions(analysis: dict[str, Any]) -> list[str]:
    usage = analysis["features"].get("function_usage", {})
    return sorted(fn for fn in usage if fn in STOCHASTIC_NATIVE or fn.startswith(STOCHASTIC_MM_PREFIXES))


def _nb_simulations_grids(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return [g for g in all_grids(analysis) if (g.get("name") or "").replace(" ", "").lower() == "nbsimulations"]


def prj_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    stochastic = _stochastic_functions(analysis)
    nb_grids = _nb_simulations_grids(analysis)
    sim_settings = []
    for g in _settings(analysis):
        sim_settings += [{"name": r["name"], "value": r["value"], "locked": r["locked"]} for r in _setting_rows(analysis, g) if any(k in r["name"].lower() for k in SIM_KEYWORDS)]
    observed = {"stochastic_functions": stochastic, "nb_simulations_grids": [f"{g['sheet']}!{g['ref']}" for g in nb_grids], "simulation_settings": sim_settings}
    if not stochastic:
        return finding("PASS", "No stochastic functions; the model runs deterministically." + (f" (#NbSimulations grid present: {observed['nb_simulations_grids']})" if nb_grids else ""), observed)
    if nb_grids or sim_settings:
        return finding("PASS", f"Stochastic functions {stochastic} with simulation configuration: #NbSimulations grid {observed['nb_simulations_grids'] or 'none'}, settings {[s['name'] for s in sim_settings] or 'none'}.", observed)
    return finding("WARNING", f"Stochastic functions {stochastic} are used but no #NbSimulations grid or simulation/seed setting is pre-set -- Mind will auto-enable stochastic mode and ask the user for a simulation count each run.", observed)


def prj_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    debug_settings = []
    for g in _settings(analysis):
        debug_settings += [{"name": r["name"], "value": r["value"], "locked": r["locked"], "hidden": r["hidden"]} for r in _setting_rows(analysis, g) if any(k in r["name"].lower() for k in DEBUG_KEYWORDS) and r["name"].lower() not in MIND_INVALID_SETTINGS]
    if debug_settings:
        return finding("PASS", f"Debug-related settings pre-set: {[d['name'] for d in debug_settings]}.", {"debug_settings": debug_settings})
    return finding("PASS", "No debug-related settings pre-set (performance profiler, stop at first NaN and Debug mode stay at defaults; enable them in Settings > Debug during development).", {"debug_settings": []}, evidence="RECOMMENDATION")


def prj_006(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """PRJ-006:  reject /ProjectSettings  setting names Mind's converter does not recognise.
    Confirmed by real Mind upload testing: declaring one of MIND_INVALID_SETTINGS makes the
    Convert step fail with 'The <name> setting does not exist'. Mind converts fine without  them."""
    grids = _settings(analysis)
    bad = []
    for g in grids:
        for r in _setting_rows(analysis, g):
            if r["name"].lower() in MIND_INVALID_SETTINGS:
                bad.append({"sheet": g["sheet"], "grid": g["ref"], "row": r["row"], "name": r["name"]})
    if bad:
        names = [b["name"] for b in bad]
        return finding("ERROR", f"{len(bad)} /ProjectSettings setting name(s) are rejected by Mind's converter as non-existent -- uploading this workbook fails the Convert step: {names}. Remove them (Mind converts fine without them).", {"invalid_settings": bad}, location=grid_location(grids[0]), evidence="DETERMINISTIC_FINDING")
    if not grids:
        return finding("PASS", "No /ProjectSettings grid, so no invalid setting names.", {"invalid_settings": []})
    return finding("PASS", "No /ProjectSettings setting names known to be rejected by Mind's converter.", {"invalid_settings": []})

