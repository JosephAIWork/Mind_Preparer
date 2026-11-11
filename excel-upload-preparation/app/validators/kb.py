"""Rules  mined directly  from the Milliman Mind knowledge base
(CompleteMindDocn.docx) that 01_Mind_Readiness_Standard.md does not cover:
documented flags, /Group syntax, backup flags, /HideRows columns,
MM_RANGE placement, unique special grids, /Translations headers and the
#NbSimulations grid. Definitions live in rules/kb-rules.yaml."""
from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

from ..formula_utils import parse_ref
from ..grids import all_grids, grid_containing, grid_flags
from ..inventory import mm_calls
from ._common import as_bool, finding, flagged_grids, fmt_grids, fmt_sites, grid_location, grid_rows, header_names, is_blank

FLAGS_PATH = Path(__file__).resolve().parents[2] / "references" / "mind-flags.yaml"
ISO_639_1 = set(
    "aa ab ae af ak am an ar as av ay az ba be bg bh bi bm bn bo br bs ca ce ch co cr cs cu cv cy da de dv dz ee el en eo es et eu fa ff fi fj fo fr fy ga gd gl gn gu gv ha he hi ho hr ht hu hy hz ia id ie ig ii ik io is it iu ja jv ka kg ki kj kk kl km kn ko kr ks ku kv kw ky la lb lg li ln lo lt lu lv mg mh mi mk ml mn mr ms mt my na nb nd ne ng nl nn no nr nv ny oc oj om or os pa pi pl ps pt qu rm rn ro ru rw sa sc sd se sg si sk sl sm sn so sq sr ss st su sv sw ta te tg th ti tk tl tn to tr ts tt tw ty ug uk ur uz ve vi vo wa wo xh yi yo za zh zu".split()
)


@lru_cache(maxsize=1)
def documented_flags() -> dict[str, dict[str, Any]]:
    if not FLAGS_PATH.exists():
        return {}
    data = yaml.safe_load(FLAGS_PATH.read_text(encoding="utf-8")) or {}
    return {f["name"].lower(): f for f in data.get("flags", [])}


def flg_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """FLG-001:  every '/Flag'  must be documented, with the documented argument count."""
    docs = documented_flags()
    if not docs:
        return finding("NOT_SUPPORTED", "references/mind-flags.yaml is missing; run tools/mine_kb_docx.py.", None)
    unknown, bad_args, used = [], [], {}
    for g in all_grids(analysis):
        for f in g["flags"] + [{"name": m["name"], "raw": m["raw"], "args": []} for m in g.get("header_markers", []) if m["name"] not in ("&&hidecolumn", "&hide")]:
            used[f["name"]] = used.get(f["name"], 0) + 1
            spec = docs.get(f["name"])
            if spec is None:
                unknown.append({"sheet": g["sheet"], "cell": g.get("title_cell") or g["anchor"], "flag": f["raw"]})
                continue
            n_args = len(f.get("args", []))
            if "args" in spec and n_args != spec["args"]:
                bad_args.append({"sheet": g["sheet"], "cell": g.get("title_cell") or g["anchor"], "flag": f["raw"], "expected": spec.get("syntax")})
            elif "max_args" in spec and n_args > spec["max_args"]:
                bad_args.append({"sheet": g["sheet"], "cell": g.get("title_cell") or g["anchor"], "flag": f["raw"], "expected": spec.get("syntax")})
    observed = {"flags_used": used, "unknown": unknown[:50], "wrong_argument_count": bad_args[:50], "documented_flag_count": len(docs)}
    if bad_args:
        return finding("ERROR", f"{len(bad_args)} flag(s) have the wrong argument shape: " + "; ".join(f"{b['sheet']}!{b['cell']} {b['flag']} (expected {b['expected']})" for b in bad_args[:5]), observed, location={"sheet": bad_args[0]["sheet"], "cell": bad_args[0]["cell"]})
    if unknown:
        return finding("WARNING", f"{len(unknown)} flag token(s) are not documented in the KB and will be ignored by Mind (typo?): " + ", ".join(f"{u['flag']} at {u['sheet']}!{u['cell']}" for u in unknown[:8]), observed, location={"sheet": unknown[0]["sheet"], "cell": unknown[0]["cell"]})
    return finding("PASS", f"All {sum(used.values())} flag usage(s) ({len(used)} distinct) are documented with the right shape." if used else "No grid flags used.", observed)


def grp_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """GRP-001:  /Group.(GroupName).x.y"""
    groups: dict[str, list[dict[str, Any]]] = {}
    problems = []
    for g in all_grids(analysis):
        for f in grid_flags(g, "group"):
            args = f["args"]
            if len(args) != 3 or not args[0]:
                problems.append({"sheet": g["sheet"], "cell": g.get("title_cell") or g["anchor"], "flag": f["raw"], "issue": "expected /Group.(GroupName).x.y"})
                continue
            if not (args[1].isdigit() and args[2].isdigit()):
                problems.append({"sheet": g["sheet"], "cell": g.get("title_cell") or g["anchor"], "flag": f["raw"], "issue": "x and y must be non-negative integers"})
                continue
            groups.setdefault(args[0], []).append({"grid": g, "pos": (int(args[1]), int(args[2]))})
    for name, members in groups.items():
        positions = [m["pos"] for m in members]
        dups = sorted({p for p in positions if positions.count(p) > 1})
        if dups:
            problems.append({"sheet": members[0]["grid"]["sheet"], "cell": members[0]["grid"]["anchor"], "flag": f"/Group.({name})", "issue": f"positions {dups} used by more than one grid"})
        if len({m["grid"]["sheet"] for m in members}) > 1:
            problems.append({"sheet": members[0]["grid"]["sheet"], "cell": members[0]["grid"]["anchor"], "flag": f"/Group.({name})", "issue": "group spans several sheets (KB recommends the same sheet to keep coherent widths/heights)"})
    observed = {"groups": {k: len(v) for k, v in groups.items()}, "problems": problems}
    if problems:
        hard = [p for p in problems if "spans" not in p["issue"]]
        return finding("ERROR" if hard else "WARNING", f"{len(problems)} /Group problem(s): " + "; ".join(f"{p['flag']}: {p['issue']}" for p in problems[:6]), observed, location={"sheet": problems[0]["sheet"], "cell": problems[0]["cell"]})
    return finding("PASS", f"{len(groups)} grid group(s) with valid /Group.(Name).x.y flags." if groups else "No /Group flags.", observed)


def bkp_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """BKP-001:  backup flags  and the MM_BACKUPBUTTON placement."""
    sources: dict[str, list[dict[str, Any]]] = {}
    dests: dict[str, list[dict[str, Any]]] = {}
    backup_grids = []
    unnamed = []
    for g in all_grids(analysis):
        for f in g["flags"]:
            if f["name"] not in ("backupsource", "backupdest", "backupdestalldimensions"):
                continue
            backup_grids.append(g)
            if not f["args"]:
                unnamed.append({"sheet": g["sheet"], "cell": g["anchor"], "flag": f["raw"]})
                continue
            (sources if f["name"] == "backupsource" else dests).setdefault(f["args"][0], []).append(g)
    buttons = mm_calls(analysis, "MM_BACKUPBUTTON")
    if not backup_grids and not buttons:
        return finding("PASS", "No backup flags or MM_BACKUPBUTTON.", {"backups": {}})
    problems = []
    for name, gs in sources.items():
        if len(gs) > 1:
            problems.append(f"backup '{name}' has {len(gs)} /BackupSource grids (only one allowed)")
        if name not in dests:
            problems.append(f"backup '{name}' has no /BackupDest grid")
    for name in dests:
        if name not in sources:
            problems.append(f"backup '{name}' has /BackupDest but no /BackupSource")
    if unnamed:
        problems.append(f"backup flags without a name: {[u['flag'] for u in unnamed]}")
    for b in buttons:
        origin = parse_ref(b["cell"])
        g = grid_containing([x for x in all_grids(analysis) if x["sheet"] == b["sheet"]], origin["r1"], origin["c1"]) if origin else None
        if g and any(f["name"] in ("backupsource", "backupdest", "backupdestalldimensions") for f in g["flags"]):
            problems.append(f"MM_BACKUPBUTTON at {b['sheet']}!{b['cell']} sits inside a backup grid")
    observed = {"backups": {n: {"sources": len(sources.get(n, [])), "destinations": len(dests.get(n, []))} for n in set(sources) | set(dests)}, "buttons": [{"sheet": b["sheet"], "cell": b["cell"]} for b in buttons], "problems": problems}
    loc = grid_location(backup_grids[0]) if backup_grids else {"sheet": buttons[0]["sheet"], "cell": buttons[0]["cell"]}
    if problems:
        return finding("ERROR", "; ".join(problems), observed, location=loc)
    if backup_grids and not buttons:
        return finding("WARNING", f"{len(observed['backups'])} backup(s) defined but no MM_BACKUPBUTTON cell triggers them.", observed, location=loc)
    return finding("PASS", f"{len(observed['backups'])} backup(s) with one source and at least one destination each; {len(buttons)} MM_BACKUPBUTTON cell(s) outside backup grids.", observed)


def hid_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """HID-001:  a /HideRows  header column holds TRUE/FALSE (or 1/0, or formulas)."""
    problems, warnings, checked = [], [], 0
    for g in all_grids(analysis):
        for m in g.get("header_markers", []):
            if m["name"] != "hiderows":
                continue
            checked += 1
            col = m["column_index"]
            if col != g["n_cols"] - 1:
                warnings.append(f"{g['display_name']}: /HideRows column is not the last column (KB: 'a new column at the right of your grid')")
            for r, row in enumerate(grid_rows(analysis, g)[1:], start=g["first_row"] + 1):
                v = row[col] if col < len(row) else None
                if is_blank(v) or (isinstance(v, str) and v.startswith("=")):
                    continue
                if as_bool(v) is None:
                    problems.append({"sheet": g["sheet"], "row": r, "grid": g["display_name"], "value": v})
    if not checked:
        return finding("PASS", "No /HideRows columns.", {"checked": 0})
    observed = {"checked": checked, "problems": problems[:50], "warnings": warnings}
    if problems:
        return finding("ERROR", f"{len(problems)} /HideRows value(s) are not TRUE/FALSE/1/0: {problems[:6]}", observed, location={"sheet": problems[0]["sheet"], "cell": f"A{problems[0]['row']}"})
    if warnings:
        return finding("WARNING", "; ".join(warnings), observed)
    return finding("PASS", f"{checked} /HideRows column(s) hold only boolean values or formulas.", observed)


def rng_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """RNG-001:  MM_RANGE alone  in a single-cell table, used at the top level."""
    calls = mm_calls(analysis, "MM_RANGE")
    if not calls:
        return finding("PASS", "No MM_RANGE calls.", {"problems": []})
    problems = []
    for c in calls:
        issues = []
        stripped = c["formula"].replace(" ", "")
        if not re.match(r"^=MM_RANGE\(", stripped, re.IGNORECASE) or not stripped.endswith(")") or stripped.upper().count("MM_RANGE(") > 1:
            issues.append("not the whole formula (MM_RANGE must be used alone, not inside another function or expression)")
        origin = parse_ref(c["cell"])
        g = grid_containing([x for x in all_grids(analysis) if x["sheet"] == c["sheet"]], origin["r1"], origin["c1"]) if origin else None
        if g and g["n_rows"] * g["n_cols"] > 1:
            issues.append(f"shares grid {g['display_name']} ({g['ref']}) with other cells")
        if issues:
            problems.append({"sheet": c["sheet"], "cell": c["cell"], "issues": issues})
    if problems:
        return finding("ERROR", f"{len(problems)} MM_RANGE call(s) misplaced: " + "; ".join(f"{p['sheet']}!{p['cell']}: {', '.join(p['issues'])}" for p in problems[:5]), {"problems": problems}, location={"sheet": problems[0]["sheet"], "cell": problems[0]["cell"]}, evidence="INFERENCE")
    return finding("PASS", f"All {len(calls)} MM_RANGE call(s) are alone in a single-cell table.", {"problems": []}, evidence="INFERENCE")


def unq_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """UNQ-001:  special grids  that must be unique per project."""
    docs = documented_flags()
    unique_flags = sorted(name for name, spec in docs.items() if spec.get("unique"))
    dups = {}
    for name in unique_flags:
        gs = flagged_grids(analysis, name)
        if len(gs) > 1:
            dups[name] = gs
    if dups:
        first = next(iter(dups.values()))[0]
        return finding("ERROR", "; ".join(f"/{docs[n]['name']} appears on {len(gs)} grids ({fmt_grids(gs, 4)})" for n, gs in dups.items()), {"duplicates": {n: [f"{g['sheet']}!{g['ref']}" for g in gs] for n, gs in dups.items()}, "unique_flags": unique_flags}, location=grid_location(first))
    return finding("PASS", f"No duplicated special grid ({len(unique_flags)} unique-flag kinds checked).", {"duplicates": {}, "unique_flags": unique_flags})


def trn_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = flagged_grids(analysis, "translations")
    if not grids:
        return finding("PASS", "No /Translations grid.", {"languages": []})
    g = grids[0]
    headers = header_names(g)
    bad = [h for h in headers if h.strip().lower() not in ISO_639_1]
    if bad or len(headers) < 2:
        return finding("ERROR", f"/Translations headers must be ISO 639-1 two-letter codes with at least two columns; invalid: {bad}", {"languages": headers}, location=grid_location(g))
    return finding("PASS", f"/Translations grid with languages {headers} (first column = design language).", {"languages": headers})


def sim_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = [g for g in all_grids(analysis) if (g.get("name") or "").replace(" ", "").lower() == "nbsimulations"]
    if not grids:
        return finding("PASS", "No #NbSimulations grid (the simulation count stays a user input in Mind).", {"grids": []})
    problems = []
    for g in grids:
        rows = grid_rows(analysis, g)
        values = [v for row in rows for v in row if not is_blank(v)]
        if len(values) != 1 or not isinstance(values[0], (int, float)) or isinstance(values[0], bool) or values[0] < 1 or int(values[0]) != values[0]:
            problems.append({"sheet": g["sheet"], "cell": g["anchor"], "values": values[:5]})
    if problems:
        return finding("ERROR", f"#NbSimulations grid must contain exactly one positive integer: {problems}", {"grids": [f"{g['sheet']}!{g['ref']}" for g in grids], "problems": problems}, location={"sheet": problems[0]["sheet"], "cell": problems[0]["cell"]})
    return finding("PASS", f"#NbSimulations grid locks the simulation count: {grid_rows(analysis, grids[0])[0]}.", {"grids": [f"{g['sheet']}!{g['ref']}" for g in grids]})
