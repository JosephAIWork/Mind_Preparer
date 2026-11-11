"""INP-*  / EXP-*  validators: Input Manager and Export Manager grids.

KB sources: 'Input manager' (/Input grids, CSV named like the grid),
'Input manager automatch' (/InputSettings: GridName | FileNamePattern with
{GridName} {InstanceName} {ModelName} {*} {???}), 'Input manager reorder
columns'  (/Reorder needs /Input, no /NoHeader, unique headers), 'Projects
settings edition' (Enable check  format in Input manager), 'Export'
(/Export, unique  /ExportSettings with its documented columns and
FileNamePattern tags).  INP-010: outside files a macro imports (the facts come
from the server,  config["outside_inputs"]; app/outside_inputs.py)."""
from __future__ import annotations

import re
from typing import Any

from ..grids import all_grids
from ..plain_codes import cite
from ._common import as_bool, finding, flagged_grids, fmt_grids, grid_location, grid_rows, header_index, header_names, is_blank

ILLEGAL_FILENAME_CHARS = set('\\/:*?"<>|')
INPUT_PATTERN_TAGS = {"gridname", "instancename", "modelname", "*"}
INPUT_PATTERN_TAG_RE = re.compile(r"\{([^{}]*)\}")
EXPORT_COLUMNS = [
    "GridName", "ExportByInstance", "ExportNoInstanceKeys", "ExportByLoop", "ExportNoLoopKeys", "Separator", "Culture", "Headers", "SubHeaders",
    "IncludeHiddenData", "FileNamePattern", "FileExtension", "Locked", "AllowScientificFormat",
]
EXPORT_BOOLEAN_COLUMNS = {"exportbyinstance", "exportnoinstancekeys", "exportbyloop", "exportnoloopkeys", "headers", "subheaders", "includehiddendata", "locked", "allowscientificformat"}
EXPORT_PATTERN_TAGS = {"modelname", "gridname", "instancename", "loopslabels", "yyyymmdd"}


def _grid_summary(g: dict[str, Any]) -> dict[str, Any]:
    return {"sheet": g["sheet"], "ref": g["ref"], "name": g["display_name"], "flags": g["flag_names"], "rows": g["n_rows"], "cols": g["n_cols"]}


# ---  INP ---------------------------------------------------------------------
def inp_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    inputs = flagged_grids(analysis, "input")
    links = flagged_grids(analysis, "inputlink")
    return finding(
        "PASS",
        f"{len(inputs)} /Input grid(s) and {len(links)} /InputLink grid(s)." + ("" if inputs else " No Input Manager imports are configured."),
        {"input_grids": [_grid_summary(g) for g in inputs], "inputlink_grids": [_grid_summary(g) for g in links]},
    )


def inp_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """INP-002:  the Input  Manager matches CSV files to grids by grid name."""
    inputs = flagged_grids(analysis, "input")
    if not inputs:
        return finding("PASS", "No /Input grids.", {"issues": []})
    issues = []
    seen: dict[str, dict[str, Any]] = {}
    for g in inputs:
        name = g.get("name") or ""
        if not name:
            issues.append({"sheet": g["sheet"], "cell": g["anchor"], "issue": "input grid has no '#Name' -- no CSV can match it"})
            continue
        if ILLEGAL_FILENAME_CHARS & set(name):
            issues.append({"sheet": g["sheet"], "cell": g["anchor"], "issue": f"name '{name}' contains characters unusable in a CSV file name"})
        key = name.strip().lower()
        if key in seen:
            issues.append({"sheet": g["sheet"], "cell": g["anchor"], "issue": f"duplicate input grid name '{name}' (also {seen[key]['sheet']}!{seen[key]['anchor']})"})
        else:
            seen[key] = g
    if issues:
        return finding("ERROR", f"{len(issues)} /Input naming issue(s): " + "; ".join(f"{i['sheet']}!{i['cell']}: {i['issue']}" for i in issues[:6]), {"issues": issues}, location={"sheet": issues[0]["sheet"], "cell": issues[0]["cell"]})
    return finding("PASS", f"All {len(inputs)} /Input grid names are unique and file-name safe (expected CSV: '<GridName>.csv', or '<InstanceName>-<GridName>.csv' with instances).", {"input_grid_names": sorted(seen)})


def _input_settings(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return flagged_grids(analysis, "inputsettings")


def inp_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _input_settings(analysis)
    if not grids:
        return finding("PASS", "No /InputSettings grid (optional: files are then matched by grid name).", {"inputsettings_grids": []})
    if len(grids) > 1:
        return finding("ERROR", f"{len(grids)} /InputSettings grids -- it must be unique: {fmt_grids(grids)}", {"inputsettings_grids": [_grid_summary(g) for g in grids]}, location=grid_location(grids[0]))
    g = grids[0]
    headers = [h.lower() for h in header_names(g)]
    if len(headers) < 2 or headers[0] != "gridname" or headers[1] != "filenamepattern":
        return finding("ERROR", f"/InputSettings grid headers must be 'GridName' then 'FileNamePattern' (found {header_names(g)[:4]}).", {"inputsettings_grids": [_grid_summary(g)], "headers": header_names(g)}, location=grid_location(g))
    input_names = {(x.get("name") or "").strip().lower() for x in flagged_grids(analysis, "input")}
    unknown = [str(row[0]) for row in grid_rows(analysis, g)[1:] if not is_blank(row[0]) and str(row[0]).strip().lower() not in input_names]
    if unknown:
        return finding("WARNING", f"/InputSettings names {unknown[:10]} do not match any /Input grid name.", {"inputsettings_grids": [_grid_summary(g)], "unknown_grid_names": unknown}, location=grid_location(g))
    return finding("PASS", f"/InputSettings grid at {g['sheet']}!{g['ref']} has the documented columns and every GridName matches an /Input grid.", {"inputsettings_grids": [_grid_summary(g)]})


def inp_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _input_settings(analysis)
    if not grids:
        return finding("PASS", "No /InputSettings grid, so no filename patterns to validate.", {"patterns": []})
    g = grids[0]
    col = header_index(g, "FileNamePattern")
    if col is None:
        return finding("ERROR", "/InputSettings grid has no FileNamePattern column.", {"patterns": []}, location=grid_location(g))
    bad, patterns = [], []
    for row in grid_rows(analysis, g)[1:]:
        if is_blank(row[0]):
            continue
        pattern = row[col] if col < len(row) else None
        patterns.append({"grid": row[0], "pattern": pattern})
        if is_blank(pattern):
            bad.append({"grid": row[0], "pattern": pattern, "issue": "empty pattern"})
            continue
        for tag in INPUT_PATTERN_TAG_RE.findall(str(pattern)):
            if tag.lower() not in INPUT_PATTERN_TAGS and not re.fullmatch(r"\?+", tag):
                bad.append({"grid": row[0], "pattern": pattern, "issue": f"unknown tag {{{tag}}}"})
    if bad:
        return finding("ERROR", f"{len(bad)} FileNamePattern problem(s) (allowed tags: {{GridName}} {{InstanceName}} {{ModelName}} {{*}} {{???}}): {bad[:6]}", {"patterns": patterns, "problems": bad}, location=grid_location(g))
    return finding("PASS", f"All {len(patterns)} FileNamePattern value(s) use documented tags only.", {"patterns": patterns})


def inp_005(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = flagged_grids(analysis, "reorder")
    if not grids:
        return finding("PASS", "No /Reorder grids.", {"reorder_grids": []})
    problems = []
    for g in grids:
        flags = set(g["flag_names"])
        issues = []
        if "input" not in flags:
            issues.append("missing /Input")
        if "noheader" in flags or "noheaders" in flags:
            issues.append("has /NoHeader")
        names = [h.strip().lower() for h in header_names(g)]
        dups = sorted({n for n in names if n and names.count(n) > 1})
        if dups:
            issues.append(f"duplicate headers {dups}")
        if not g["header_is_all_text"]:
            issues.append("first row is not text-only (headers are needed to match columns)")
        if issues:
            problems.append({"sheet": g["sheet"], "cell": g["anchor"], "grid": g["display_name"], "issues": issues})
    if problems:
        return finding("ERROR", f"{len(problems)} /Reorder grid(s) break the documented requirements: " + "; ".join(f"{p['grid']}: {', '.join(p['issues'])}" for p in problems[:6]), {"reorder_grids": [_grid_summary(g) for g in grids], "problems": problems}, location={"sheet": problems[0]["sheet"], "cell": problems[0]["cell"]})
    return finding("PASS", f"All {len(grids)} /Reorder grid(s) have /Input, no /NoHeader and unique text headers.", {"reorder_grids": [_grid_summary(g) for g in grids]})


def inp_006(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    inputs = flagged_grids(analysis, "input")
    if not inputs:
        return finding("PASS", "No /Input grids; the Input Manager format check is not relevant.", {"input_grids": 0})
    # Input  format checking  is toggled per grid in Mind's Input Manager at import time -- it is NOT a
    # workbook-declared setting. Declaring an 'EnableCheckFormatInputManager' /ProjectSettings setting
    # makes Mind's converter fail (proven by real upload testing; flagged  by PRJ-006).
    return finding(
        "PASS",
        f"{len(inputs)} /Input grid(s): enable format checking per grid in Mind's Input Manager at import "
        "time (it validates imported CSVs against each grid's Excel format). Do not declare an "
        f"'EnableCheckFormatInputManager' /ProjectSettings setting -- Mind's converter rejects it: see {cite('PRJ-006')}.",
        {"input_grids": len(inputs)}, evidence="RECOMMENDATION")


def _file_label(row: dict[str, Any]) -> str:
    """How  a row  of the import table is named to the user: the file it expects, else its key, else its cell."""
    return str(row.get("expected") or "/".join(str(k) for k in row.get("keys") or []) or row.get("cell") or "?")


def inp_010(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    """INP-010:  the macros  open other files and paste their data in (app/outside_inputs.py). Mind
    runs no macro and opens no outside file, so each file must become a Mind input the Mind team
    swaps in Mind's Input  Manager. The facts come from the server, in config["outside_inputs"]:
    what the upload's  macros open, the import table, the  files the user picked with the workbook
    and  whether this version descends from one a finished bring-in made, with  the files picked now
    ("brought_in"; when not, "brought_in_reason"  says why for a line that went through a bring-in:
    the input files changed since, or it did not finish). Without them, or when the macros open
    nothing, there is nothing to do (PASS: the Status enum has no NOT_APPLICABLE)."""
    oi = (config or {}).get("outside_inputs")
    if not isinstance(oi, dict) or not oi.get("imports"):
        return finding("PASS", "Nothing to bring in: the macros open no outside file.", {"imports": False})
    rows = [r for r in oi.get("rows") or [] if isinstance(r, dict)]
    macro = oi.get("macro")
    missing = [r for r in rows if not r.get("file")]
    reason = None if oi.get("brought_in") else (str(oi.get("brought_in_reason") or "").strip().rstrip(".") or None)
    observed = {
        "imports": True, "macro": macro, "calls": oi.get("calls") or [], "files": len(rows), "matched": int(oi.get("matched") or 0),
        "missing": [_file_label(r) for r in missing], "extra": list(oi.get("extra") or []), "brought_in": bool(oi.get("brought_in")),
        "brought_in_reason": reason,
        "sites": [{"sheet": str(r["cell"]).rsplit("!", 1)[0], "cell": str(r["cell"]).rsplit("!", 1)[1]} for r in rows if "!" in str(r.get("cell") or "")],
    }
    location = observed["sites"][0] if observed["sites"] else None
    made = [str(x) for x in oi.get("stand_ins") or []]
    observed["stand_ins"] = made
    if oi.get("brought_in") and made:
        # brought in, but from files the assistant made from the model's own numbers: works, and the
        # Mind team can swap the CSVs -- yet the  actual files are what the model was built for
        return finding(
            "WARNING",
            f"The outside files are brought in, but {len(made)} of {len(rows) or len(made)} {'is a stand-in' if len(made) == 1 else 'are stand-ins'} the assistant made "
            f"from the numbers the model already holds ({', '.join(made[:6])}{' ...' if len(made) > 6 else ''}). "
            "Best results come with the actual input files: add them, then use Bring in the outside files again.",
            observed, location)
    if oi.get("brought_in"):
        return finding("PASS", "The outside files are brought in: Mind needs no macro for them.", observed, location)
    if not macro:
        return finding("WARNING", "The macros open outside files, but which macro brings them in is not clear: the data they brought in last stays in the workbook as it is.", observed, location)
    if not rows:
        return finding("WARNING", f"The macro {macro} opens outside files, but no cell names them: the data it brought in last stays in the workbook as it is, and Mind cannot update it.", observed, location)
    n = len(rows)
    picked = f"{observed['matched']} of {n} picked"
    if missing:
        names = [_file_label(r) for r in missing]
        picked += f", missing: {', '.join(names[:6])}" + (f" (+{len(names) - 6} more)" if len(names) > 6 else "")
    if reason:  # brought in before, and that no longer counts: the files changed since, or it stopped half-way
        return finding(
            "ERROR",
            f"The macro {macro} fills the workbook from {n} outside file{'' if n == 1 else 's'}, and Mind runs no macro: {reason} ({picked}). "
            "Use Bring in the outside files to bring them in again.",
            observed, location)
    return finding(
        "ERROR",
        f"The macro {macro} fills the workbook from {n} outside file{'' if n == 1 else 's'}, and Mind runs no macro: {picked}. Use Bring in the outside files.",
        observed, location)


# --- EXP ---------------------------------------------------------------------
def exp_001(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    exports = flagged_grids(analysis, "export")
    flag_only = [g for g in all_grids(analysis) if {"exportbyinstance", "exportbyloop", "exportnoinstancekeys", "exportnoloopkeys"} & set(g["flag_names"]) and "export" not in g["flag_names"]]
    if flag_only:
        return finding("WARNING", f"{len(flag_only)} grid(s) carry export option flags without /Export (they are not exported): {fmt_grids(flag_only)}", {"export_grids": [_grid_summary(g) for g in exports], "option_flags_without_export": [_grid_summary(g) for g in flag_only]}, location=grid_location(flag_only[0]))
    return finding("PASS", f"{len(exports)} /Export grid(s)." + ("" if exports else " The Export Manager will list no grids."), {"export_grids": [_grid_summary(g) for g in exports]})


def _export_settings(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    return flagged_grids(analysis, "exportsettings")


def exp_002(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _export_settings(analysis)
    if len(grids) > 1:
        return finding("ERROR", f"{len(grids)} /ExportSettings grids -- it must be unique: {fmt_grids(grids)}", {"exportsettings_grids": [_grid_summary(g) for g in grids]}, location=grid_location(grids[0]))
    return finding("PASS", f"{len(grids)} /ExportSettings grid(s)." + ("" if grids else " Global export settings apply."), {"exportsettings_grids": [_grid_summary(g) for g in grids]})


def exp_003(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _export_settings(analysis)
    if not grids:
        return finding("PASS", "No /ExportSettings grid to validate.", {"problems": []})
    g = grids[0]
    headers = header_names(g)
    lower = [h.lower() for h in headers]
    allowed = {c.lower(): c for c in EXPORT_COLUMNS}
    problems = []
    unknown = [h for h in headers if h.lower() not in allowed]
    if unknown:
        problems.append({"issue": f"unknown column(s) {unknown}; documented: {EXPORT_COLUMNS}"})
    if "gridname" not in lower:
        problems.append({"issue": "GridName column missing"})
    else:
        export_names = {(x.get("name") or "").strip().lower() for x in flagged_grids(analysis, "export")}
        gcol = lower.index("gridname")
        for r, row in enumerate(grid_rows(analysis, g)[1:], start=g["first_row"] + 1):
            name = row[gcol] if gcol < len(row) else None
            if is_blank(name):
                continue
            if str(name).strip().lower() not in export_names:
                problems.append({"row": r, "issue": f"GridName '{name}' is not an /Export grid"})
            for i, h in enumerate(lower):
                v = row[i] if i < len(row) else None
                if is_blank(v) or (isinstance(v, str) and v.startswith("=")):
                    continue
                if h in EXPORT_BOOLEAN_COLUMNS and as_bool(v) is None:
                    problems.append({"row": r, "issue": f"{headers[i]}='{v}' is not true/false/1/0"})
                if h == "fileextension" and str(v).strip().lower() not in (".csv", ".txt", "csv", "txt"):
                    problems.append({"row": r, "issue": f"FileExtension='{v}' must be .csv or .txt"})
    if problems:
        return finding("ERROR", f"{len(problems)} /ExportSettings problem(s): {problems[:6]}", {"headers": headers, "problems": problems}, location=grid_location(g))
    return finding("PASS", f"/ExportSettings grid at {g['sheet']}!{g['ref']}: {len(headers)} documented column(s), every GridName is an /Export grid, boolean/extension values valid.", {"headers": headers, "problems": []})


def exp_004(rule, analysis: dict[str, Any], config: dict[str, Any]) -> dict:
    grids = _export_settings(analysis)
    if not grids:
        return finding("PASS", "No /ExportSettings grid, so no FileNamePattern to validate.", {"patterns": []})
    g = grids[0]
    col = header_index(g, "FileNamePattern")
    if col is None:
        return finding("PASS", "/ExportSettings has no FileNamePattern column (global pattern applies).", {"patterns": []})
    bad, patterns = [], []
    for row in grid_rows(analysis, g)[1:]:
        pattern = row[col] if col < len(row) else None
        if is_blank(pattern):
            continue
        patterns.append(pattern)
        for tag in INPUT_PATTERN_TAG_RE.findall(str(pattern)):
            if tag.lower() not in EXPORT_PATTERN_TAGS:
                bad.append({"pattern": pattern, "unknown_tag": tag})
    if bad:
        return finding("WARNING", f"{len(bad)} FileNamePattern tag(s) are not documented ({{ModelName}} {{GridName}} {{InstanceName}} {{LoopsLabels}} {{yyyyMMdd}}): {bad[:6]}", {"patterns": patterns, "problems": bad}, location=grid_location(g))
    return finding("PASS", f"All {len(patterns)} export FileNamePattern value(s) use documented tags.", {"patterns": patterns})
