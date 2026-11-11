"""Next  quarter's Mind  input files -- from the client's RAW input files to the files Mind's Inputs
window takes.

Once the outside files are brought in (app/outside_inputs.py), the model in Mind reads each outside
file's data from  an `IN_<sheet> /Input` grid. Each quarter the Mind team gives Mind  new data for
those grids in the  project's **Inputs** window. Mind checks every file  against its grid: the same
name, the same header, the same columns.  A raw input file (the quarterly cash-flow  file itself, as
the client's tool opens it) does not fit: Mind refuses it ("The file 2026_Q1 - BBA_01.csv doesn't
have proper headers, there are 1340 columns while grid 'IN_BBA_01' of instance Original_Instance has
93 columns"). The grid holds only the block the model's import macro pulls out of the raw file,
often turned (items across, months down).

So the app does what the client's tool  does, then writes the grids out:

1. the picked files are sorted (`classify_files`): a file already shaped like a grid (its header is
   the grid's header) goes into the zip as it is, under the grid's name; any other file is a raw
   input file, matched to the model's import table like the files picked with the model (a file
   named like a grid but shaped otherwise is read as a raw file, and said so). A grid file must be
   UTF-8 text; a comma-separated `.txt`  goes in as `<grid>.csv` (Mind  reads a `.txt` as TAB-separated);
2. the model's  OWN import macro  runs on  a copy of the version the outside files were first brought
   into, with the new raw files exactly as picked (`outside_inputs.run_import`: an Excel of the app's
   own, message boxes answered). The macro opens  each file its own way, as at the bring-in (its
   OpenText settings,  its checks of the name): the  values are exactly as the client's tool  puts them.
   A raw .csv is not touched first: Excel keeps 15  significant digits of its numbers, as in  the
   client's tool, and the bring-in from the same file gives the same values;
3. each grid's block is read from that copy (Excel, values only) and written with the grid's own
   header, under the grid's own name (`IN_BBA_01.csv`; TAB `.txt` when a value holds a comma, as at
   the bring-in);
4. one zip: `README.md` (which file goes into which grid, and why raw files cannot  go into Mind) and
   the folder `Mind inputs` with the files. No new model upload: Mind keeps the model, the files
   only replace the grids' data. A grid without a new file is not in the zip: Mind keeps the data it
   has.
"""
from __future__ import annotations

import codecs
import csv
import io
import math
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any, Callable, Iterable

from . import outside_inputs
from .outside_inputs import normalize_name, rect_ref, write_input_csv

ProgressFn = Callable[[str], None]

ZIP_NAME = "Next quarter Mind inputs.zip"
FOLDER_IN_ZIP = "Mind inputs"
README_NAME = "README.md"
GRID_FILE_SUFFIXES = (".csv", ".txt")

RAW_FILES_LINE = ("Raw input files cannot go into Mind directly. Mind takes one file per input grid, "
                  "shaped like the grid: the same name, the same header, the same columns.")


# ---  the grids  of the model ----------------------------------------------------------------------
def grid_name_of(title: Any) -> str | None:
    """'#IN_BBA_01  /Input' ->  'IN_BBA_01' (None for anything else)."""
    m = re.fullmatch(r"\s*#(.+?)\s+/Input\s*", str(title or ""), re.IGNORECASE)
    return m.group(1).strip() if m else None


def read_first_row(path: Path, max_bytes: int = 16 << 20) -> dict[str, Any]:
    """The  first line  of a text file as cells, the way Mind counts a file's columns: TAB-separated in a
    `.txt` (Mind reads a .txt as TAB-separated, even one with commas), comma-separated in any other
    file (quotes  as in any CSV). UTF-8 (a BOM is dropped); a file  that is not UTF-8 is read  as well as
    it can be. -> {"cells":  [...], "separator": "TAB" | "comma"}"""
    path = Path(path)
    with open(path, "rb") as fh:
        line = fh.readline(max_bytes).rstrip(b"\r\n")
    text = line.decode("utf-8-sig", errors="replace")
    if path.suffix.lower() == ".txt":
        return {"cells": text.split("\t"), "separator": "TAB"}
    return {"cells": _comma_cells(text), "separator": "comma"}


def _comma_cells(text: str) -> list[str]:
    try:
        return next(csv.reader(io.StringIO(text)), [])
    except csv.Error:
        return text.split(",")


def is_utf8(path: Path, chunk: int = 1 << 20) -> bool:
    """Whether  the whole  file is UTF-8 text (the files Mind's Inputs window gets from the app are)."""
    dec = codecs.getincrementaldecoder("utf-8")()
    try:
        with open(path, "rb") as fh:
            while True:
                block = fh.read(chunk)
                if not block:
                    dec.decode(b"", final=True)
                    return True
                dec.decode(block)
    except UnicodeDecodeError:
        return False


def _holds_any(path: Path, needles: tuple[bytes, ...], chunk: int = 1 << 20) -> bool:
    with open(path, "rb") as fh:
        while True:
            block = fh.read(chunk)
            if not block:
                return False
            if any(n in block for n in needles):
                return True


def grids_of(change_log: Iterable[dict[str, Any]], csv_dir: Path | str | None) -> list[dict[str, Any]]:
    """The  Mind input  grids a bring-in made, from the change log of its "mind_inputs" version (ops
    `make_mind_input`: the area's sheet and range, `after` = the grid's title) and the files it wrote
    (`csv_dir`: one per grid,  named like it; their first row is the grid's header, exactly  as Mind
    has it). -> [{name, sheet,  ref, r1, c1, r2, c2, rows, cols,  header, file}]"""
    folder = Path(csv_dir) if csv_dir else None
    files = {normalize_name(p.stem): p for p in (folder.iterdir() if folder and folder.is_dir() else [])
             if p.is_file() and p.suffix.lower() in GRID_FILE_SUFFIXES}
    out: list[dict[str, Any]] = []
    for op in change_log or []:
        if op.get("op") != "make_mind_input":
            continue
        name = grid_name_of(op.get("after"))
        rect = outside_inputs._addr_rect(str(op.get("cell") or ""))
        if not name or rect is None or not op.get("sheet"):
            continue
        r1, c1, r2, c2 = rect
        f = files.get(normalize_name(name))
        header = read_first_row(f)["cells"] if f else None
        out.append({"name": name, "sheet": str(op["sheet"]), "ref": rect_ref(rect), "r1": r1, "c1": c1, "r2": r2, "c2": c2,
                    "rows": r2 - r1 + 1, "cols": c2 - c1 + 1, "header": header, "file": str(f) if f else None})
    return out


# --- what the user picked ------------------------------------------------------------------------
def _same_header(a: list[str], b: list[str] | None) -> bool:
    return b is not None and len(a) == len(b) and all(normalize_name(x) == normalize_name(y) for x, y in zip(a, b))


def classify_files(paths: Iterable[Path], grids: list[dict[str, Any]], table: list[dict[str, Any]]) -> dict[str, Any]:
    """Sort  the picked  files.

    -> {"grid_files": [{file, path, grid, renamed, as_csv?}] files already shaped like a grid (its
        header; as_csv: a comma-separated .txt, which goes into the zip as <grid>.csv),
        "rows": match_files rows for the raw files  (the import table, each row with its file or None),
        "raw": [names  of the raw files given to  the import table],
        "notes": [{file, why}] a file  named like a grid but shaped otherwise (read as a raw  file),
        "unused": [{file, why}] files nothing takes,
        "ambiguous": [names]  raw files that fit several rows of the import table}"""
    by_name = {normalize_name(g["name"]): g for g in grids}
    grid_files: list[dict[str, Any]] = []
    raw: list[Path] = []
    notes: list[dict[str, Any]] = []
    unused: list[dict[str, Any]] = []
    for p in paths:
        p = Path(p)
        if p.suffix.lower() in GRID_FILE_SUFFIXES:
            try:
                first = read_first_row(p)["cells"]
            except OSError as exc:
                unused.append({"file": p.name, "why": f"it could not be read ({exc.strerror or exc})"})
                continue
            fits = [g for g in grids if _same_header(first, g.get("header"))]
            as_csv = False
            if not fits and p.suffix.lower() == ".txt" and len(first) == 1 and "," in first[0]:
                # a .txt separated by commas: Mind would read one column (it reads a .txt as TAB-separated)
                commas = _comma_cells(first[0])
                by_comma = [g for g in grids if _same_header(commas, g.get("header"))]
                if by_comma:
                    fits, first, as_csv = by_comma, commas, True
            named = by_name.get(normalize_name(p.stem))
            if named is not None and named in fits:
                fits = [named]
            try:
                if len(fits) == 1:
                    g = fits[0]
                    if not is_utf8(p):
                        unused.append({"file": p.name, "why": not_utf8_line(p.name, g["name"])})
                    elif as_csv and _holds_any(p, (b'"', b"\t")):
                        unused.append({"file": p.name, "why": (f"{p.name} is separated by commas, but Mind reads a .txt file as TAB-separated, "
                                                                "and this one holds quotes or TABs. Save it as .csv, or with TABs, and pick it again.")})
                    else:
                        gf = {"file": p.name, "path": str(p), "grid": g["name"], "renamed": normalize_name(p.stem) != normalize_name(g["name"])}
                        if as_csv:
                            gf["as_csv"] = True
                        grid_files.append(gf)
                    continue
                if named is not None and len(first) == named["cols"] and not is_utf8(p):
                    # its header did not read as UTF-8: shaped like the grid, saved in another encoding
                    unused.append({"file": p.name, "why": not_utf8_line(p.name, named["name"])})
                    continue
            except OSError as exc:
                unused.append({"file": p.name, "why": f"it could not be read ({exc.strerror or exc})"})
                continue
            if named is not None:
                notes.append({"file": p.name, "why": raw_file_line(p.name, len(first), named["name"], named["cols"])})
        raw.append(p)
    match = outside_inputs.match_files(table, [p.name for p in raw])
    for f in match["extra"]:
        unused.append({"file": f, "why": "no row of the model's import table names it, and it is not shaped like a Mind input grid"})
    seen: set[str] = set()
    for gf in list(grid_files):  # one file per grid: two files for the same grid are both refused
        if gf["grid"] in seen:
            continue
        twins = [x for x in grid_files if x["grid"] == gf["grid"]]
        if len(twins) > 1:
            for x in twins:
                grid_files.remove(x)
                unused.append({"file": x["file"], "why": f"{len(twins)} files are shaped like the grid {gf['grid']}: pick one"})
        seen.add(gf["grid"])
    return {"grid_files": grid_files, "rows": match["rows"], "raw": [r["file"] for r in match["rows"] if r.get("file")],
            "notes": notes, "unused": unused, "ambiguous": list(match.get("ambiguous") or [])}


def not_utf8_line(file: str, grid: str) -> str:
    """One  plain sentence  for a grid's file saved in another encoding than UTF-8."""
    return (f"{file} is shaped like the Mind input grid {grid}, but it is not UTF-8 text, so the app cannot be sure Mind reads it right. "
            "Save it again as UTF-8 (in Excel: \"CSV UTF-8\") and pick it again.")


def raw_file_line(file: str, cols: int, grid: str, grid_cols: int) -> str:
    """One  plain sentence  for a raw file given where a grid's file is expected."""
    return (f"{file} has {cols:,} column{'' if cols == 1 else 's'}; the Mind input grid {grid} has {grid_cols:,}. "
            "It looks like a raw input file, so the app reads it with the model's import macro.")


_HEADER_ERROR = re.compile(
    r"The file (?P<file>.+?) doesn'?t have proper headers,\s*there are (?P<n>\d+) columns while grid '(?P<grid>[^']+)'"
    r"(?: of instance (?P<inst>\S+))? has (?P<m>\d+) columns", re.IGNORECASE)


def explain_mind_input_error(text: str) -> str | None:
    """What  to do  when Mind's Inputs window refuses a file because its columns are not the grid's
    (a raw input file given in a grid's place). None when the text holds no such refusal."""
    m = _HEADER_ERROR.search(str(text or ""))
    if not m:
        return None
    file, n, grid, cols = m.group("file").strip(), int(m.group("n")), m.group("grid"), int(m.group("m"))
    if n > cols:
        what = "It looks like a raw input file. "
    else:
        what = "Its columns are not the grid's. "
    return (f"Mind refused {file}: it has {n:,} columns, the grid {grid} has {cols:,}. {what}{RAW_FILES_LINE} "
            "In Mind Ready, press \"Make next quarter's Mind input files\" on the Outside files card and pick the raw files: "
            "it gives one file per grid to put in Mind's Inputs window.")


# --- which  grid each new file fills ------------------------------------------------------------
def _rect(d: dict[str, Any]) -> tuple[int, int, int, int]:
    return int(d["r1"]), int(d["c1"]), int(d["r2"]), int(d["c2"])


def grids_filled(report: dict[str, Any], grids: list[dict[str, Any]]) -> dict[str, Any]:
    """Which  grid each  file the macro opened filled (an area on the grid's sheet that meets the grid's
    range). An area larger than its grid cannot go into Mind (Mind's grid keeps the model's shape):
    that  grid is refused, with why. -> {"filled": {grid: file}, "too_big": {grid:  why}, "outside": [{file, sheet, ref}]}"""
    filled: dict[str, str] = {}
    too_big: dict[str, str] = {}
    outside: list[dict[str, Any]] = []
    for e in report.get("files") or []:
        if not e.get("file") or not e.get("opened"):
            continue
        for a in e.get("areas") or []:
            ar = _rect(a)
            hit = False
            for g in grids:
                if normalize_name(g["sheet"]) != normalize_name(a["sheet"]) or not outside_inputs._intersects(ar, _rect(g)):
                    continue
                hit = True
                if outside_inputs._contains(_rect(g), ar):
                    filled[g["name"]] = e["file"]
                else:
                    too_big[g["name"]] = (f"{e['file']} filled {a['sheet']}!{a['ref']}, more than the Mind input grid {g['name']} holds "
                                          f"({g['sheet']}!{g['ref']}): Mind's grid keeps the model's shape, so this file cannot replace it")
            if not hit:
                outside.append({"file": e["file"], "sheet": a["sheet"], "ref": a["ref"]})
    for g in too_big:
        filled.pop(g, None)
    return {"filled": filled, "too_big": too_big, "outside": outside}


# --- writing  the files -----------------------------------------------------------------------------
def export_grids(workbook: Path, grids: list[dict[str, Any]], out_dir: Path, progress: ProgressFn | None = None) -> list[dict[str, Any]]:
    """Each  grid's block  of `workbook` (values exactly as Excel holds them: Value2 through an Excel of
    the app's own, never the user's) written with the grid's header under the grid's name.
    -> [{grid, path, rows, cols,  separator}]"""
    import pythoncom

    say = progress or (lambda _m: None)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    made: list[dict[str, Any]] = []
    pythoncom.CoInitialize()
    xl = wb = None
    pid = None
    try:
        xl, pid = outside_inputs._new_excel(macros=False)
        try:
            xl.EnableEvents = False
        except Exception:
            pass
        say("Reading the new data of each Mind input grid")
        wb = xl.Workbooks.Open(str(workbook), UpdateLinks=0, ReadOnly=True)
        for g in grids:
            data = wb.Worksheets(g["sheet"]).Range(g["ref"]).Value2
            rows = [list(r) for r in (data if isinstance(data, tuple) else ((data,),))]
            header = list(g.get("header") or [f"C{i + 1}" for i in range(g["cols"])])
            info = write_input_csv(out_dir / g["name"], header, rows)
            made.append({"grid": g["name"], "path": info["path"], "rows": info["rows"], "cols": info["cols"], "separator": info["separator"]})
        wb.Close(SaveChanges=False)
        wb = None
        return made
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        wb = None
        outside_inputs._close_excel(xl, None, pid)
        xl = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# --- numbers past Excel's 15 digits in a raw CSV ------------------------------------------------------
_DECIMAL_RE = re.compile(r"\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*")


def _significant_digits(text: str) -> int:
    mant = re.split(r"[eE]", text.strip())[0].lstrip("+-").replace(".", "").lstrip("0")
    return len(mant)


def long_numbers(csv_path: Path) -> int:
    """How  many numbers  of a CSV are written with more than 15 significant digits. Excel keeps 15 when
    it opens a CSV (144686640.71679688 comes in as 144686640.716796), in the client's tool as here: the
    app says  so, it does not change what the macro reads."""
    n = 0
    with open(csv_path, encoding="utf-8-sig", errors="replace", newline="") as fh:
        for row in csv.reader(fh):
            n += sum(1 for t in row if t and len(t) > 15 and _DECIMAL_RE.fullmatch(t) and _significant_digits(t) > 15)
    return n


def _read_rows(path: Path) -> list[list[str]]:
    sep = "\t" if Path(path).suffix.lower() == ".txt" else ","
    with open(path, encoding="utf-8-sig", errors="replace", newline="") as fh:
        return [ln.rstrip("\r\n").split(sep) for ln in fh]


def changed_cells(old: Path | str | None, new: Path | str) -> int | None:
    """How  many data  cells differ between two files of the same grid (the header is not counted; a
    number written two ways counts as the same, and so does a number that differs only past  its 15th
    significant digit: Excel keeps 15 digits of a number when  it opens a CSV, so the  same data as .csv
    and as .xlsx can  differ there). None when there is no old file to compare  with."""
    if not old or not Path(old).is_file():
        return None
    a, b = _read_rows(Path(old))[1:], _read_rows(Path(new))[1:]
    n = 0
    for i in range(max(len(a), len(b))):
        ra = a[i] if i < len(a) else []
        rb = b[i] if i < len(b) else []
        for j in range(max(len(ra), len(rb))):
            x = ra[j] if j < len(ra) else ""
            y = rb[j] if j < len(rb) else ""
            if x == y:
                continue
            try:
                fx, fy = float(x), float(y)
                if fx == fy or math.isclose(fx, fy, rel_tol=2e-14):
                    continue
            except ValueError:
                pass
            n += 1
    return n


def readme(result: dict[str, Any]) -> str:
    """The  zip's README:  which file goes into which grid, how to put them in Mind, why raw files cannot."""
    lines = [
        "# Next quarter's Mind input files",
        "",
        f"Made by Mind Ready from the model {result.get('model') or ''} on {result.get('made_at') or ''}.",
        "",
        "## What to do",
        "",
        "1. Unzip this file.",
        "2. Open the project in Mind and click **Inputs** (left of Run).",
        f"3. Click **Browse** and pick the folder `{FOLDER_IN_ZIP}` (or drag the files in it onto the window).",
        "4. Mind matches each file to the grid of the same name. Wait until every file shows SUCCESS.",
        "5. Close the window and click **Run**. Nothing to save, and no new model upload.",
        "",
        "## Which file goes into which grid",
        "",
        "| File | Mind input grid | Made from |",
        "| --- | --- | --- |",
    ]
    for f in result.get("files") or []:
        src = f.get("from") or ""
        how = "given as it is" if f.get("source") == "as_given" else "the model's import macro"
        lines.append(f"| `{f['file']}` | {f['grid']} | {src} ({how}) |")
    if result.get("not_in_zip"):
        lines += ["", "## Not in this zip", ""]
        lines += [f"- {x['grid']}: {x['why']}" for x in result["not_in_zip"]]
    lines += [
        "",
        "## Why not the raw files",
        "",
        RAW_FILES_LINE + " A raw input file holds much more (all its items and months, its own header rows),",
        "so Mind refuses it: \"doesn't have proper headers, there are 1340 columns while grid ... has 93 columns\".",
        "Each file here holds only the block the model's import macro takes from the raw file, laid out as the grid.",
        "",
    ]
    return "\n".join(lines)


def write_zip(zip_path: Path, files: list[dict[str, Any]], readme_text: str) -> Path:
    """README.md  at the  root, the grid files in the folder `Mind inputs` (the folder to pick in Mind)."""
    zip_path = Path(zip_path)
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(README_NAME, readme_text.replace("\n", "\r\n").encode("utf-8"))
        for f in files:
            zf.write(f["path"], arcname=f"{FOLDER_IN_ZIP}/{f['file']}")
    return zip_path


def make_next_quarter(base_workbook: Path, table: list[dict[str, Any]], picked: list[Path], grids: list[dict[str, Any]], macro: str,
                      work_dir: Path, messages: list[dict[str, Any]] | None = None, progress: ProgressFn | None = None,
                      model_name: str | None = None, made_at: str | None = None) -> dict[str, Any]:
    """The  whole job  (see the module text). `base_workbook`: the version the outside files were first
    brought into (it still has the import macro and its own areas); `table`: its import table;
    `grids`: grids_of(...) of the line's  "mind_inputs" version.

    -> {"ok", "files": [{file, grid, from, source: macro | as_given,  rows, cols, separator, changed, path}],
        "not_in_zip": [{grid,  why}], "notes": [str], "unused": [{file, why}], "import":  run_import's report
        (None when no raw file), "zip": path, "readme": text, "summary",  "error"}"""
    say = progress or (lambda _m: None)
    work_dir = Path(work_dir)
    out_dir = work_dir / "files"
    if out_dir.exists():
        shutil.rmtree(out_dir, ignore_errors=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    sorted_ = classify_files(picked, grids, table)
    by_grid = {g["name"]: g for g in grids}
    result: dict[str, Any] = {"ok": False, "files": [], "not_in_zip": [], "notes": [n["why"] for n in sorted_["notes"]],
                              "unused": sorted_["unused"], "import": None, "zip": None, "readme": None, "summary": "", "error": None,
                              "model": model_name, "made_at": made_at}
    if sorted_["ambiguous"]:
        result["notes"].append("Fits several rows of the model's import table, so not used: " + ", ".join(sorted_["ambiguous"]) + ".")
    files: dict[str, dict[str, Any]] = {}
    raw_rows = [r for r in sorted_["rows"] if r.get("file")]
    if raw_rows:
        paths = {p.name: Path(p) for p in picked}
        # the files exactly as picked: the macro  opens each its own way (as at the bring-in)
        chosen = {f"{r['sheet']}!{r['cell']}": paths[r["file"]] for r in raw_rows}
        for p in chosen.values():
            if p.suffix.lower() == ".csv":
                n = long_numbers(p)
                if n:
                    result["notes"].append(f"{p.name} has {n:,} number{'' if n == 1 else 's'} with more than 15 digits. Excel keeps 15 when it opens "
                                           "a CSV, as in the client's tool, so the Mind input files hold them that way.")
        say(f"Running the model's import macro {macro} with {len(chosen)} new file{'' if len(chosen) == 1 else 's'}")
        out_wb = work_dir / "import" / Path(base_workbook).name
        report = outside_inputs.run_import(Path(base_workbook), sorted_["rows"], chosen, macro, out_wb, work_dir, progress=say, messages=messages or [])
        result["import"] = report
        ok = report.get("ok") if "ok" in report else (report.get("saved") and not report.get("stopped_early"))
        if not ok or not out_wb.is_file():
            said = [str(d.get("message_repaired") or d.get("message") or "").strip() for d in report.get("dialogs") or []]
            said = [x for x in said if x]
            not_opened = [str(e["file"]) for e in report.get("files") or [] if e.get("file") and not e.get("opened")]
            text = "The model's import macro did not take the new files"
            if said:
                text += f': "{said[0][:200]}"'
            elif report.get("error"):
                text += f": {str(report['error'])[:200]}"
            if not_opened:
                text += f". Not read: {', '.join(not_opened[:6])}"
            result["error"] = text.rstrip(".") + ". No file was made: check the input files, then try again."
            return result
        where = grids_filled(report, grids)
        for g, why in where["too_big"].items():
            result["not_in_zip"].append({"grid": g, "why": why})
        for o in where["outside"]:
            result["notes"].append(f"{o['file']} also filled {o['sheet']}!{o['ref']}, which stays inside the model in Mind (no Mind input grid holds it).")
        fill = [by_grid[g] for g in where["filled"]]
        if fill:
            made = export_grids(out_wb, fill, out_dir, progress=say)
            for m in made:
                files[m["grid"]] = {"file": Path(m["path"]).name, "grid": m["grid"], "from": where["filled"][m["grid"]], "source": "macro",
                                    "rows": m["rows"], "cols": m["cols"], "separator": m["separator"], "path": m["path"]}
    for gf in sorted_["grid_files"]:
        if gf["grid"] in files:
            result["notes"].append(f"{gf['file']} was left out: {files[gf['grid']]['from']} fills the same grid {gf['grid']}.")
            continue
        src = Path(gf["path"])
        dest = out_dir / (gf["grid"] + (".csv" if gf.get("as_csv") else src.suffix.lower()))
        shutil.copyfile(src, dest)
        g = by_grid[gf["grid"]]
        files[gf["grid"]] = {"file": dest.name, "grid": gf["grid"], "from": gf["file"], "source": "as_given", "rows": max(0, len(_read_rows(dest)) - 1),
                             "cols": g["cols"], "separator": read_first_row(dest)["separator"], "path": str(dest)}
        if gf.get("as_csv"):
            result["notes"].append(f"{gf['file']} is separated by commas, and Mind reads a .txt file as TAB-separated: it goes into the zip as {dest.name}.")
        elif gf["renamed"]:
            result["notes"].append(f"{gf['file']} is shaped like the grid {gf['grid']}: it goes into the zip as {dest.name}.")
    for g in grids:
        f = files.get(g["name"])
        if f is not None:
            f["changed"] = changed_cells(g.get("file"), f["path"])
        elif not any(x["grid"] == g["name"] for x in result["not_in_zip"]):
            result["not_in_zip"].append({"grid": g["name"], "why": "no new file for it: Mind keeps the data it has"})
    result["files"] = [files[g["name"]] for g in grids if g["name"] in files]
    if not result["files"]:
        why = "; ".join(u["why"] for u in result["unused"][:2]) or "none of the picked files fills a Mind input grid"
        result["error"] = f"No Mind input file could be made: {why}."
        return result
    result["readme"] = readme(result)
    result["zip"] = str(write_zip(work_dir / ZIP_NAME, result["files"], result["readme"]))
    result["ok"] = True
    n = len(result["files"])
    result["summary"] = (f"{n} Mind input file{'' if n == 1 else 's'} made ({', '.join(f['file'] for f in result['files'][:8])}"
                         f"{', ...' if n > 8 else ''}): put {'it' if n == 1 else 'them'} in Mind's Inputs window, then Run.")
    return result
