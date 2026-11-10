"""Next quarter's Mind input files (app/next_quarter_inputs.py): the grids a bring-in made, the picked
files sorted (raw input files for the model's import macro, files already shaped like a grid, files
nothing takes), which grid each new file fills, the files written with the grid's own header and
name, the zip (README.md + the folder "Mind inputs"), and Mind's refusal of a raw file in words.
Excel steps replaced here, except the one marked needs_excel."""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest

from app import next_quarter_inputs as nq
from app import outside_inputs
from app.excel_com import com_available

needs_excel = pytest.mark.skipif(not com_available(), reason="needs a real Excel")

# Joseph's case (10 November 2026): the raw quarterly file given to Mind's Inputs window
MIND_REFUSAL = ("Failed: Input control error: The file 2026_Q1 - BBA_01.csv doesn't have proper headers, there are 1340 columns "
                "while grid 'IN_BBA_01' of instance Original_Instance has 93 columns.")

TABLE = [  # the model's import table (outside_inputs.find_import_table rows)
    {"sheet": "Ctrl_01", "cell": "H35", "row": 35, "keys": ["BBA_01"], "expected": "2026Q1_2026_BBA_01.xlsm", "target_sheet": "BBA_01"},
    {"sheet": "Ctrl_01", "cell": "H36", "row": 36, "keys": ["VFA_01"], "expected": "2026Q1_2026_VFA_01.xlsm", "target_sheet": "VFA_01"},
]
LOG = [  # the change log of the "mind_inputs" version (server._mind_inputs_log)
    {"op": "make_mind_input", "sheet": "BBA_01", "cell": "AC3:AE5", "after": "#IN_BBA_01 /Input"},
    {"op": "make_mind_input", "sheet": "VFA_01", "cell": "B3:C4", "after": "#IN_VFA_01 /Input"},
    {"op": "set_formula", "sheet": "Calc", "cell": "A1", "after": "=1"},
]


def _bring_in_files(folder: Path) -> Path:
    """The files the bring-in wrote: a CSV, and a TAB .txt (a header holds a comma)."""
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "IN_BBA_01.csv").write_text("Lapse,Death,Surrender\n1,2,3\n4,5,6\n7,8,9\n", encoding="utf-8")
    (folder / "IN_VFA_01.txt").write_text("Accrued interest, end\tFund\n10\t20\n30\t40\n", encoding="utf-8")
    return folder


@pytest.fixture
def grids(tmp_path):
    return nq.grids_of(LOG, _bring_in_files(tmp_path / "bring_in_csv"))


def _raw_csv(path: Path, cols: int = 1340, lines: int = 5) -> Path:
    """A raw quarterly file saved as CSV: as wide as the client's export, with a BOM."""
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = [",".join([""] * cols)] + [",".join(["", "VariableName", "Description"] + ["0"] * (cols - 3))] * (lines - 1)
    path.write_bytes(b"\xef\xbb\xbf" + "\r\n".join(rows).encode("utf-8"))
    return path


# --- the grids of the model ----------------------------------------------------------------------
def test_the_grids_come_from_the_bring_in_with_the_header_mind_has(grids, tmp_path):
    assert [g["name"] for g in grids] == ["IN_BBA_01", "IN_VFA_01"]
    bba, vfa = grids
    assert (bba["sheet"], bba["ref"], bba["rows"], bba["cols"], bba["r1"], bba["c1"]) == ("BBA_01", "AC3:AE5", 3, 3, 3, 29)
    assert bba["header"] == ["Lapse", "Death", "Surrender"] and Path(bba["file"]).name == "IN_BBA_01.csv"
    assert vfa["header"] == ["Accrued interest, end", "Fund"] and Path(vfa["file"]).name == "IN_VFA_01.txt"
    # a grid whose file is gone: no header (the files are then written with C1..Cn)
    assert nq.grids_of(LOG[:1], tmp_path / "nowhere")[0]["header"] is None
    assert nq.grid_name_of("#IN_BBA_01 /Input") == "IN_BBA_01" and nq.grid_name_of("#Totals") is None and nq.grid_name_of(None) is None


def test_the_first_row_is_read_as_mind_counts_columns(tmp_path):
    raw = _raw_csv(tmp_path / "2026_Q1 - BBA_01.csv")
    assert len(nq.read_first_row(raw)["cells"]) == 1340
    quoted = tmp_path / "q.csv"
    quoted.write_text('"a,b",c\n1,2\n', encoding="utf-8")
    assert nq.read_first_row(quoted) == {"cells": ["a,b", "c"], "separator": "comma"}
    tab = tmp_path / "t.txt"
    tab.write_text("x, y\tz\n", encoding="utf-8")
    assert nq.read_first_row(tab) == {"cells": ["x, y", "z"], "separator": "TAB"}


# --- sorting the picked files ------------------------------------------------------------------------
def test_raw_files_go_to_the_import_table_and_grid_files_go_in_as_they_are(grids, tmp_path):
    raw = _raw_csv(tmp_path / "in" / "2026_Q1 - BBA_01.csv")
    grid_file = tmp_path / "in" / "vfa next quarter.txt"  # shaped like IN_VFA_01 (its header), named otherwise
    grid_file.write_text("Accrued interest, end\tFund\n11\t21\n31\t41\n", encoding="utf-8")
    other = tmp_path / "in" / "notes.csv"
    other.write_text("a,b\n1,2\n", encoding="utf-8")
    out = nq.classify_files([raw, grid_file, other], grids, TABLE)
    assert out["raw"] == ["2026_Q1 - BBA_01.csv"]
    assert [r.get("file") for r in out["rows"]] == ["2026_Q1 - BBA_01.csv", None]
    assert out["grid_files"] == [{"file": "vfa next quarter.txt", "path": str(grid_file), "grid": "IN_VFA_01", "renamed": True}]
    assert out["unused"] == [{"file": "notes.csv", "why": "no row of the model's import table names it, and it is not shaped like a Mind input grid"}]
    assert out["notes"] == [] and out["ambiguous"] == []


def test_a_raw_file_named_like_a_grid_is_recognised_and_read_with_the_macro(grids, tmp_path):
    raw = _raw_csv(tmp_path / "IN_BBA_01.csv")  # the raw file renamed after the grid: Mind would refuse it
    out = nq.classify_files([raw], grids, TABLE)
    assert out["grid_files"] == [] and out["raw"] == ["IN_BBA_01.csv"]
    assert out["notes"] == [{"file": "IN_BBA_01.csv", "why": (
        "IN_BBA_01.csv has 1,340 columns; the Mind input grid IN_BBA_01 has 3. "
        "It looks like a raw input file, so the app reads it with the model's import macro.")}]


def test_two_files_shaped_like_the_same_grid_are_both_refused(grids, tmp_path):
    a, b = tmp_path / "a.csv", tmp_path / "IN_BBA_01.csv"
    for p in (a, b):
        p.write_text("Lapse,Death,Surrender\n1,2,3\n", encoding="utf-8")
    out = nq.classify_files([a, b], grids, TABLE)
    assert out["grid_files"] == [] and out["raw"] == []
    assert [u["file"] for u in out["unused"]] == ["a.csv", "IN_BBA_01.csv"]
    assert all(u["why"] == "2 files are shaped like the grid IN_BBA_01: pick one" for u in out["unused"])


def test_mind_refusing_a_raw_file_is_said_with_what_to_do():
    text = nq.explain_mind_input_error("IN_BBA_01  " + MIND_REFUSAL)
    assert text == (
        "Mind refused 2026_Q1 - BBA_01.csv: it has 1,340 columns, the grid IN_BBA_01 has 93. It looks like a raw input file. "
        "Raw input files cannot go into Mind directly. Mind takes one file per input grid, shaped like the grid: the same name, "
        "the same header, the same columns. In Mind Ready, press \"Make next quarter's Mind input files\" on the Outside files card "
        "and pick the raw files: it gives one file per grid to put in Mind's Inputs window.")
    fewer = nq.explain_mind_input_error("The file IN_BBA_01.csv doesn't have proper headers, there are 90 columns while grid 'IN_BBA_01' has 93 columns")
    assert fewer.startswith("Mind refused IN_BBA_01.csv: it has 90 columns, the grid IN_BBA_01 has 93. Its columns are not the grid's.")
    assert nq.explain_mind_input_error("IN_BBA_01 SUCCESS") is None and nq.explain_mind_input_error(None) is None


# --- which grid each new file fills ----------------------------------------------------------------
def _entry(file, sheet, r1, c1, r2, c2, opened=True):
    return {"file": file, "opened": opened, "areas": [{"sheet": sheet, "ref": outside_inputs.rect_ref((r1, c1, r2, c2)), "r1": r1, "c1": c1, "r2": r2, "c2": c2}]}


def test_each_file_fills_the_grid_its_area_lies_in_and_a_bigger_area_is_refused(grids):
    report = {"files": [_entry("b.xlsx", "BBA_01", 3, 29, 5, 31), _entry("v.xlsx", "VFA_01", 3, 2, 9, 3), _entry("x.xlsx", "Other", 1, 1, 2, 2),
                        {"file": None, "opened": False, "areas": []}]}
    out = nq.grids_filled(report, grids)
    assert out["filled"] == {"IN_BBA_01": "b.xlsx"}
    assert out["too_big"] == {"IN_VFA_01": ("v.xlsx filled VFA_01!B3:C9, more than the Mind input grid IN_VFA_01 holds (VFA_01!B3:C4): "
                                            "Mind's grid keeps the model's shape, so this file cannot replace it")}
    assert out["outside"] == [{"file": "x.xlsx", "sheet": "Other", "ref": "A1:B2"}]
    # a file with fewer items fills part of its grid: still its grid
    assert nq.grids_filled({"files": [_entry("b.xlsx", "bba_01", 3, 29, 4, 30)]}, grids)["filled"] == {"IN_BBA_01": "b.xlsx"}


def test_changed_cells_counts_data_cells_only_and_numbers_by_value(tmp_path):
    old, new = tmp_path / "old.csv", tmp_path / "new.csv"
    old.write_text("A,B\n1,2\n3,4\n", encoding="utf-8")
    new.write_text("A2,B\n1.0,2\n3,5\n6,\n", encoding="utf-8")
    assert nq.changed_cells(old, new) == 2  # 4 -> 5, and the new row's 6 (the header and 1 vs 1.0 do not count)
    assert nq.changed_cells(None, new) is None and nq.changed_cells(tmp_path / "gone.csv", new) is None


# --- the whole job, Excel replaced -------------------------------------------------------------------
@pytest.fixture
def fake_excel(monkeypatch):
    calls: dict[str, list] = {"import": [], "export": []}

    def run_import(workbook, rows, files, macro, out_path, work_dir, timeout_s=900, progress=None, messages=None, restore_rows=None):
        calls["import"].append({"workbook": Path(workbook), "files": {k: Path(v).name for k, v in files.items()}, "macro": macro, "messages": messages})
        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(workbook, out_path)
        entries = []
        for r in rows:
            key = f"{r['sheet']}!{r['cell']}"
            p = files.get(key)
            area = _entry(p.name, "BBA_01", 3, 29, 5, 31)["areas"] if p and key.endswith("H35") else []
            entries.append({"cell": key, "file": p.name if p else None, "opened": bool(p), "areas": area})
        return {"macro": macro, "seconds": 1.0, "files": entries, "dialogs": [], "stopped_early": False, "error": None, "saved": str(out_path), "ok": True}

    def export_grids(workbook, grids, out_dir, progress=None):
        calls["export"].append({"workbook": Path(workbook), "grids": [g["name"] for g in grids]})
        made = []
        for g in grids:
            info = outside_inputs.write_input_csv(Path(out_dir) / g["name"], g["header"], [[1, 2, 3], [4, 5, 60], [7, 8, 9]])
            made.append({"grid": g["name"], "path": info["path"], "rows": info["rows"], "cols": info["cols"], "separator": info["separator"]})
        return made

    def exact_csv_workbook(csv_path, out_dir, progress=None):
        calls.setdefault("exact", []).append(Path(csv_path).name)
        out = Path(out_dir) / (Path(csv_path).stem + ".xlsx")
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(csv_path, out)
        return {"path": str(out), "fixed": 7}

    monkeypatch.setattr(outside_inputs, "run_import", run_import)
    monkeypatch.setattr(nq, "export_grids", export_grids)
    monkeypatch.setattr(nq, "exact_csv_workbook", exact_csv_workbook)
    return calls


def _model(tmp_path) -> Path:
    p = tmp_path / "model.xlsm"
    p.write_bytes(b"the model")
    return p


def test_raw_files_become_grid_files_in_a_zip_with_a_readme(grids, fake_excel, tmp_path):
    raw = _raw_csv(tmp_path / "in" / "2026_Q2 - BBA_01.csv")
    res = nq.make_next_quarter(_model(tmp_path), TABLE, [raw], grids, "calcCF", tmp_path / "work", messages=[{"line": 1}],
                               model_name="model.xlsm", made_at="2026-11-10 23:00")
    assert res["ok"] is True and res["error"] is None
    (imp,) = fake_excel["import"]
    # the raw .csv goes to the macro as a copy with every digit of its numbers (Excel keeps 15 in a CSV)
    assert fake_excel["exact"] == ["2026_Q2 - BBA_01.csv"]
    assert imp["macro"] == "calcCF" and imp["files"] == {"Ctrl_01!H35": "2026_Q2 - BBA_01.xlsx"} and imp["messages"] == [{"line": 1}]
    assert res["notes"] == ["2026_Q2 - BBA_01.csv: Excel keeps 15 digits of a number when it opens a CSV; the app gave 7 numbers back their full digits."]
    assert fake_excel["export"] == [{"workbook": tmp_path / "work" / "import" / "model.xlsm", "grids": ["IN_BBA_01"]}]
    (f,) = res["files"]
    assert {k: f[k] for k in ("file", "grid", "from", "source", "rows", "cols", "separator", "changed")} == {
        "file": "IN_BBA_01.csv", "grid": "IN_BBA_01", "from": "2026_Q2 - BBA_01.csv", "source": "macro", "rows": 3, "cols": 3, "separator": "comma", "changed": 1}
    assert Path(f["path"]).read_text(encoding="utf-8").splitlines()[0] == "Lapse,Death,Surrender"  # the grid's own header
    assert res["not_in_zip"] == [{"grid": "IN_VFA_01", "why": "no new file for it: Mind keeps the data it has"}]
    assert res["summary"] == "1 Mind input file made (IN_BBA_01.csv): put it in Mind's Inputs window, then Run."
    with zipfile.ZipFile(res["zip"]) as zf:
        assert Path(res["zip"]).name == "Next quarter Mind inputs.zip"
        assert zf.namelist() == ["README.md", "Mind inputs/IN_BBA_01.csv"]
        text = zf.read("README.md").decode("utf-8")
    assert "| `IN_BBA_01.csv` | IN_BBA_01 | 2026_Q2 - BBA_01.csv (the model's import macro) |" in text
    assert "pick the folder `Mind inputs`" in text and "Raw input files cannot go into Mind directly." in text
    assert "- IN_VFA_01: no new file for it: Mind keeps the data it has" in text
    assert "model.xlsm on 2026-11-10 23:00" in text


def test_a_grid_file_goes_in_as_it_is_without_running_the_macro(grids, fake_excel, tmp_path):
    given = tmp_path / "in" / "IN_VFA_01.txt"
    given.parent.mkdir(parents=True)
    given.write_text("Accrued interest, end\tFund\n10\t20\n30\t41\n", encoding="utf-8")
    res = nq.make_next_quarter(_model(tmp_path), TABLE, [given], grids, "calcCF", tmp_path / "work")
    assert res["ok"] and fake_excel["import"] == [] and res["import"] is None
    (f,) = res["files"]
    assert (f["file"], f["grid"], f["source"], f["rows"], f["separator"], f["changed"]) == ("IN_VFA_01.txt", "IN_VFA_01", "as_given", 2, "TAB", 1)
    assert res["notes"] == []


def test_a_raw_file_and_a_grid_file_for_the_same_grid_the_macro_wins(grids, fake_excel, tmp_path):
    raw = _raw_csv(tmp_path / "in" / "2026_Q2 - BBA_01.csv")
    given = tmp_path / "in" / "old.csv"
    given.write_text("Lapse,Death,Surrender\n1,2,3\n", encoding="utf-8")
    res = nq.make_next_quarter(_model(tmp_path), TABLE, [raw, given], grids, "calcCF", tmp_path / "work")
    assert [f["source"] for f in res["files"]] == ["macro"]
    assert res["notes"][-1] == "old.csv was left out: 2026_Q2 - BBA_01.csv fills the same grid IN_BBA_01."


def test_a_macro_that_stops_makes_no_file_and_says_its_message(grids, monkeypatch, tmp_path):
    def run_import(workbook, rows, files, macro, out_path, work_dir, **_kw):
        return {"macro": macro, "files": [{"cell": "Ctrl_01!H35", "file": "2026_Q2 - BBA_01.csv", "opened": True, "areas": []}],
                "dialogs": [{"message": "x", "message_repaired": "BBA_01のinput領域には93項目まで入力可能です"}], "stopped_early": True, "error": None, "saved": None, "ok": False}

    monkeypatch.setattr(outside_inputs, "run_import", run_import)
    raw = _raw_csv(tmp_path / "in" / "2026_Q2 - BBA_01.csv")
    res = nq.make_next_quarter(_model(tmp_path), TABLE, [raw], grids, "calcCF", tmp_path / "work")
    assert res["ok"] is False and res["files"] == [] and res["zip"] is None
    assert res["error"] == ("The model's import macro did not take the new files: \"BBA_01のinput領域には93項目まで入力可能です\". "
                            "No file was made: check the input files, then try again.")


def test_files_nothing_takes_make_no_zip(grids, fake_excel, tmp_path):
    other = tmp_path / "notes.csv"
    other.write_text("a,b\n", encoding="utf-8")
    res = nq.make_next_quarter(_model(tmp_path), TABLE, [other], grids, "calcCF", tmp_path / "work")
    assert res["ok"] is False and res["zip"] is None and fake_excel["import"] == []
    assert res["error"] == "No Mind input file could be made: no row of the model's import table names it, and it is not shaped like a Mind input grid."


# --- the values, through a real Excel -----------------------------------------------------------------
@needs_excel
def test_export_writes_each_grid_block_exactly_as_excel_holds_it(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BBA_01"
    # values whose decimal text Excel and Python read the same way (openpyxl writes text that Excel parses
    # itself: a 17-digit value may come back 1 ulp off, which is why the app reads values through Excel)
    for r, row in enumerate([[0.1, 2 ** -16, 0], [1234567.875, None, -2], ["text", 3, 1e16]], start=3):
        for c, v in enumerate(row, start=29):
            ws.cell(r, c, v)
    model = tmp_path / "model.xlsx"
    wb.save(model)
    grid = {"name": "IN_BBA_01", "sheet": "BBA_01", "ref": "AC3:AE5", "cols": 3, "header": ["Lapse", "Death", "Surrender"]}
    (made,) = nq.export_grids(model, [grid], tmp_path / "out")
    assert made["grid"] == "IN_BBA_01" and made["rows"] == 3 and made["separator"] == "comma"
    lines = Path(made["path"]).read_text(encoding="utf-8").splitlines()
    assert lines == ["Lapse,Death,Surrender", "0.1,0.0000152587890625,0", "1234567.875,,-2", "text,3,10000000000000000"]


@needs_excel
def test_a_raw_csv_reaches_the_macro_with_every_digit(tmp_path):
    """Excel opens a CSV keeping 15 significant digits (144686640.71679688 -> 144686640.716796): the copy
    the macro opens has the exact numbers; dates and text stay as Excel reads them."""
    src = tmp_path / "2026_Q1 - BBA_01.csv"
    src.write_bytes("\ufeff,VariableName,Description,2026-06-30 00:00:00\r\n,s2IRLapseRisk,Lapse,144686640.71679688\r\n,x,y,0.1\r\n"
                    ",z,w,144686640.71679688,txt,1.2345678901234567\r\n".encode("utf-8"))
    out = nq.exact_csv_workbook(src, tmp_path / "exact")
    assert Path(out["path"]).name == "2026_Q1 - BBA_01.xlsx" and out["fixed"] == 3
    wb = openpyxl.load_workbook(out["path"])
    ws = wb.worksheets[0]
    assert ws["D2"].value == 144686640.71679688 and ws["D3"].value == 0.1 and ws["C2"].value == "Lapse"
    # a row with text between its long numbers: the numbers are written one by one, the text stays
    assert ws["D4"].value == 144686640.71679688 and ws["E4"].value == "txt" and ws["F4"].value == 1.2345678901234567
    assert ws["D1"].is_date  # Excel's own reading of a date stays
