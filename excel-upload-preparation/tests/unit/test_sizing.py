"""1.6.6: the upload size gate (app/sizing.py -- sizes and the sheet list
read from the package alone, .xlsb included), sheet skipping in the
inventory (app/inventory.py load_workbook_selective / build_analysis
ignore_sheets) and the scan progress callback (app/progress.py)."""
from __future__ import annotations

import struct
import zipfile
from pathlib import Path

import openpyxl
import pytest
from local_usecases import usecase
from openpyxl.workbook.defined_name import DefinedName

from app import sizing
from app.config import load_config
from app.inventory import build_analysis, cell_window, load_workbook_selective
from app.modes import plan_mode
from app.progress import overall_progress


def _three_sheet_book(tmp_path: Path) -> Path:
    """Alpha (a titled grid), Big (hidden, 300 rows of formulas -- the heavy
    one), Gamma (a formula referencing both) plus a global and a sheet-scoped
    defined name on Gamma (index 2)."""
    wb = openpyxl.Workbook()
    a = wb.active
    a.title = "Alpha"
    a["A1"] = "SYNTHETIC_TEST_FIXTURE"
    a["A3"] = "#Inputs /Input"
    a["A4"], a["B4"] = "h1", "h2"
    a["A5"], a["B5"] = 1, "=A5*2"
    b = wb.create_sheet("Big")
    b.sheet_state = "hidden"
    for r in range(1, 300):
        b.cell(r, 1, r)
        b.cell(r, 2, f"=A{r}+1")
    c = wb.create_sheet("Gamma")
    c["A1"] = "x"
    c["B2"] = "=Big!A1+Alpha!A5"
    wb.defined_names["GlobalName"] = DefinedName("GlobalName", attr_text="Gamma!$A$1")
    c.defined_names["OnGamma"] = DefinedName("OnGamma", attr_text="Gamma!$B$2", localSheetId=2)
    path = tmp_path / "three_sheets.xlsx"
    wb.save(path)
    wb.close()
    return path


# --- app/sizing.py ---------------------------------------------------------------------
def test_inspect_container_reads_sizes_and_sheets_from_the_package(tmp_path):
    p = _three_sheet_book(tmp_path)
    info = sizing.inspect_container(p)
    assert info["format"] == "xlsx" and info["sheet_list_source"] == "xl/workbook.xml" and info["warnings"] == []
    assert [s["name"] for s in info["sheets"]] == ["Alpha", "Big", "Gamma"]
    assert [s["index"] for s in info["sheets"]] == [0, 1, 2]
    big = next(s for s in info["sheets"] if s["name"] == "Big")
    assert big["state"] == "hidden" and big["part"] == "xl/worksheets/sheet2.xml"
    with zipfile.ZipFile(p) as zf:
        total = sum(i.file_size for i in zf.infolist())
    assert info["decompressed_bytes"] == total > info["file_bytes"] > 0
    assert big["bytes"] == max(s["bytes"] for s in info["sheets"]) and big["share"] > 0.5
    assert abs(sum(s["share"] for s in info["sheets"]) - 1) < 0.01
    assert info["sheet_bytes"] == sum(s["bytes"] for s in info["sheets"])
    assert info["largest_parts"][0]["bytes"] >= info["largest_parts"][-1]["bytes"]


def test_size_gate_threshold_comes_from_env_then_config_then_default(tmp_path, monkeypatch):
    p = _three_sheet_book(tmp_path)
    monkeypatch.delenv("MIND_READY_SIZE_THRESHOLD_MB", raising=False)
    assert sizing.size_threshold_bytes({"upload_size_threshold_mb": 3}) == 3 * sizing.MB
    assert sizing.size_threshold_bytes({}) == int(sizing.DEFAULT_THRESHOLD_MB * sizing.MB)
    assert sizing.size_threshold_bytes({"upload_size_threshold_mb": "nonsense"}) == int(sizing.DEFAULT_THRESHOLD_MB * sizing.MB)
    assert load_config()["upload_size_threshold_mb"] == 25
    below = sizing.size_gate(p, {"upload_size_threshold_mb": 25})
    assert below["above_threshold"] is False and below["measure"] == "decompressed" and below["threshold_mb"] == 25.0
    assert "under the 25 MB limit" in below["message"]
    monkeypatch.setenv("MIND_READY_SIZE_THRESHOLD_MB", "0.001")
    above = sizing.size_gate(p, {"upload_size_threshold_mb": 25})
    assert above["above_threshold"] is True and above["threshold_bytes"] == int(0.001 * sizing.MB)
    assert "skip sheets" in above["message"] and above["decompressed_mb"] >= 0
    monkeypatch.setenv("MIND_READY_SIZE_THRESHOLD_MB", "garbage")
    assert sizing.size_threshold_bytes({"upload_size_threshold_mb": 7}) == int(sizing.DEFAULT_THRESHOLD_MB * sizing.MB)


def test_inspect_container_survives_a_file_that_is_not_a_package(tmp_path):
    p = tmp_path / "x.xlsx"
    p.write_bytes(b"not a zip")
    info = sizing.inspect_container(p)
    assert info["sheets"] == [] and info["warnings"] and info["decompressed_bytes"] == info["file_bytes"] == 9


def _biff12_record(rtype: int, payload: bytes) -> bytes:
    head = bytes([rtype]) if rtype < 0x80 else bytes([(rtype & 0x7F) | 0x80, rtype >> 7])
    size = len(payload)
    size_bytes = b""
    while True:
        b, size = size & 0x7F, size >> 7
        size_bytes += bytes([b | (0x80 if size else 0)])
        if not size:
            break
    return head + size_bytes + payload


def _wide(text: str | None) -> bytes:
    if text is None:
        return struct.pack("<I", 0xFFFFFFFF)
    return struct.pack("<I", len(text)) + text.encode("utf-16-le")


def test_xlsb_sheet_list_is_read_from_the_binary_workbook_part(tmp_path):
    """A synthetic .xlsb container: BrtBundleSh records (MS-XLSB) with a
    visible, a hidden and a very-hidden sheet, a record of another type in
    between, joined to their parts through workbook.bin.rels."""
    sheets = [("Inputs", 0, "rId1", "worksheets/sheet1.bin"), ("Big Data", 1, "rId2", "worksheets/sheet2.bin"), ("Secret", 2, "rId3", "/xl/worksheets/sheet3.bin")]
    body = _biff12_record(0x0083, b"")  # BrtBeginBook
    body += _biff12_record(0x0099, b"\x00" * 12)  # something else (BrtWbProp), skipped
    for i, (name, state, rid, _) in enumerate(sheets):
        body += _biff12_record(0x009C, struct.pack("<II", state, i + 1) + _wide(rid) + _wide(name))
    body += _biff12_record(0x0084, b"")  # BrtEndBook
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    rels += "".join(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{target}"/>' for _, _, rid, target in sheets)
    rels += "</Relationships>"
    p = tmp_path / "synthetic.xlsb"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.bin", body)
        zf.writestr("xl/_rels/workbook.bin.rels", rels)
        zf.writestr("xl/worksheets/sheet1.bin", b"\x01" * 100)
        zf.writestr("xl/worksheets/sheet2.bin", b"\x02" * 10000)
        zf.writestr("xl/worksheets/sheet3.bin", b"\x03" * 400)
    info = sizing.inspect_container(p)
    assert info["format"] == "xlsb" and info["sheet_list_source"] == "xl/workbook.bin" and info["warnings"] == []
    assert [(s["name"], s["state"], s["bytes"]) for s in info["sheets"]] == [("Inputs", "visible", 100), ("Big Data", "hidden", 10000), ("Secret", "veryHidden", 400)]
    assert info["decompressed_bytes"] > info["file_bytes"]  # the binary parts deflate well: unpacked is the honest size
    assert info["sheets"][1]["share"] > 0.9


@pytest.mark.skipif(
    usecase("UC-6.xlsb") is None or usecase("UC-6.xlsm") is None,
    reason="needs a client model's local .xlsb/.xlsm pair (UC-6, tests/local_usecases.py)",
)
def test_real_xlsb_sheet_list_matches_its_xlsm_twin():
    b = sizing.inspect_container(usecase("UC-6.xlsb"))
    x = sizing.inspect_container(usecase("UC-6.xlsm"))
    assert b["sheet_list_source"] == "xl/workbook.bin" and x["sheet_list_source"] == "xl/workbook.xml"
    assert [(s["name"], s["state"]) for s in b["sheets"]] == [(s["name"], s["state"]) for s in x["sheets"]]
    assert all(s["bytes"] > 0 for s in b["sheets"])


# --- app/inventory.py: sheet skipping + progress -----------------------------------------
def test_selective_loader_keeps_positions_names_states_and_sheet_scoped_names(tmp_path):
    p = _three_sheet_book(tmp_path)
    wb = load_workbook_selective(p, ["Big"])
    assert wb.sheetnames == ["Alpha", "Big", "Gamma"]
    assert [ws.sheet_state for ws in wb.worksheets] == ["visible", "hidden", "visible"]
    assert wb["Big"].max_row == 1 and wb["Big"]["A1"].value is None and getattr(wb["Big"], "mind_ready_ignored", False) is True
    assert wb["Gamma"]["B2"].value == "=Big!A1+Alpha!A5" and list(wb["Gamma"].defined_names) == ["OnGamma"]
    assert "GlobalName" in wb.defined_names
    plain = load_workbook_selective(p)
    assert plain["Big"].max_row == 299  # nothing to ignore, no progress hook: openpyxl as usual


def test_build_analysis_skips_ignored_sheets_and_reports_every_stage(tmp_path):
    p = _three_sheet_book(tmp_path)
    events: list[tuple] = []
    analysis = build_analysis(p, tmp_path / "an", "t", ignore_sheets=["Big"], progress=lambda st, msg, fr, **f: events.append((st, msg, fr, f)))
    wb0 = analysis["workbooks"][0]
    assert [s["name"] for s in wb0["sheets"]] == ["Alpha", "Gamma"]
    assert wb0["ignored_sheets"] == [{"name": "Big", "state": "hidden", "index": 1}]
    assert sorted({f["sheet"] for f in wb0["formulas"]}) == ["Alpha", "Gamma"]
    f = analysis["features"]
    assert (f["sheet_count"], f["ignored_sheet_count"], f["total_sheet_count"], f["hidden_sheet_count"]) == (2, 1, 3, 0)
    assert [r["type"] for r in analysis["risks"]] == ["SHEETS_IGNORED"] and "Big" in analysis["risks"][0]["message"]
    assert [(d["name"], d["local_sheet_id"]) for d in wb0["defined_names"]] == [("GlobalName", None), ("OnGamma", 2)]
    stages = [e[0] for e in events]
    assert stages[0] == "copy" and stages[-1] == "names"
    assert [e[1] for e in events if e[0] == "load"] == ["Opening the workbook", "Reading sheet 1/3: Alpha", "Skipping sheet 2/3: Big", "Reading sheet 3/3: Gamma"]
    assert [e[1] for e in events if e[0] == "inventory"] == ["Scanning sheet 1/3: Alpha", "Scanning sheet 3/3: Gamma"]
    assert all(0 <= e[2] <= 1 for e in events if e[2] is not None)
    # the workbook view still works (the skipped sheet is an empty placeholder)
    assert cell_window(analysis, "Gamma", "B2", 0, 0)["rows"][0]["cells"][0]["formula"] == "=Big!A1+Alpha!A5"
    assert cell_window(analysis, "Big", "A1", 0, 0)["rows"][0]["cells"][0]["value"] is None


def test_plan_mode_ignore_list_hides_the_sheet_from_every_rule(tmp_path):
    p = _three_sheet_book(tmp_path)
    config = load_config()
    seen = []
    with_big = plan_mode.run(p, tmp_path / "full", config)
    without = plan_mode.run(p, tmp_path / "part", config, ignore_sheets=["Big"], progress=lambda st, *a, **k: seen.append(st))
    status = lambda res, rid: next(f["status"] for f in res["validation_report"]["findings"] if f["rule_id"] == rid)  # noqa: E731
    assert status(with_big, "STR-007") != "PASS"  # a hidden sheet without the &&Hide marker
    assert status(without, "STR-007") == "PASS"  # ... that the scan was told to skip
    assert len(with_big["validation_report"]["findings"]) == len(without["validation_report"]["findings"])
    assert seen.count("rules") == len(without["validation_report"]["findings"]) and seen[-1] == "report"
    assert without["workbook_analysis"]["features"]["formula_count"] == 2


# --- app/progress.py ---------------------------------------------------------------------
def test_overall_progress_is_monotonic_over_the_stages():
    order = ["copy", "load", "inventory", "names", "rules", "report", "plan"]
    values = [overall_progress(st, 0.0, False) for st in order]
    assert values == sorted(values) and values[0] == 0.0
    assert overall_progress("load", 1.0, False) <= overall_progress("inventory", 0.0, False) + 1e-9
    assert overall_progress("done", None, False) == 1.0 and overall_progress("nope", 0.5, False) == 0.0
    assert overall_progress("copy", 0.0, True) > 0.0  # after the convert stage
    assert overall_progress("load", None, False) == overall_progress("load", 0.5, False)
