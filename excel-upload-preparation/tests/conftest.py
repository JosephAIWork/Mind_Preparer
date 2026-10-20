"""Synthetic workbook builders, per tests/README.md ("Synthetic workbooks
must be clearly labeled"). Every fixture writes 'SYNTHETIC_TEST_FIXTURE' into
A1 of its first sheet and is built fresh into pytest's tmp_path -- no binary
.xlsx files are checked into the repo.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

pytest_plugins = ["kb_fixtures"]  # 1.3.0 synthetic Mind-style models (tests/kb_fixtures.py)

LABEL = "SYNTHETIC_TEST_FIXTURE"


def _label(ws) -> None:
    ws["A1"] = LABEL


@pytest.fixture
def plain_grid_xlsx(tmp_path: Path) -> Path:
    """A clean grid, no MM functions, no VBA -- should pass every implemented check."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    _label(ws)
    ws["A2"] = "Header1"
    ws["B2"] = "Header2"
    ws["A3"] = 1
    ws["B3"] = 2
    ws["A4"] = 3
    ws["B4"] = "=A4+B3"
    path = tmp_path / "plain_grid.xlsx"
    wb.save(path)
    wb.close()
    return path


@pytest.fixture
def loops_ok_xlsx(tmp_path: Path) -> Path:
    """Consistent MM_LOOP / MM_LOOPLABELS / MM_RESULT usage -- should pass the loop checks."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Loops"
    _label(ws)
    for i in range(1, 6):
        ws.cell(row=1 + i, column=1, value=i)
    ws["B1"] = '=MM_LOOP("Scenario", A2:A6)'
    ws["B2"] = '=MM_LOOP("Scenario", A2:A6)'  # repeated definition, same length -> OK
    for i in range(1, 6):
        ws.cell(row=6 + i, column=1, value=f"Label{i}")
    ws["C1"] = '=MM_LOOPLABELS("Scenario", A7:A11)'  # length matches the loop -> OK
    ws["D1"] = "=MM_RESULT(B1, \"Scenario\", 1)"
    path = tmp_path / "loops_ok.xlsx"
    wb.save(path)
    wb.close()
    return path


@pytest.fixture
def loops_broken_xlsx(tmp_path: Path) -> Path:
    """Same-case repeated definition with mismatched length (LOOP-003), a
    case-variant of that same name (LOOP-002), and a mismatched MM_LOOPLABELS
    size (LBL-001) -- should fail all three loop checks independently."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Loops"
    _label(ws)
    for i in range(1, 11):
        ws.cell(row=1 + i, column=1, value=i)
    ws["B1"] = '=MM_LOOP("Scenario", A2:A6)'  # length 5
    ws["B2"] = '=MM_LOOP("Scenario", A2:A11)'  # same exact name, length 10 -> LOOP-003 mismatch
    ws["B3"] = '=MM_LOOP("scenario", A2:A6)'  # case variant -> LOOP-002 collision
    for i in range(1, 4):
        ws.cell(row=12 + i, column=1, value=f"Label{i}")
    ws["C1"] = '=MM_LOOPLABELS("Scenario", A13:A15)'  # length 3, does not match first "Scenario" def (length 5)
    path = tmp_path / "loops_broken.xlsx"
    wb.save(path)
    wb.close()
    return path


@pytest.fixture
def registered_functions_only_xlsx(tmp_path: Path) -> Path:
    """Only calls MM_ functions confirmed present in the mined registry
    (MM_LOOP, MM_RESULT) -- MM_LOOPLABELS has no entry in
    02_MM_Function_Registry.md, so it's deliberately excluded here; see
    test_unsupported_functions_passes_on_registered_calls."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Loops"
    _label(ws)
    for i in range(1, 6):
        ws.cell(row=1 + i, column=1, value=i)
    ws["B1"] = '=MM_LOOP("Scenario", A2:A6)'
    ws["C1"] = '=MM_RESULT(B1, "Scenario", 1)'
    path = tmp_path / "registered_functions_only.xlsx"
    wb.save(path)
    wb.close()
    return path


@pytest.fixture
def unregistered_function_xlsx(tmp_path: Path) -> Path:
    """Calls an MM_-prefixed function that isn't in the mined registry."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bad"
    _label(ws)
    ws["A1"] = "=MM_TOTALLYMADEUP(1,2)"
    path = tmp_path / "unregistered_function.xlsx"
    wb.save(path)
    wb.close()
    return path


@pytest.fixture
def messy_workbook_xlsx(tmp_path: Path) -> Path:
    """Merged cells, a locked cell, a comment, a hyperlink, a hidden sheet without
    the '&&Hide' marker -- exercises the format/structure/risk inventory validators."""
    from openpyxl.comments import Comment
    from openpyxl.worksheet.protection import SheetProtection

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    _label(ws)
    ws.merge_cells("B2:C2")
    ws["B2"] = "Merged"
    ws["D2"] = "locked"
    ws["D2"].protection = openpyxl.styles.Protection(locked=True)
    ws.protection.sheet = True  # KB: locking only takes effect on a protected sheet
    ws["E2"].comment = Comment("a note", "tester")
    ws["F2"].hyperlink = "https://example.invalid/"
    ws["F2"].value = "link"

    ws.sheet_properties.tabColor = "FF0000"  # tab colour: openpyxl hands back an RGB object, must stay JSON-safe
    hidden = wb.create_sheet("Notes")  # no '&&Hide' marker, but sheet_state hidden
    hidden.sheet_state = "hidden"
    _label(hidden)

    path = tmp_path / "messy_workbook.xlsx"
    wb.save(path)
    wb.close()
    return path


@pytest.fixture
def macro_xlsm(tmp_path: Path) -> Path:
    """An .xlsm with a synthetic (non-functional) vbaProject.bin package part,
    to exercise VBA-presence detection without any macro actually running."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    _label(ws)
    path = tmp_path / "macro.xlsm"
    wb.save(path)
    wb.close()

    # openpyxl won't add a real vbaProject.bin; inject a placeholder part directly
    # so has_vba detection (which looks for this exact package part) can be tested.
    tmp_zip = tmp_path / "macro_with_vba.xlsm"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            dst.writestr(item, src.read(item.filename))
        dst.writestr("xl/vbaProject.bin", b"SYNTHETIC_NONFUNCTIONAL_VBA_PLACEHOLDER")
    return tmp_zip

@pytest.fixture(autouse=True)
def _stop_excel_workers():
    """1.11.9: a test never leaves a session's Excel worker (and its Excel) behind."""
    yield
    try:
        from app.excel_worker import stop_all_workers
    except Exception:
        return
    stop_all_workers()
