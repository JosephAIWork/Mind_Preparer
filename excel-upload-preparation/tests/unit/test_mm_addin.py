"""1.10.1: the MM for Excel add-in is loaded into the app's background Excel."""
from __future__ import annotations

import pytest

from app import excel_com
from app.excel_com import com_available, excel_session, mm_addin_loaded, mm_for_excel_xll


def test_env_override_names_or_removes_the_add_in(tmp_path, monkeypatch):
    monkeypatch.setenv(excel_com.MM_XLL_ENV, "off")
    assert mm_for_excel_xll() is None
    monkeypatch.setenv(excel_com.MM_XLL_ENV, str(tmp_path / "missing.xll"))
    assert mm_for_excel_xll() is None
    fake = tmp_path / "fake.xll"
    fake.write_bytes(b"")
    monkeypatch.setenv(excel_com.MM_XLL_ENV, str(fake))
    assert mm_for_excel_xll() == fake


def test_without_env_the_machine_s_add_in_is_found_or_none(monkeypatch):
    monkeypatch.delenv(excel_com.MM_XLL_ENV, raising=False)
    found = mm_for_excel_xll()
    assert found is None or (found.is_file() and found.suffix.lower() == ".xll")


@pytest.mark.skipif(not com_available(), reason="needs pywin32 + Excel")
def test_background_excel_evaluates_mm_functions_with_the_add_in(monkeypatch):
    monkeypatch.delenv(excel_com.MM_XLL_ENV, raising=False)
    if mm_for_excel_xll() is None:
        pytest.skip("no MM for Excel add-in on this machine")

    def run() -> tuple[bool, object, object]:
        with excel_session() as excel:
            loaded = mm_addin_loaded(excel)
            wb = excel.Workbooks.Add()
            try:
                ws = wb.Worksheets(1)
                ws.Range("A1").Value = 10
                ws.Range("A2").Value = 20
                ws.Range("B1").Formula = '=MM_LOOP("k",A1:A2)'
                ws.Range("C1").Formula = '=MM_DIMSIZE("k")'
                excel.CalculateFull()
                return loaded, ws.Range("B1").Value, ws.Range("C1").Value
            finally:
                wb.Close(False)

    loaded, first, size = run()
    assert loaded is True
    assert first == 10 and size == 1  # one pass locally: the first loop value


@pytest.mark.skipif(not com_available(), reason="needs pywin32 + Excel")
def test_the_add_in_can_be_left_out():
    def run() -> bool:
        with excel_session(mm_addin=False) as excel:
            return mm_addin_loaded(excel)

    assert run() is False
