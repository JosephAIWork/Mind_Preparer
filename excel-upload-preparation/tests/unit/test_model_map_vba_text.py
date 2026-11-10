"""The macro's text in its own code page (model_map.read_vba / decode_vba_source).

The VBE stores module text in the ANSI code page of the PC that saved the workbook, and Excel 365
records 0 ("the system's") as the project's code page; olevba then decodes as UTF-8 with replacement
characters, so Japanese comments and string literals -- Find("年月"), MsgBox "完了しました" -- arrived as
mojibake. The fixture tests/unit/fixtures/vba_japanese_cp932.xlsm is a small row-per-record model
(the loop builder's shape) whose module holds cp932 bytes with code page 0 recorded, as a workbook
from a Japanese PC has (made once: written through Excel, the module stream's source replaced by the
cp932 bytes at the same length; see the scratch script named in the CHANGELOG)."""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from app.instance_builder import build_instance_plan
from app.inventory import build_analysis
from app.loop_builder import build_loop_plan
from app.model_map import build_model_map, decode_vba_source, read_vba

FIXTURE = Path(__file__).parent / "fixtures" / "vba_japanese_cp932.xlsm"

JP = "' 年月を探す\r\nSet hdr = Rows(1).Find(\"年月\")\r\nMsgBox \"完了しました\"\r\n"
HE = "' שלום עולם\r\nMsgBox \"בדיקה\"\r\n"
FR = "' Résumé des données\r\nMsgBox \"terminé\"\r\n"
KO = "' 한국어 주석\r\nMsgBox \"완료\"\r\n"
ZH = "' 中文注释\r\nMsgBox \"完成\"\r\n"


def test_the_fixture_is_read_as_japanese_not_mojibake():
    modules = dict(read_vba(FIXTURE))
    code = modules["Module1"]
    assert "�" not in code
    assert 'Find("年月")' in code and 'MsgBox "完了しました"' in code
    assert "' 全契約の計算を実行する" in code
    # the shape the builders read is intact
    assert "For ins = first_row To last_row" in code and 'Worksheets("Policy Data").Range("A" & ins & ":C" & ins).Value' in code


def test_the_loop_and_instance_builders_still_find_the_loop_of_the_japanese_module(tmp_path):
    p = tmp_path / FIXTURE.name
    shutil.copy(FIXTURE, p)
    analysis = build_analysis(p, tmp_path / "w", "jp")
    mp = build_model_map(analysis)
    assert mp["has_macros"] and mp["loop"]["found"]
    d = mp["loop"]["drivers"][0]
    assert (d["cell"], d["macro"], d["loop_var"]) == ("Calculation!A2:C2", "Run_All", "ins")
    plan = build_loop_plan(analysis, mp)
    assert plan["found"] and plan["loops"][0]["status"] == "ready"
    assert any(o["op"] == "set_formula" and "MM_LOOP(" in o["after"] for o in plan["operations"])
    inst = build_instance_plan(analysis, mp, loop_plan=plan)
    assert inst["found"] is True


@pytest.mark.parametrize("text,encoding,stored,system,expected", [
    (JP, "cp932", 0, 1255, "cp932"),  # a Japanese PC's workbook opened on a Hebrew PC (the Tokyo case)
    (JP, "cp932", 0, 1252, "cp932"),
    (JP, "cp932", 932, 1255, "cp932"),  # the stored code page, when there is one, is trusted
    (HE, "cp1255", 0, 1255, "cp1255"),  # the system's own code page
    (HE, "cp1255", 1255, 1252, "cp1255"),
    (FR, "cp1252", 0, 1252, "cp1252"),
    # the limit: a Western file on a Hebrew PC with no code page recorded reads its accented letters
    # as Hebrew ones (every byte is kept; no single-byte code page can be told from another)
    (FR, "cp1252", 0, 1255, "cp1255"),
    (FR, "cp1252", 1252, 1255, "cp1252"),
    (KO, "cp949", 0, 1255, "cp949"),
    (ZH, "cp936", 0, 936, "cp936"),  # Chinese has no kana / Hangul to tell it by: the PC's own code page, or the stored one
    (ZH, "cp936", 936, 1252, "cp936"),
    (JP, "utf-8", 0, 1255, "utf-8"),
])
def test_each_script_comes_back_as_written(text, encoding, stored, system, expected):
    raw = text.encode(encoding)
    out, codec = decode_vba_source(raw, stored, system)
    assert codec == expected
    if expected in (encoding, "utf-8"):
        assert out == text
    else:
        assert out.encode(expected) == raw  # another code page that holds every byte: nothing is lost


def test_plain_ascii_and_the_last_resort():
    assert decode_vba_source(b"Sub A()\r\nEnd Sub\r\n", 0, 1255) == ("Sub A()\r\nEnd Sub\r\n", "ascii")
    # bytes no listed code page holds: cp1252 with replacement characters, never an exception
    out, codec = decode_vba_source(b"' \x81\x8d\x8f\x90\x9d\r\n", 0, None)
    assert codec == "cp1252" and "�" in out


def test_a_workbook_without_macros_has_no_modules(tmp_path):
    import openpyxl

    p = tmp_path / "plain.xlsx"
    openpyxl.Workbook().save(p)
    assert read_vba(p) == []
