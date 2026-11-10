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
# Japanese with no kana at all: a search word, a sheet name, a comment of kanji only (common: 年月, 計算, 契約)
JP_KANJI = "' 年月\r\nSet hdr = Rows(1).Find(\"年月\")\r\nRange(\"計算\").Value = 1\r\n"
JP_KANJI2 = "' 全契約計算実行\r\nMsgBox \"計算完了\"\r\nSet ws = Worksheets(\"前提\")\r\n"
# half-width katakana only (old Japanese message boxes): one byte each, as any single-byte script
JP_HALF = "MsgBox \"ｴﾗｰ\"\r\n"


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
    (KO, "cp949", 0, 932, "cp949"),  # a Korean file on a Japanese PC: the PC's code page is checked before it is trusted
    (ZH, "cp936", 0, 936, "cp936"),  # Chinese has no kana / Hangul to tell it by: the PC's own code page, or the stored one
    (ZH, "cp936", 936, 1252, "cp936"),
    (ZH, "cp936", 0, 932, "cp936"),  # a Chinese file on a Japanese PC: ideographs, before the PC's code page is trusted
    (JP, "utf-8", 0, 1255, "utf-8"),
    # kanji only, no kana: still Japanese (not Korean, not Chinese), on a Hebrew or a Western PC
    (JP_KANJI, "cp932", 0, 1255, "cp932"),
    (JP_KANJI, "cp932", 0, 1252, "cp932"),
    (JP_KANJI, "cp932", 932, 932, "cp932"),
    (JP_KANJI2, "cp932", 0, 1255, "cp932"),
    (JP_KANJI2, "cp932", 0, 1252, "cp932"),
    # half-width katakana only: the stored code page is trusted, and so is the PC's own
    (JP_HALF, "cp932", 932, 1255, "cp932"),
    (JP_HALF, "cp932", 0, 932, "cp932"),
])
def test_each_script_comes_back_as_written(text, encoding, stored, system, expected):
    raw = text.encode(encoding)
    out, codec = decode_vba_source(raw, stored, system)
    assert codec == expected
    if expected in (encoding, "utf-8"):
        assert out == text
    else:
        assert out.encode(expected) == raw  # another code page that holds every byte: nothing is lost


def test_kanji_only_japanese_is_not_read_as_korean_or_chinese():
    """The first version asked cp932 for a quarter of kana, so Find("年月") came back as Korean (cp949 gave
    clean-looking Hangul) and 全契約計算実行 as Chinese. The common kanji count as much as kana now."""
    for text in (JP_KANJI, JP_KANJI2):
        for system in (1255, 1252, 932):
            out, codec = decode_vba_source(text.encode("cp932"), 0, system)
            assert (codec, out) == ("cp932", text), (text, system, codec)
    out, codec = decode_vba_source(JP_HALF.encode("cp932"), 932, 1255)
    assert (codec, out) == ("cp932", JP_HALF)
    # the limit: half-width katakana alone, with no code page recorded, on a Hebrew PC read in the PC's code page
    assert decode_vba_source(JP_HALF.encode("cp932"), 0, 1255)[1] == "cp1255"


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
