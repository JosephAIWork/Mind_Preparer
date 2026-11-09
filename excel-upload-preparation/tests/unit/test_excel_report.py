from pathlib import Path

import openpyxl
import pytest

from app.config import load_config
from app.excel_com import com_available, verify_opens_in_excel
from app.excel_report import HEADERS, REPORT_SHEET_NAME, SUMMARY_SHEET_NAME, build_report_workbook, build_standalone_report, findings_rows
from app.modes import plan_mode
from app.plain_codes import rule_name
from app.rules_engine import RulesEngine

needs_excel = pytest.mark.skipif(not com_available(), reason="requires pywin32 + an installed Excel")


def _run(tmp_path, source):
    config = load_config()
    engine = RulesEngine()
    result = plan_mode.run(source, tmp_path / "work", config, engine=engine)
    return result, Path(result["workbook_analysis"]["source"]["copy_path"]), engine


def test_standalone_report_is_a_valid_workbook_with_all_findings(tmp_path, loops_broken_xlsx):
    result, _, engine = _run(tmp_path, loops_broken_xlsx)
    output = build_standalone_report(result["validation_report"], tmp_path / "standalone.xlsx", rules_by_id=engine.rules, source_name="x.xlsx")
    wb = openpyxl.load_workbook(output)
    assert wb.sheetnames == ["Summary", "Findings"]
    ws = wb["Findings"]
    assert [c.value for c in ws[1]] == HEADERS
    assert ws.max_row == result["validation_report"]["summary"]["finding_count"] + 1
    assert wb["Summary"]["B4"].value == result["validation_report"]["status"]
    # the plain name next to each code, and the status still coloured in its own column
    plain = HEADERS.index("In plain words")
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert plain == 1 and all(row[plain] and row[plain] == rule_name(row[0]) for row in rows)
    status = HEADERS.index("Status")
    assert ws.cell(row=2, column=status + 1).fill.fgColor.rgb.endswith(("C6EFCE", "FFEB9C", "FFC7CE", "FFD966", "D9D9D9"))


def test_report_rows_say_the_code_in_plain_words_and_each_error_value_once():
    report = {"findings": [{"rule_id": "REF-001", "status": "ERROR", "location": {"sheet": "S", "cell": "A1"},
                            "message": "3 cell(s) contain broken references: S!A1 =#REF!+1, S!A2 =SUM(#REF!), S!A3 =#N/A"}]}
    (row,) = findings_rows(report, {})
    assert row[:2] == ["REF-001", "a broken reference to a cell or name"]
    assert row[HEADERS.index("Message")] == (
        "3 cell(s) contain broken references: S!A1 =#REF!+1, S!A2 =SUM(#REF!), S!A3 =#N/A\n"
        "(#REF! means the formula points at a cell that no longer exists; #N/A means the value is not available (often: a lookup found nothing).)"
    )
    # the hint that cites other rules says them with their names
    hint = findings_rows({"findings": [{"rule_id": "FRM-001", "status": "PASS", "message": "inventory"}]}, {})[0][HEADERS.index("Recommended Fix")]
    assert "FRM-002 (a function Mind may not support) and FORMULA-002 (a formula calls a custom macro function)" in hint


def test_openpyxl_fallback_appends_sheets_and_says_so(tmp_path, plain_grid_xlsx):
    result, copy_path, engine = _run(tmp_path, plain_grid_xlsx)
    built = build_report_workbook(copy_path, result["validation_report"], tmp_path / "report.xlsx", rules_by_id=engine.rules, prefer_excel=False, verify=False)
    assert built.method == "openpyxl"
    assert any("openpyxl" in w for w in built.warnings)
    wb = openpyxl.load_workbook(built.path)
    assert REPORT_SHEET_NAME in wb.sheetnames and SUMMARY_SHEET_NAME in wb.sheetnames
    assert wb["Data"]["A2"].value == "Header1"  # the original sheet, untouched
    assert [c.value for c in wb[REPORT_SHEET_NAME][1]] == HEADERS
    # the readiness sheets are hidden so an upload of this copy ignores them; a visible sheet stays active
    assert wb[REPORT_SHEET_NAME].sheet_state == "hidden" and wb[SUMMARY_SHEET_NAME].sheet_state == "hidden"
    assert wb.active.sheet_state == "visible"


def test_the_report_warnings_say_excels_errors_in_plain_words(tmp_path, plain_grid_xlsx, monkeypatch):
    """Deliver shows the warnings as they are: Excel's error in them is a plain sentence, never
    pywin32's raw tuple -- also when the check that the file opens kept it cut short."""
    from contextlib import contextmanager

    from app import excel_report

    class ComError(Exception):  # pywin32's com_error: str() is the raw tuple
        def __init__(self, hresult, message, excepinfo):
            super().__init__(hresult, message, excepinfo, None)
            self.hresult, self.strerror, self.excepinfo = hresult, message, excepinfo

    @contextmanager
    def excel_closes():
        raise ComError(-2147023174, "The RPC server is unavailable.", None)
        yield  # pragma: no cover -- a generator, for contextmanager

    member_not_found = "(-2147352567, 'Exception occurred.', (0, None, 'Member not found.', None, 0, -2146788248), None)"
    result, copy_path, engine = _run(tmp_path, plain_grid_xlsx)
    monkeypatch.setattr(excel_report, "com_available", lambda: True)
    monkeypatch.setattr(excel_report, "excel_session", excel_closes)
    monkeypatch.setattr(excel_report, "verify_opens_in_excel", lambda path: {"opens": False, "error": member_not_found[:-12]})
    built = build_report_workbook(copy_path, result["validation_report"], tmp_path / "report.xlsx", rules_by_id=engine.rules)
    assert built.method == "openpyxl" and built.verified_opens_in_excel is False
    assert built.warnings[0] == (
        'Excel COM report writing failed: Excel closed or stopped answering while the app was using it '
        '(Excel said: "The RPC server is unavailable."); fell back to openpyxl.'
    )
    assert built.warnings[-1] == (
        'Verification failed: Excel could not open the output -- Excel does not have the command the app asked for '
        '(Excel said: "Member not found.").'
    )
    assert not any("-21473" in w for w in built.warnings)


@needs_excel
def test_excel_written_report_opens_in_excel(tmp_path, loops_broken_xlsx):
    """The 1.3.0 guarantee: the downloadable workbook copy is written by Excel
    and re-opened by Excel to prove it is a real, working file."""
    result, copy_path, engine = _run(tmp_path, loops_broken_xlsx)
    built = build_report_workbook(copy_path, result["validation_report"], tmp_path / "report.xlsx", rules_by_id=engine.rules)
    assert built.method == "excel_com"
    assert built.verified_opens_in_excel is True
    assert built.warnings == []
    wb = openpyxl.load_workbook(built.path)
    assert wb.sheetnames[0] == SUMMARY_SHEET_NAME and wb.sheetnames[-1] == REPORT_SHEET_NAME
    ws = wb[REPORT_SHEET_NAME]
    assert [c.value for c in ws[1]] == HEADERS
    assert ws.max_row == result["validation_report"]["summary"]["finding_count"] + 1
    assert ws.sheet_state == "hidden" and wb[SUMMARY_SHEET_NAME].sheet_state == "hidden" and wb.active.sheet_state == "visible"
    assert verify_opens_in_excel(built.path)["opens"] is True
