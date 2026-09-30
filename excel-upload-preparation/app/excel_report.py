"""Builds the reviewable Excel outputs.

Two products (CHANGELOG 1.3.0):

* **Standalone report** (`build_standalone_report`): a brand-new .xlsx with a
  Summary sheet and a Findings sheet, written with openpyxl. It contains no
  part of the analyzed workbook, so it is always a valid file regardless of
  what the source contains.
* **Workbook copy + report sheets** (`build_report_workbook`): the analyzed
  workbook (already an immutable copy -- see app/inventory.py) with
  `Mind_Readiness_Report` and `..._Summary` sheets appended. This is written
  by **Excel itself through COM** whenever Excel is available, because a
  pure-Python (openpyxl) save drops package parts it doesn't model -- Power
  Pivot data models, customXml, slicers, controls -- and the result then
  doesn't open in Excel at all (verified on a real 13 MB Mind workbook). The
  openpyxl path remains only as a fallback for machines without Excel and is
  labelled as reduced fidelity in the returned result.

Both are verified by re-opening the output in Excel when Excel is available
(`ReportBuildResult.verified_opens_in_excel`) -- that check is what "a real
working Excel file" means in this app.

JSON stays available for scripting via scripts/validate_workbook.py.
"""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .excel_com import bgr, close_quietly, com_available, excel_session, excel_text, open_for_write, save_in_place, verify_opens_in_excel

REPORT_SHEET_NAME = "Mind_Readiness_Report"
SUMMARY_SHEET_NAME = f"{REPORT_SHEET_NAME}_Summary"

HEADERS = [
    "Rule ID",
    "Category",
    "Severity",
    "Status",
    "Location",
    "Message",
    "Recommended Fix",
    "Can Auto-Fix?",
]

STATUS_RGB = {
    "PASS": "C6EFCE",
    "WARNING": "FFEB9C",
    "ERROR": "FFC7CE",
    "REQUIRES_USER_INPUT": "FFD966",
    "NOT_SUPPORTED": "D9D9D9",
}
STATUS_FILL = {k: PatternFill("solid", fgColor=v) for k, v in STATUS_RGB.items()}
HEADER_RGB = "305496"

# Deterministic remediation text, keyed by rule id first, falling back to
# category. Not LLM-generated -- a report has to render every time, with no
# extra latency or failure mode; see app/llm.py for the separate, narrower
# on-demand "suggest a fix" feature used only in the UI.
FIX_HINTS: dict[str, str] = {
    "FRM-004": "Not a blocker -- Mind ignores VBA it can't run. Confirm nothing needed depends on the macro actually executing; if something does, reimplement that logic as native formulas/MM_ functions.",
    "FRM-002": "An MM_ function missing from the MM_ registry must be corrected. A native Excel function the Mind documentation does not list is neither documented as supported nor as refused: convert the model in Mind to know, and replace the function only if Mind refuses it.",
    "FRM-001": "Inventory only -- review the function mix; the per-function verdicts are in FRM-002 / FORMULA-002.",
    "FRM-005": "Break each circular reference by hand: a value that feeds the next round is what MM_ITERATIONS with /iterationinput and /iterationoutput is for; otherwise reference the previous period or a fixed starting value. Turn Excel's iterative calculation off afterwards.",
    "FRM-006": "Replace each 'First:Last'!cell reference by the explicit list of the sheets it spans (the finding proposes the formula), by hand or with the assistant.",
    "FRM-003": "Remove the implicit-intersection '@' where the formula returns a single value; replace spilled/array results Mind must resize with MM_RANGE (single-cell table) or fixed-size array formulas.",
    "FORMULA-002": "Re-implement the VBA user-defined function as native formulas / MM_ functions, or register it as a custom C# formula in Mind -- VBA does not run in Mind.",
    "LOOP-002": "Use one consistent capitalization for this loop name everywhere it's defined (loop names are case-sensitive in Mind).",
    "LOOP-003": "Make every repeated definition of this loop use the same range length.",
    "LOOP-004": "MM_LOOPINSTANCE(Name, MaxSizeRange, InstanceSizeRange, [Index], [IndexStart]) -- give it at least the 3 required arguments and a one-row / one-column range.",
    "LBL-001": "Make the MM_LOOPLABELS range the same length as the loop it labels.",
    "RES-002": "Reference loops by their exact (case-sensitive) MM_LOOP / MM_LOOPINSTANCE name; referencing a loop that was never created is the #1 cause of the 'Limit of 100 calculation iteration' run error (KB FAQ).",
    "RES-003": "MM_RESULT sums every loop you don't name -- confirm each omitted dimension is meant to be summed, or name it with an index.",
    "RES-004": "Name each loop at most once per MM_RESULT call; to sum two dimensions of one loop, add two MM_RESULT calls.",
    "RES-005": "Confirm 'SIM' references are meant to pick one simulation index (MM_RESULT) or the simulation count (MM_DIMSIZE).",
    "RZS-002": "Write resize formulas as '=<formula> + MM_SETSIZE(NbRows, NbCols)' (or '& IF(MM_SETSIZE(r,c)=0,\"\",0)' for text); MM_SETSIZE must not be nested in another function or a SetSize of a SetSize.",
    "RZS-003": "Clear the cells MM_SETSIZE will fill (down/right of the formula cell): it never overwrites existing content, so leftovers silently break the resize.",
    "RZS-004": "Resize flags need something to drive them -- an MM_SETSIZE reference grid or an /Input grid resized by CSV import.",
    "RZS-005": "A named resize group (/Resize.Name) needs at least two grids carrying the same name to have anything to synchronise.",
    "RZS-006": "Formulas that read a resizable grid must reach its last row/column with MM_LASTROWCELL / MM_LASTCOLUMNCELL (or MM_TABLE / MM_COLUMN / MM_ROW) instead of a fixed A1:A10 range.",
    "RSK-001": "Give each grid a single resize driver: one MM_SETSIZE cell, one resize flag family.",
    "RSK-002": "Keep array results inside the grid boundary; MM_RANGE must be alone in a single-cell table.",
    "LKP-002": "Pass MM_READTABLE a range that starts on the grid's header row (or MM_TABLE(<header cell>)).",
    "LKP-003": "The Header argument must match exactly one header of the range.",
    "LKP-004": "MM_READTABLE returns the first matching row -- de-duplicate the key columns or confirm first-match is intended.",
    "LKP-005": "MM_READTABLE returns a text error message when no row matches; use MM_READTABLENAN if downstream arithmetic expects NaN instead.",
    "STR-001": "Separate grids with an empty row/column; a '#Title' cell inside a detected grid means two grids were merged into one.",
    "STR-002": "Mind ignores a text cell with nothing around it; put labels in a grid (or accept that they don't appear).",
    "STR-003": "Make the first row of each grid text-only if it is meant to be the header row; a number or formula in it turns the headers into 'Col n'.",
    "STR-004": "Name grids with a '#GridName' cell directly above them; unnamed grids appear as 'untitled <cell>' in Mind's navigation.",
    "STR-005": "Rename default sheet names (Sheet1, Feuil1...) -- they become folder names in Mind.",
    "STR-006": "Rename the file: the workbook name is a navigation folder in Mind and an instance-key header.",
    "STR-007": "Add the '&&Hide' marker to intentionally-hidden sheet names (or unhide the sheet if it wasn't meant to be hidden).",
    "GRID-004": "Move the StepLabels grid to its own sheet, alone.",
    "FMT-001": "Only Standard/Number/Text/Boolean/Date formats (plus data validation and hyperlinks) are preserved; percentage, currency, accounting, scientific, fraction and custom formats may display differently in Mind.",
    "FMT-002": "Replace theme colours with explicit RGB colours -- Mind does not support theme colours.",
    "FMT-003": "Empty cells lose Excel styling in Mind; put an apostrophe + space in styled empty cells you want to keep.",
    "FMT-004": "Confirm the merged cells are intentional; unmerge if they aren't required.",
    "FMT-005": "Locked cells on a protected sheet become non-clickable in Mind -- confirm that is intended.",
    "FMT-006": "Outlines import fine (grouping works per depth level); confirm the grouping is intentional.",
    "FMT-007": "Confirm whether these cell comments should be kept before upload.",
    "FORMAT-001": "Preservation inventory only -- compare against the post-fix workbook if a transformation was applied.",
    "RSK-003": "Review each hyperlink for upload risk; remove if not needed.",
    "RSK-004": "LET() is not on the KB supported list -- rewrite without it.",
    "RSK-005": "Confirm external named ranges/links are intended; Mind can preserve or cut links at upload.",
    "READY-001": "Recalculate the workbook (via the app's Recalculate step, or open and recalculate in Excel) before treating it as upload-ready.",
    "FILE-002": "Convert the file to .xlsx or .xlsm before upload -- this app can do that for you automatically.",
    "MMX-001": "Maintain the workbook on a machine with the MMForExcel add-in installed; '_xludf.' prefixes mean it was last saved without it.",
    "MMX-002": "Maintain the workbook in Windows Excel (KB requirement 'Available for Microsoft Excel 2007+ for Windows').",
    "CAL-002": "/CalculationSteps: column 1 = workbook name, column 2 = positive integer step starting at 1; equal steps run in parallel.",
    "CAL-004": "MM_ITERATIONS(\"Name\", Size) may be used only once per workbook.",
    "CAL-005": "Every /iterationinput.(name) needs a matching /iterationoutput.(name) and vice versa.",
    "DBG-003": "Replace whole-column / very large range references with grid-bounded ranges (MM_TABLE / MM_COLUMN) -- calculation time scales with cells evaluated.",
    "DBG-004": "Prefer OFFSET/INDEX, MM_READTABLE as an array formula, or MM_*OPTIMIZED lookups over VLOOKUP/LOOKUP/MATCH on large ranges (KB performance tips).",
    "INP-002": "Input grid names must be unique and usable as CSV file names (the Input Manager matches files by grid name).",
    "INP-005": "/Reorder grids must carry /Input, must not carry /NoHeader and must have unique headers.",
    "EXP-002": "Keep exactly one /ExportSettings grid.",
    "EXP-003": "/ExportSettings columns must be from the documented set, GridName must name an /Export grid, boolean columns must be true/false/1/0, FileExtension .csv/.txt.",
    "EXP-004": "FileNamePattern tags must be {ModelName} {GridName} {InstanceName} {LoopsLabels} {yyyyMMdd}.",
    "INP-004": "FileNamePattern tags must be {GridName} {InstanceName} {ModelName} {*} or {???}.",
    "PAR-002": "/Parameters grid columns: Label | Type | PossibleValues | one or more value columns.",
    "PAR-003": "Type must be one of number, text, switch, checkbox, radio, slider, dropdown, percentage, date.",
    "PAR-004": "PossibleValues: 'min|max' or 'min|max|step' for number/slider/percentage; 'a|b|c' for checkbox/radio/dropdown; empty for text/switch/date.",
    "PRJ-001": "Keep at most one /ProjectSettings grid.",
    "PRJ-002": "/ProjectSettings columns: Name | Value | Locked | Hidden.",
    "PRJ-004": "Stochastic formulas detected: set the simulation count (#NbSimulations grid or project settings) deliberately.",
    "FLG-001": "Use only documented flags (references/mind-flags.yaml); an unknown '/Flag' token is ignored by Mind.",
    "GRP-001": "/Group.(GroupName).x.y needs a group name and two non-negative integer positions.",
    "BKP-001": "Each backup name needs exactly one /BackupSource and at least one /BackupDest; MM_BACKUPBUTTON must sit outside backup grids.",
    "HID-001": "/HideRows column values must be TRUE/FALSE or 1/0 (formulas allowed).",
    "RNG-001": "MM_RANGE must be the only cell of its grid.",
    "UNQ-001": "Special grids (/ProjectSettings, /StepLabels, /InputSettings, /ExportSettings, /AOCOrder, /InstanceKeys, /CalculationSteps, /Translations) must be unique.",
    "TRN-001": "/Translations headers must be ISO 639-1 two-letter language codes.",
    "SIM-001": "#NbSimulations grid must hold one positive integer.",
}
CATEGORY_FIX_HINTS: dict[str, str] = {
    "formula": "Review this formula for Mind compatibility.",
    "loop": "Review this MM_LOOP/MM_RESULT/MM_SETSIZE usage.",
    "structure": "Review workbook/sheet/grid structure.",
    "format": "Confirm this formatting choice is intentional before upload.",
    "risk": "Review for upload risk.",
    "io": "Review the /Input or /Export configuration.",
    "lookup": "Review the MM_READTABLE usage.",
    "parameter": "Review the /Parameters grid.",
    "project": "Review the /ProjectSettings grid.",
    "iteration": "Review the iteration / calculation-step setup.",
    "performance": "Recommendation only.",
    "environment": "Review the authoring environment.",
    "kb": "Review against the Milliman Mind knowledge base.",
}

# Rule IDs app/change_apply.py actually has an automatic handler for.
AUTO_FIXABLE_RULE_IDS = {"FILE-002", "LOOP-002"}

COLUMN_WIDTHS = [12, 14, 12, 20, 22, 60, 60, 14]


@dataclass
class ReportBuildResult:
    path: Path
    method: str  # "excel_com" | "openpyxl"
    verified_opens_in_excel: bool | None = None
    warnings: list[str] = field(default_factory=list)

    def __fspath__(self) -> str:
        return str(self.path)


def _fix_hint(rule_id: str, category: str, message: str) -> str:
    return FIX_HINTS.get(rule_id) or CATEGORY_FIX_HINTS.get(category) or message


def _location_text(location: dict[str, Any]) -> str:
    if not location:
        return ""
    if "sheet" in location and "cell" in location:
        return f"{location['sheet']}!{location['cell']}"
    return ", ".join(f"{k}={v}" for k, v in location.items())


def findings_rows(validation_report: dict[str, Any], rules_by_id: dict[str, dict[str, Any]] | None = None) -> list[list[Any]]:
    rules_by_id = rules_by_id or {}
    findings = sorted(
        validation_report["findings"],
        key=lambda f: (f["status"] != "ERROR", f["status"] != "REQUIRES_USER_INPUT", f["status"], f["rule_id"]),
    )
    rows = []
    for finding in findings:
        rule_id = finding["rule_id"]
        category = rules_by_id.get(rule_id, {}).get("category", "")
        rows.append(
            [
                rule_id,
                category,
                finding.get("severity", ""),
                finding["status"],
                _location_text(finding.get("location") or {}),
                finding["message"],
                _fix_hint(rule_id, category, finding["message"]),
                "Yes" if rule_id in AUTO_FIXABLE_RULE_IDS else "No",
            ]
        )
    return rows


def summary_rows(validation_report: dict[str, Any], source_name: str = "", method: str = "") -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["Milliman Mind Readiness Report", ""],
        ["", ""],
        ["Source workbook", source_name],
        ["Overall status", validation_report["status"]],
        ["Total findings", validation_report["summary"]["finding_count"]],
        ["", ""],
        ["Status", "Count"],
    ]
    for status, count in validation_report["summary"]["status_counts"].items():
        rows.append([status, count])
    rows.append(["", ""])
    rows.append(["Report written by", method or "openpyxl"])
    rows.append(["Note", "PASS is only reported after a real, clean recalculation. Statuses: NOT_SUPPORTED > ERROR > REQUIRES_USER_INPUT > WARNING > PASS."])
    return rows


# --- openpyxl writers (standalone report + fallback) --------------------------
def _style_header(ws: Worksheet) -> None:
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor=HEADER_RGB)
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"


def _write_findings_sheet_openpyxl(ws: Worksheet, rows: list[list[Any]]) -> None:
    _style_header(ws)
    for r, values in enumerate(rows, start=2):
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=col_idx in (6, 7))
        fill = STATUS_FILL.get(values[3])
        if fill:
            ws.cell(row=r, column=4).fill = fill
    for col_idx, width in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _write_summary_sheet_openpyxl(ws: Worksheet, rows: list[list[Any]], overall_status: str) -> None:
    for r, values in enumerate(rows, start=1):
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A4"].font = Font(bold=True)
    ws["B4"].fill = STATUS_FILL.get(overall_status, PatternFill())
    ws["A7"].font = Font(bold=True)
    ws["B7"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 60


def build_standalone_report(
    validation_report: dict[str, Any],
    output_path: Path,
    rules_by_id: dict[str, dict[str, Any]] | None = None,
    source_name: str = "",
) -> Path:
    """A fresh .xlsx holding only the report (Summary + Findings). Always valid."""
    wb = openpyxl.Workbook()
    summary_ws = wb.active
    summary_ws.title = "Summary"
    _write_summary_sheet_openpyxl(summary_ws, summary_rows(validation_report, source_name, "openpyxl (standalone report)"), validation_report["status"])
    findings_ws = wb.create_sheet("Findings")
    _write_findings_sheet_openpyxl(findings_ws, findings_rows(validation_report, rules_by_id))
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path


def _append_report_openpyxl(copy_path: Path, validation_report: dict[str, Any], output_path: Path, rules_by_id, source_name: str) -> None:
    wb = openpyxl.load_workbook(copy_path, data_only=False, keep_vba=copy_path.suffix.lower() == ".xlsm")
    for name in (REPORT_SHEET_NAME, SUMMARY_SHEET_NAME):
        if name in wb.sheetnames:
            del wb[name]
    ws = wb.create_sheet(REPORT_SHEET_NAME)
    _write_findings_sheet_openpyxl(ws, findings_rows(validation_report, rules_by_id))
    summary_ws = wb.create_sheet(SUMMARY_SHEET_NAME, 0)
    _write_summary_sheet_openpyxl(summary_ws, summary_rows(validation_report, source_name, "openpyxl (reduced fidelity fallback)"), validation_report["status"])
    # The readiness sheets are for people, not for Mind: hidden so an upload of this copy ignores them.
    ws.sheet_state = "hidden"
    summary_ws.sheet_state = "hidden"
    visible = [i for i, w in enumerate(wb.worksheets) if w.sheet_state == "visible"]
    if visible:
        wb.active = visible[0]
    wb.save(output_path)
    wb.close()


# --- Excel COM writer ------------------------------------------------------------
XL_TOP = -4160


def _append_report_excel(copy_path: Path, validation_report: dict[str, Any], output_path: Path, rules_by_id, source_name: str) -> None:
    shutil.copy2(copy_path, output_path)  # Excel edits the output copy, never the analysis copy
    rows = findings_rows(validation_report, rules_by_id)
    summary = summary_rows(validation_report, source_name, "Excel (COM automation, full fidelity)")
    def _write(excel: Any) -> None:
        wb = open_for_write(excel, output_path)
        try:
            for name in (REPORT_SHEET_NAME, SUMMARY_SHEET_NAME):
                try:
                    wb.Worksheets(name).Delete()
                except Exception:
                    pass
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = REPORT_SHEET_NAME
            data = [list(HEADERS)] + rows
            n_rows, n_cols = len(data), len(HEADERS)
            safe = tuple(tuple(excel_text(v) for v in row) for row in data)
            ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).Value2 = safe
            header = ws.Range(ws.Cells(1, 1), ws.Cells(1, n_cols))
            header.Font.Bold = True
            header.Font.Color = bgr("FFFFFF")
            header.Interior.Color = bgr(HEADER_RGB)
            header.WrapText = True
            for i, row in enumerate(rows, start=2):
                color = STATUS_RGB.get(row[3])
                if color:
                    ws.Cells(i, 4).Interior.Color = bgr(color)
            for col_idx, width in enumerate(COLUMN_WIDTHS, start=1):
                ws.Columns(col_idx).ColumnWidth = width
            ws.Columns(6).WrapText = True
            ws.Columns(7).WrapText = True
            try:
                ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).VerticalAlignment = XL_TOP
                ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).AutoFilter(1)
            except Exception:
                pass
            try:
                ws.Activate()
                excel.ActiveWindow.FreezePanes = False
                excel.ActiveWindow.SplitColumn = 0
                excel.ActiveWindow.SplitRow = 1
                excel.ActiveWindow.FreezePanes = True
            except Exception:
                pass

            sws = wb.Worksheets.Add(Before=wb.Worksheets(1))
            sws.Name = SUMMARY_SHEET_NAME
            srows = tuple(tuple(excel_text(v) for v in row) for row in summary)
            sws.Range(sws.Cells(1, 1), sws.Cells(len(srows), 2)).Value2 = srows
            sws.Cells(1, 1).Font.Bold = True
            sws.Cells(1, 1).Font.Size = 14
            sws.Cells(4, 1).Font.Bold = True
            color = STATUS_RGB.get(validation_report["status"])
            if color:
                sws.Cells(4, 2).Interior.Color = bgr(color)
            sws.Cells(7, 1).Font.Bold = True
            sws.Cells(7, 2).Font.Bold = True
            sws.Columns(1).ColumnWidth = 24
            sws.Columns(2).ColumnWidth = 60
            # Hidden (xlSheetHidden = 0): the readiness sheets are for people, not for Mind, so an
            # upload of this copy ignores them. Excel needs a visible active sheet first.
            for i in range(1, wb.Worksheets.Count + 1):
                w = wb.Worksheets(i)
                if w.Name not in (REPORT_SHEET_NAME, SUMMARY_SHEET_NAME) and w.Visible == -1:
                    try:
                        w.Activate()
                    except Exception:
                        pass
                    w = None
                    break
                w = None
            ws.Visible = 0
            sws.Visible = 0
            save_in_place(wb, output_path)
        finally:
            close_quietly(wb)

    with excel_session() as excel:
        _write(excel)


def build_report_workbook(
    copy_path: Path,
    validation_report: dict[str, Any],
    output_path: Path,
    rules_by_id: dict[str, dict[str, Any]] | None = None,
    prefer_excel: bool = True,
    verify: bool = True,
    source_name: str = "",
) -> ReportBuildResult:
    """Appends the report sheets to a copy of the analyzed workbook and saves
    it as `output_path`. `copy_path` is the immutable copy already produced
    by app/inventory.py -- never the user's original file. See the module
    docstring for why Excel does the writing when it can."""
    copy_path = Path(copy_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    source_name = source_name or copy_path.name
    warnings: list[str] = []

    method = "openpyxl"
    if prefer_excel and com_available():
        try:
            _append_report_excel(copy_path, validation_report, output_path, rules_by_id, source_name)
            method = "excel_com"
        except Exception as exc:  # Excel not installed / COM failure -> fall back, but say so
            warnings.append(f"Excel COM report writing failed ({str(exc)[:160]}); fell back to openpyxl.")
    if method == "openpyxl":
        _append_report_openpyxl(copy_path, validation_report, output_path, rules_by_id, source_name)
        warnings.append(
            "Written with openpyxl: package parts it cannot model (data model, customXml, slicers, controls...) "
            "are dropped, and the file may not open in Excel. Use the standalone report if in doubt."
        )

    verified: bool | None = None
    if verify and com_available():
        info = verify_opens_in_excel(output_path)
        verified = info["opens"]
        if verified is False:
            warnings.append(f"Verification failed: Excel could not open the output ({info['error']}).")
    return ReportBuildResult(path=output_path, method=method, verified_opens_in_excel=verified, warnings=warnings)
