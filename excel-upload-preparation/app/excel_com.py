"""Shared Excel COM session helper (Windows + pywin32 + an installed Excel).

Why every *output* workbook goes through Excel rather than openpyxl
(CHANGELOG 1.3.0): openpyxl round-trips only the parts it understands. Real
Mind workbooks carry parts it doesn't -- a Power Pivot data model
(`xl/model/item.data`, 13 MB in the test workbook), `customXml/` items,
slicers, form controls... -- and it drops them silently. The resulting file
then fails to open in Excel at all ("Open method of Workbooks class
failed"), confirmed on this machine against a real workbook. Excel itself
is the only writer that preserves everything, so it is used whenever it is
available; the openpyxl path is kept only as an explicitly-labelled
reduced-fidelity fallback.

Every session: hidden, alerts off, macros force-disabled
(`AutomationSecurity = 3`), links never updated, events off.
"""
from __future__ import annotations

import contextlib
import gc
from pathlib import Path
from typing import Any, Iterator

FILE_FORMAT_CODES = {"xlsx": 51, "xlsm": 52, "xlsb": 50}  # xlOpenXMLWorkbook / MacroEnabled / Binary
XL_CELL_TYPE_FORMULAS = -4123


def com_available() -> bool:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


@contextlib.contextmanager
def excel_session() -> Iterator[Any]:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable -- never run macros
        excel.AskToUpdateLinks = False
        try:
            excel.EnableEvents = False
        except Exception:
            pass
        yield excel
    finally:
        # Release every COM proxy the caller created *before* Excel goes away:
        # a proxy collected after Quit()/CoUninitialize() raises
        # RPC_E_DISCONNECTED (0x80010108) inside the garbage collector, which
        # faulthandler reports as a "Windows fatal exception" (harmless but
        # alarming). Callers keep proxies inside an inner function scope and
        # we collect here.
        gc.collect()
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def open_workbook(excel: Any, path: Path, read_only: bool = True) -> Any:
    # IgnoreReadOnlyRecommended: a workbook saved with "read-only recommended"
    # is otherwise opened read-only whenever alerts are off (the prompt's
    # default answer), whatever ReadOnly says.
    return excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=read_only, IgnoreReadOnlyRecommended=True, CorruptLoad=0)


class NotSavedError(RuntimeError):
    """Excel did not write the workbook to the file it was opened from."""


def open_for_write(excel: Any, path: Path) -> Any:
    """Open `path` to change it and save it in place. Refuses a workbook Excel
    opened read-only all the same (the file is locked by another program, or
    carries a password to modify): with alerts off, Save() on a read-only
    workbook writes a copy into the user's default folder and reports success,
    leaving `path` untouched."""
    wb = open_workbook(excel, path, read_only=False)
    try:
        read_only = bool(wb.ReadOnly)
    except Exception:
        read_only = False
    if read_only:
        close_quietly(wb)
        wb = None
        raise NotSavedError(f"Excel opened {Path(path).name} read-only (it is open in another program, or it has a password to modify): it cannot be changed in place")
    return wb


def _file_stamp(path: Path) -> tuple[int, int]:
    st = Path(path).stat()
    return st.st_mtime_ns, st.st_size


def save_in_place(wb: Any, path: Path) -> None:
    """wb.Save(), then the proof that `path` itself was written: a save that
    went elsewhere (or nowhere) leaves the file's time and size as they were,
    and must never be reported as done."""
    before = _file_stamp(path)
    wb.Save()
    if _file_stamp(path) == before:
        raise NotSavedError(f"Excel reported the save as done but {Path(path).name} was not written")


def close_quietly(wb: Any) -> None:
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass


def verify_opens_in_excel(path: Path) -> dict[str, Any]:
    """Open `path` read-only in a fresh Excel instance and report whether it
    loads. This is the app's definition of "a real working Excel file"."""
    if not com_available():
        return {"available": False, "opens": None, "sheet_count": None, "error": "pywin32 not available"}
    def _probe(excel: Any) -> dict[str, Any]:
        wb = None
        try:
            wb = open_workbook(excel, path, read_only=True)
            return {"available": True, "opens": True, "sheet_count": int(wb.Worksheets.Count), "error": None}
        except Exception as exc:
            return {"available": True, "opens": False, "sheet_count": None, "error": str(exc)[:300]}
        finally:
            close_quietly(wb)

    try:
        with excel_session() as excel:
            info = _probe(excel)
    except Exception as exc:  # Excel itself not installed / COM registration broken
        info = {"available": False, "opens": None, "sheet_count": None, "error": str(exc)[:300]}
    return info


def excel_text(value: Any) -> Any:
    """Make a Python value safe to assign through Range.Value2: strings that
    Excel would parse as a formula get the text-prefix apostrophe, None
    becomes ''."""
    if value is None:
        return ""
    if isinstance(value, str):
        if value[:1] in ("=", "+", "@") or value[:1] == "-" and len(value) > 1 and not value[1:2].isdigit():
            value = "'" + value
        if len(value) > 32000:
            value = value[:32000] + "..."
        return value
    return value


def bgr(hex_rgb: str) -> int:
    """'305496' (RRGGBB) -> the BGR integer Excel's Interior.Color expects."""
    r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
    return b * 65536 + g * 256 + r
