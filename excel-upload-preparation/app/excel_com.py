"""Shared Excel COM session helper (Windows + pywin32 + an installed Excel).

Why every *output* workbook goes through Excel rather than openpyxl:
openpyxl round-trips only the parts it understands. Real
Mind workbooks carry parts it doesn't -- a Power Pivot data model
(`xl/model/item.data`, 13 MB in the test workbook), `customXml/` items,
slicers, form controls... -- and it drops them silently. The resulting file
then fails to open in Excel at all ("Open method of Workbooks class
failed"), confirmed on this machine against a real workbook. Excel itself
is the only writer that preserves everything, so it is used whenever it is
available; the openpyxl path is kept only as an explicitly-labelled
reduced-fidelity fallback.

Every session: hidden, alerts off, macros force-disabled
(`AutomationSecurity = 3`), links never updated, events off, and the
MM for Excel add-in loaded when the machine has it (`mm_for_excel_xll`).
"""
from __future__ import annotations

import contextlib
import gc
import os
import re
import sys
from pathlib import Path
from typing import Any, Iterator

from .settings import SETTINGS

FILE_FORMAT_CODES = {"xlsx": 51, "xlsm": 52, "xlsb": 50}  # xlOpenXMLWorkbook / MacroEnabled / Binary
XL_CELL_TYPE_FORMULAS = -4123

# The MM for Excel add-in (Milliman Mind's Excel add-in, an XLL). Excel started by
# automation never loads the add-ins of the user's startup list (the `OPEN` values under
# HKCU\...\Excel\Options), so a background Excel shows #NAME? on every MM_ call unless the
# XLL is registered into that instance explicitly (`Application.RegisterXLL`). With it loaded,
# MM_LOOP / MM_RESULT evaluate as in the user's own Excel: one pass (the first loop value);
# the loop itself runs in Mind.
MM_XLL_DEFAULT = r"C:\Program Files (x86)\Milliman\MM for Excel\XLLDistributions-AddIn64-packed.xll"
MM_XLL_ENV = "MINDREADY_MM_XLL"  # a path, or 0 / off to leave the add-in out
MM_ADDIN_MARKERS = ("xlldistributions", "mm for excel", "mmforexcel")


def mm_for_excel_xll() -> Path | None:
    """The MM for Excel XLL on this machine (the one Excel's own startup list names, else the
    default install path), or None. `mm_for_excel_addin` in settings.yaml names a path instead;
    `MINDREADY_MM_XLL` overrides both: a path, or `0` / `off`."""
    env = os.environ.get(MM_XLL_ENV)
    if env is not None:
        if env.strip().lower() in ("", "0", "off", "no", "none", "false"):
            return None
        path = Path(env.strip().strip('"'))
        return path if path.is_file() else None
    if SETTINGS.mm_for_excel_addin:
        path = Path(SETTINGS.mm_for_excel_addin.strip('"'))
        return path if path.is_file() else None
    candidates: list[Path] = []
    if sys.platform == "win32":
        try:
            import winreg

            for version in ("16.0", "15.0"):
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"Software\Microsoft\Office\{version}\Excel\Options") as key:
                        index = 0
                        while True:
                            try:
                                name, value, _ = winreg.EnumValue(key, index)
                            except OSError:
                                break
                            index += 1
                            if name.upper().startswith("OPEN") and isinstance(value, str):
                                found = re.search(r'"([^"]+\.xll)"', value, re.I)
                                if found:
                                    candidates.append(Path(found.group(1)))
                except OSError:
                    pass
        except ImportError:
            pass
    candidates.append(Path(MM_XLL_DEFAULT))
    for path in candidates:
        if any(marker in str(path).lower() for marker in MM_ADDIN_MARKERS) and path.is_file():
            return path
    return None


def load_mm_addin(excel: Any) -> Path | None:
    """Registers the MM for Excel XLL into `excel` (an automation instance). Returns the path
    loaded, or None when the machine has no add-in or Excel refused it."""
    xll = mm_for_excel_xll()
    if xll is None:
        return None
    try:
        return xll if excel.RegisterXLL(str(xll)) else None
    except Exception:
        return None


def mm_addin_loaded(excel: Any) -> bool:
    """True when an MM for Excel add-in is open in `excel` (XLLs show up in `AddIns2`)."""
    try:
        for i in range(1, excel.AddIns2.Count + 1):
            addin = excel.AddIns2.Item(i)
            try:
                text = f"{addin.Name} {addin.FullName}".lower()
                if addin.IsOpen and any(marker in text for marker in MM_ADDIN_MARKERS):
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def com_available() -> bool:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def start_excel(mm_addin: bool = True) -> Any:
    """A hidden Excel of its own, set up as every session needs it (on a thread that called
    CoInitialize). `mm_addin`: load the MM for Excel add-in into it when the machine has
    one, so MM_ functions evaluate instead of showing #NAME?."""
    import win32com.client

    from .perf import stage

    with stage("excel_start", mm_addin=mm_addin):
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable -- never run macros
        excel.AskToUpdateLinks = False
        try:
            excel.EnableEvents = False
        except Exception:
            pass
        if mm_addin:
            load_mm_addin(excel)
    return excel


def quit_excel(excel: Any) -> None:
    """Quit `excel`, every COM proxy collected first and after.

    A proxy collected after Quit()/CoUninitialize() raises RPC_E_DISCONNECTED (0x80010108)
    inside the garbage collector, which faulthandler reports as a "Windows fatal exception"
    (harmless but alarming). Callers keep proxies inside an inner function scope and we
    collect here."""
    gc.collect()
    try:
        if excel is not None:
            excel.Quit()
    except Exception:
        pass
    excel = None
    gc.collect()


@contextlib.contextmanager
def excel_session(mm_addin: bool = True) -> Iterator[Any]:
    """A hidden Excel of its own for the duration of the block (start_excel / quit_excel)."""
    import pythoncom

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = start_excel(mm_addin)
        yield excel
    finally:
        quit_excel(excel)
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()


# What the per-session Excel worker (app/excel_worker.py) needs to tell a dead or hung
# Excel from a refused operation, and to stop one that hangs.
EXCEL_GONE_HRESULTS = {
    -2147417848,  # RPC_E_DISCONNECTED: the object invoked has disconnected from its clients
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # RPC_S_CALL_FAILED
    -2147417851,  # RPC_E_SERVERFAULT
    -2146959355,  # CO_E_SERVER_EXEC_FAILURE
}


def excel_pid(excel: Any) -> int | None:
    """The process id of an automation Excel (to stop it when it hangs or dies half-way)."""
    try:
        import win32process

        return int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
    except Exception:
        return None


def kill_excel(pid: int | None) -> None:
    if not pid:
        return
    try:
        import ctypes

        handle = ctypes.windll.kernel32.OpenProcess(0x0001, False, int(pid))  # PROCESS_TERMINATE
        if handle:
            ctypes.windll.kernel32.TerminateProcess(handle, 1)
            ctypes.windll.kernel32.CloseHandle(handle)
    except Exception:
        pass


def excel_gone(exc: BaseException) -> bool:
    """True when `exc` says the Excel process is gone or unreachable (not that it refused a call)."""
    code = getattr(exc, "hresult", None)
    if code is None and getattr(exc, "args", None):
        code = exc.args[0]
    return code in EXCEL_GONE_HRESULTS


def excel_alive(excel: Any) -> bool:
    try:
        _ = excel.Workbooks.Count
        return True
    except Exception:
        return False


def open_check(excel: Any, path: Path) -> tuple[dict[str, Any], Any]:
    """Open `path` to change it, as an apply would, and answer as verify_opens_in_excel does.
    -> (the answer, the open workbook or None). A file Excel opens read-only (a password to
    modify) still opens: True, with no workbook to keep."""
    try:
        wb = open_for_write(excel, path)
    except NotSavedError:
        return {"available": True, "opens": True, "sheet_count": None, "error": None}, None
    except Exception as exc:
        return {"available": True, "opens": False, "sheet_count": None, "error": str(exc)[:300]}, None
    try:
        count = int(wb.Worksheets.Count)
    except Exception:
        count = None
    return {"available": True, "opens": True, "sheet_count": count, "error": None}, wb


def open_workbook(excel: Any, path: Path, read_only: bool = True) -> Any:
    # IgnoreReadOnlyRecommended: a workbook saved with "read-only recommended"
    # is otherwise opened read-only whenever alerts are off (the prompt's
    # default answer), whatever ReadOnly says.
    from .perf import stage

    with stage("excel_open", file=Path(path).name, read_only=read_only):
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
    from .perf import stage

    before = _file_stamp(path)
    with stage("excel_save", file=Path(path).name):
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

    from .perf import stage

    try:
        with stage("verify_opens", file=Path(path).name), excel_session() as excel:
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
