"""Check app.recalc._error_cells (count, then Go To Special) against reading every used range.

    python scripts/check_recalc_errors.py <workbook>

Opens a copy in a hidden Excel, recalculates it (full rebuild), and per sheet finds the error formula
cells both ways, timing each. Exit 0 when both ways find the same cells, in the same order.
"""
from __future__ import annotations

import json
import shutil
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.excel_com import close_quietly, excel_session, open_workbook  # noqa: E402
from app.recalc import _error_cells, _errors_in_block  # noqa: E402
from app.settings import SETTINGS  # noqa: E402


def main() -> int:
    src = Path(sys.argv[1]).resolve()
    work = SETTINGS.data_folder / "perf" / "work"
    work.mkdir(parents=True, exist_ok=True)
    copy = work / f"errors_{src.name}"
    shutil.copy2(src, copy)
    report = {"workbook": str(src), "sheets": {}, "full_s": 0.0, "fast_s": 0.0, "errors": 0, "same": True}
    with excel_session() as excel:
        def run() -> None:
            wb = open_workbook(excel, copy, read_only=False)
            try:
                excel.CalculateFullRebuild()
                for sheet in wb.Worksheets:
                    ur = sheet.UsedRange
                    t = time.perf_counter()
                    full: list = []
                    _errors_in_block(ur, full)
                    t_full = time.perf_counter() - t
                    t = time.perf_counter()
                    fast = _error_cells(sheet, ur)
                    t_fast = time.perf_counter() - t
                    same = full == fast
                    report["sheets"][str(sheet.Name)] = {"errors": len(full), "full_s": round(t_full, 2), "fast_s": round(t_fast, 2), "same": same}
                    report["full_s"] += t_full
                    report["fast_s"] += t_fast
                    report["errors"] += len(full)
                    report["same"] = report["same"] and same
                    ur = None
                    sheet = None
            finally:
                close_quietly(wb)

        run()
    copy.unlink(missing_ok=True)
    report["full_s"], report["fast_s"] = round(report["full_s"], 1), round(report["fast_s"], 1)
    print(json.dumps(report, indent=1, ensure_ascii=True))
    return 0 if report["same"] else 1


if __name__ == "__main__":
    sys.exit(main())
