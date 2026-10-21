"""Write what a scan found, in a stable form, so two versions of the app can be compared.

    python scripts/scan_fingerprint.py <workbook> <out.json> [--ignore "Sheet A,Sheet B"]

The analysis (sheets, grids, formulas, names, features, style statistics), every rule's finding and the
prep plan (actions, counts, operations), with the run-specific parts (paths, ids, times) left out, plus
the seconds the scan took. Two outputs from the same workbook must be identical apart from "seconds".
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path.cwd()))  # run from the app folder to compare another copy of it

from app.config import load_config  # noqa: E402
from app.formula_utils import formula_memo  # noqa: E402
from app.modes import plan_mode  # noqa: E402
from app.prep import plan_actions  # noqa: E402
from app.rules_engine import RulesEngine  # noqa: E402

DROP = {"path", "copy_path", "analysis_id", "request_id", "source_path", "output_path", "generated_at", "created_at", "timestamp", "duration_ms", "_version_id", "report_id"}


def clean(x):
    if isinstance(x, dict):
        return {k: clean(v) for k, v in sorted(x.items()) if k not in DROP}
    if isinstance(x, (list, tuple)):
        return [clean(v) for v in x]
    return x


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("out")
    ap.add_argument("--ignore", default="")
    a = ap.parse_args()
    ignore = [x for x in a.ignore.split(",") if x] or None
    t = time.perf_counter()
    with formula_memo():  # as the server's scan (1.11.10)
        res = plan_mode.run(Path(a.workbook), Path(tempfile.mkdtemp(prefix="fp_")), load_config(), engine=RulesEngine(), ignore_sheets=ignore)
        plan = plan_actions(res["workbook_analysis"], res["validation_report"], None)
    seconds = round(time.perf_counter() - t, 1)
    out = {"seconds": seconds, "analysis": clean(res["workbook_analysis"]), "report": clean(res["validation_report"]), "plan": clean(plan)}
    Path(a.out).write_text(json.dumps(out, indent=0, sort_keys=True, default=str, ensure_ascii=False), encoding="utf-8")
    print(f"{ascii(a.workbook)}: {seconds} s", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
