"""Time the app's scan (the analysis every upload and every Apply runs), stage by stage and rule by rule.

    python scripts/perf_scan.py <workbook> [--out perf_scan.json] [--profile prof.txt]

Runs app.modes.plan_mode exactly as the server does, with a progress hook that stamps every stage
(copy, load, inventory, names, rules, report) and every rule, plus the plan (app.prep.plan_actions).
`--profile` also writes the top cumulative functions (cProfile). Peak memory of the process is reported.
"""
from __future__ import annotations

import argparse
import cProfile
import io
import json
import os
import pstats
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.config import load_config  # noqa: E402
from app.formula_utils import formula_memo  # noqa: E402
from app.modes import plan_mode  # noqa: E402
from app.prep import plan_actions  # noqa: E402
from app.rules_engine import RulesEngine  # noqa: E402


def peak_mb() -> float | None:
    try:
        import psutil

        info = psutil.Process().memory_info()
        return round(getattr(info, "peak_wset", info.rss) / 2**20)
    except Exception:
        return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--out", default=None)
    ap.add_argument("--profile", default=None)
    ap.add_argument("--ignore", default="", help="sheets to skip, comma separated (as the server's ignore_sheets)")
    a = ap.parse_args()
    path = Path(a.workbook).resolve()
    stages: dict[str, float] = {}
    rules: dict[str, float] = {}
    last = {"stage": None, "rule": None, "t": time.perf_counter(), "rt": None}

    def progress(stage, message, fraction=None, **facts):
        now = time.perf_counter()
        if last["stage"] is not None:
            stages[last["stage"]] = stages.get(last["stage"], 0.0) + now - last["t"]
        if last["rule"] is not None:
            rules[last["rule"]] = rules.get(last["rule"], 0.0) + now - last["rt"]
            last["rule"] = None
        if stage == "rules" and facts.get("rule_id"):
            last["rule"], last["rt"] = facts["rule_id"], now
        last["stage"], last["t"] = stage, now

    engine = RulesEngine()
    work = Path(tempfile.mkdtemp(prefix="perf_scan_"))
    prof = cProfile.Profile() if a.profile else None
    t0 = time.perf_counter()
    if prof:
        prof.enable()
    ignore = [s.strip() for s in a.ignore.split(",") if s.strip()] or None
    with formula_memo():  # as the server's scan (1.11.10)
        result = plan_mode.run(path, work, load_config(), engine=engine, ignore_sheets=ignore, progress=progress)
        progress("plan", "plan")
        plan = plan_actions(result["workbook_analysis"], result["validation_report"], None)
    progress("end", "end")
    if prof:
        prof.disable()
    total = time.perf_counter() - t0
    wa = result["workbook_analysis"]["workbooks"][0]
    out = {
        "workbook": str(path), "mb": round(path.stat().st_size / 2**20, 1), "total_s": round(total, 1), "peak_mb": peak_mb(),
        "stages_s": {k: round(v, 1) for k, v in stages.items()},
        "slowest_rules_s": {k: round(v, 1) for k, v in sorted(rules.items(), key=lambda kv: -kv[1])[:15]},
        "rules_total_s": round(sum(rules.values()), 1),
        "formulas": len(wa.get("formulas", [])), "sheets": len(wa.get("sheets", [])),
        "plan_actions": sum(1 for x in plan if x.get("count")),
    }
    if prof:
        s = io.StringIO()
        pstats.Stats(prof, stream=s).sort_stats("cumulative").print_stats(60)
        Path(a.profile).write_text(s.getvalue(), encoding="utf-8")
    text = json.dumps(out, indent=1)
    if a.out:
        Path(a.out).parent.mkdir(parents=True, exist_ok=True)
        Path(a.out).write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
