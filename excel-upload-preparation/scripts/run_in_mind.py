#!/usr/bin/env python3
"""Let the app run itself against real Milliman Mind until everything looks
good and the numbers match (app/mind_loop.py). No person and no assistant is
in the loop while it runs.

    python scripts/run_in_mind.py model.xlsm
    python scripts/run_in_mind.py model.xlsm --work-dir runs\\model --max-iterations 4
    python scripts/run_in_mind.py model.xlsm --no-run          # convert only, no model run
    python scripts/run_in_mind.py model.xlsm --skip-mind       # local gates only (no browser)
    python scripts/run_in_mind.py model.xlsm --enable fix_broken_refs

Needs: Excel (the prep and the numbers gate go through it), a Mind session
(`python -m app.mind_client check` -> authenticated: true), and the AI key
saved in the web app for assistant grid names (optional; deterministic names
otherwise). Writes <work-dir>/loop_report.json after every iteration.
Exit code: 0 converged, 1 stuck/exhausted, 2 bad input.
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.mind_loop import LoopConfig, run_loop, summarize  # noqa: E402
from app.settings import SETTINGS  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input")
    p.add_argument("--work-dir", help="where iterations, copies, screenshots and loop_report.json go (default: runs/<stem>_<stamp>)")
    p.add_argument("--max-iterations", type=int, default=4)
    p.add_argument("--project-prefix", default=SETTINGS.mind_project_prefix, help="sandbox project names start with this (must start with the prefix in settings.yaml)")
    p.add_argument("--enable", action="append", default=[], help="opt-in prep action to force on (repeatable)")
    p.add_argument("--disable", action="append", default=[], help="prep action never to use (repeatable)")
    p.add_argument("--no-assistant", action="store_true", help="deterministic grid names only")
    p.add_argument("--no-run", action="store_true", help="convert in Mind but do not run the model")
    p.add_argument("--no-numbers", action="store_true", help="skip the Excel value comparison")
    p.add_argument("--skip-mind", action="store_true", help="local gates only; never opens a browser")
    p.add_argument("--keep-projects", action="store_true", help="do not try to delete the sandbox projects afterwards")
    p.add_argument("--headed", action="store_true", help="show the Edge window")
    args = p.parse_args()

    source = Path(args.input)
    if not source.is_file():
        print(json.dumps({"status": "ERROR", "message": f"file not found: {source}"}, indent=2))
        return 2
    if source.suffix.lower() not in (".xlsx", ".xlsm", ".xlsb"):
        print(json.dumps({"status": "ERROR", "message": "expected an .xlsx / .xlsm / .xlsb workbook (.xlsb is converted through Excel first)"}, indent=2))
        return 2
    if not args.project_prefix.startswith(SETTINGS.mind_project_prefix):
        print(json.dumps({"status": "ERROR", "message": f"--project-prefix must start with {SETTINGS.mind_project_prefix} (settings.yaml: the only projects the app may delete)"}, indent=2))
        return 2

    work = Path(args.work_dir) if args.work_dir else Path("runs") / f"{source.stem}_{time.strftime('%Y%m%d_%H%M%S')}"
    cfg = LoopConfig(
        source=source,
        work_dir=work,
        max_iterations=args.max_iterations,
        project_prefix=args.project_prefix,
        headed=args.headed,
        use_assistant=not args.no_assistant,
        run_model=not args.no_run,
        check_numbers=not args.no_numbers,
        enable=list(args.enable),
        disable=list(args.disable),
        delete_projects=not args.keep_projects,
        skip_mind=args.skip_mind,
    )

    def progress(e: dict) -> None:
        print(f"[{time.strftime('%H:%M:%S')}] {e.get('message', json.dumps(e))}", flush=True)

    print(f"[run_in_mind] source={source}\n[run_in_mind] work={work}", flush=True)
    report = run_loop(cfg, progress)
    print()
    print(summarize(report), flush=True)
    print(f"\nreport: {work / 'loop_report.json'}", flush=True)
    return 0 if report.get("verdict") == "converged" else 1


if __name__ == "__main__":
    raise SystemExit(main())
