#!/usr/bin/env python3
"""CLI for the narrow, real fixers in app/change_apply.py (.xlsb / format
conversion and loop-name casing -- the only genuinely automatic corrections the
mined rule set permits; everything else needs a human decision and is
surfaced instead as a one-click apply in the UI)."""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.change_apply import convert_output_format, fix_loop_case_mismatch  # noqa: E402
from app.inventory import build_analysis  # noqa: E402

ACTIONS = ("convert_output_format", "fix_loop_case_mismatch")


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input", nargs="?")
    p.add_argument("--action", choices=ACTIONS, help="Which fix to apply")
    p.add_argument("--target-format", choices=("xlsx", "xlsm"), default="xlsx", help="convert_output_format only")
    p.add_argument("--output")
    args = p.parse_args()

    if not args.input or not args.action:
        payload = {
            "status": "NOT_SUPPORTED",
            "message": (
                "Generic change-set application (arbitrary edits) is not implemented -- only the "
                f"narrow, genuinely-safe fixes are: {ACTIONS}. Pass --action to use one."
            ),
            "input": args.input,
        }
        print(json.dumps(payload, indent=2))
        return 2

    input_path = Path(args.input)
    if not input_path.is_file():
        print(json.dumps({"status": "ERROR", "message": f"file not found: {input_path}"}, indent=2))
        return 2

    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(tmp)
        if args.action == "convert_output_format":
            result = convert_output_format(input_path, work_dir, args.target_format)
        else:
            analysis = build_analysis(input_path, work_dir / "analysis", "cli")
            result = fix_loop_case_mismatch(input_path, work_dir / "fix", analysis)

        if result.get("status") == "APPLIED" and args.output:
            Path(args.output).write_bytes(Path(result["output_path"]).read_bytes())
            result["output_path"] = args.output

        payload = json.dumps({k: str(v) if isinstance(v, Path) else v for k, v in result.items()}, indent=2)

    print(payload)
    status = result.get("status")
    if status in ("APPLIED", "NOT_APPLICABLE"):
        return 0
    return 2 if status == "NOT_SUPPORTED" else 1


if __name__ == "__main__":
    raise SystemExit(main())
