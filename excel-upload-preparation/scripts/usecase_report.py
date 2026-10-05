"""Turn the output of scripts/run_usecases.py into a readable report.

    python scripts/usecase_report.py <run dir> [--title "..."] [--out docs/AUTOFIX_USE_CASES.md]

One section per workbook: what each step of the flow did and how long it took,
what the numbers check said, what the automatic fixer fixed (the biggest kinds
of error, with the reason and one before -> after) and what it left for a person.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


def clock(seconds: float | None) -> str:
    s = int(round(seconds or 0))
    return f"{s} s" if s < 60 else f"{s // 60} min {s % 60:02d} s"


def n(x: Any) -> str:
    return f"{x:,}" if isinstance(x, int) else str(x)


def section(result: dict[str, Any], fix: dict[str, Any] | None) -> list[str]:
    name = Path(result["workbook"]).stem
    steps = {s["step"]: s for s in result.get("steps", [])}
    out = [f"### {name}", ""]
    verdict = "**clean** -- every formula recalculates without an error" if result.get("clean") else f"**{n(result.get('errors_after'))} error cell(s) left**"
    if not result.get("ok"):
        verdict = f"**the run failed**: {result.get('failure')}"
    out.append(f"{result.get('size_mb')} MB · result: {verdict} · total time {clock(result.get('total_seconds'))} · versions: {' → '.join(v.split(' — ')[0] for v in result.get('versions') or [])}")
    out.append("")
    out.append("| Step | Time | What happened |")
    out.append("|---|---|---|")
    for s in result.get("steps", []):
        what = ""
        step = s["step"]
        if step == "upload+analyze":
            skipped = s.get("sheets_left_out_of_the_scan") or []
            what = f"unpacks to {s.get('unpacked_mb')} MB" + (f"; left out of the scan: {', '.join(skipped)}" if skipped else "; every sheet scanned") + f"; verdict *{s.get('readiness')}*, {s.get('blocking')} blocking"
        elif step.startswith("prep round") and "skipped" not in step:
            what = f"{n(s.get('applied'))} change(s) applied ({', '.join(f'{k} {v}' for k, v in (s.get('by_action') or {}).items())})" + (f", {s.get('failed')} refused" if s.get("failed") else "") + f"; verdict *{s.get('readiness')}*"
        elif step.startswith("prep round"):
            what = str(s.get("reason"))
        elif step.startswith("numbers check"):
            c = s.get("counts") or {}
            what = str(s.get("verdict")) + (f" ({n(c.get('error_to_error', 0))} cells are errors before and after)" if c else "")
        elif step == "recalculate":
            what = f"{n(s.get('formula_errors'))} formula error cell(s)" + (f": {', '.join(f'{k} {n(v)}' for k, v in (s.get('by_error') or {}).items())}" if s.get("by_error") else "")
        elif step == "auto-fix":
            what = (
                f"errors {n(s.get('errors_before'))} → {n(s.get('errors_after'))} in {s.get('passes')} pass(es); {n(s.get('cells_rewritten'))} cell(s) rewritten; "
                f"{n(s.get('rolled_back'))} fix(es) undone by the numbers gate"
                + (f"; {n(s.get('broken_by_prep'))} error(s) made by Prep left untouched" if s.get("broken_by_prep") else "")
            )
        elif step.startswith("recalculate (confirmation)"):
            what = f"independent recalculation: {n(s.get('formula_errors'))} formula error cell(s) -- {s.get('status')}"
        out.append(f"| {step} | {clock(s.get('seconds'))} | {what} |")
    out.append("")
    if fix:
        res = fix.get("result") or {}
        fixes = res.get("fixes") or []
        if fixes:
            out.append(f"**What the fixer wrote** ({n(res.get('fix_groups'))} kinds of error; the largest):")
            out.append("")
            out.append("| Cells | Where | Error | Reason | Before → after |")
            out.append("|---|---|---|---|---|")
            for g in fixes[:8]:
                before = str(g.get("before"))[:70].replace("|", "\\|")
                after = str(g.get("after"))[:80].replace("|", "\\|")
                out.append(f"| {n(g['cells'])} | `{g['sheet']}!{g['ranges'][0]}` | {g['error']} | {str(g['reason'])[:230].replace('|', '/')} | `{before}` → `{after}` |")
            out.append("")
        left = res.get("left") or []
        if left:
            out.append(f"**Left for a person** ({n(res.get('left_count'))} cell(s) in {n(res.get('left_groups'))} kinds; the largest):")
            out.append("")
            out.append("| Cells | Where | Error | Why it was not fixed |")
            out.append("|---|---|---|---|")
            for g in left[:10]:
                out.append(f"| {n(g['count'])} | `{g['sheet']}!{g['cell']}` | {g['error']} | {str(g['why'])[:330].replace('|', '/')} |")
            out.append("")
        if res.get("good_values_changed"):
            out.append(f"**Values that changed** (the owner's rule lifted for this run): {n(res['good_values_changed'])}.")
            out.append("")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("run_dir", type=Path)
    ap.add_argument("--title", default="Use cases")
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()
    lines = [f"## {args.title}", ""]
    for d in sorted(p for p in args.run_dir.iterdir() if (p / "result.json").exists()):
        result = json.loads((d / "result.json").read_text(encoding="utf-8"))
        fix = json.loads((d / "autofix_result.json").read_text(encoding="utf-8")) if (d / "autofix_result.json").exists() else None
        lines += section(result, fix)
    text = "\n".join(lines) + "\n"
    if args.out:
        args.out.write_text(text, encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
