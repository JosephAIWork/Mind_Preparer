#!/usr/bin/env python3
"""Audit the grid names of a workbook the way docs/RUN_IN_MIND_NAMING.md
describes: how many grids Mind would list as 'Untitled(row,col)', how many
carry a name that means nothing ('Cashflows C4', 'I'), what the deterministic
heuristics propose, and -- with --suggest -- what the APIM assistant proposes
instead.

Read-only by default: it analyses an immutable copy and writes no workbook.
`--context` additionally dumps the cells around each unnamed/weak grid, which
is what you read when judging a name yourself.

    python scripts/audit_grid_names.py model.xlsm
    python scripts/audit_grid_names.py model.xlsm --suggest --output names.json
    python scripts/audit_grid_names.py model.xlsm --context context.txt
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import prep  # noqa: E402
from app.grid_naming import build_grid_context, suggest_names  # noqa: E402
from app.grids import all_grids  # noqa: E402
from app.modes import plan_mode  # noqa: E402
from app.rules_engine import RulesEngine  # noqa: E402


def render_context(context: dict) -> str:
    lines = [f"--- {context['id']}  ({context['size']})  current={context['current_name']!r}"]
    if context["flags"]:
        lines.append(f" flags: {', '.join(context['flags'])}")
    if context["above"]:
        lines.append(" ABOVE:")
        lines += [f"   {line}" for line in context["above"]]
    if context["left"]:
        lines.append(" LEFT:")
        lines += [f"   {line}" for line in context["left"]]
    lines.append(" DATA:")
    lines += [f"   {line}" for line in context["sample"]]
    return "\n".join(lines)


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input")
    p.add_argument("--suggest", action="store_true", help="also ask the APIM assistant for names (needs the AI key saved in the web app)")
    p.add_argument("--context", help="write the surrounding cells of every unnamed/weak grid to this file")
    p.add_argument("--output", help="write the full audit as JSON to this file (default: summary to stdout)")
    p.add_argument("--work-dir", help="where to place the immutable copy (default: a temp dir)")
    args = p.parse_args()

    source = Path(args.input)
    if not source.is_file():
        print(json.dumps({"status": "ERROR", "message": f"file not found: {source}"}, indent=2))
        return 2

    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(args.work_dir) if args.work_dir else Path(tmp)
        result = plan_mode.run(source, work_dir, {}, RulesEngine())
        analysis, report = result["workbook_analysis"], result["validation_report"]

        grids = all_grids(analysis)
        labels = prep.standalone_labels(analysis)

        rows = []
        for g in grids:
            deterministic, source_of = prep.deterministic_name(analysis, g, labels)
            rows.append(
                {
                    "grid": f"{g['sheet']}!{g['ref']}",
                    "sheet": g["sheet"],
                    "ref": g["ref"],
                    "size": f"{g['n_rows']}x{g['n_cols']}",
                    "current_name": g.get("name"),
                    "unnamed": not g.get("name"),
                    "deterministic": deterministic,
                    "deterministic_source": source_of,
                    "weak": prep.is_weak_name(g.get("name") or deterministic, g["sheet"]),
                }
            )

        weak = [r for r in rows if r["weak"]]
        unnamed = [r for r in rows if r["unnamed"]]

        if args.context:
            by_id = {f"{g['sheet']}!{g['ref']}": g for g in grids}
            blocks = [render_context(build_grid_context(analysis, by_id[r["grid"]])) for r in weak]
            Path(args.context).write_text("\n".join(blocks), encoding="utf-8")

        if args.suggest and weak:
            by_id = {f"{g['sheet']}!{g['ref']}": g for g in grids}
            out = suggest_names([build_grid_context(analysis, by_id[r["grid"]]) for r in weak])
            if out["available"]:
                proposed = out["names"]
                for r in rows:
                    r["suggested"] = proposed.get(r["grid"])
            else:
                print(f"[assistant unavailable] {out.get('message')}", file=sys.stderr)

        # what the prep plan would do about it, and what it cannot do
        plan = {a["id"]: a for a in prep.plan_actions(analysis, report)}
        titles = plan["create_grid_titles"]
        blocked = [s for s in titles["skipped"] if "standalone text" not in s]

        payload = {
            "source": str(source),
            "totals": {
                "grids": len(grids),
                "unnamed": len(unnamed),
                "weak": len(weak),
                "title_operations": titles["count"],
                "blocked": len(blocked),
            },
            "blocked_reasons": blocked,
            "grids": rows,
        }

        if args.output:
            Path(args.output).write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        else:
            print(json.dumps(payload["totals"], indent=2))
            for reason in blocked[:20]:
                print(f"  blocked: {reason}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
