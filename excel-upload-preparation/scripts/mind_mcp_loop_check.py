"""Check a Mind loop as DATA through Mind's MCP server (optional; needs a token, see app/mind_mcp.py).

    python scripts/mind_mcp_loop_check.py <project link or GUID> --source "Calc!AY35:BH35" [--loop Policies]
        [--check check.json] [--expect-size N] [--no-run] [--save] [--out result.json]

By default it starts a classic run first (a project's saved results are often none: the browser runs have
"Save after run" off), reports it (success, consistency, calculations, and on a failure the log's error
lines with every Mind cell mapped to its Excel address), then reads the loop's size and every instance of
every cell of --source (the cells MM_RESULT reads). --check is scripts/check_loop_plan.py's check.json:
instance k is compared with the macro's pass k; instances beyond the macro's passes (rows added to the
table after the macro ran) are listed, not judged.
Without a token (or with an expired one) it says so and exits 3: use scripts/mind_loop_probe.py (browser).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import mind_mcp  # noqa: E402


def print_report(rep: dict) -> None:
    run = rep.get("run")
    if run is not None:
        row = run.get("run") or {}
        print(f"run: {'started' if run.get('started') else 'NOT started: ' + str(run.get('why'))} | result {row.get('Result')} | consistency {row.get('ConsistencyStatus')} | {run.get('calculations')} calculations in {run.get('seconds')} s")
        for e in (run.get("errors") or {}).get("lines", [])[:15]:
            print("  log:", e[:200])
        for c in (run.get("errors") or {}).get("cells", [])[:30]:
            print(f"  error cell: {c['mind']} -> {c.get('excel')}")
    print(f"dimensions: {[(d.get('Name'), d.get('Length')) for d in rep.get('dimensions') or []]}")
    if rep.get("why"):
        print(rep["why"])
    print(f"loop {rep.get('loop')}: size {rep.get('size')}")
    for r in rep.get("instances") or []:
        tag = "beyond the macro's passes (a row added to the table)" if r["beyond_macro"] else ("DIFFERS in " + ", ".join(r["differs"]) if r["differs"] else "")
        print(f"  instance {r['instance']}: {r['values'][:10]} {tag}")
    if rep.get("cells"):
        print(f"macro comparison: {rep.get('same')} cells the same, {rep.get('differ')} differ")
    print("RESULT:", "OK" if rep.get("ok") else "NOT OK")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("project")
    ap.add_argument("--source", default=None, help="Sheet!A1:B1, the cells carrying the loop")
    ap.add_argument("--loop", default=None)
    ap.add_argument("--check", default=None)
    ap.add_argument("--expect-size", type=int, default=None)
    ap.add_argument("--no-run", action="store_true")
    ap.add_argument("--save", action="store_true", help="save the project after the run (default off, like the browser)")
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    ok, why = mind_mcp.available()
    if not ok:
        print(f"Mind MCP not available: {why}. Use scripts/mind_loop_probe.py (the browser) instead.")
        return 3
    expected = json.loads(Path(a.check).read_text(encoding="utf-8"))["expected"].get(a.source) if a.check and a.source else None
    rep = mind_mcp.check_loop(a.project, a.source, a.loop, expected, a.expect_size, do_run=not a.no_run, save=a.save)
    if a.expect_size is not None:
        print(f"expected size {a.expect_size}")
    print_report(rep)
    if a.out:
        Path(a.out).write_text(json.dumps(rep, indent=1, ensure_ascii=False, default=str), encoding="utf-8")
        print("written", a.out)
    return 0 if rep.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
