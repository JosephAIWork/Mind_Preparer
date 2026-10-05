"""Drive the running web app through the whole flow for one or more workbooks,
exactly as a user would with the buttons:

    upload -> standard Prep (the actions that are on by default, repeated while
    Prep still plans something) -> Recalculate -> Fix all automatically ->
    Recalculate again (the independent confirmation)

    python scripts/run_usecases.py --base http://127.0.0.1:8601 --out <dir> a.xlsm b.xlsm

Writes <out>/<stem>/result.json (every step, timings, error counts), the final
workbook, and <out>/summary.json. Nothing is uploaded anywhere but the local app.
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any

import requests

LONG = 6 * 3600  # one call on a 40 MB workbook can take many minutes


def _post(base: str, path: str, **kw: Any) -> dict[str, Any]:
    r = requests.post(base + path, timeout=LONG, **kw)
    if r.status_code >= 400:
        raise RuntimeError(f"POST {path} -> {r.status_code}: {r.text[:500]}")
    return r.json()


def _get(base: str, path: str) -> dict[str, Any]:
    r = requests.get(base + path, timeout=600)
    if r.status_code >= 400:
        raise RuntimeError(f"GET {path} -> {r.status_code}: {r.text[:500]}")
    return r.json()


def _default_ops(plan: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """What the Prep screen sends when the user presses Apply without touching a checkbox."""
    chosen = [a for a in plan if a.get("default_on") and a.get("count", 0) > 0]
    return [o for a in chosen for o in a["operations"]], {a["id"]: a["count"] for a in chosen}


def _recalc_facts(rec: dict[str, Any]) -> dict[str, Any]:
    by_error: dict[str, int] = {}
    for e in rec.get("formula_errors", []):
        by_error[e.get("error", "?")] = by_error.get(e.get("error", "?"), 0) + 1
    return {
        "status": rec.get("status"), "ran": rec.get("ran"), "version_id": rec.get("version_id"),
        "formula_errors": len(rec.get("formula_errors", [])), "addin_gap_errors": len(rec.get("addin_gap_errors", [])),
        "by_error": by_error, "groups": len(rec.get("groups", [])), "message": str(rec.get("message", ""))[:300],
    }


def flow(base: str, workbook: Path, out_dir: Path, prep_rounds: int, autofix: bool, autofix_options: dict[str, Any]) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    result: dict[str, Any] = {"workbook": str(workbook), "size_mb": round(workbook.stat().st_size / 1e6, 1), "base": base, "steps": [], "started": time.strftime("%Y-%m-%d %H:%M:%S")}

    def step(name: str, started: float, **facts: Any) -> None:
        entry = {"step": name, "seconds": round(time.time() - started, 1), **facts}
        result["steps"].append(entry)
        print(f"[{workbook.stem}] {name}: " + json.dumps({k: v for k, v in entry.items() if k != 'step'}, ensure_ascii=False)[:600], flush=True)
        (out_dir / "result.json").write_text(json.dumps(result, indent=1, ensure_ascii=False), encoding="utf-8")

    try:
        result["app"] = _get(base, "/api/health")
        t = time.time()
        with workbook.open("rb") as f:
            body = _post(base, "/api/sessions", files={"file": (workbook.name, f, "application/octet-stream")}, data={"mode": "plan"})
        sid = body["sessionId"]
        result["session"] = sid
        plan = body["plan"]
        step("upload+analyze", t, report_status=body["report"]["status"], status_counts=body["report"]["summary"]["status_counts"],
             readiness=(body.get("readiness") or {}).get("state"), blocking=(body.get("readiness") or {}).get("blocking_count"))

        previous = None
        for n in range(1, prep_rounds + 1):
            ops, by_action = _default_ops(plan)
            if not ops:
                break
            if previous is not None and len(ops) >= previous:
                step(f"prep round {n} skipped", time.time(), reason=f"Prep still plans {len(ops)} change(s), no fewer than before ({previous}): stalled", planned=by_action)
                break
            previous = len(ops)
            t = time.time()
            out = _post(base, f"/api/sessions/{sid}/apply", json={"operations": ops, "reanalyze": True})
            plan = out.get("plan", plan)
            res = out["result"]
            step(f"prep round {n}", t, sent=len(ops), by_action=by_action, applied=len(res.get("applied", [])), failed=len(res.get("failed", [])),
                 failed_sample=[f"{f.get('sheet')}!{f.get('cell') or f.get('range') or ''}: {str(f.get('error'))[:160]}" for f in res.get("failed", [])[:5]],
                 version=out["version"]["label"], verified=res.get("verified_opens_in_excel"),
                 readiness=(out.get("readiness") or {}).get("state"), blocking=(out.get("readiness") or {}).get("blocking_count"))

        t = time.time()
        rec = _post(base, f"/api/sessions/{sid}/recalculate")
        step("recalculate", t, **_recalc_facts(rec))
        result["errors_before_autofix"] = len(rec.get("formula_errors", []))

        if autofix and rec.get("formula_errors"):
            t = time.time()
            _post(base, f"/api/sessions/{sid}/auto-fix", json=autofix_options)
            last = ""
            while True:
                time.sleep(3)
                st = _get(base, f"/api/sessions/{sid}/auto-fix")
                line = f"{st.get('state')} | {st.get('title') or ''} | {st.get('message') or ''}"
                if line != last:
                    print(f"[{workbook.stem}]   auto-fix: {line[:220]}", flush=True)
                    last = line
                if st.get("state") in ("done", "error", "idle"):
                    break
            fix = st.get("result") or {}
            (out_dir / "autofix_result.json").write_text(json.dumps(st, indent=1, ensure_ascii=False), encoding="utf-8")
            step("auto-fix", t, state=st.get("state"), error=st.get("error"), status=fix.get("status"), passes=fix.get("passes"),
                 errors_before=fix.get("errors_before"), errors_after=fix.get("errors_after"), cells_fixed=fix.get("cells_fixed"),
                 cells_rewritten=fix.get("cells_rewritten"), left_for_a_person=fix.get("left_count"), rolled_back=fix.get("rolled_back"),
                 by_strategy=fix.get("by_strategy"), version=(st.get("version") or {}).get("label"), summary=fix.get("summary"))

            t = time.time()
            rec = _post(base, f"/api/sessions/{sid}/recalculate")
            step("recalculate (confirmation)", t, **_recalc_facts(rec))

        result["errors_after"] = len(rec.get("formula_errors", []))
        result["clean"] = bool(rec.get("ran")) and not rec.get("formula_errors")
        ready = _get(base, f"/api/sessions/{sid}/readiness")
        result["readiness"] = {k: (ready.get("readiness") or {}).get(k) for k in ("state", "headline", "blocking_count", "manual_blocking_count", "optional_count")}
        versions = _get(base, f"/api/sessions/{sid}/versions")
        vlist = versions if isinstance(versions, list) else versions.get("versions", [])
        result["versions"] = [v["label"] for v in vlist]
        if vlist:
            final = vlist[-1]
            r = requests.get(f"{base}/api/sessions/{sid}/files/{final['file_name']}", timeout=600)
            if r.ok:
                dest = out_dir / final["file_name"]
                dest.write_bytes(r.content)
                result["final_workbook"] = str(dest)
        result["ok"] = True
    except Exception as exc:  # one workbook failing must not stop the others
        result["ok"] = False
        result["failure"] = f"{type(exc).__name__}: {exc}"
        print(f"[{workbook.stem}] FAILED: {result['failure'][:600]}", flush=True)
    result["total_seconds"] = round(time.time() - t0, 1)
    (out_dir / "result.json").write_text(json.dumps(result, indent=1, ensure_ascii=False), encoding="utf-8")
    return result


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbooks", nargs="+", type=Path)
    ap.add_argument("--base", default="http://127.0.0.1:8600")
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--prep-rounds", type=int, default=3)
    ap.add_argument("--no-autofix", action="store_true", help="stop after the first recalculation (a baseline of what Prep leaves)")
    ap.add_argument("--no-assistant", action="store_true", help="auto-fix without asking the assistant (the built-in strategies only)")
    args = ap.parse_args()
    options = {"use_assistant": not args.no_assistant}
    summary = []
    for wb in args.workbooks:
        res = flow(args.base, wb, args.out / wb.stem, args.prep_rounds, not args.no_autofix, options)
        summary.append({k: res.get(k) for k in ("workbook", "ok", "clean", "errors_before_autofix", "errors_after", "total_seconds", "failure", "versions", "readiness")})
        (args.out / "summary.json").write_text(json.dumps(summary, indent=1, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=1, ensure_ascii=False))
    return 0 if all(s.get("ok") and s.get("clean") for s in summary) else 1


if __name__ == "__main__":
    sys.exit(main())
