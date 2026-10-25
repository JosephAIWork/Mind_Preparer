"""Run the main screen's "Do everything" chain through the app's API, headless, and keep the result.

    python scripts/night_chain.py <workbook> <out dir> [--port 8600] [--first 2] [--skip-autofix] [--ignore "Sheet A,Sheet B"]

`--ignore`: the sheets the scan skips (what the big-file sheet chooser does: a deferred upload, then
the scan with those sheets left out).

upload -> Fix all issues (auto-fix, polled) -> Fix the formulas and the structure (every blocking Prep action on by
default, in rounds) -> Build the
Mind loop (first value as given, else the plan's) -> Recalculate -> download the final version and write
chain.json (every step's result) into <out dir>. Mirrors MindReadyUI/src/app/pages/home.page.ts runAll().
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import requests

FORMULA_ACTION_IDS = {
    "fix_broken_refs", "fix_broken_refs_whole", "expand_3d_references", "freeze_external_links", "direct_references",
    "freeze_spill_refs", "guard_blank_arithmetic", "freeze_environment_formulas", "line_breaks_as_char", "keep_error_cells", "empty_cells_as_zero",
}


def say(*a) -> None:
    print(time.strftime("%H:%M:%S"), *a, flush=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("out")
    ap.add_argument("--port", type=int, default=8600)
    ap.add_argument("--first", default=None)
    ap.add_argument("--skip-autofix", action="store_true")
    ap.add_argument("--ignore", default="")
    ap.add_argument("--no-assistant", action="store_true", help="fix without the assistant's advice (deterministic, for before/after comparisons)")
    ap.add_argument("--pivot-mode", default=None, help="blocking / formulas / values: what Prep does with the pivot tables (Advanced > Prep)")
    a = ap.parse_args()
    B = f"http://127.0.0.1:{a.port}/api"
    wb = Path(a.workbook)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    chain: dict = {"workbook": str(wb), "steps": {}}
    t0 = time.time()

    say("upload", wb.name)
    ignore = [x.strip() for x in a.ignore.split(",") if x.strip()]
    with open(wb, "rb") as fh:
        data = {"mode": "plan", **({"defer": "1"} if ignore else {})}
        r = requests.post(f"{B}/sessions", files={"file": (wb.name, fh)}, data=data, timeout=3600)
    r.raise_for_status()
    sid = r.json()["sessionId"]
    if ignore:
        say("deferred upload; scanning without", ignore)
        r = requests.post(f"{B}/sessions/{sid}/analyze", json={"ignore_sheets": ignore}, timeout=3600)
        r.raise_for_status()
    chain["session"] = sid
    say("session", sid, f"({time.time() - t0:.0f}s)", "findings", r.json()["report"].get("counts"))

    def state():
        return requests.get(f"{B}/sessions/{sid}", timeout=300).json()

    if a.pivot_mode:
        r = requests.post(f"{B}/sessions/{sid}/pivot-mode", json={"mode": a.pivot_mode}, timeout=600)
        r.raise_for_status()
        chain["pivot_mode"] = a.pivot_mode
        say("pivot setting:", a.pivot_mode)
    # what the main screen shows before Do everything is clicked, to compare with the real time
    try:
        est = requests.get(f"{B}/sessions/{sid}/estimate", timeout=300)
        chain["estimate"] = est.json() if est.status_code == 200 else {"error": est.text[:300]}
        say("estimate:", json.dumps(chain["estimate"])[:300])
    except requests.RequestException as exc:
        chain["estimate"] = {"error": str(exc)[:300]}
    t_chain = time.time()

    if not a.skip_autofix:
        say("fix all issues: start")
        r = requests.post(f"{B}/sessions/{sid}/auto-fix", json={"use_assistant": not a.no_assistant, "keep_good_values": True}, timeout=300)
        if r.status_code != 200:
            say("auto-fix refused:", r.text[:300])
            chain["steps"]["fixAll"] = {"ok": False, "detail": r.text[:300]}
        else:
            last = ""
            while True:
                time.sleep(10)
                s = requests.get(f"{B}/sessions/{sid}/auto-fix", timeout=120).json()
                line = f"{s.get('title')} · {s.get('message')} · pass {s.get('pass_no')} · errors {s.get('errors_before')} -> {s.get('errors')}"
                if line != last:
                    say("  ", line[:160])
                    last = line
                if s.get("state") != "running":
                    break
            full = requests.get(f"{B}/sessions/{sid}/auto-fix", params={"full": "1"}, timeout=300).json()
            res = full.get("result") or {}
            left = res.get("left") or []
            # "unchanged" with no error left is fine: the workbook already recalculated clean (a client model had no error)
            ok = res.get("status") == "clean" or (res.get("status") in ("partial", "unchanged") and not left and (res.get("errors_after") or 0) == 0)
            # 1.11.6, as the main screen: errors the workbook already had (no regression, no Stop) let the chain go on
            ok = ok or (res.get("status") in ("partial", "stuck", "unchanged") and not res.get("stopped") and not res.get("regression_cells"))
            chain["steps"]["fixAll"] = {"ok": ok, "status": res.get("status"), "errors_before": res.get("errors_before"), "errors_after": res.get("errors_after"), "cells_fixed": res.get("cells_fixed"), "left": len(left), "version": (full.get("version") or {}).get("label"), "error": full.get("error")}
            say("fix all issues:", json.dumps(chain["steps"]["fixAll"])[:300])

    # Fix the formulas and the structure: every blocking Prep action that is on by default, in rounds
    # (an applied action can surface another one), until none is left or 3 rounds ran.
    step: dict = {"ok": True, "rounds": [], "applied": 0, "failed": 0}
    for round_no in range(1, 4):
        st = state()
        acts = [a for a in st["plan"] if a["level"] == "blocking" and a["default_on"] and a["count"] > 0]
        ops = [o for a in acts for o in a["operations"]]
        say(f"prepare round {round_no}: {[(a['id'], a['count']) for a in acts]} -> {len(ops)} operations")
        if not ops:
            break
        errors_before = sum(1 for f in st["report"]["findings"] if f["status"] == "ERROR")
        r = requests.post(f"{B}/sessions/{sid}/apply", json={"operations": ops, "reanalyze": True, "versionStep": "major"}, timeout=7200)
        r.raise_for_status()
        o = r.json()
        res = o["result"]
        applied = len(res["applied"]) if isinstance(res.get("applied"), list) else res.get("applied") or 0
        failed = len(res["failed"]) if isinstance(res.get("failed"), list) else res.get("failed") or 0
        errors_after = sum(1 for f in (o.get("report") or {}).get("findings", []) if f["status"] == "ERROR")
        step["rounds"].append({"actions": [a["id"] for a in acts], "applied": applied, "failed": failed, "errors_before": errors_before, "errors_after": errors_after, "version": o["version"]["label"]})
        step["applied"] += applied
        step["failed"] += failed
        step["version"] = o["version"]["label"]
        if failed:
            step["ok"] = False
            break
    st = state()
    left = [a["id"] for a in st["plan"] if a["level"] == "blocking" and a["default_on"] and a["count"] > 0]
    step["left"] = left
    if left:
        step["ok"] = False
    chain["steps"]["prepare"] = step
    say("prepare:", json.dumps(step)[:400])

    plan = requests.get(f"{B}/sessions/{sid}/mind-loop-plan", timeout=1800).json()["plan"]
    (out / "loop_plan.json").write_text(json.dumps(plan, indent=1, ensure_ascii=False), encoding="utf-8")
    ready = [l for l in plan.get("loops", []) if l["status"] == "ready"]
    header = any(l.get("first") == 1 and any("header" in (n or "").lower() for n in l.get("notes", [])) for l in plan.get("loops", []))
    first = a.first or ("2" if header else None)
    say("loop plan: found", plan.get("found"), "| ready", [l["name"] for l in ready], "| ops", plan.get("operation_count"), "| first", first)
    if plan.get("found") and ready and plan.get("operation_count"):
        body = {"reanalyze": True}
        if first:
            body["first"] = first
        r = requests.post(f"{B}/sessions/{sid}/mind-loop-plan/apply", json=body, timeout=7200)
        if r.status_code != 200:
            chain["steps"]["loop"] = {"ok": False, "detail": r.text[:300]}
        else:
            o = r.json()
            res = o["result"]
            applied = len(res["applied"]) if isinstance(res.get("applied"), list) else res.get("applied")
            failed = len(res["failed"]) if isinstance(res.get("failed"), list) else res.get("failed")
            chain["steps"]["loop"] = {"ok": not failed, "applied": applied, "failed": failed, "version": o["version"]["label"], "how": o["plan"].get("how_it_runs")}
    else:
        chain["steps"]["loop"] = {"ok": None, "skipped": "no ready loop"}
    say("loop:", json.dumps(chain["steps"]["loop"])[:300])

    say("recalculate")
    r = requests.post(f"{B}/sessions/{sid}/recalculate", timeout=7200)
    r.raise_for_status()
    rc = r.json()
    chain["steps"]["recalc"] = {"ok": rc["status"] == "PASS" or (rc.get("new_errors") == 0), "status": rc["status"], "message": rc["message"][:200], "formula_errors": len(rc.get("formula_errors", [])), "new_errors": rc.get("new_errors"), "preexisting_errors": rc.get("preexisting_errors"), "addin_gap": len(rc.get("addin_gap_errors", []))}
    say("recalc:", json.dumps(chain["steps"]["recalc"])[:300])

    st = state()
    v = st["version"]
    chain["final_version"] = v
    chain["readiness"] = st.get("readiness")
    r = requests.get(f"{B}/sessions/{sid}/files/{requests.utils.quote(v['file_name'])}", timeout=3600)
    r.raise_for_status()
    final = out / v["file_name"]
    final.write_bytes(r.content)
    chain["final_file"] = str(final)
    chain["seconds"] = round(time.time() - t0)
    chain["seconds_after_estimate"] = round(time.time() - t_chain)  # what the estimate is for
    (out / "chain.json").write_text(json.dumps(chain, indent=1, ensure_ascii=False, default=str), encoding="utf-8")
    say("final:", v["label"], "->", final, f"({chain['seconds']} s)")
    return 0 if all((s.get("ok") in (True, None)) for s in chain["steps"].values()) else 1


if __name__ == "__main__":
    sys.exit(main())
