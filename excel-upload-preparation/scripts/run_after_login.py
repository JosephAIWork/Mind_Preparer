"""Unattended: wait for a person to sign in to Milliman Mind once, then push a
workbook through the running Mind Ready app's run-in-Mind loop.

    python scripts/run_after_login.py "C:\\path\\Model.xlsm" [--port 8600]
        [--max-iterations 4] [--check-numbers] [--delete-projects]
        [--login-timeout-h 24] [--out DIR]

What it does, in order:
  1. Opens a headed Edge window (the app's dedicated MindReady profile) on the
     Mind project URL and waits until the page is an authenticated Mind page
     (or until the login timeout). Nothing else is done before that.
  2. Uploads the workbook to the app (POST /api/sessions, one-shot scan) and
     waits for the analysis.
  3. Starts the loop (POST /api/sessions/{id}/mind-loop) and polls it to the
     end, echoing every event to the console and to <out>/run_after_login.log.
  4. Writes <out>/summary.txt (verdict, reason, final workbook, leftover Mind
     projects) and copies loop_report.json next to it.

Defaults are the "get it into Mind" ones the runbook does not use: the numbers
gate is OFF (--check-numbers turns it on) and the sandbox projects are KEPT
(--delete-projects deletes the deletable ones), so the converted model stays
visible in Mind's MindReady_Sandbox folder.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import mind_client as mc  # noqa: E402


def _log(out: Path, msg: str) -> None:
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    with (out / "run_after_login.log").open("a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def wait_for_login(out: Path, timeout_s: float, poll_s: float = 3.0) -> bool:
    """Headed Edge on the saved project URL until it shows an authenticated
    Mind page for 8 s in a row. Saves the resolved URL to mind_session.json."""
    sess = mc.load_session()
    url = sess.get("project_url") or mc.MAIN_URL
    _log(out, f"opening Edge (MindReady profile) at {url} -- please sign in to Milliman Mind in that window")
    with mc.mind_session(headed=True) as (_ctx, page):
        page.goto(url, wait_until="domcontentloaded")
        deadline = time.time() + timeout_s
        last_url, stable_since = "", None
        while time.time() < deadline:
            u = page.url
            if u != last_url:
                _log(out, f"browser at {u[:140]}")
                last_url, stable_since = u, None
            if mc._on_app(u) and not mc._is_login_url(u):
                stable_since = stable_since or time.time()
                if time.time() - stable_since >= 8:
                    mc._save_session({"project_url": u, "last_login": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())})
                    _log(out, "signed in -- Mind session saved; closing the login window")
                    return True
            try:
                page.wait_for_timeout(int(poll_s * 1000))
            except Exception:
                time.sleep(poll_s)
    _log(out, "login timed out; nothing was uploaded")
    return False


def _json(req: urllib.request.Request, timeout: float = 600.0) -> dict:
    with urllib.request.urlopen(req, timeout=timeout) as res:
        return json.loads(res.read().decode("utf-8"))


def api_get(base: str, path: str, timeout: float = 60.0) -> dict:
    return _json(urllib.request.Request(base + path), timeout)


def api_post_json(base: str, path: str, body: dict, timeout: float = 600.0) -> dict:
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(base + path, data=data, headers={"Content-Type": "application/json"}, method="POST")
    return _json(req, timeout)


def api_upload(base: str, workbook: Path, mode: str = "plan", timeout: float = 1800.0) -> dict:
    boundary = "----MindReadyRunAfterLogin"
    body = bytearray()
    body += f"--{boundary}\r\nContent-Disposition: form-data; name=\"mode\"\r\n\r\n{mode}\r\n".encode()
    body += f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{workbook.name}\"\r\nContent-Type: application/octet-stream\r\n\r\n".encode()
    body += workbook.read_bytes()
    body += f"\r\n--{boundary}--\r\n".encode()
    req = urllib.request.Request(base + "/api/sessions", data=bytes(body), headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}, method="POST")
    return _json(req, timeout)


def run(args: argparse.Namespace) -> int:
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    workbook = Path(args.workbook).resolve()
    if not workbook.is_file():
        print(f"no such workbook: {workbook}")
        return 2
    base = f"http://127.0.0.1:{args.port}"
    try:
        health = api_get(base, "/api/health")
    except Exception as exc:
        _log(out, f"the app is not answering on {base} ({exc}); start run_mind_ready_web.bat first")
        return 2
    _log(out, f"app {health.get('version')} on {base}: excel={health.get('excel')} assistant={health.get('assistant')}")

    if not wait_for_login(out, args.login_timeout_h * 3600):
        return 1

    _log(out, f"uploading {workbook.name} ({workbook.stat().st_size / 1048576:.1f} MB) and scanning it -- this takes a minute or two")
    try:
        started = api_upload(base, workbook)
    except urllib.error.HTTPError as exc:
        _log(out, f"upload failed: HTTP {exc.code} {exc.read().decode('utf-8', 'replace')[:300]}")
        return 1
    sid = started["sessionId"]
    counts = (started.get("report") or {}).get("summary", {}).get("status_counts", {})
    _log(out, f"session {sid}: findings {counts}")

    payload = {
        "maxIterations": args.max_iterations,
        "useAssistant": True,
        "runModel": True,
        "checkNumbers": bool(args.check_numbers),
        "deleteProjects": bool(args.delete_projects),
        "skipMind": False,
    }
    _log(out, f"starting the run-in-Mind loop with {payload}")
    try:
        api_post_json(base, f"/api/sessions/{sid}/mind-loop", payload, timeout=60)
    except urllib.error.HTTPError as exc:
        _log(out, f"could not start the loop: HTTP {exc.code} {exc.read().decode('utf-8', 'replace')[:300]}")
        return 1

    after, state, status = 0, "running", {}
    while state == "running":
        time.sleep(5)
        try:
            status = api_get(base, f"/api/sessions/{sid}/mind-loop?after={after}")
        except Exception as exc:
            _log(out, f"(poll failed: {exc})")
            continue
        for e in status.get("events", []):
            _log(out, e.get("message") or json.dumps(e))
        after = status.get("next", after)
        state = status.get("state", "running")

    report = status.get("report") or {}
    work_dir = status.get("work_dir")
    if work_dir and (Path(work_dir) / "loop_report.json").is_file():
        shutil.copy2(Path(work_dir) / "loop_report.json", out / "loop_report.json")
    lines = [
        f"state: {state}",
        f"verdict: {report.get('verdict')}",
        f"reason: {report.get('reason')}",
        f"final workbook: {report.get('final_workbook')}",
        f"Mind projects: {', '.join(p.get('name', '?') + ' (' + str(p.get('status')) + ')' for p in report.get('projects', [])) or 'none'}",
        f"leftover projects (still in Mind): {', '.join(report.get('leftover_projects') or []) or 'none'}",
        f"work dir: {work_dir}",
        f"error: {status.get('error')}",
        "",
        status.get("summary") or "",
    ]
    (out / "summary.txt").write_text("\n".join(lines), encoding="utf-8")
    for line in lines:
        _log(out, line)
    return 0 if report.get("verdict") == "converged" else 1


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("workbook")
    p.add_argument("--port", type=int, default=8600)
    p.add_argument("--max-iterations", type=int, default=4)
    p.add_argument("--check-numbers", action="store_true", help="turn the numbers gate on (off by default here)")
    p.add_argument("--delete-projects", action="store_true", help="delete the deletable sandbox projects afterwards (kept by default here)")
    p.add_argument("--login-timeout-h", type=float, default=24.0)
    p.add_argument("--out", default=str(ROOT / "runs" / f"after_login_{time.strftime('%Y%m%d_%H%M%S')}"))
    raise SystemExit(run(p.parse_args()))


if __name__ == "__main__":
    main()
