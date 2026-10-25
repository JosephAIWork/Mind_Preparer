"""Upload one workbook to real Milliman Mind, convert it, optionally run it, and print what Mind said.

    python scripts/mind_upload.py <workbook> [--project <prefix>_<stem>] [--run] [--headed] [--keep]

Uses the saved sign-in of app/mind_client.py (python -m app.mind_client check / login <init url>).
Projects are created in the sandbox folder with a name starting with the project prefix (both in
settings.yaml) and deleted afterwards
unless --keep is given. Nothing else in Mind is touched.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.mind_client import (  # noqa: E402
    PROJECT_PREFIX,
    SANDBOX_FOLDER,
    add_template,
    create_blank_project,
    delete_project,
    ensure_folder,
    is_authenticated,
    mind_session,
    open_manager,
    open_project,
    run_model,
    upload_and_convert,
)


def read_deterministic_differences(page) -> list[str]:
    """After "Upload successful (with warnings)": Display Logs -> the "Deterministic Differences" tab ->
    every line "<wb> > <sheet> > <cell>: Excel value: x Mind value: y"."""
    from app.mind_client import body_text, screenshot

    btn = page.get_by_role("button", name=re.compile(r"Display Logs|View logs|Display Error Logs", re.I))
    if not btn.count():
        return []
    try:
        btn.first.click(timeout=8000)
    except Exception:
        return []
    time.sleep(3)
    tab = page.get_by_role("tab", name=re.compile(r"Deterministic", re.I))
    if not tab.count():
        tab = page.get_by_text(re.compile(r"Deterministic Differences", re.I))
    if not tab.count():
        return []
    tab.first.click(timeout=8000)
    time.sleep(3)
    lines = [ln.strip() for ln in body_text(page).splitlines() if "Mind value" in ln]
    try:
        page.get_by_role("button", name=re.compile(r"Hide Logs|Close", re.I)).first.click(timeout=4000)
    except Exception:
        pass
    return lines


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--project", default=None)
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--headed", action="store_true")
    ap.add_argument("--keep", action="store_true", help="leave the project in Mind afterwards")
    ap.add_argument("--timeout", type=int, default=900, help="seconds to wait for the conversion")
    args = ap.parse_args()
    workbook = Path(args.workbook).resolve()
    name = args.project or (f"{PROJECT_PREFIX}_" + re.sub(r"[^A-Za-z0-9]+", "_", workbook.stem)[:40])
    if not name.startswith(PROJECT_PREFIX):
        print(f"the project name must start with {PROJECT_PREFIX} (sandbox only; settings.yaml)")
        return 2
    auth = is_authenticated()
    if not auth.get("authenticated"):
        print("not signed in to Mind:", json.dumps(auth)[:200])
        return 3
    shots = workbook.parent / f"{workbook.stem}_mind_shots"
    shots.mkdir(exist_ok=True)
    result: dict = {"workbook": str(workbook), "project": name}
    t0 = time.time()
    with mind_session(headed=args.headed) as (_ctx, page):
        if not open_manager(page):
            print("could not open the Project manager")
            return 4
        ensure_folder(page, SANDBOX_FOLDER)
        init = create_blank_project(page, name)
        result["init_url"] = init
        if not init:
            print("could not create the project", name)
            return 5
        state = open_project(page, init)
        result["project_state"] = state
        if state != "empty":
            print(f"project {name} is not empty ({state}); use --project with a new name")
            return 6
        conv = upload_and_convert(page, workbook, timeout_s=args.timeout, shots=shots)
        result["conversion"] = {k: conv.get(k) for k in ("success", "steps", "errors", "timed_out")}
        print(f"conversion: {'OK' if conv['success'] else 'FAILED'} after {time.time() - t0:.0f}s | steps {conv.get('steps')}")
        for e in conv.get("errors") or []:
            print("  error:", e[:300])
        if conv["success"]:
            lines = read_deterministic_differences(page)
            result["deterministic_differences"] = len(lines)
            (shots / "deterministic.txt").write_text("\n".join(lines), encoding="utf-8")
            print(f"deterministic differences: {len(lines)} cell(s)" + (f" -> {shots / 'deterministic.txt'}" if lines else " (none: Mind agrees with Excel)"))
        if conv["success"] and args.run:
            added = add_template(page)
            result["template_added"] = added
            if added:
                run = run_model(page, shots=shots)
                result["run"] = {k: run.get(k) for k in run if k != "text"}
                print("run:", json.dumps(result["run"])[:400])
        if not args.keep:
            try:
                result["deleted"] = delete_project(page, name)
            except Exception as exc:
                result["deleted"] = f"not deleted: {exc}"
    (shots / "result.json").write_text(json.dumps(result, indent=1, ensure_ascii=False), encoding="utf-8")
    print("screenshots and result.json in", shots)
    return 0 if result["conversion"]["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
