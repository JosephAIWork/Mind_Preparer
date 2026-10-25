"""Upload a workbook to a sandbox project and look at what Mind shows after the conversion:
the buttons, the "View logs" dialog and its tabs ("Deterministic differences"), with screenshots
and the page text saved -- to teach scripts/mind_upload.py where the differences are.

    python scripts/mind_explore_logs.py <workbook> [--keep]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.mind_client import PROJECT_PREFIX, SANDBOX_FOLDER, body_text, create_blank_project, delete_project, ensure_folder, is_authenticated, mind_session, open_manager, open_project, screenshot, step_status  # noqa: E402


def names(page, role):
    out = []
    loc = page.get_by_role(role)
    for i in range(min(loc.count(), 60)):
        try:
            t = loc.nth(i).inner_text(timeout=1000).strip()
        except Exception:
            t = "?"
        out.append(t[:60])
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--keep", action="store_true")
    args = ap.parse_args()
    workbook = Path(args.workbook).resolve()
    name = f"{PROJECT_PREFIX}_explore_" + re.sub(r"[^A-Za-z0-9]+", "_", workbook.stem)[:30]
    shots = workbook.parent / f"{workbook.stem}_explore"
    shots.mkdir(exist_ok=True)
    if not is_authenticated().get("authenticated"):
        print("not signed in")
        return 3
    log: dict = {"steps": []}
    with mind_session() as (_ctx, page):
        open_manager(page)
        ensure_folder(page, SANDBOX_FOLDER)
        init = create_blank_project(page, name)
        state = open_project(page, init)
        log["project"] = {"name": name, "init": init, "state": state}
        page.set_input_files("input#file", str(workbook))
        end = time.time() + 900
        while time.time() < end:
            text = body_text(page)
            steps = step_status(page)
            if "Upload successful" in text or "Conversion or tests failed" in text or page.get_by_role("button", name=re.compile(r"Error Logs|View logs", re.I)).count():
                break
            time.sleep(5)
        time.sleep(3)
        screenshot(page, shots / "1_after_conversion.png")
        log["steps"] = step_status(page)
        log["buttons_after_conversion"] = names(page, "button")
        log["text_after_conversion"] = body_text(page)[:3000]
        print("steps:", log["steps"])
        print("buttons:", log["buttons_after_conversion"])
        btn = page.get_by_role("button", name=re.compile(r"logs", re.I))
        if btn.count():
            btn.first.click(timeout=8000)
            time.sleep(3)
            screenshot(page, shots / "2_logs_open.png")
            log["tabs"] = names(page, "tab")
            log["buttons_in_logs"] = names(page, "button")
            log["text_logs"] = body_text(page)[:6000]
            print("tabs:", log["tabs"])
            print("buttons in logs:", log["buttons_in_logs"])
            det = page.get_by_role("tab", name=re.compile(r"determin", re.I))
            if not det.count():
                det = page.get_by_text(re.compile(r"Deterministic", re.I))
            if det.count():
                det.first.click(timeout=8000)
                time.sleep(3)
                screenshot(page, shots / "3_deterministic.png")
                t = body_text(page)
                log["text_deterministic"] = t[:20000]
                lines = [ln for ln in t.splitlines() if "Mind value" in ln]
                log["difference_lines"] = lines
                print(f"deterministic tab: {len(lines)} line(s) with 'Mind value'; first:", lines[:3])
                (shots / "deterministic.txt").write_text("\n".join(lines), encoding="utf-8")
            else:
                print("no 'Deterministic' tab found")
        else:
            print("no logs button")
        if not args.keep:
            try:
                log["deleted"] = delete_project(page, name)
            except Exception as exc:
                log["deleted"] = f"not deleted: {str(exc)[:200]}"
    (shots / "explore.json").write_text(json.dumps(log, indent=1, ensure_ascii=False), encoding="utf-8")
    print("saved to", shots)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
