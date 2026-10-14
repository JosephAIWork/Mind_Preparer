"""Ask Mind's own AI Assistant (the panel inside an open project) a question and record the answer.

    python scripts/mind_ask.py <project init url> "<question>" [--out <dir>] [--wait 180]

Opens the project in the dedicated Edge profile (signed in), types the question into the assistant's
box, presses Ask, waits for the reply to settle, and writes answer.txt + a screenshot into --out.
Read-only: the assistant has no agentic features; nothing in the project changes.
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.mind_client import body_text, is_authenticated, kill_tour, mind_session, open_project, screenshot  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("init_url")
    ap.add_argument("question")
    ap.add_argument("--out", default=None)
    ap.add_argument("--wait", type=int, default=180)
    a = ap.parse_args()
    out = Path(a.out) if a.out else Path.cwd() / "mind_ask"
    out.mkdir(parents=True, exist_ok=True)
    auth = is_authenticated()
    if not auth.get("authenticated"):
        print("not signed in to Mind")
        return 3
    with mind_session(headed=False) as (_ctx, page):
        state = open_project(page, a.init_url)
        print("project:", state, flush=True)
        kill_tour(page)
        try:
            close = page.get_by_role("button", name=re.compile(r"Close modal", re.I))
            if close.count():
                close.first.click(timeout=3000)
        except Exception:
            pass
        # the assistant panel: open it when its box is not on the page yet
        box = page.locator("textarea, input[type=text]").filter(has_not=page.locator("[type=file]"))
        found = None
        for _ in range(3):
            for i in range(box.count()):
                try:
                    ph = (box.nth(i).get_attribute("placeholder") or "")
                    if re.search(r"ask|question|help", ph, re.I) and box.nth(i).is_visible():
                        found = box.nth(i)
                        break
                except Exception:
                    continue
            if found:
                break
            for name in (r"AI Assistant", r"Assistant"):
                try:
                    btn = page.get_by_role("button", name=re.compile(name, re.I))
                    if btn.count():
                        btn.first.click(timeout=3000)
                        page.wait_for_timeout(1500)
                        break
                except Exception:
                    continue
            try:
                icon = page.locator("[title*='Assistant' i],[aria-label*='Assistant' i]")
                if icon.count():
                    icon.first.click(timeout=3000)
                    page.wait_for_timeout(1500)
            except Exception:
                pass
        if not found:
            screenshot(page, out / "no_assistant.png")
            print("assistant box not found; see", out / "no_assistant.png")
            return 4
        before = body_text(page)
        found.click(timeout=3000)
        found.fill(a.question, timeout=5000)
        page.wait_for_timeout(300)
        sent = False
        try:
            ask = page.get_by_role("button", name=re.compile(r"^\s*Ask\s*$", re.I))
            if ask.count():
                ask.first.click(timeout=3000)
                sent = True
        except Exception:
            pass
        if not sent:
            found.press("Enter")
        screenshot(page, out / "asked.png")
        t0 = time.time()
        last = ""
        stable = 0
        while time.time() - t0 < a.wait:
            page.wait_for_timeout(3000)
            now = body_text(page)
            if now != last:
                last = now
                stable = 0
            else:
                stable += 1
            if stable >= 3 and len(now) > len(before) + len(a.question):
                break
        screenshot(page, out / "answer.png")
        after = body_text(page)
        new_lines = [ln for ln in after.splitlines() if ln.strip() and ln not in before]
        answer = "\n".join(new_lines)
        (out / "answer.txt").write_text(answer, encoding="utf-8")
        (out / "page_after.txt").write_text(after, encoding="utf-8")
        print("=== answer (new text on the page) ===")
        print(answer[:6000])
        print("=== saved in", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
