"""Click through the real screens the way a user would, for one workbook:

    upload -> Recalculate -> "Fix all automatically" -> wait -> read the result card

    python scripts/ui_autofix_check.py --base http://127.0.0.1:8601 --out <dir> model.xlsm

Needs Playwright (pip install playwright) and a browser: Edge is used when
Chromium is not installed. Saves a screenshot before and after, and prints
what the result card says. It checks the *screen*; scripts/run_usecases.py
checks the numbers.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--base", default="http://127.0.0.1:8600")
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--headed", action="store_true")
    args = ap.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    minutes = 60_000
    with sync_playwright() as pw:
        try:
            browser = pw.chromium.launch(headless=not args.headed)
        except Exception:
            browser = pw.chromium.launch(headless=not args.headed, channel="msedge")
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(args.base, wait_until="networkidle")
        page.set_input_files("input[type=file]", str(args.workbook))
        page.get_by_role("button", name="Run analysis").click()
        # a workbook above the size limit stops at the sheet choice (scan everything here); the analysis ends on Findings
        page.wait_for_function("() => document.body.innerText.includes('Scan everything') || location.pathname.startsWith('/findings')", timeout=10 * minutes)
        if not page.url.rstrip("/").endswith("/findings"):
            page.get_by_role("button", name="Scan everything").click()
            page.wait_for_url("**/findings", timeout=120 * minutes)
        page.screenshot(path=str(args.out / "0_findings.png"), full_page=False)
        page.get_by_role("link", name="Recalculate").click(timeout=5 * minutes)
        page.get_by_role("button", name="Recalculate now").click(timeout=30 * minutes)
        page.get_by_text(re.compile(r"FORMULA ERRORS \(\d")).first.wait_for(timeout=30 * minutes)  # the result table, not the page's blurb
        page.screenshot(path=str(args.out / "1_before.png"), full_page=True)
        panel = page.locator("[aria-label='Fix all automatically']")
        button = panel.get_by_role("button", name=re.compile("^Fix all")).first
        print("button:", button.inner_text())
        button.click()
        panel.get_by_text(re.compile(r"errors? left")).first.wait_for(timeout=5 * minutes)  # the live progress line
        print("recalculate button disabled while fixing:", page.get_by_role("button", name=re.compile("^Recalculate")).first.is_disabled())
        page.screenshot(path=str(args.out / "2_running.png"), full_page=True)
        page.locator("[aria-label='Fix all automatically'] >> text=/Clean:|left for a person|Nothing could be fixed|No formula error/").first.wait_for(timeout=120 * minutes)
        page.wait_for_timeout(1500)
        page.screenshot(path=str(args.out / "3_after.png"), full_page=True)
        card = page.locator("[aria-label='Fix all automatically']").inner_text()
        print(card[:2500])
        ok = "Clean:" in card
        browser.close()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
