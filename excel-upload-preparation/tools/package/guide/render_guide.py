"""The guide that goes with the app (guide.html -> PDF, through Edge).

    python tools/package/guide/render_guide.py [<out.pdf>]

The version comes from VERSION and the feedback address from settings.yaml (feedback.email).
Default output: %LOCALAPPDATA%\\MindReadyBuild\\MindReady_Guide.pdf, next to the built exe.
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import yaml
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
APP = HERE.parents[2]  # excel-upload-preparation


def main() -> int:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(os.environ["LOCALAPPDATA"]) / "MindReadyBuild" / "MindReady_Guide.pdf"
    version = (APP / "VERSION").read_text(encoding="utf-8").strip()
    settings = yaml.safe_load((APP / "settings.yaml").read_text(encoding="utf-8")) or {}
    email = ((settings.get("feedback") or {}).get("email")) or "the Milliman Tel Aviv team"
    html = (HERE / "guide.html").read_text(encoding="utf-8")
    html = html.replace("{{VERSION}}", version).replace("{{DATE}}", time.strftime("%d %B %Y").lstrip("0")).replace("{{EMAIL}}", email)
    filled = out.with_suffix(".html")
    out.parent.mkdir(parents=True, exist_ok=True)
    filled.write_text(html, encoding="utf-8")
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=True)
        page = browser.new_page()
        page.goto(filled.as_uri())
        page.pdf(
            path=str(out), format="A4", print_background=True, display_header_footer=True,
            header_template="<span></span>",
            footer_template='<div style="font-size:8px;color:#8d8d8d;width:100%;text-align:center;font-family:Segoe UI">'
                            'Mind Ready ' + version + ' &middot; page <span class="pageNumber"></span> of <span class="totalPages"></span></div>',
        )
        browser.close()
    filled.unlink()
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
