"""Drive the real Milliman Mind web app in a dedicated, Claude-controlled Edge
browser so the app's outputs can be uploaded and their results read back
automatically.

Design notes
------------
- A *separate* persistent Edge profile under <data folder>\\edge-profile (settings.yaml)
  (never OneDrive-synced, never committed) keeps the login session. This is
  distinct from the user's normal Edge profile, so the user's own Edge does not
  need to be closed.
- The user logs in **once** (interactively, in a headed window). Cookies persist
  in the profile, so later runs can be headless.
- The shared ``/init/...`` link is treated as one-time: after the first login we
  save the *resolved* project URL to ``mind_session.json`` and reuse that.

This module is deliberately small and dependency-light (only Playwright).

Project operations -- the selectors below were mapped on the live site and
proven by a full closed loop on a real model; every one takes the Playwright ``page``
from ``mind_session`` and returns plain dicts so ``app/mind_loop.py`` can run
unattended:

  open_manager -> ensure_folder -> create_blank_project -> open_project
  -> upload_and_convert -> add_template -> run_model -> template_names
  -> export_summary -> delete_project (best effort)
"""
from __future__ import annotations

import json
import re
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from playwright.sync_api import Page, sync_playwright

from .settings import SETTINGS

# --- stable, non-synced locations --------------------------------------------
_ROOT = SETTINGS.data_folder
PROFILE_DIR = _ROOT / "edge-profile"
DOWNLOAD_DIR = _ROOT / "downloads"
RECON_DIR = _ROOT / "recon"
SESSION_FILE = _ROOT / "mind_session.json"
SIGNAL_FILE = _ROOT / "login_done.signal"  # optional manual "I'm logged in" marker

APP_HOST_SUFFIX = SETTINGS.mind_host_suffix
# URL fragments that mean "still on an identity-provider / login page"
_LOGIN_HINTS = ("login", "signin", "sign-in", "sso", "/auth", "oauth", "adfs",
                "microsoftonline", "okta", "b2clogin", "authorize", "logon", "idp")
_LAUNCH_ARGS = ["--no-first-run", "--no-default-browser-check", "--start-maximized"]


def _ensure_dirs() -> None:
    for d in (PROFILE_DIR, DOWNLOAD_DIR, RECON_DIR):
        d.mkdir(parents=True, exist_ok=True)


@contextmanager
def mind_session(headed: bool = False) -> Iterator[tuple[Any, Page]]:
    """Open the persistent Edge context and hand back (context, page).

    Everything Playwright is scoped here so the browser always closes cleanly.
    """
    _ensure_dirs()
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            channel="msedge",
            headless=not headed,
            accept_downloads=True,
            downloads_path=str(DOWNLOAD_DIR),
            args=_LAUNCH_ARGS,
            no_viewport=headed,  # let the real maximized window drive the size when headed
        )
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.set_default_timeout(60_000)
            yield ctx, page
        finally:
            ctx.close()


# --- session bookkeeping ------------------------------------------------------
def load_session() -> dict[str, Any]:
    if SESSION_FILE.is_file():
        try:
            return json.loads(SESSION_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_session(data: dict[str, Any]) -> None:
    _ensure_dirs()
    merged = load_session()
    merged.update(data)
    SESSION_FILE.write_text(json.dumps(merged, indent=2), encoding="utf-8")


def _is_login_url(url: str) -> bool:
    u = url.lower()
    return any(h in u for h in _LOGIN_HINTS)


def _on_app(url: str) -> bool:
    try:
        host = re.sub(r"^https?://", "", url).split("/", 1)[0].lower()
    except Exception:
        return False
    return host.endswith(APP_HOST_SUFFIX)


def _looks_authenticated(page: Page, init_url: str | None) -> bool:
    """On the Mind app host, not an identity-provider page, and not still sitting
    on the one-time init link."""
    url = page.url
    if not _on_app(url) or _is_login_url(url):
        return False
    if init_url and url.rstrip("/") == init_url.rstrip("/"):
        return False
    return True


# --- element inventory / capture (used by the Phase A tour) -------------------
_INVENTORY_JS = r"""
() => {
  const vis = (el) => {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  const txt = (el) => (el.innerText || el.value || el.getAttribute('aria-label') ||
                       el.getAttribute('title') || el.getAttribute('placeholder') || '').trim().slice(0, 80);
  const grab = (sel) => Array.from(document.querySelectorAll(sel)).filter(vis).map((el) => ({
    tag: el.tagName.toLowerCase(),
    type: el.getAttribute('type'),
    role: el.getAttribute('role'),
    id: el.id || null,
    name: el.getAttribute('name'),
    href: el.getAttribute('href'),
    accept: el.getAttribute('accept'),
    text: txt(el),
  })).filter((e, i, a) => a.findIndex((x) => JSON.stringify(x) === JSON.stringify(e)) === i);
  return {
    title: document.title,
    url: location.href,
    headings: Array.from(document.querySelectorAll('h1,h2,h3')).filter(vis).map((h) => h.innerText.trim().slice(0, 80)).slice(0, 40),
    buttons: grab('button, [role=button], input[type=button], input[type=submit]').slice(0, 80),
    links: grab('a[href]').slice(0, 80),
    tabs: grab('[role=tab], nav a, [class*=nav] a').slice(0, 60),
    inputs: grab('input, textarea, select').slice(0, 60),
    fileInputs: grab('input[type=file]'),
    iframes: Array.from(document.querySelectorAll('iframe')).map((f) => f.getAttribute('src')).slice(0, 20),
  };
}
"""


def capture(page: Page, tag: str) -> dict[str, Any]:
    """Full-page screenshot + a JSON inventory of the interactive elements on the
    current page, written under RECON_DIR. Returns the inventory dict."""
    _ensure_dirs()
    safe = re.sub(r"[^A-Za-z0-9_.-]", "_", tag)[:60]
    shot = RECON_DIR / f"{safe}.png"
    try:
        page.screenshot(path=str(shot), full_page=True)
    except Exception:
        page.screenshot(path=str(shot))  # some apps break full_page; fall back to viewport
    try:
        inv = page.evaluate(_INVENTORY_JS)
    except Exception as e:
        inv = {"error": str(e)[:200], "url": page.url, "title": page.title()}
    inv["_screenshot"] = str(shot)
    (RECON_DIR / f"{safe}.json").write_text(json.dumps(inv, indent=2, ensure_ascii=False), encoding="utf-8")
    try:
        html = page.content()
        (RECON_DIR / f"{safe}.html").write_text(html[:1_000_000], encoding="utf-8")
    except Exception:
        pass
    return inv


# --- interactive login + tour -------------------------------------------------
def login(init_url: str, timeout_s: int = 900, poll_s: float = 3.0) -> dict[str, Any]:
    """Open a headed Edge window at ``init_url`` and wait for the user to finish
    logging in. Success = the browser lands on a Mind app page that is neither an
    identity-provider page nor the one-time init link (or the SIGNAL_FILE
    appears). Saves the resolved project URL and captures the landing page.
    """
    _ensure_dirs()
    if SIGNAL_FILE.exists():
        SIGNAL_FILE.unlink()
    result: dict[str, Any] = {"ok": False}
    with mind_session(headed=True) as (_ctx, page):
        page.goto(init_url, wait_until="domcontentloaded")
        print(f"[login] opened {init_url}", flush=True)
        print("[login] waiting for you to complete sign-in in the Edge window…", flush=True)
        deadline = time.time() + timeout_s
        last_url = ""
        stable_since: float | None = None
        while time.time() < deadline:
            url = page.url
            if url != last_url:
                print(f"[login] url: {url}", flush=True)
                last_url = url
                stable_since = None
            authed = _looks_authenticated(page, init_url)
            if SIGNAL_FILE.exists():
                print("[login] signal file seen — treating as logged in", flush=True)
                break
            if authed:
                # require the app URL to hold steady briefly so we don't catch a redirect hop
                stable_since = stable_since or time.time()
                if time.time() - stable_since >= 8:
                    print("[login] authenticated app page is stable", flush=True)
                    break
            try:
                page.wait_for_timeout(int(poll_s * 1000))
            except Exception:
                time.sleep(poll_s)
        else:
            print("[login] timed out waiting for sign-in", flush=True)

        # settle, then capture wherever we ended up
        try:
            page.wait_for_load_state("networkidle", timeout=15_000)
        except Exception:
            pass
        final_url = page.url
        authed = _looks_authenticated(page, init_url)
        inv = capture(page, "landing")
        result = {
            "ok": bool(authed),
            "resolved_url": final_url,
            "title": inv.get("title"),
            "init_url": init_url,
            "captured_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        }
        if authed:
            _save_session({"project_url": final_url, "init_url": init_url, "last_login": result["captured_at"]})
        print("[login] result: " + json.dumps(result), flush=True)
    return result


def is_authenticated() -> dict[str, Any]:
    """Headless check that the saved project URL loads without bouncing to login."""
    sess = load_session()
    url = sess.get("project_url")
    if not url:
        return {"authenticated": False, "reason": "no saved project_url; run login first"}
    with mind_session(headed=False) as (_ctx, page):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45_000)
            try:
                page.wait_for_load_state("networkidle", timeout=10_000)
            except Exception:
                pass
            ok = _looks_authenticated(page, sess.get("init_url"))
            return {"authenticated": bool(ok), "url": page.url, "title": page.title()}
        except Exception as e:
            return {"authenticated": False, "error": str(e)[:200]}


# --- project operations (mapped on the live site) -----------------------------
MAIN_URL = SETTINGS.mind_main_url
MANAGER_HOST = "prod01"
# the sandbox folder the app creates its projects in, and the only place it deletes from
SANDBOX_FOLDER = SETTINGS.mind_sandbox_folder
PROJECT_PREFIX = SETTINGS.mind_project_prefix
# Angular ngx-guided-tour: a full-screen mask that intercepts every click and
# re-arms on each page load. Removed before any interaction.
_TOUR_MASK = "ngx-guided-tour,.guided-tour-user-input-mask,.guided-tour-spotlight-overlay"
_UNTITLED = re.compile(r"^Untitled\((\d+),(\d+)\)$")
_STEP_STATUS_JS = r"""()=>{const o={};[...document.querySelectorAll('[class*=progress-step]')].forEach(b=>{const t=(b.innerText||'').split('\n').map(s=>s.trim()).filter(Boolean);if(t.length>=2)o[t[1]]=t[0];});return o;}"""
_PROJECT_HREF_JS = "(nm)=>{const a=[...document.querySelectorAll('a.holder')].find(a=>{const p=(a.innerText||'').trim().split(String.fromCharCode(10));return p[p.length-1].trim()===nm;});return a?a.getAttribute('href'):null;}"
# The tile's own '...' menu: a descendant of the matched tile, or of a wrapper
# that holds exactly that one tile. Never the nearest button on the page -- a
# neighbour's menu must not be opened on this project's behalf.
_MARK_MENU_JS = r"""(nm)=>{
  const a=[...document.querySelectorAll('a.holder')].find(h=>{const p=(h.innerText||'').trim().split(String.fromCharCode(10));return p[p.length-1].trim()===nm;});
  if(!a) return 'nf';
  let best=a.querySelector('.menu-button');
  if(!best){const w=a.closest('app-manager-project-item')||a.parentElement;
    if(w&&w.querySelectorAll('a.holder').length===1) best=w.querySelector('.menu-button');}
  if(!best||best.getBoundingClientRect().width===0) return 'no-menu';
  document.querySelectorAll('[data-del]').forEach(e=>e.removeAttribute('data-del'));
  best.setAttribute('data-del','1'); return 'ok';}"""


def kill_tour(page: Page) -> None:
    try:
        page.evaluate(f"()=>document.querySelectorAll('{_TOUR_MASK}').forEach(e=>e.remove())")
    except Exception:
        pass


def body_text(page: Page) -> str:
    try:
        return page.evaluate("()=>(document.body.innerText||'').trim()")
    except Exception:
        return ""


def _wait(page: Page, ms: int) -> None:
    try:
        page.wait_for_timeout(ms)
    except Exception:
        time.sleep(ms / 1000)


def _settle(page: Page, host: str, timeout_s: int = 55) -> bool:
    end = time.time() + timeout_s
    while time.time() < end:
        if host in page.url and "code=" not in page.url and len(body_text(page)) > 40:
            _wait(page, 1500)
            return True
        _wait(page, 1500)
    return False


def screenshot(page: Page, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(path), full_page=True)
    except Exception:
        try:
            page.screenshot(path=str(path))
        except Exception:
            pass


def open_manager(page: Page) -> bool:
    """shared.milliman-mind.com -> 'Access the Project manager' -> the production
    Project Manager (prod01). Lands on the workspace root ('Projects')."""
    page.goto(MAIN_URL, wait_until="domcontentloaded")
    _wait(page, 1200)
    try:
        page.get_by_text(re.compile(r"Access the Project manager", re.I)).first.click(timeout=8000)
    except Exception:
        pass
    ok = _settle(page, MANAGER_HOST)
    _wait(page, 1200)
    kill_tour(page)
    try:
        page.get_by_role("link", name=re.compile(r"^\s*Projects\s*$")).first.click(timeout=4000)
        _wait(page, 1000)
        kill_tour(page)
    except Exception:
        pass
    return ok


def ensure_folder(page: Page, folder: str = SANDBOX_FOLDER) -> bool:
    """Open `folder` in the manager, creating it when it does not exist."""
    exists = False
    for _ in range(6):  # the manager's list may still be loading right after a run; look again before creating anything
        exists = page.evaluate("(f)=>[...document.querySelectorAll('a')].some(e=>((e.innerText||'').trim().startsWith(f)))", folder)
        if exists:
            break
        _wait(page, 1500)
    if not exists:
        try:
            page.get_by_role("button", name=re.compile(r"^\s*New folder\s*$", re.I)).first.click(timeout=6000)
            _wait(page, 1200)
            kill_tour(page)
            page.locator('input[placeholder="Choose a folder name"]').first.fill(folder, timeout=5000)
            page.get_by_role("button", name=re.compile(r"^\s*Create new folder\s*$", re.I)).first.click(timeout=5000)
            _wait(page, 2500)
            kill_tour(page)
        except Exception:
            return False
    try:
        page.get_by_role("link", name=re.compile(rf"^\s*{re.escape(folder)}")).first.click(timeout=6000)
        _wait(page, 1500)
        kill_tour(page)
        return True
    except Exception:
        return False


def create_blank_project(page: Page, name: str) -> str | None:
    """New project -> name -> Create project. Returns the project's /init link
    (the tile is target=_blank, so it is opened by goto, never by click)."""
    existing = page.evaluate(_PROJECT_HREF_JS, name)
    if existing:
        return existing
    page.get_by_role("button", name=re.compile(r"^\s*New project\s*$", re.I)).first.click(timeout=6000)
    _wait(page, 1200)
    kill_tour(page)
    page.locator('input[placeholder="Choose a project name"]').first.fill(name, timeout=5000)
    page.get_by_role("button", name=re.compile(r"^\s*Create project\s*$", re.I)).first.click(timeout=5000)
    _wait(page, 4000)
    kill_tour(page)
    return page.evaluate(_PROJECT_HREF_JS, name)


def open_project(page: Page, init_url: str, timeout_s: int = 90) -> str:
    """Load a project and wait until it is usable. Returns 'empty' (drop zone
    present), 'model' (Run button present) or 'unknown'."""
    page.goto(init_url, wait_until="domcontentloaded")
    end = time.time() + timeout_s
    while time.time() < end:
        if page.locator("input#file").count():
            _wait(page, 1500)
            kill_tour(page)
            return "empty"
        if page.get_by_role("button", name=re.compile(r"^\s*Run\s*$")).count():
            _wait(page, 1500)
            kill_tour(page)
            return "model"
        _wait(page, 1500)
    kill_tour(page)
    return "unknown"


def step_status(page: Page) -> dict[str, str]:
    """{'Upload': 'Complete', 'Convert': 'Invalid', 'Test': 'Incomplete'} from the wizard."""
    try:
        return page.evaluate(_STEP_STATUS_JS)
    except Exception:
        return {}


def read_error_logs(page: Page) -> list[str]:
    """Click 'Display Error Logs' / 'View logs' and return the errors, each as one line:
    '<kind>: <message lines> | MODEL ... | <model> > <instance> > <sheet>!<cell> | <mind ref>'
    (the whole block of each error, not just its first line)."""
    btn = page.get_by_role("button", name=re.compile(r"Display Error Logs|View logs", re.I))
    if not btn.count():
        return []
    try:
        btn.first.click(timeout=5000)
    except Exception:
        pass
    _wait(page, 2500)
    text = body_text(page)
    out: list[str] = []
    seen: set[str] = set()
    for block in re.split(r"(?=#\d+\s+[A-Za-z][^\n]*\n)", text):
        m = re.match(r"#\d+\s+([^\n]+)\n(.*)", block, re.S)
        if not m:
            continue
        body = re.split(r"\n\s*(?:Close|Hide Error Logs)\s*(?:\n|$)", m.group(2))[0]
        lines = [ln.strip() for ln in body.splitlines() if ln.strip()]
        msg = (m.group(1).strip() + ": " + " | ".join(lines)) if lines else m.group(1).strip()
        if msg not in seen:
            seen.add(msg)
            out.append(msg)
    return out


def upload_and_convert(page: Page, workbook: Path, timeout_s: int = 420, shots: Path | None = None) -> dict[str, Any]:
    """Drop `workbook` on the empty project and wait for the Upload -> Convert
    -> Test wizard. {'success', 'steps', 'errors', 'text'}."""
    page.set_input_files("input#file", str(workbook))
    end = time.time() + timeout_s
    failed = False
    while time.time() < end:
        text = body_text(page)
        steps = step_status(page)
        if page.get_by_role("button", name=re.compile(r"Display Error Logs", re.I)).count() or "Conversion or tests failed" in text:
            failed = True
            break
        if "Upload successful" in text or (steps and len(steps) >= 3 and all(v.lower() in ("complete", "success", "valid") for v in steps.values())):
            break
        _wait(page, 5000)
    _wait(page, 1500)
    if shots:
        screenshot(page, shots / "convert.png")
    steps = step_status(page)
    errors = read_error_logs(page) if failed else []
    if shots and failed:
        screenshot(page, shots / "convert_errors.png")
    text = body_text(page)
    success = (not failed) and ("Upload successful" in text or (bool(steps) and all(v.lower() in ("complete", "success", "valid") for v in steps.values())))
    return {"success": success, "steps": steps, "errors": errors, "timed_out": time.time() >= end, "text": text[:600]}


def add_template(page: Page, timeout_s: int = 150) -> bool:
    """'Add template to Milliman Mind' after a successful conversion; True once
    the project shows its Run button."""
    try:
        page.get_by_role("button", name=re.compile(r"Add template to Milliman Mind", re.I)).first.click(timeout=8000)
    except Exception:
        return False
    end = time.time() + timeout_s
    while time.time() < end:
        if page.get_by_role("button", name=re.compile(r"^\s*Run\s*$")).count():
            _wait(page, 1500)
            kill_tour(page)
            return True
        _wait(page, 3000)
    return False


def run_model(page: Page, timeout_s: int = 300, shots: Path | None = None) -> dict[str, Any]:
    """Click Run and wait. {'completed', 'audit_consistent', 'seconds', 'text'}."""
    btn = page.get_by_role("button", name=re.compile(r"^\s*Run\s*$"))
    if not btn.count():
        return {"completed": False, "audit_consistent": False, "seconds": None, "text": "no Run button"}
    btn.first.click(timeout=8000)
    _wait(page, 2000)
    kill_tour(page)
    if page.locator(".cds--modal.is-visible").count():
        for pat in (r"^\s*Run\s*$", r"Start", r"Confirm", r"^\s*OK\s*$", r"Launch"):
            try:
                b = page.locator(".cds--modal.is-visible").get_by_role("button", name=re.compile(pat, re.I))
                if b.count() and b.first.is_visible():
                    b.first.click(timeout=4000)
                    break
            except Exception:
                pass
    end = time.time() + timeout_s
    text = ""
    while time.time() < end:
        text = body_text(page)
        if re.search(r"calculations completed|run completed|consistent with the audit", text, re.I):
            break
        if re.search(r"\bfailed\b|\berror\b", text, re.I) and "Display" not in text:
            break
        _wait(page, 4000)
    _wait(page, 1500)
    if shots:
        screenshot(page, shots / "run.png")
    text = body_text(page)
    m = re.search(r"completed in ([\d.]+) seconds", text, re.I)
    return {
        "completed": bool(re.search(r"calculations completed|run completed", text, re.I)),
        "audit_consistent": bool(re.search(r"consistent with the audit trail", text, re.I)),
        "seconds": float(m.group(1)) if m else None,
        "text": text[:600],
    }


def template_names(page: Page, stem: str) -> list[str]:
    """Every entry of the template selector (the `<stem> ▾` button in the top
    bar): templates and their grids by name -- 'Untitled(row,col)' for a grid
    with no title. Empty list when the selector cannot be opened."""
    opened = False
    for attempt in range(4):
        kill_tour(page)
        for loc in (
            page.get_by_role("button", name=re.compile(rf"^\s*{re.escape(stem)}\s*$", re.I)),
            page.locator(f'button:has-text("{stem}")'),
            page.get_by_text(re.compile(rf"^\s*{re.escape(stem)}\s*$")),
        ):
            try:
                if loc.count():
                    loc.first.click(timeout=4000)
                    opened = True
                    break
            except Exception:
                continue
        if opened:
            break
        _wait(page, 2500)
    if not opened:
        return []
    _wait(page, 1500)
    try:
        names = page.evaluate(
            r"""()=>[...document.querySelectorAll('[class*=menu] *,[role=menuitem],li,option')].map(e=>(e.innerText||'').trim()).filter(t=>t&&t.length<80&&!t.includes(String.fromCharCode(10)))"""
        )
    except Exception:
        names = []
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    seen: set[str] = set()
    out: list[str] = []
    for n in names:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def untitled_entries(names: list[str]) -> list[dict[str, Any]]:
    """The 'Untitled(row,col)' entries of a template list, decoded to a cell ref."""
    out = []
    for n in names:
        m = _UNTITLED.match(n.strip())
        if m:
            row, col = int(m.group(1)), int(m.group(2))
            letters = ""
            c = col
            while c > 0:
                c, rem = divmod(c - 1, 26)
                letters = chr(65 + rem) + letters
            out.append({"label": n, "row": row, "col": col, "cell": f"{letters}{row}"})
    return out


def export_summary(page: Page) -> dict[str, Any]:
    """Open the Export manager (top-right 'Exports' after a run) and count the
    exportable elements. {'opened', 'elements', 'tabs'}."""
    exp = page.get_by_role("button", name=re.compile(r"^\s*Exports?\s*$", re.I))
    if not exp.count():
        exp = page.locator("[aria-label*='Export' i],[title*='Export' i]")
    if not exp.count():
        return {"opened": False, "elements": None, "tabs": []}
    try:
        exp.first.click(timeout=6000)
    except Exception:
        return {"opened": False, "elements": None, "tabs": []}
    _wait(page, 2500)
    kill_tour(page)
    tabs = []
    elements = 0
    for tab in ("Export flat files", "Export Excel files"):
        try:
            t = page.get_by_text(re.compile(rf"^\s*{re.escape(tab)}\s*$", re.I))
            if t.count():
                t.first.click(timeout=3000)
                _wait(page, 1500)
        except Exception:
            pass
        try:
            rows = page.evaluate("()=>[...document.querySelectorAll('table tr')].length")
        except Exception:
            rows = 0
        tabs.append({"tab": tab, "data_rows": max(rows - 1, 0)})
        elements = max(elements, max(rows - 1, 0))
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    return {"opened": True, "elements": elements, "tabs": tabs}


def project_exists(page: Page, name: str) -> bool:
    return bool(page.evaluate(_PROJECT_HREF_JS, name))


def delete_project(page: Page, name: str, folder: str = SANDBOX_FOLDER, attempts: int = 2) -> str:
    """Best-effort delete of a sandbox project through the tile menu and the
    type-'delete'-to-confirm modal. Refuses any name without the sandbox
    project prefix the app is allowed to create and delete. Returns 'deleted' | 'still present' | 'not found' | 'refused'."""
    if not name.startswith(PROJECT_PREFIX):
        return "refused"
    result = "not found"
    for _ in range(attempts):
        open_manager(page)
        ensure_folder(page, folder)
        if not project_exists(page, name):
            return result if result != "not found" else "not found"
        st = page.evaluate(_MARK_MENU_JS, name)
        if st != "ok":
            result = "still present"
            continue
        try:
            page.locator("[data-del='1']").first.click(timeout=4000)
            _wait(page, 1000)
            page.get_by_text(re.compile(r"^\s*Delete( project)?\s*$", re.I)).first.click(timeout=3000)
            _wait(page, 1200)
            # the confirmation modal names its target ('Delete "<name>"'); only
            # that modal, and only with this exact name, gets the word 'delete'
            modal = page.locator(".cds--modal.is-visible")
            try:
                modal_text = modal.first.inner_text(timeout=3000) if modal.count() else ""
            except Exception:
                modal_text = ""
            if not re.search(rf"(?<![\w.-]){re.escape(name)}(?![\w.-])", modal_text):
                page.keyboard.press("Escape")
                _wait(page, 600)
                return "refused"
            inp = page.locator(".cds--modal.is-visible input")
            if inp.count():
                inp.first.click()
                inp.first.fill("delete", timeout=3000)
                _wait(page, 400)
            danger = page.locator(".cds--modal.is-visible button.cds--btn--danger")
            for _ in range(20):
                if danger.count() and danger.first.is_enabled():
                    break
                _wait(page, 400)
            if not (danger.count() and danger.first.is_enabled()):
                page.keyboard.press("Escape")
                result = "still present"
                continue
            danger.first.click(timeout=3000)
            _wait(page, 2500)
        except Exception:
            result = "still present"
            continue
        open_manager(page)
        ensure_folder(page, folder)
        if not project_exists(page, name):
            return "deleted"
        result = "still present"
    return result


if __name__ == "__main__":
    import sys

    cmd = sys.argv[1] if len(sys.argv) > 1 else "help"
    if cmd == "login":
        if len(sys.argv) < 3:
            print("usage: python -m app.mind_client login <init_url> [timeout_s]")
            raise SystemExit(2)
        t = int(sys.argv[3]) if len(sys.argv) > 3 else 900
        login(sys.argv[2], timeout_s=t)
    elif cmd == "check":
        print(json.dumps(is_authenticated(), indent=2))
    elif cmd == "session":
        print(json.dumps(load_session(), indent=2))
    else:
        print("commands: login <url> [timeout_s] | check | session")
