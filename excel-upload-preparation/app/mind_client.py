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
  -> upload_and_convert -> add_template -> upload_inputs (a model with Mind input
  files only) -> template_names -> run_model -> export_summary -> delete_project
  (best effort)
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
import time
import unicodedata
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


def sign_in(timeout_s: int = 600, progress: Any = None, should_stop: Any = None) -> dict[str, Any]:
    """The app's own "Sign in to Mind" (1.12.1; before, only `login <init link>` from a command line,
    which the packaged program does not have). Opens Edge, visible, with the app's profile at Mind's
    main page and waits until the user has signed in: the page is on the Mind host, not an
    identity-provider page, for 8 s in a row. Saves the main page as the URL later checks load, then
    closes the window. -> {"ok", "reason", "url"}. `progress(message)` hears what happens;
    `should_stop()` ends the wait early (the window is closed either way)."""
    tell = progress or (lambda _m: None)
    stop = should_stop or (lambda: False)
    _ensure_dirs()
    with mind_session(headed=True) as (_ctx, page):
        page.goto(MAIN_URL, wait_until="domcontentloaded")
        tell("An Edge window is open on Mind: sign in there (it closes by itself when you are in).")
        deadline = time.time() + timeout_s
        stable_since: float | None = None
        while time.time() < deadline:
            if stop():
                return {"ok": False, "reason": "stopped", "url": None}
            try:
                url = page.url
            except Exception:  # the user closed the window
                return {"ok": False, "reason": "the Edge window was closed before the sign-in finished", "url": None}
            if _looks_authenticated(page, None):
                stable_since = stable_since or time.time()
                if time.time() - stable_since >= 8:
                    stamp = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
                    _save_session({"project_url": MAIN_URL, "last_login": stamp})
                    tell("Signed in to Mind.")
                    return {"ok": True, "reason": None, "url": url, "last_login": stamp}
            else:
                stable_since = None
            try:
                page.wait_for_timeout(2000)
            except Exception:
                return {"ok": False, "reason": "the Edge window was closed before the sign-in finished", "url": None}
        return {"ok": False, "reason": f"no sign-in within {timeout_s // 60} minutes", "url": None}


def is_authenticated() -> dict[str, Any]:
    """Headless check that the saved project URL (Mind's main page when none was saved) loads
    without bouncing to a sign-in page."""
    sess = load_session()
    url = sess.get("project_url") or MAIN_URL
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
    # since November 2026 Mind asks first for the type of project (Milliman Mind / Milliman OpS)
    name_box = page.locator('input[placeholder="Choose a project name"]').first
    if not name_box.is_visible():
        kind = page.get_by_text(re.compile(r"Milliman Mind blank project", re.I)).first
        try:
            kind.click(timeout=6000)
            _wait(page, 1500)
            kill_tour(page)
        except Exception:
            pass  # no type picker: the name box is just slow to appear
    name_box.fill(name, timeout=15000)
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


_ADD_TEMPLATE = re.compile(r"Add templates? to Milliman Mind", re.I)  # plural when several workbooks went up


def add_template(page: Page, timeout_s: int = 150) -> bool:
    """'Add template to Milliman Mind' ('Add templates ...' after several workbooks) once the
    conversion succeeded; True once the project shows its Run button."""
    try:
        page.get_by_role("button", name=_ADD_TEMPLATE).first.click(timeout=8000)
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


# --- the Input Manager: Mind input files (outside files made Mind inputs) ----------------------
# Mapped on the live site 2026-11-08: a model with /Input grids has an "Inputs" button left of Run;
# it opens a window "Inputs" listing model > sheet > grid with Status and File columns, Culture and
# CSV Separator boxes and a drop zone whose <input type=file webkitdirectory> takes a FOLDER. Mind
# matches each file to the grid of the same name (grid IN_X <- IN_X.csv), the row then shows
# SUCCESS and the file name; nothing to save, the next run uses the imported values.
INPUT_UPLOAD_SUFFIXES = (".csv", ".txt")
_INPUTS_BUTTON = re.compile(r"^\s*Inputs\s*$")
_INPUTS_DROP = "Upload inputs or Drag"  # the window's drop zone: present while the window is open
_UNMATCHED = "Match unmatched"  # Mind's popup for a file no grid name matches
_INPUTS_HEADER = re.compile(r"^\s*Element\b.*\bStatus\b.*\bFile\b")
_INPUT_FILE_NAME = re.compile(r"\.(?:csv|txt|zip|parquet)$", re.I)
# a grid row's status while it has no file or Mind is still taking it
INPUT_PENDING_STATUSES = frozenset({"", "-", "NO FILE", "PENDING", "UPLOADING", "IN PROGRESS", "PROCESSING", "LOADING",
                                    "QUEUED", "WAITING", "VALIDATING", "IMPORTING", "RUNNING", "STARTED"})
_INPUT_STATUS_WORDS = INPUT_PENDING_STATUSES | {"SUCCESS", "ERROR", "WARNING", "FAILED", "FAILURE", "INVALID",
                                                "CANCELLED", "CANCELED", "DONE", "COMPLETE", "COMPLETED"}
# cultures that read numbers as the input files write them (a decimal point, no thousands mark;
# proven with English - United Kingdom). Under another one (German: '.' groups thousands) 1000.5
# could silently become 10005, so the files are not uploaded at all.
_POINT_CULTURE = re.compile(r"^\s*(?:English(?!.*South Africa)|Japanese|Chinese|Korean|Hebrew)\b", re.I)
_VIS_JS = "const vis=(e)=>{const r=e.getBoundingClientRect();const s=getComputedStyle(e);return r.width>0&&r.height>0&&s.visibility!=='hidden'&&s.display!=='none';};"
# the smallest visible window (modal / dialog) whose text holds `marker`
_BOX_FIND_JS = ("const boxes=[...document.querySelectorAll('.cds--modal,[role=dialog],[role=alertdialog]')].filter(vis)"
                ".filter(b=>(b.innerText||'').toLowerCase().includes(marker.toLowerCase()));"
                "boxes.sort((a,b)=>(a.innerText||'').length-(b.innerText||'').length);const box=boxes[0]||null;")
# its text, and (mark=true) its Cancel / Close button tagged for a click -- never any other button
_BOX_JS = ("([marker,mark])=>{" + _VIS_JS + _BOX_FIND_JS +
           "if(!box) return {text:null,marked:'nf'};let marked='no';"
           "if(mark){const bs=[...box.querySelectorAll('button')].filter(vis);"
           "const pick=bs.find(b=>/^\\s*(cancel|close)\\s*$/i.test(b.innerText||''))||"
           "bs.find(b=>/close/i.test((b.getAttribute('aria-label')||'')+' '+(b.getAttribute('title')||'')));"
           "document.querySelectorAll('[data-mr-close]').forEach(e=>e.removeAttribute('data-mr-close'));"
           "if(pick){pick.setAttribute('data-mr-close','1');marked='ok';}}"
           "return {text:box.innerText||'',marked};}")
# the Inputs window's table: one list of cell texts per visible row (the header row included)
_INPUT_ROWS_JS = ("()=>{" + _VIS_JS + "const marker='" + _INPUTS_DROP + "';" + _BOX_FIND_JS +
                  "const root=box||document.body;let rows=[...root.querySelectorAll('tr,[role=row]')];"
                  "if(!rows.length) rows=[...root.querySelectorAll('*')].filter(e=>getComputedStyle(e).display==='table-row');"
                  "return rows.filter(vis).map(r=>[...r.children].filter(c=>/^(TD|TH)$/.test(c.tagName)||"
                  "/cell|header/.test(c.getAttribute('role')||'')||getComputedStyle(c).display==='table-cell')"
                  ".map(c=>(c.innerText||'').replace(/\\s+/g,' ').trim()));}")
_INPUT_SETTINGS_JS = ("()=>{const t=(id)=>{const s=document.getElementById(id);if(!s||s.selectedIndex<0)return null;"
                      "const o=s.options[s.selectedIndex];return o?(o.text||'').trim():null;};"
                      "return {culture:t('culture'),separator:t('delimiter')};}")


def _norm(text: Any) -> str:
    return unicodedata.normalize("NFKC", str(text or "")).casefold().strip()


def _has(text: str, marker: str) -> bool:
    return marker.casefold() in (text or "").casefold()


def input_files(folder: Path | str | None) -> list[Path]:
    """The Mind input files of `folder`: its .csv / .txt files (not in subfolders, no lock or
    hidden files), by name."""
    if not folder:
        return []
    p = Path(folder)
    if not p.is_dir():
        return []
    return sorted((f for f in p.iterdir() if f.is_file() and f.suffix.lower() in INPUT_UPLOAD_SUFFIXES
                   and not f.name.startswith(("~$", "."))), key=lambda f: f.name.casefold())


def _settled(status: str | None) -> bool:
    s = (status or "").strip().upper()
    return s not in INPUT_PENDING_STATUSES and not s.endswith(("...", "…"))


def parse_input_rows(rows: list[list[Any]] | None) -> list[dict[str, Any]]:
    """The Inputs window's table (cell texts per row, header first) -> one dict per grid row:
    {model, sheet, grid, status, file}. The header names the columns (Element, Status, File); a
    row with neither a status nor a file is a model or sheet heading (the first one the model)."""
    out: list[dict[str, Any]] = []
    cols: tuple[int, int, int] | None = None
    model = sheet = None
    for raw in rows or []:
        cells = [" ".join(str(c or "").split()) for c in (raw or [])]
        low = [c.casefold() for c in cells]
        if cols is None:
            if "status" in low and "file" in low:
                cols = (low.index("element") if "element" in low else 0, low.index("status"), low.index("file"))
            continue
        e, s, f = cols
        name, status, file = (cells[i] if i < len(cells) else "" for i in (e, s, f))
        if status.casefold() == "status" and file.casefold() == "file":
            continue  # the header again (a sticky copy)
        if not (status or file):
            heading = next((c for c in cells if c), "")
            if heading:
                if model is None:
                    model = heading
                else:
                    sheet = heading
            continue
        if name:
            out.append({"model": model, "sheet": sheet, "grid": name, "status": status,
                        "file": None if file.casefold() in ("", "-", "no file") else file})
    return out


def parse_inputs_text(text: str) -> list[dict[str, Any]]:
    """The same rows read from the page text (when the table cannot be read): after the
    'Element  DL  Status  File ...' header, a grid name is followed by its status (SUCCESS, -,
    ...) and its file ('No file' or a name); the other names are the model and sheet headings."""
    lines = (text or "").splitlines()
    start = next((i for i, ln in enumerate(lines) if _INPUTS_HEADER.match(ln)), None)
    if start is None:
        return []
    tokens: list[str] = []
    for ln in lines[start + 1:]:
        if ln.strip() in ("Culture", "CSV Separator") or _has(ln, _INPUTS_DROP):
            break
        tokens += [t.strip() for t in ln.split("\t") if t.strip()]

    def is_status(t: str) -> bool:
        return t == "-" or (t == t.upper() and t in _INPUT_STATUS_WORDS)

    out: list[dict[str, Any]] = []
    model = sheet = None
    i = 0
    while i < len(tokens):
        if i + 1 < len(tokens) and is_status(tokens[i + 1]):
            file, j = None, i + 2
            if j < len(tokens) and (tokens[j].casefold() == "no file" or _INPUT_FILE_NAME.search(tokens[j])):
                file = None if tokens[j].casefold() == "no file" else tokens[j]
                j += 1
            out.append({"model": model, "sheet": sheet, "grid": tokens[i], "status": tokens[i + 1], "file": file})
            i = j
            continue
        if model is None:
            model = tokens[i]
        else:
            sheet = tokens[i]
        i += 1
    return out


def _read_input_rows(page: Page, text: str | None = None) -> list[dict[str, Any]]:
    try:
        rows = parse_input_rows(page.evaluate(_INPUT_ROWS_JS))
    except Exception:
        rows = []
    return rows or parse_inputs_text(body_text(page) if text is None else text)


def _row_for(rows: list[dict[str, Any]], name: str) -> dict[str, Any] | None:
    """The grid row that took the file `name`: by the name in its File column (with or without
    the extension), else the grid of the same name (Mind matches a file to the grid it is named
    after) unless that grid shows another file."""
    full, stem = _norm(name), _norm(Path(name).stem)
    for r in rows:
        f = re.split(r"[\\/]", str(r.get("file") or ""))[-1]
        if f and (_norm(f) == full or _norm(Path(f).stem) == stem):
            return r
    for r in rows:
        f = re.split(r"[\\/]", str(r.get("file") or ""))[-1]
        if _norm(r.get("grid")) == stem and not _INPUT_FILE_NAME.search(f):
            return r
    return None


def _inputs_settled(rows: list[dict[str, Any]], expect: list[str]) -> bool:
    """Mind has finished with the files: each one that has a row (a row showing its name, or the
    grid named after it -- listed from the moment the window opens) shows a final status. A file
    with no row has no grid to go to: waiting longer changes nothing."""
    return bool(rows) and all((r := _row_for(rows, n)) is None or _settled(r["status"]) for n in expect)


def _inputs_box(page: Page, marker: str, mark: bool = False) -> dict[str, Any]:
    try:
        return page.evaluate(_BOX_JS, [marker, mark]) or {}
    except Exception:
        return {}


def _close_box(page: Page, marker: str, tries: int = 3) -> bool:
    """Close the window whose text holds `marker` with its own Cancel / Close button (Escape
    when it has none). True once that text has left the page."""
    for _ in range(tries):
        if not _has(body_text(page), marker):
            return True
        kill_tour(page)
        clicked = False
        if _inputs_box(page, marker, mark=True).get("marked") == "ok":
            try:
                page.locator("[data-mr-close='1']").first.click(timeout=4000)
                clicked = True
            except Exception:
                pass
        if not clicked:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
        _wait(page, 1500)
    return not _has(body_text(page), marker)


def _inputs_excerpt(text: str, popup: str = "") -> str:
    """The window's rows as one line (the Culture list left out), the popup's text first."""
    lines = (text or "").splitlines()
    start = next((i for i, ln in enumerate(lines) if _INPUTS_HEADER.match(ln)), 0)
    keep = []
    for ln in lines[start:]:
        if ln.strip() in ("Culture", "CSV Separator") or _has(ln, _INPUTS_DROP):
            break
        if ln.strip():
            keep.append(" ".join(ln.split()))
    s = " | ".join(keep)
    if popup:
        s = " ".join(popup.split())[:300] + " || " + s
    return s[:600]


def upload_inputs(page: Page, folder: Path, timeout_s: int = 300, shots: Path | None = None) -> dict[str, Any]:
    """Upload the Mind input files of `folder` (.csv / .txt) through the project's Inputs window
    (the Input Manager), once the model has been added. Culture and CSV Separator are never
    changed; when they would not read the files as written (a decimal point; Default separator)
    nothing is uploaded. Mind matches each file to the /Input grid of the same name; a file it
    cannot match opens "Match unmatched inputs": the file is reported and the popup cancelled --
    never a guessed match. The window is closed at the end (nothing to save).

    -> {"ok": every file is on its grid with SUCCESS, "statuses": {grid: status} for every grid
    listed, "unmatched": [files no grid took], "failed": {grid: status} for a file Mind did not
    accept, "waiting": [files still not taken at the time limit], "rows": [{model, sheet, grid,
    status, file}], "files": [names uploaded], "reason": one plain sentence when not ok, "text": the
    window's rows (first 600 characters), "timed_out", "closed", "seconds", "culture", "separator"}."""
    t0 = time.time()
    folder = Path(folder)
    files = input_files(folder)
    res: dict[str, Any] = {"ok": False, "statuses": {}, "unmatched": [], "failed": {}, "waiting": [], "rows": [],
                           "files": [f.name for f in files], "reason": None, "text": "", "timed_out": False,
                           "closed": True, "seconds": None, "culture": None, "separator": None}
    if not files:
        res.update(reason=f"no .csv or .txt file in {folder}", seconds=0.0)
        return res
    # the input box takes a whole folder: anything else in it (a zip, a subfolder) would go up too,
    # so a folder holding more than the input files is copied, input files only, to a clean one
    stage = None
    upload_dir = folder
    if any(p not in files for p in folder.iterdir()):
        stage = tempfile.TemporaryDirectory(prefix="mind_inputs_", ignore_cleanup_errors=True)
        upload_dir = Path(stage.name) / (folder.name or "inputs")
        upload_dir.mkdir(parents=True)
        for f in files:
            shutil.copy2(f, upload_dir / f.name)
    try:
        _upload_inputs(page, upload_dir, res, t0 + timeout_s, timeout_s, shots)
    finally:
        res["seconds"] = round(time.time() - t0, 1)
        if stage is not None:  # only now: the browser reads the files while it uploads them
            stage.cleanup()
    return res


def _upload_inputs(page: Page, upload_dir: Path, res: dict[str, Any], deadline: float, timeout_s: int, shots: Path | None) -> None:
    names: list[str] = res["files"]
    # 1. open the window: the Inputs button (only a model with /Input grids has one)
    open_by = min(deadline, time.time() + 30)
    seen_button = False
    last_click = None
    while not _has(body_text(page), _INPUTS_DROP):
        if time.time() >= open_by:
            res["closed"] = _close_box(page, _INPUTS_DROP)
            res["reason"] = ("Mind's Inputs window did not open" if seen_button else
                             "Mind shows no Inputs button: the model in Mind has no /Input grid")
            return
        kill_tour(page)
        btn = page.get_by_role("button", name=_INPUTS_BUTTON)
        if btn.count():
            seen_button = True
            if last_click is None or time.time() - last_click >= 8:
                last_click = time.time()
                try:
                    btn.first.click(timeout=5000)
                except Exception:
                    pass
        _wait(page, 2000)
    kill_tour(page)
    try:
        sel = page.evaluate(_INPUT_SETTINGS_JS) or {}
        res["culture"], res["separator"] = sel.get("culture"), sel.get("separator")
    except Exception:
        pass
    if shots:
        screenshot(page, shots / "inputs_before.png")
    # Culture and CSV Separator stay as Mind shows them; the files are only handed over when those
    # read them as written (one header row, a decimal point, commas -- TAB in a .txt)
    culture, separator = res["culture"], res["separator"]
    if (culture and not _POINT_CULTURE.match(culture)) or (separator and separator.strip().casefold() != "default"):
        res["reason"] = (f"Mind's Inputs window would read the files with Culture '{culture or '?'}' and CSV Separator "
                         f"'{separator or '?'}': the input files are written for English - United Kingdom and Default, "
                         "so they were not uploaded (their values could change)")
        res["closed"] = _close_box(page, _INPUTS_DROP)
        return
    # 2. hand Mind the folder (the input has webkitdirectory: it must be a folder, not files)
    try:
        page.set_input_files("input[type=file][webkitdirectory]", str(upload_dir))
    except Exception as exc:
        res["reason"] = f"could not hand the folder of input files to Mind: {exc}"[:300]
        res["closed"] = _close_box(page, _INPUTS_DROP)
        return
    # 3. wait until Mind has taken every file (or says it cannot match one)
    expect = list(names)
    popup = ""
    popups = 0
    stuck = None  # why the wait ended before Mind had finished
    rows: list[dict[str, Any]] = []
    text = ""
    while True:
        _wait(page, 3000)
        kill_tour(page)
        text = body_text(page)
        if _has(text, _UNMATCHED):
            popups += 1
            # the files the popup names (its own text only: the page's text also lists matched files)
            now = _inputs_box(page, _UNMATCHED).get("text") or ""
            at = re.search(re.escape(_UNMATCHED), text, re.I)
            popup = popup or now or (text[at.start():at.start() + 300] if at else "")
            expect = [n for n in expect if not re.search(rf"(?<![\w.-]){re.escape(_norm(n))}(?![\w.-])", _norm(now))]
            if shots:
                screenshot(page, shots / "inputs_unmatched.png")
            closed = _close_box(page, _UNMATCHED)  # Cancel: the app never guesses a match
            if not closed and popups >= 3:
                stuck = "Mind's 'Match unmatched inputs' window would not close"
                break
        elif not _has(text, _INPUTS_DROP):
            stuck = "the Inputs window closed before Mind had shown every file"
            break
        else:
            rows = _read_input_rows(page, text)
            if _inputs_settled(rows, expect):
                break
        if time.time() >= deadline:
            res["timed_out"] = True
            break
    if shots:
        screenshot(page, shots / "inputs_after.png")
    # 4. what Mind shows, file by file
    for r in rows:
        key = r["grid"] if r["grid"] not in res["statuses"] else f"{r.get('sheet') or '?'} > {r['grid']}"
        res["statuses"][key] = r["status"] or "-"
    res["rows"] = rows
    unfinished = res["timed_out"] or stuck is not None
    for n in names:
        r = _row_for(rows, n)
        if n not in expect:
            res["unmatched"].append(n)  # named by the popup
        elif r is None or not _settled(r["status"]):
            (res["waiting"] if unfinished else res["unmatched"]).append(n)
        elif r["status"].strip().upper() != "SUCCESS":
            res["failed"][r["grid"]] = r["status"]
    res["text"] = _inputs_excerpt(text, popup)
    parts = [stuck] if stuck else []
    if res["unmatched"]:
        parts.append("Mind found no /Input grid named like " + ", ".join(res["unmatched"]))
    if res["failed"]:
        parts.append("Mind did not accept " + ", ".join(f"{g} ({s})" for g, s in res["failed"].items()))
    if res["timed_out"]:
        parts.append(f"Mind had not finished after {timeout_s} s")
    if res["waiting"]:
        parts.append("still waiting for " + ", ".join(res["waiting"]))
    if not parts and not rows:
        parts.append("the Inputs window's rows could not be read")
    res["reason"] = "; ".join(parts) or None
    res["ok"] = res["reason"] is None and all(
        (r := _row_for(rows, n)) is not None and r["status"].strip().upper() == "SUCCESS" for n in names)
    # 5. close the window: nothing to save, the next run uses the imported values
    if _has(body_text(page), _UNMATCHED):
        _close_box(page, _UNMATCHED)
    res["closed"] = _close_box(page, _INPUTS_DROP)


def close_open_windows(page: Page, tries: int = 4) -> bool:
    """Close any window still covering the project (the Inputs window can stay open: its header
    then takes the click meant for Run). Clicks the VISIBLE "Close modal" button -- a hidden modal
    of the page has one too -- else presses Escape. True once no full-screen window is left."""
    for _ in range(tries):
        if not page.locator(".cds--modal.is-visible, cds-modal.full-screen-cds-modal .cds--modal-header").count():
            return True
        clicked = False
        closers = page.get_by_role("button", name=re.compile(r"Close modal", re.I))
        for j in range(closers.count()):
            try:
                if closers.nth(j).is_visible():
                    closers.nth(j).click(timeout=4000)
                    clicked = True
                    break
            except Exception:
                pass
        if not clicked:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
        _wait(page, 1500)
    return not page.locator(".cds--modal.is-visible").count()


def run_model(page: Page, timeout_s: int = 300, shots: Path | None = None) -> dict[str, Any]:
    """Click Run and wait. {'completed', 'audit_consistent', 'seconds', 'text'}."""
    close_open_windows(page)
    btn = page.get_by_role("button", name=re.compile(r"^\s*Run\s*$"))
    if not btn.count():
        return {"completed": False, "audit_consistent": False, "seconds": None, "text": "no Run button"}
    try:
        btn.first.click(timeout=8000)
    except Exception:
        close_open_windows(page)
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
