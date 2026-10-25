"""Upload a Mind-loop version to a sandbox project in real Mind, run it, and look for the loop.

    python scripts/mind_loop_probe.py <loop version .xlsm> [--project <prefix>_...] [--expected check.json]
        [--parts OUTPUT,Untitled] [--no-run] [--init-url <existing project link>] [--timeout 900]

What it records, under <workbook>_mind_probe/:
  - the conversion wizard (steps, errors), the full logs dialog text, the Deterministic Differences
    lines (Mind's value vs the workbook's cached value, cell by cell: with the MM for Excel add-in
    the cached values are the FIRST loop pass, so every other pass shows up here with Mind's value);
  - the run (completed, consistent with the audit trail);
  - the template selector's entries, and for every entry matching --parts: a screenshot of the
    opened part, every button on the page (text, aria-label, title, class), and a click on each
    control that looks like "expand" / "dimension" / "loop" / "instance", with the text that appears;
  - the Audit trail, Project Edition and Grid validation manager screens, as text + screenshot;
  - when --expected is given (scripts/check_loop_plan.py's check.json): Mind's values from the log
    compared with the macro's values, pass by pass.
With a Mind MCP token (app/mind_mcp.py) the run and the loop reading go through MCP after the upload:
the loop's size and every instance of --source as data, compared with --expected; the run's errors mapped
to Excel cells. Without one (or with --browser) the browser runs the model and reads the parts as before.
Nothing is deleted; the project stays in Mind (sandbox names only: the project prefix in settings.yaml).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import mind_mcp  # noqa: E402
from app.mind_client import (  # noqa: E402
    PROJECT_PREFIX,
    SANDBOX_FOLDER,
    add_template,
    body_text,
    create_blank_project,
    ensure_folder,
    is_authenticated,
    kill_tour,
    mind_session,
    open_manager,
    open_project,
    run_model,
    screenshot,
    template_names,
    upload_and_convert,
)

EXPAND_RE = re.compile(r"expand|dimension|loop|instance|unfold|detail|enlarge|open|maximi|chevron|arrow|more", re.I)
BUTTONS_JS = r"""()=>[...document.querySelectorAll('button,[role=button],a[class*=btn],[class*=icon-button],[class*=expand]')].map((b,i)=>{const r=b.getBoundingClientRect();return {i,t:(b.innerText||'').trim().slice(0,60),a:b.getAttribute('aria-label')||'',ti:b.getAttribute('title')||'',c:(b.className&&b.className.baseVal!==undefined?b.className.baseVal:(b.className||'')).toString().slice(0,80),x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height),vis:r.width>0&&r.height>0}}).filter(b=>b.vis)"""
DIFF_RE = re.compile(r"^(?P<wb>.+?)\s*>\s*(?P<sheet>[^>]+?)\s*>\s*(?P<cell>[A-Z]{1,3}\d+)\s*:\s*Excel value:\s*(?P<excel>.*?)\s*Mind value:\s*(?P<mind>.*)$", re.I)


def _log(out: Path, name: str, data) -> None:
    out.mkdir(parents=True, exist_ok=True)
    (out / name).write_text(json.dumps(data, indent=1, ensure_ascii=False) if not isinstance(data, str) else data, encoding="utf-8")


def _wait(page, ms: int) -> None:
    try:
        page.wait_for_timeout(ms)
    except Exception:
        time.sleep(ms / 1000)


def read_logs(page, out: Path) -> dict:
    """Display Logs -> the inline log panel (Stochastic / Deterministic / Assumptions counts, and the
    Deterministic Differences when the panel lists them). Closes it with "Hide Logs" only: the wizard's
    "Close" button next to it would drop the converted template before it is added."""
    res: dict = {"opened": False, "summary": {}, "differences": [], "text": ""}
    btn = page.get_by_role("button", name=re.compile(r"^\s*Display Logs\s*$|^\s*View logs\s*$|Display Error Logs", re.I))
    if not btn.count():
        return res
    try:
        btn.first.click(timeout=8000)
    except Exception:
        return res
    _wait(page, 3000)
    res["opened"] = True
    screenshot(page, out / "logs_open.png")
    text = body_text(page)
    res["text"] = text[:30000]
    for m in re.finditer(r"(Stochastic|Deterministic|Assumptions)\s*\n\s*(\d+)\s+success\s*\n\s*([^\n]*)\n\s*(?:[^\n]*\n\s*)?(\d+)\s+difference", text):
        res["summary"][m.group(1)] = {"success": int(m.group(2)), "state": m.group(3).strip(), "differences": int(m.group(4))}
    for ln in text.splitlines():
        ln = ln.strip()
        m = DIFF_RE.match(ln)
        if m:
            res["differences"].append({"sheet": m.group("sheet").strip(), "cell": m.group("cell").upper(), "excel": m.group("excel").strip(), "mind": m.group("mind").strip(), "raw": ln})
        elif "Mind value" in ln:
            res["differences"].append({"raw": ln})
    # sections that can be clicked open (the Deterministic block, the differences list)
    for label in ("Deterministic", "difference"):
        try:
            loc = page.get_by_text(re.compile(rf"^\s*\d*\s*{label}", re.I))
            if loc.count():
                loc.first.click(timeout=3000)
                _wait(page, 2000)
                more = body_text(page)
                if len(more) > len(text):
                    res["text_after_click_" + label] = more[:30000]
                    for ln in more.splitlines():
                        m = DIFF_RE.match(ln.strip())
                        if m and not any(d.get("raw") == ln.strip() for d in res["differences"]):
                            res["differences"].append({"sheet": m.group("sheet").strip(), "cell": m.group("cell").upper(), "excel": m.group("excel").strip(), "mind": m.group("mind").strip(), "raw": ln.strip()})
                    screenshot(page, out / f"logs_{label}.png")
        except Exception:
            pass
    try:
        hide = page.get_by_role("button", name=re.compile(r"^\s*Hide Logs\s*$", re.I))
        if hide.count():
            hide.first.click(timeout=4000)
    except Exception:
        pass
    _wait(page, 800)
    return res



def _open_selector(page, stem: str) -> bool:
    """Click the template selector in the project bar: its text is the workbook name, often cut with
    an ellipsis, so match the start of the name (then shorter and shorter prefixes)."""
    kill_tour(page)
    for n in (len(stem), 24, 18, 12, 8):
        prefix = stem[:n].strip()
        if not prefix:
            continue
        loc = page.locator("button", has_text=prefix)
        try:
            if loc.count():
                loc.first.click(timeout=4000)
                _wait(page, 1200)
                return True
        except Exception:
            continue
    return False


def selector_entries(page, stem: str, out: Path) -> list[str]:
    """Every entry of the open template selector menu (templates and grids), as text."""
    if not _open_selector(page, stem):
        return []
    _wait(page, 1200)
    screenshot(page, out / "selector_menu.png")
    try:
        names = page.evaluate(
            r"""()=>[...document.querySelectorAll('[role=menu] *,[role=menuitem],[role=listbox] *,[role=option],[class*=menu] *,[class*=dropdown] *,[class*=overlay] *,li')].map(e=>(e.innerText||'').trim()).filter(t=>t&&t.length<90&&!t.includes(String.fromCharCode(10)))"""
        )
    except Exception:
        names = []
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    _wait(page, 500)
    seen, out_names = set(), []
    for n in names:
        if n not in seen:
            seen.add(n)
            out_names.append(n)
    return out_names


def open_part(page, stem: str, entry: str, out: Path, tag: str) -> dict:
    """Open the template selector and click `entry`; record what the page shows and every control."""
    rec: dict = {"entry": entry, "opened": False}
    if not _open_selector(page, stem):
        rec["why"] = "selector did not open"
        return rec
    _wait(page, 1500)
    try:
        item = page.get_by_text(re.compile(rf"^\s*{re.escape(entry)}\s*$")).first
        item.click(timeout=5000)
        rec["opened"] = True
    except Exception as exc:
        rec["why"] = f"entry not clickable: {exc}"[:200]
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return rec
    _wait(page, 3500)
    kill_tour(page)
    screenshot(page, out / f"part_{tag}.png")
    rec["text"] = body_text(page)[:6000]
    try:
        rec["buttons"] = page.evaluate(BUTTONS_JS)
    except Exception:
        rec["buttons"] = []
    # every control that looks like an expand / dimension / loop control: click it, record what appears
    rec["clicks"] = []
    seen = set()
    for b in rec["buttons"]:
        key = (b["t"], b["a"], b["ti"], b["c"])
        label = " ".join(x for x in (b["t"], b["a"], b["ti"], b["c"]) if x)
        if key in seen or not EXPAND_RE.search(label) or len(rec["clicks"]) >= 12:
            continue
        seen.add(key)
        before = body_text(page)
        try:
            page.mouse.click(b["x"] + b["w"] / 2, b["y"] + b["h"] / 2)
            _wait(page, 2500)
            kill_tour(page)
            after = body_text(page)
            shot = out / f"part_{tag}_click{len(rec['clicks'])}.png"
            screenshot(page, shot)
            new_lines = [ln for ln in after.splitlines() if ln.strip() and ln not in before]
            rec["clicks"].append({"button": {k: b[k] for k in ("t", "a", "ti", "c", "x", "y")}, "new_text": "\n".join(new_lines)[:3000], "shot": shot.name})
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            _wait(page, 600)
        except Exception as exc:
            rec["clicks"].append({"button": {k: b[k] for k in ("t", "a", "ti", "c")}, "error": str(exc)[:200]})
    return rec


def open_screen(page, name_re: str, out: Path, tag: str) -> dict:
    """A top-bar screen (Audit trail, Project Edition, ...): click, record text + screenshot, Escape."""
    rec: dict = {"screen": tag, "opened": False}
    kill_tour(page)
    loc = page.get_by_role("button", name=re.compile(name_re, re.I))
    if not loc.count():
        loc = page.locator(f"[aria-label*='{tag}' i],[title*='{tag}' i]")
    if not loc.count():
        loc = page.get_by_text(re.compile(name_re, re.I))
    if not loc.count():
        rec["why"] = "not found"
        return rec
    try:
        loc.first.click(timeout=6000)
        rec["opened"] = True
    except Exception as exc:
        rec["why"] = str(exc)[:200]
        return rec
    _wait(page, 3000)
    kill_tour(page)
    screenshot(page, out / f"screen_{tag}.png")
    rec["text"] = body_text(page)[:8000]
    try:
        rec["buttons"] = [b for b in page.evaluate(BUTTONS_JS) if b["t"] or b["a"] or b["ti"]][:80]
    except Exception:
        rec["buttons"] = []
    # tabs inside the screen (Audit trail has Changes / Runs)
    rec["tabs"] = {}
    try:
        for t in page.get_by_role("tab").all()[:8]:
            label = t.inner_text().strip()
            t.click(timeout=3000)
            _wait(page, 2000)
            rec["tabs"][label] = body_text(page)[:4000]
            screenshot(page, out / f"screen_{tag}_{re.sub(r'[^A-Za-z0-9]+', '_', label)[:20]}.png")
    except Exception:
        pass
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    _wait(page, 800)
    return rec


def compare(expected_path: Path, differences: list[dict]) -> dict:
    """Mind's OUTPUT values from the log vs the macro's values (check.json: expected[target] rows per pass)."""
    chk = json.loads(expected_path.read_text(encoding="utf-8"))
    exp = chk.get("expected", {})
    mind = {(d.get("sheet", "").strip().upper(), d.get("cell", "")): d.get("mind") for d in differences if d.get("cell")}
    rows = []
    for target, passes in exp.items():
        m = re.match(r"^(.+?)!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", target)
        if not m:
            continue
        rows.append({"source": target, "passes": len(passes), "cells": len(passes[0]) if passes else 0})
    # the OUTPUT block rows are known from the plan: OUTPUT!A<k>:J<k> for pass k -- caller matches by sheet/cell
    out_cells = {k: v for k, v in mind.items() if k[0] == "OUTPUT"}
    return {"mind_output_cells": len(out_cells), "sources": rows, "output_values": {f"{k[0]}!{k[1]}": v for k, v in sorted(out_cells.items())}}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--project", default=None)
    ap.add_argument("--init-url", default=None, help="probe an existing project instead of uploading")
    ap.add_argument("--expected", default=None, help="check.json from scripts/check_loop_plan.py")
    ap.add_argument("--parts", default="OUTPUT,Untitled,CALCUL", help="template-selector entries to open (comma separated prefixes)")
    ap.add_argument("--no-run", action="store_true")
    ap.add_argument("--timeout", type=int, default=900)
    ap.add_argument("--source", default=None, help="MCP: the block MM_RESULT reads (Sheet!A1:J1), read per instance")
    ap.add_argument("--loop", default=None, help="MCP: the loop name (default: the project's loop)")
    ap.add_argument("--expect-size", type=int, default=None, help="MCP: the loop size to expect")
    ap.add_argument("--browser", action="store_true", help="use the browser for the run and the reading even when Mind MCP is available")
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
    out = workbook.parent / f"{workbook.stem}_mind_probe"
    out.mkdir(exist_ok=True)
    result: dict = {"workbook": str(workbook), "project": name, "started": time.strftime("%Y-%m-%d %H:%M:%S")}
    t0 = time.time()
    with mind_session(headed=False) as (_ctx, page):
        init = args.init_url
        if not init:
            if not open_manager(page):
                print("could not open the Project manager")
                return 4
            ensure_folder(page, SANDBOX_FOLDER)
            init = create_blank_project(page, name)
            if not init:
                print("could not create the project", name)
                return 5
        result["init_url"] = init
        state = open_project(page, init)
        result["project_state"] = state
        print(f"project {name}: {state} ({init})", flush=True)
        if state == "empty":
            conv = upload_and_convert(page, workbook, timeout_s=args.timeout, shots=out)
            result["conversion"] = {k: conv.get(k) for k in ("success", "steps", "errors", "timed_out")}
            print(f"conversion: {'OK' if conv['success'] else 'FAILED'} after {time.time() - t0:.0f}s | steps {conv.get('steps')}", flush=True)
            for e in conv.get("errors") or []:
                print("  error:", e[:300], flush=True)
            logs = read_logs(page, out)
            result["logs"] = {"opened": logs["opened"], "summary": logs.get("summary"), "differences": len(logs["differences"])}
            _log(out, "logs.json", logs)
            print(f"logs: {logs.get('summary')} | difference lines: {len(logs['differences'])}", flush=True)
            if args.expected and logs["differences"]:
                cmp = compare(Path(args.expected), logs["differences"])
                _log(out, "compare.json", cmp)
                print(f"Mind OUTPUT values in the log: {cmp['mind_output_cells']}", flush=True)
            if not conv["success"]:
                _log(out, "result.json", result)
                return 1
            if args.no_run:
                _log(out, "result.json", result)
                return 0
            added = add_template(page)
            result["template_added"] = added
            print("template added:", added, flush=True)
            if not added:
                _log(out, "result.json", result)
                return 1
        elif state != "model":
            print("project is neither empty nor a model; stopping")
            _log(out, "result.json", result)
            return 6
        use_mcp, mcp_why = (False, "--browser asked for") if args.browser else mind_mcp.available()
        result["mcp"] = {"used": use_mcp, "why": mcp_why}
        if use_mcp:
            print("Mind MCP available: the run and the loop check go through MCP once the browser closes", flush=True)
        else:
            print(f"Mind MCP not used ({mcp_why}): the browser runs the model and reads the parts", flush=True)
        if not use_mcp:
            run = run_model(page, shots=out)
            result["run"] = {k: run.get(k) for k in run if k != "text"}
            print("run:", json.dumps(result["run"])[:300], flush=True)
            screenshot(page, out / "after_run.png")
            result["after_run_text"] = body_text(page)[:3000]
            try:
                result["after_run_buttons"] = [b for b in page.evaluate(BUTTONS_JS) if b["t"] or b["a"] or b["ti"]][:120]
            except Exception:
                result["after_run_buttons"] = []
            stem = workbook.stem
            try:
                close = page.get_by_role("button", name=re.compile(r"Close modal", re.I))
                if close.count():
                    close.first.click(timeout=3000)
                    _wait(page, 800)
            except Exception:
                pass
            names = selector_entries(page, stem, out)
            if not names:
                names = template_names(page, stem)
            result["template_entries"] = names
            print(f"template entries ({len(names)}):", names[:40], flush=True)
            wanted = [p.strip() for p in args.parts.split(",") if p.strip()]
            parts = []
            for n in names:
                if any(n.startswith(w) for w in wanted) and len(parts) < 6:
                    parts.append(n)
            result["parts"] = []
            for i, entry in enumerate(parts):
                rec = open_part(page, stem, entry, out, f"{i}_{re.sub(r'[^A-Za-z0-9]+', '_', entry)[:24]}")
                result["parts"].append({k: v for k, v in rec.items() if k != "buttons"})
                _log(out, f"part_{i}.json", rec)
                print(f"part {entry}: opened={rec.get('opened')} clicks={len(rec.get('clicks', []))}", flush=True)
            for tag, pat in (("audit", r"Audit trail"), ("edition", r"Project Edition"), ("settings", r"Project Settings"), ("gridvalidation", r"Grid validation"), ("exports", r"^\s*Exports?\s*$")):
                rec = open_screen(page, pat, out, tag)
                result[f"screen_{tag}"] = {k: v for k, v in rec.items() if k not in ("buttons",)}
                _log(out, f"screen_{tag}.json", rec)
                print(f"screen {tag}: opened={rec.get('opened')}", flush=True)
    if result.get("mcp", {}).get("used"):
        expected = None
        if args.expected and args.source:
            expected = json.loads(Path(args.expected).read_text(encoding="utf-8"))["expected"].get(args.source)
        try:
            rep = mind_mcp.check_loop(result["init_url"], args.source, args.loop, expected, args.expect_size, do_run=True)
            result["mcp_check"] = rep
            _log(out, "mcp_check.json", rep)
            run = rep.get("run") or {}
            row = run.get("run") or {}
            print(f"MCP run: {row.get('Result')} | consistency {row.get('ConsistencyStatus')} | {run.get('calculations')} calculations in {run.get('seconds')} s", flush=True)
            for c in (run.get("errors") or {}).get("cells", [])[:20]:
                print(f"  error cell: {c['mind']} -> {c.get('excel')}", flush=True)
            print(f"MCP loop: {rep.get('loop')} size {rep.get('size')}" + (f" | macro comparison: {rep.get('same')} same, {rep.get('differ')} differ" if rep.get("cells") else ""), flush=True)
        except mind_mcp.MindMcpError as exc:
            result["mcp_check"] = {"error": str(exc)}
            print(f"MCP check failed ({exc}); rerun with --browser to use the browser instead", flush=True)
    result["seconds"] = round(time.time() - t0)
    _log(out, "result.json", result)
    print("everything in", out, flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
