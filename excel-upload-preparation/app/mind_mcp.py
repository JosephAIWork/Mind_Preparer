"""Milliman Mind's MCP server, as an optional extra.

Mind exposes its AI Assistant's tools over MCP (https://vm-prod.milliman-mind.com/api/chatbot/mcp, a
bearer token per user, issued by the Mind team; the one seen lasts 24 h). With a token the app reads a
project as data: its loops and their sizes, every instance's cell values, the run log, Mind's grid
cells mapped back to Excel addresses. Without one, nothing here is required: `available()` says why and
the callers fall back to the browser automation (app/mind_client.py).

Token: env MIND_MCP_TOKEN, else `mind_mcp_token` in the data folder (settings.yaml) (one line; never in the repo).
The company network re-signs TLS, so the Windows certificate store is used (truststore) when present.
"""
from __future__ import annotations

import base64
import json
import os
import re
import ssl
import time
from typing import Any

from .settings import SETTINGS

MCP_URL = os.environ.get("MIND_MCP_URL", SETTINGS.mind_mcp_url)
TOKEN_ENV = "MIND_MCP_TOKEN"
TOKEN_FILE = SETTINGS.data_folder / "mind_mcp_token"
_INIT_RE = re.compile(r"/init/\d+/([0-9a-fA-F-]{36})")


class MindMcpError(RuntimeError):
    pass


def read_token() -> str | None:
    tok = os.environ.get(TOKEN_ENV) or (TOKEN_FILE.read_text(encoding="utf-8").strip() if TOKEN_FILE.is_file() else "")
    return tok or None


def token_expiry(token: str) -> float | None:
    """The token's `exp` (epoch seconds), read from its payload without verifying it."""
    try:
        part = token.split(".")[1]
        return float(json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))["exp"])
    except Exception:
        return None


def available() -> tuple[bool, str]:
    """(True, "") when a token is present and not expired; else (False, the reason, in plain words)."""
    tok = read_token()
    if not tok:
        return False, f"no Mind MCP token (set {TOKEN_ENV} or write it to {TOKEN_FILE})"
    exp = token_expiry(tok)
    if exp and exp < time.time() + 60:
        return False, f"the Mind MCP token expired at {time.strftime('%Y-%m-%d %H:%M', time.localtime(exp))}"
    return True, ""


def project_id_from(link_or_id: str) -> str:
    """The project GUID from a .../init/<company>/<guid> link, or the GUID itself."""
    m = _INIT_RE.search(link_or_id or "")
    if m:
        return m.group(1)
    if re.fullmatch(r"[0-9a-fA-F-]{36}", (link_or_id or "").strip()):
        return link_or_id.strip()
    raise MindMcpError(f"not a Mind project link or GUID: {link_or_id!r}")


def _ssl_context() -> ssl.SSLContext:
    try:
        import truststore

        return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except ImportError:
        return ssl.create_default_context()


class MindMcp:
    """One MCP connection; `open(project)` opens a project session that every tool call then uses.

        with MindMcp() as mcp:
            mcp.open("https://shared.milliman-mind.com/init/<company>/<guid>")
            mcp.call("get_project_structure")
    """

    def __init__(self, token: str | None = None, url: str = MCP_URL, timeout: float = 120) -> None:
        import httpx

        self.token = token or read_token()
        if not self.token:
            raise MindMcpError(available()[1])
        self.url = url
        self.headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json", "Accept": "application/json, text/event-stream"}
        self.client = httpx.Client(timeout=timeout, verify=_ssl_context())
        self.session_id: str | None = None  # the PROJECT session (open_project), not the MCP transport session
        self._id = 0
        self._rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "mind-ready", "version": "1.11.2"}})
        self._notify("notifications/initialized")

    # ---- transport ----------------------------------------------------------------------------
    def _post(self, body: dict[str, Any]) -> Any:
        r = self.client.post(self.url, headers=self.headers, json=body)
        if r.status_code in (401, 403):
            raise MindMcpError(f"Mind refused the MCP token ({r.status_code}): it may have expired")
        if r.headers.get("mcp-session-id"):
            self.headers["mcp-session-id"] = r.headers["mcp-session-id"]
        if r.status_code >= 400:
            raise MindMcpError(f"MCP HTTP {r.status_code}: {r.text[:300]}")
        if not r.text:
            return None
        if "event-stream" in r.headers.get("content-type", ""):
            data = [ln[5:].strip() for ln in r.text.splitlines() if ln.startswith("data:")]
            return json.loads(data[-1]) if data else None
        return r.json()

    def _rpc(self, method: str, params: dict[str, Any] | None = None) -> Any:
        self._id += 1
        out = self._post({"jsonrpc": "2.0", "id": self._id, "method": method, "params": params or {}})
        if out is None:
            return None
        if "error" in out:
            raise MindMcpError(f"{method}: {out['error']}")
        return out.get("result")

    def _notify(self, method: str) -> None:
        self._post({"jsonrpc": "2.0", "method": method})

    # ---- tools ----------------------------------------------------------------------------------
    def call(self, tool: str, **args: Any) -> Any:
        """Call a tool; the project session is passed for you. Returns the parsed JSON content when the
        tool answers JSON, else its text."""
        if self.session_id and tool not in ("open_project",):
            args.setdefault("sessionId", self.session_id)
        res = self._rpc("tools/call", {"name": tool, "arguments": {k: v for k, v in args.items() if v is not None}})
        content = (res or {}).get("content") or []
        text = "\n".join(c.get("text", "") for c in content if c.get("type") == "text")
        if (res or {}).get("isError"):
            raise MindMcpError(f"{tool}: {text[:400]}")
        try:
            return json.loads(text)
        except (ValueError, TypeError):
            return text

    def open(self, link_or_id: str) -> str:
        out = self.call("open_project", projectId=project_id_from(link_or_id))
        sid = out.get("sessionId") or out.get("SessionId") if isinstance(out, dict) else None
        if not sid and isinstance(out, str):
            m = re.search(r"[Ss]ession[Ii]d\W+([0-9A-Za-z-]{8,})", out)
            sid = m.group(1) if m else None
        if not sid:
            raise MindMcpError(f"open_project gave no session: {str(out)[:300]}")
        self.session_id = sid
        return sid

    def close(self) -> None:
        try:
            if self.session_id:
                self.call("close_project")
        except Exception:
            pass
        self.session_id = None
        try:
            self.client.close()
        except Exception:
            pass

    def __enter__(self) -> "MindMcp":
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()


# ---- the tools' answers: a compact text format ("Rows[1]{A,B}:" then indented CSV lines) ----------------
def _csv(line: str) -> list[str]:
    import csv

    return next(csv.reader([line.strip()], skipinitialspace=True), [])


def _scalar(v: str) -> Any:
    v = v.strip()
    if v in ("null", ""):
        return None
    if v in ("true", "false"):
        return v == "true"
    if len(v) >= 2 and v[0] == v[-1] == '"':
        return v[1:-1]
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return v


def table(text: str, key: str) -> list[dict[str, Any]]:
    """Rows of a `Key[n]{ColA,ColB}:` block, as dicts."""
    lines = str(text or "").splitlines()
    for i, ln in enumerate(lines):
        m = re.match(rf"^(\s*){re.escape(key)}\[\d+\]\{{([^}}]*)\}}:\s*$", ln)
        if not m:
            continue
        indent, cols = len(m.group(1)), [c.strip() for c in m.group(2).split(",")]
        out = []
        for row in lines[i + 1:]:
            if not row.strip() or len(row) - len(row.lstrip()) <= indent:
                break
            vals = _csv(row)
            out.append({c: _scalar(v) for c, v in zip(cols, vals)})
        return out
    return []


def field(text: str, key: str) -> Any:
    """A `Key: value` line (the first one)."""
    m = re.search(rf"^\s*{re.escape(key)}:\s*(.*)$", str(text or ""), re.M)
    return _scalar(m.group(1)) if m else None


def values(text: str) -> list[dict[str, Any]]:
    """`Values[n]:` then `- [k]: Dim1,...,Values` lines: the first line names the columns."""
    lines = str(text or "").splitlines()
    start = next((i for i, ln in enumerate(lines) if re.match(r"^\s*Values\[\d+\]:\s*$", ln)), None)
    if start is None:
        return []
    rows = [_csv(re.sub(r"^\s*-\s*\[\d+\]:\s*", "", ln)) for ln in lines[start + 1:] if re.match(r"^\s*-\s*\[\d+\]:", ln)]
    if not rows:
        return []
    head, out = rows[0], []
    for r in rows[1:]:
        out.append({h: _scalar(v) for h, v in zip(head, r)})
    return out


# ---- 1. the loop as data ------------------------------------------------------------------------------
def dimensions(mcp: MindMcp) -> list[dict[str, Any]]:
    """[{Name, Length, IsLoopInstance, HasLabels}] -- empty until the session has a run (start one)."""
    return table(mcp.call("get_project_structure"), "Dimensions")


def first_instance(mcp: MindMcp) -> tuple[int, str]:
    st = str(mcp.call("get_project_structure"))
    m = re.search(r"ModelIndex:\s*(\d+)[\s\S]*?Instances\[\d+\]\{Id\}:\s*\n\s*([^\n]+)", st)
    return (int(m.group(1)), m.group(2).strip()) if m else (0, "Original_Instance")


def cell_instances(mcp: MindMcp, sheet: str, address: str, loop: str | None = None, model: tuple[int, str] | None = None) -> dict[str, Any]:
    """Every instance's value of one Excel cell, read in Mind: {grid, row, col, values: {index: value}}."""
    mi, inst = model or first_instance(mcp)
    r = mcp.call("resolve_excel_address", modelIndex=mi, instanceId=inst, spreadsheetName=sheet, cellAddress=address)
    if field(r, "OutOfGrids"):
        return {"cell": f"{sheet}!{address}", "out_of_grids": True, "values": {}}
    g, row, col = field(r, "GridIndex"), field(r, "Row"), field(r, "Col")
    if field(r, "IsHeaderRow") or not isinstance(row, int) or row < 1:
        # a grid's header row (Mind's row 0): the cell-values tool reads data rows only
        return {"cell": f"{sheet}!{address}", "grid": field(r, "GridName"), "header_row": True, "dimensioned": False, "scalar": None, "values": {}}
    v = mcp.call("get_cell_values", modelIndex=mi, instanceId=inst, gridIndex=g, row=row, col=col)
    vals: dict[int, Any] = {}
    recs = values(v)
    dimensioned = bool(recs) and any(k != "Values" for k in recs[0])
    scalar = None
    for rec in recs:
        key = rec.get(loop) if loop and loop in rec else next((rec[k] for k in rec if k != "Values"), None)
        if isinstance(key, int):
            vals[key] = rec.get("Values")
    if not dimensioned:
        # a cell without the loop (a constant, a formula not reading the driver): one value for every instance
        scalar = recs[0].get("Values") if recs else _scalar(str(field(v, "Value") or ""))
    return {"cell": f"{sheet}!{address}", "grid": field(r, "GridName"), "grid_index": g, "row": row, "col": col, "total": field(v, "TotalValuesCount"),
            "dimensioned": dimensioned, "scalar": scalar, "values": vals}


def instance_value(cell: dict[str, Any], k: int) -> Any:
    """Instance k of a cell read by `cell_instances`: its own value, or the single value of a cell the loop does not reach."""
    return cell["values"].get(k) if cell.get("dimensioned") else cell.get("scalar")


# ---- 2. run, and the run's errors mapped back to Excel ------------------------------------------------
_MIND_REF = re.compile(r"model_(\d+)_grid_(\d+)_R(\d+)_C(\d+)")
_UNTITLED_REF = re.compile(r"Untitled\(\d+,\d+\)\((\d+),(\d+)\)")


def run(mcp: MindMcp, save: bool = False, timeout_s: int = 1800, poll_s: float = 3.0) -> dict[str, Any]:
    """Start a classic run of the whole project, wait for it, and return its history row and, on a
    failure, the log with every Mind cell it names mapped to its Excel address."""
    opts = mcp.call("get_run_options")
    if field(opts, "RunAllowed") is False:
        return {"started": False, "why": str(opts)[:400]}
    started = mcp.call("start_run", runType="classic", saveAfterRun=save)
    refused = field(started, "Refused")
    if refused:
        return {"started": False, "why": str(refused)}
    t0 = time.time()
    status = ""
    time.sleep(poll_s)
    while time.time() - t0 < timeout_s:
        status = mcp.call("get_run_status")
        if field(status, "Done") is True:
            break
        time.sleep(poll_s)
    hist = table(mcp.call("get_run_history"), "Rows")
    last = hist[0] if hist else {}
    out = {"started": True, "done": field(status, "Done") is True, "has_error": bool(field(status, "HasError")), "seconds": round(time.time() - t0, 1),
           "calculations": field(status, "NbCalculation"), "run": last}
    if out["has_error"] or (last and last.get("Result") not in ("Success", None)):
        out["errors"] = run_errors(mcp, str(last.get("RunId") or ""))
    return out


def run_errors(mcp: MindMcp, run_id: str) -> dict[str, Any]:
    """The run log's error lines, with every Mind cell reference turned into its Excel address."""
    if not run_id:
        return {"log": "", "lines": [], "cells": []}
    log = str(mcp.call("get_run_log", runId=run_id))
    lines = [ln.strip() for ln in log.splitlines() if re.search(r"error|fail|exception|#(REF|VALUE|DIV|N/A|NAME)", ln, re.I)]
    cells, seen = [], set()
    model = None
    for m in _MIND_REF.finditer(log):
        key = m.group(0)
        if key in seen:
            continue
        seen.add(key)
        model = model or first_instance(mcp)
        try:
            a = mcp.call("get_cell_excel_address", modelIndex=int(m.group(1)), instanceId=model[1], gridIndex=int(m.group(2)), row=int(m.group(3)), col=int(m.group(4)))
            cells.append({"mind": key, "excel": f"{field(a, 'SpreadsheetName')}!{field(a, 'Address')}"})
        except MindMcpError as exc:
            cells.append({"mind": key, "excel": None, "why": str(exc)[:200]})
    for m in _UNTITLED_REF.finditer(log):  # Untitled(r,c)(R,C): R,C are the Excel row and column already
        r, c = int(m.group(1)), int(m.group(2))
        col = ""
        while c:
            c, rem = divmod(c - 1, 26)
            col = chr(65 + rem) + col
        if m.group(0) not in seen:
            seen.add(m.group(0))
            cells.append({"mind": m.group(0), "excel": f"{col}{r}"})
    return {"log": log[-6000:], "lines": lines[:50], "cells": cells}


# ---- 1 + 2 together: run, then read the loop and compare with the macro ------------------------------
def _same(a: Any, b: Any) -> bool:
    import math

    if a is None or a == "":
        return b in (None, "", 0)
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-6)
    if isinstance(a, str) and re.match(r"^\d{4}-\d{2}-\d{2}", a):
        from datetime import date

        y, m, d = (int(x) for x in a[:10].split("-"))
        return str(b).startswith(a[:10]) or (isinstance(b, (int, float)) and abs(b - (date(y, m, d) - date(1899, 12, 30)).days) < 1e-6)
    return str(a).strip() == str(b).strip()


def check_loop(project: str, source: str | None = None, loop: str | None = None, expected: list[list[Any]] | None = None,
               expect_size: int | None = None, do_run: bool = True, save: bool = False) -> dict[str, Any]:
    """Open `project`, run it (by default), read the loop's size and every instance of the `source`
    block (Sheet!A1:J1, the cells MM_RESULT reads), compare instance k with the macro's pass k."""
    from .formula_utils import num_to_col, parse_ref

    rep: dict[str, Any] = {"project": project, "source": source, "via": "mcp", "ok": True}
    with MindMcp() as mcp:
        mcp.open(project)
        if do_run:
            rep["run"] = run(mcp, save=save)
            if not rep["run"].get("started") or rep["run"].get("has_error"):
                rep["ok"] = False
        dims = dimensions(mcp)
        rep["dimensions"] = dims
        name = loop or next((d["Name"] for d in dims if d.get("Name") != "SIM"), None)
        size = next((d.get("Length") for d in dims if d.get("Name") == name), None)
        rep.update(loop=name, size=size)
        if not dims:
            rep.update(ok=False, why="no dimension: the session holds no run results")
        if expect_size is not None and size != expect_size:
            rep["ok"] = False
        if not source:
            return rep
        sheet, ref = source.split("!", 1)
        sheet = sheet.strip("'")
        r = parse_ref(ref)
        cells = [f"{num_to_col(c)}{row}" for row in range(r["r1"], r["r2"] + 1) for c in range(r["c1"], r["c2"] + 1)]
        model = first_instance(mcp)
        got = [cell_instances(mcp, sheet, c, name, model) for c in cells]
    rows, same, differ = [], 0, 0
    for k in range(1, (size or 0) + 1):
        line = [instance_value(g, k) for g in got]
        bad: list[str] = []
        if expected and k <= len(expected):
            bad = [cells[j] for j, (m, e) in enumerate(zip(line, expected[k - 1])) if not _same(e, m)]
            same += len(cells) - len(bad)
            differ += len(bad)
        rows.append({"instance": k, "values": line, "differs": bad, "beyond_macro": bool(expected) and k > len(expected)})
    rep.update(cells=cells, instances=rows, same=same, differ=differ)
    if differ:
        rep["ok"] = False
    return rep
