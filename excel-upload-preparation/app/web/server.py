"""FastAPI backend of the Mind Ready web app (the front end is the Angular + Carbon
app in `MindReadyUI`).

One endpoint per call of the front end's API service, on top of the analysis
engine (app/modes, app/prep, app/chat_context, app/recalc, app/excel_report,
app/change_apply). Nothing here re-implements a rule or a change; it only maps
the engine's dicts onto the front end's TypeScript contracts and keeps a per-session
**version lineage**: every write is a new file, verified to open in Excel,
with its change log.

Sessions live in memory (this is a local, single-user tool); their files
live in a temp directory per session. The built front-end (`dist/`) is
served from the same process with an SPA fallback, so one URL does it all:

    python -m app.web.server            (or run_mind_ready_web.bat)
    http://localhost:8600
"""
from __future__ import annotations

import contextlib
import hashlib
import json
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from fastapi import Body, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.exception_handlers import http_exception_handler
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from starlette.exceptions import HTTPException as StarletteHTTPException

PACKAGE_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PACKAGE_ROOT))

from app.autofix import AutoFixConfig, a1, run_autofix  # noqa: E402
from app.autofix_advisor import make_advisor  # noqa: E402
from app.numbers_check import Moves, compare_workbooks  # noqa: E402
from app.change_apply import convert_output_format  # noqa: E402
from app.chat_context import answer_question, recalculation_context  # noqa: E402
from app.config import load_config  # noqa: E402
from app.excel_com import com_available  # noqa: E402
from app.perf import record as perf_record  # noqa: E402
from app.perf import stage as perf_stage  # noqa: E402
from app import excel_worker  # noqa: E402
from app.excel_report import build_report_workbook, build_standalone_report  # noqa: E402
from app.inventory import cell_window, has_vba_project, make_immutable_copy  # noqa: E402
from app import ai_key, feedback  # noqa: E402
from app import expiry  # noqa: E402  (through the module: tests replace its checker)
from app.settings import SETTINGS  # noqa: E402
from app.llm import DEFAULT_MODEL_ID, check_api_key, extract_formula, llm_available, suggest_formula_fix  # noqa: E402
from app.modes import fix_incompatible_formulas, plan_mode, prep_mind_loops, structure_fix  # noqa: E402
from app.grid_naming import build_grid_context, suggest_names  # noqa: E402
from app.estimate import MIND_STAGE_TITLES, estimate_chain, estimate_mind_run, mind_run_remaining  # noqa: E402
from app.mind_loop import LoopConfig, run_loop, summarize  # noqa: E402
from app.model_map import build_model_map
from app import model_map as model_map_module  # noqa: E402  (read_vba through the module: tests replace it)
from app import outside_inputs  # noqa: E402  (through the module: tests replace run_import / apply_mind_inputs)
from app import stand_in_inputs  # noqa: E402  (through the module: tests replace make_stand_ins / assistant_ask)
from app.loop_builder import _add_repairs, build_loop_plan  # noqa: E402
from app.instance_builder import build_instance_plan, write_main_workbook, zip_for_mind  # noqa: E402
from app.instance_finder import EMPTY_ASSISTANT, find_instances, make_finder_advisor, merge_assistant  # noqa: E402
from app.grids import all_grids, json_safe  # noqa: E402
from app.models import Finding, Status, aggregate_status  # noqa: E402
from app.plain_codes import cite_all, glossary, is_com_error, plain_com_error, plain_com_errors_in, plain_error, rule_name  # noqa: E402
from app.validators import run_rule  # noqa: E402
from app.readiness import OUTSIDE_RULE, STATUS_RANK, apply_outcome, apply_record, compute_readiness, plan_counts, prep_progress, recalc_state  # noqa: E402
from app.formula_utils import col_to_num, formula_memo  # noqa: E402
from app.prep import ASSISTANT_OPS, DEFAULT_PIVOT_MODE, PIVOT_MODES, STRUCTURAL_OPS, apply_operations, deterministic_name, is_weak_name, plan_actions, plan_create_grid_titles, plan_named_areas, standalone_labels  # noqa: E402
from app.validators.kb import documented_flags  # noqa: E402
from app.validators.mind_values import environment_values  # noqa: E402
from app.recalc import classify_against_original, classify_errors, formulas_with_broken_refs, group_errors, recalculate  # noqa: E402
from app.prep import _cells_of as _array_cells  # noqa: E402
from app.prep import array_size_hint  # noqa: E402
from app.rules_engine import RulesEngine  # noqa: E402
from app.progress import APPLY_STAGE_TITLES, STAGE_TITLES, overall_progress  # noqa: E402
from app.sizing import MB, size_gate  # noqa: E402

MODE_MAP = {
    "plan": ("Plan (analyze everything)", plan_mode),
    "prep_loops": ("Prep Mind Loops", prep_mind_loops),
    "fix_formulas": ("Fix Incompatible Formulas", fix_incompatible_formulas),
    "structure_fix": ("Structure Fix", structure_fix),
}
ALLOWED_SUFFIXES = {".xlsx", ".xlsm", ".xlsb"}
# The front end is the Angular + Carbon app in MindReadyUI (its build is committed). Another
# build can be served with MIND_READY_DIST=<path to its dist folder>.
DEFAULT_DIST = PACKAGE_ROOT.parent / "MindReadyUI" / "dist" / "mind-ready-ui" / "browser"
DIST_DIR = Path(os.environ.get("MIND_READY_DIST", str(DEFAULT_DIST)))
GRID_FIELDS = (
    "sheet", "title_cell", "title", "name", "display_name", "flags", "flag_names", "anchor", "ref",
    "first_row", "last_row", "first_col", "last_col", "n_rows", "n_cols", "header_values", "header_is_all_text",
    "inner_title_cells", "formula_count",
)

# The API's own pages (/docs, /redoc, /openapi.json) only when settings.yaml asks for them.
_API_PAGES = {} if SETTINGS.api_pages else {"docs_url": None, "redoc_url": None, "openapi_url": None}
@contextlib.asynccontextmanager
async def _lifespan(_app: FastAPI):
    """At start: end the hidden Excels a Mind Ready stopped abruptly left behind, and make closing the
    window end this one's first (app/excel_com.py). A shared copy with an end date also asks the
    internet's date, in the background (app/expiry.py; nothing for the developer version)."""
    from app.excel_com import clean_leftover_excels, on_window_close

    expiry.start_network_check()
    ended = clean_leftover_excels()
    if ended:
        print(f"Closed {ended} hidden Excel(s) left behind the last time Mind Ready was stopped.", flush=True)
    on_window_close()
    yield


class _ExpiryGate:
    """After a shared copy's end date (app/expiry.py) every API call but /api/health answers 410 with
    the plain message, which the screens show instead of the app; the screens' own files still load, to
    show it. A plain ASGI middleware: streamed answers and downloads pass through it untouched. The
    decision is reused for a minute, so a call costs nothing more."""

    def __init__(self, inner):
        self.inner = inner

    async def __call__(self, scope, receive, send):
        path = scope.get("path", "") if scope["type"] == "http" else ""
        if path.startswith("/api/") and path != "/api/health":
            verdict = expiry.current()
            if verdict.expired:
                await JSONResponse(verdict.refusal(), status_code=410)(scope, receive, send)
                return
        await self.inner(scope, receive, send)


app = FastAPI(lifespan=_lifespan, title="Mind Ready API", version=(PACKAGE_ROOT / "VERSION").read_text(encoding="utf-8").strip() if (PACKAGE_ROOT / "VERSION").exists() else "dev", **_API_PAGES)
app.add_middleware(_ExpiryGate)
_ENGINE: RulesEngine | None = None


def engine() -> RulesEngine:
    global _ENGINE
    if _ENGINE is None:
        _ENGINE = RulesEngine()
    return _ENGINE


@dataclass
class Session:
    id: str
    work_dir: Path
    mode: str
    versions: list[dict[str, Any]] = field(default_factory=list)
    paths: dict[str, Path] = field(default_factory=dict)  # version id -> file
    files: dict[str, Path] = field(default_factory=dict)  # download name -> file
    current_version_id: str = ""
    result: dict[str, Any] | None = None
    plan: list[dict[str, Any]] = field(default_factory=list)
    last_delta: dict[str, Any] | None = None  # what the last analysis changed vs the previous one
    recalc: dict[str, Any] | None = None  # last real recalculation of the current version
    recalc_version_id: str | None = None
    recalc_path: Path | None = None  # the recalculated copy: freshest cell values for the workbook view
    grid_names: dict[str, str] = field(default_factory=dict)  # "Sheet!Ref" -> assistant-proposed grid name
    created_at: float = field(default_factory=time.time)
    # Upload size gate + scan status
    size: dict[str, Any] | None = None  # app.sizing.size_gate of the uploaded file
    ignore_sheets: list[str] = field(default_factory=list)  # sheets the user chose not to scan
    pivot_mode: str = DEFAULT_PIVOT_MODE  # what Prep (and Do everything) does with the pivot tables
    pending: bool = False  # uploaded and inspected, not analysed yet (deferred upload)
    scan: dict[str, Any] = field(default_factory=dict)  # live status of the current/last scan
    # One entry per analysis: how much prep work each version still needs (the Prep gauge)
    plan_history: list[dict[str, Any]] = field(default_factory=list)
    # One entry per Apply: the changes really written into the file (the applied-changes gauge)
    apply_history: list[dict[str, Any]] = field(default_factory=list)
    # The ORIGINAL after a full Excel recalculation: its error cells and its #REF! formulas
    # (the baseline every later recalculation is compared with: the upload, or the version the
    # outside files were brought into -- `baselines` keeps one per file, computed once)
    original_baseline: dict[str, Any] | None = None
    baselines: dict[str, dict[str, Any]] = field(default_factory=dict)  # baseline file -> its recalculation
    # (sheet, cell) of the original's formulas that contain #REF!, taken from its own scan
    # (None until the original was scanned with no sheet skipped: then the baseline reads the file);
    # `ref_formulas` the same for every baseline file (the upload and each "import" version)
    original_ref_formulas: set[tuple[str, str]] | None = None
    ref_formulas: dict[str, set[tuple[str, str]]] = field(default_factory=dict)
    # (sheet, cell) -> what each CELL / INFO formula shows in the original upload, as the user's
    # own Excel saved it (FRM-012 freezes these, not the app's working copy's)
    original_environment: dict[tuple[str, str], Any] | None = None
    # Live status of the current/last Apply (GET /status, field "apply")
    apply_status: dict[str, Any] = field(default_factory=dict)
    # The model map (inputs, loop driver, run settings, outputs) of the analysed version
    model_map: dict[str, Any] | None = None
    model_map_version_id: str | None = None
    # Outside input files (app/outside_inputs.py): the files picked with the workbook, what its
    # macros open, the job that brings them in, and the Mind inputs (CSV) that job made
    inputs: list[dict[str, Any]] = field(default_factory=list)  # {name, path, bytes, sha256}, one per name
    input_seq: int = 0  # <work>/inputs/<nn>/<name>: one folder per file received
    outside_vba: dict[str, Any] | None = None  # what the upload's macros open (read once)
    outside_inputs: dict[str, Any] | None = None  # {version_id, vba, table, match}; None: the macros open no outside file
    outside_run: dict[str, Any] = field(default_factory=dict)  # the bring-in job (GET .../outside-inputs)
    mind_inputs_dir: str | None = None  # the CSV folder of the last successful bring-in
    mind_inputs_dirs: dict[str, str] = field(default_factory=dict)  # "mind_inputs" version id -> its CSV folder
    # the bring-ins that FINISHED: the last version each made ("mind_inputs", or "import" when no Mind
    # input was made) -> the fingerprint of the picked files it read (_inputs_fingerprint). A line counts
    # as brought in only through one of these, and only while the picked files are still those.
    brought_in_files: dict[str, str] = field(default_factory=dict)
    # The instance finder's result (where instances or a loop would help) for the analysed version, and
    # the assistant's block on it for that same version once it was asked and answered
    instance_finder: dict[str, Any] | None = None
    instance_finder_version_id: str | None = None
    instance_finder_assistant: dict[str, Any] | None = None
    # The instance builder's plan with no key chosen and keys=present, built once per version: the
    # panel's preview and the finder's "builder" line both ask for exactly this plan
    instance_plan_default: dict[str, Any] | None = None
    instance_plan_default_version_id: str | None = None
    # The zip of the last "Build the Mind instances" (the model + its main workbook), None until one is built
    instances_zip: Path | None = None
    lock: threading.Lock = field(default_factory=threading.Lock, repr=False, compare=False)
    # one Apply at a time: two at once would both copy the same version and the last one to finish
    # would become current without the other's changes
    apply_lock: threading.Lock = field(default_factory=threading.Lock, repr=False, compare=False)

    def readiness(self) -> dict[str, Any] | None:
        """The verdict for the current version (None before the first analysis)."""
        if not self.result:
            return None
        return compute_readiness(self.result["validation_report"], self.plan, self.recalc, self.current_version_id)

    @property
    def current_path(self) -> Path:
        return self.paths[self.current_version_id]


SESSIONS: dict[str, Session] = {}


# --- helpers -------------------------------------------------------------------------
def _session(session_id: str) -> Session:
    s = SESSIONS.get(session_id)
    if s is None:
        raise HTTPException(status_code=404, detail=f"unknown session {session_id}")
    return s


def _register_file(s: Session, path: Path) -> str:
    """Download name for a file, unique within the session."""
    path = Path(path)
    name = path.name
    if s.files.get(name) not in (None, path):
        stem, suffix = path.stem, path.suffix
        n = 2
        while s.files.get(f"{stem}_v{n}{suffix}") not in (None, path):
            n += 1
        name = f"{stem}_v{n}{suffix}"
    s.files[name] = path
    return name


def _add_version(
    s: Session, path: Path, source: str, label_body: str, change_log: list[dict[str, Any]], verified: bool | None, sha256: str, minor: bool = False
) -> dict[str, Any]:
    """Versions are numbered major.minor. A fix made on the Recalculate
    step (`minor`) is the next minor of the current major (v2 -> v2.1 -> v2.2);
    everything else -- and a recalculation after such fixes -- starts the next
    major (v3). Ids follow the number: ver-002, ver-002.1, ..."""
    current = next((v for v in s.versions if v["id"] == s.current_version_id), None)
    if minor and current is not None:
        major = current.get("major", 1)
        n_minor = 1 + max((v.get("minor", 0) for v in s.versions if v.get("major") == major), default=0)
    else:
        major = 1 + max((v.get("major", 0) for v in s.versions), default=0)
        n_minor = 0
    number = f"{major}.{n_minor}" if n_minor else f"{major}"
    vid = f"ver-{major:03d}" + (f".{n_minor}" if n_minor else "")
    download_name = _register_file(s, path)
    version = {
        "id": vid,
        "parent": s.current_version_id or None,  # the version this one was made from
        "major": major,
        "minor": n_minor,
        "label": f"v{number} — {label_body}",
        "file_name": download_name,
        "sha256": sha256[:12],
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "source": source,
        "change_log": [_public_op(o) for o in change_log],
        "verified_opens_in_excel": verified,
    }
    s.versions.append(version)
    s.paths[vid] = Path(path)
    s.current_version_id = vid
    return version


def _public_op(o: dict[str, Any]) -> dict[str, Any]:
    """An operation as the screen and the change log get it: long lists cut short (the first 50
    cells; the first 20 values of a set_range_values, with their count)."""
    out = {k: v for k, v in o.items() if k not in ("cells", "values")}
    if "cells" in o:
        out["cells"] = list(o["cells"])[:50]
    if "values" in o:
        values = list(o["values"])
        out["values"] = values[:20]
        out["values_count"] = len(values)
    return out


def _summary(analysis: dict[str, Any]) -> dict[str, Any]:
    wb = analysis["workbooks"][0]
    f = analysis["features"]
    sheets = []
    for sh in wb["sheets"]:
        sheets.append(
            {
                "name": sh["name"],
                "state": sh["state"],
                "dimensions": sh.get("dimensions") or "",
                "grids": [{k: g.get(k) for k in GRID_FIELDS} for g in sh["grids"]],
                "standalone_text_cells": [{"cell": c["cell"], "text": str(c["text"])} for c in sh.get("standalone_text_cells", [])],
                "formula_count": sh.get("formula_count", 0),
                "array_formula_count": sh.get("array_formula_count", 0),
                "protected": bool(sh.get("protected")),
            }
        )
    return {
        "file_name": f"{wb.get('file_name')}.{wb.get('file_type')}",
        "file_type": wb.get("file_type"),
        "sha256": (wb.get("source_sha256") or analysis["source"].get("sha256") or "")[:12],
        "application": f.get("application"),
        "app_version": f.get("app_version"),
        "has_vba": bool(f.get("has_vba")),
        "sheet_count": f.get("sheet_count", 0),
        "ignored_sheet_count": f.get("ignored_sheet_count", 0),
        "total_sheet_count": f.get("total_sheet_count", f.get("sheet_count", 0)),
        "ignored_sheets": [{"name": s["name"], "state": s["state"]} for s in wb.get("ignored_sheets", [])],
        "hidden_sheet_count": f.get("hidden_sheet_count", 0),
        "grid_count": f.get("grid_count", 0),
        "flagged_grid_count": f.get("flagged_grid_count", 0),
        "formula_count": f.get("formula_count", 0),
        "array_formula_count": f.get("array_formula_count", 0),
        "defined_name_count": f.get("defined_name_count", 0),
        "external_link_part_count": f.get("external_link_part_count", 0),
        "mm_functions_used": f.get("mm_functions_used", {}),
        "native_function_usage": dict(sorted(f.get("native_function_usage", {}).items(), key=lambda kv: -kv[1])[:40]),
        "sheets": sheets,
    }


def _delta(previous: dict[str, Any] | None, current: dict[str, Any], version_id: str) -> dict[str, Any]:
    """What changed between two analyses of the same session: findings fixed
    (now PASS), improved (better status or a changed message while still not
    PASS), regressed, and the status counts side by side -- so a fix never
    just *vanishes* from a filtered table."""
    cur = {f["rule_id"]: f for f in current["findings"]}
    if not previous:
        return {"version_id": version_id, "previous_version_id": None, "fixed": [], "improved": [], "regressed": [], "previous_counts": {}, "counts": current["summary"]["status_counts"]}
    prev = {f["rule_id"]: f for f in previous["findings"]}
    fixed, improved, regressed = [], [], []
    for rid, f in cur.items():
        p = prev.get(rid)
        if p is None:
            continue
        before, after = STATUS_RANK.get(p["status"], 0), STATUS_RANK.get(f["status"], 0)
        if p["status"] != "PASS" and f["status"] == "PASS":
            fixed.append({"rule_id": rid, "rule_name": rule_name(rid), "from": p["status"], "to": f["status"]})
        elif after < before or (after == before and f["status"] != "PASS" and f["message"] != p["message"]):
            improved.append({"rule_id": rid, "rule_name": rule_name(rid), "from": p["status"], "to": f["status"], "note": "status improved" if after < before else "same status, details changed"})
        elif after > before:
            regressed.append({"rule_id": rid, "rule_name": rule_name(rid), "from": p["status"], "to": f["status"]})
    return {
        "version_id": version_id,
        "previous_version_id": previous.get("_version_id"),
        "fixed": fixed,
        "improved": improved,
        "regressed": regressed,
        "previous_counts": previous["summary"]["status_counts"],
        "counts": current["summary"]["status_counts"],
    }


# --- scan status ------------------------------------------------------------------
def _scan_start(s: Session, needs_convert: bool) -> None:
    now = time.time()
    s.scan = {
        "state": "running",
        "stage": None,
        "title": "Starting",
        "message": "Starting",
        "fraction": None,
        "overall": 0.0,
        "sheet": None,
        "rule_id": None,
        "started_at": now,
        "updated_at": now,
        "finished_at": None,
        "stage_started": now,
        "stages": [],
        "error": None,
        "version_id": s.current_version_id,
        "needs_convert": needs_convert,
        "ignore_sheets": list(s.ignore_sheets),
    }


def _progress_for(s: Session):
    def cb(stage: str, message: str, fraction: float | None = None, **facts: Any) -> None:
        sc = s.scan
        now = time.time()
        if sc.get("stage") != stage:
            if sc.get("stage"):
                sc["stages"].append({"stage": sc["stage"], "seconds": round(now - sc["stage_started"], 2)})
            sc["stage"], sc["stage_started"] = stage, now
        sc.update(
            title=STAGE_TITLES.get(stage, stage),
            message=message,
            fraction=fraction,
            updated_at=now,
            sheet=facts.get("sheet"),
            rule_id=facts.get("rule_id"),
        )
        sc["overall"] = overall_progress(stage, fraction, bool(sc.get("needs_convert")))

    return cb


def _scan_finish(s: Session, error: str | None = None) -> None:
    sc = s.scan
    now = time.time()
    if sc.get("stage"):
        sc["stages"].append({"stage": sc["stage"], "seconds": round(now - sc.get("stage_started", now), 2)})
    sc.update(state="error" if error else "ready", finished_at=now, updated_at=now, error=error, overall=sc.get("overall", 0.0) if error else 1.0)
    sc["title"] = "Failed" if error else "Done"
    sc["message"] = error or f"Scan finished in {now - sc.get('started_at', now):.1f} s"
    sc["stage"] = sc.get("stage") if error else "done"


def _scan_snapshot(s: Session) -> dict[str, Any]:
    sc = dict(s.scan) if s.scan else {"state": "pending" if s.pending else "idle", "stage": None, "title": None, "message": None, "fraction": None, "overall": 0.0, "stages": [], "error": None}
    now = time.time()
    started = sc.get("started_at")
    finished = sc.get("finished_at")
    sc["elapsed_s"] = round(((finished or now) - started), 1) if started else 0.0
    sc.pop("stage_started", None)
    sc["state"] = sc.get("state") or ("pending" if s.pending else "idle")
    sc["version_id"] = sc.get("version_id") or s.current_version_id
    sc["apply"] = _apply_snapshot(s)
    return sc


# --- apply status -----------------------------------------------------------------
def _apply_start(s: Session, total: int, reanalyze: bool) -> None:
    now = time.time()
    s.apply_status = {
        "state": "running",
        "stage": None,
        "title": "Starting",
        "message": "Starting",
        "fraction": None,
        "done": 0,
        "total": total,
        "reanalyze": reanalyze,
        "started_at": now,
        "updated_at": now,
        "finished_at": None,
        "stage_started": now,
        "stages": [],
        "applied": None,
        "failed": None,
        "error": None,
        "from_version_id": s.current_version_id,
    }


def _apply_progress_for(s: Session):
    def cb(stage: str, message: str, fraction: float | None = None, **facts: Any) -> None:
        st = s.apply_status
        now = time.time()
        if st.get("stage") != stage:
            if st.get("stage"):
                st["stages"].append({"stage": st["stage"], "seconds": round(now - st["stage_started"], 2)})
            st["stage"], st["stage_started"] = stage, now
        st.update(title=APPLY_STAGE_TITLES.get(stage, stage), message=message, fraction=fraction, updated_at=now)
        if "done" in facts:
            st["done"] = facts["done"]
        if "total" in facts:
            st["total"] = facts["total"]

    return cb


def _apply_finish(s: Session, res: dict[str, Any] | None = None, error: str | None = None) -> None:
    st = s.apply_status
    now = time.time()
    if st.get("stage"):
        st["stages"].append({"stage": st["stage"], "seconds": round(now - st.get("stage_started", now), 2)})
    st.update(state="error" if error else "done", finished_at=now, updated_at=now, error=error, fraction=None if error else 1.0)
    if res is not None:
        st["applied"], st["failed"] = len(res.get("applied", [])), len(res.get("failed", []))
    st["title"] = "Failed" if error else "Done"
    st["message"] = error or f"{st.get('applied') or 0} change(s) applied" + (f", {st['failed']} not applied" if st.get("failed") else "")
    st["stage"] = st.get("stage") if error else "done"


def _apply_snapshot(s: Session) -> dict[str, Any] | None:
    if not s.apply_status:
        return None
    st = dict(s.apply_status)
    started, finished = st.get("started_at"), st.get("finished_at")
    st["elapsed_s"] = round(((finished or time.time()) - started), 1) if started else 0.0
    st.pop("stage_started", None)
    return st


def _tracked_apply(s: Session, ops: list[dict[str, Any]], reanalyze: bool, defer_verify: bool = False, source: str | None = None) -> dict[str, Any]:
    """apply_operations on the current version, every stage and every written
    operation reported into `s.apply_status` (GET /status, field "apply")."""
    _apply_start(s, len(ops), reanalyze)
    try:
        # Through the session's Excel worker -- the copy is already open when the previous
        # version's check opened it; Excel dead or hung -> the way it was done before
        worker = excel_worker.worker_for(s.id, s.work_dir / "apply")
        res = None
        if worker is not None:
            try:
                res = worker.apply(s.current_path, ops, progress=_apply_progress_for(s), defer_verify=defer_verify, recalc_ahead=source == "loop")
            except excel_worker.WorkerUnavailable as exc:
                perf_record("excel_worker_fallback", 0.0, session=s.id, error=str(exc)[:160])
                _apply_start(s, len(ops), reanalyze)
        if res is None:
            res = apply_operations(s.current_path, s.work_dir / "apply", ops, progress=_apply_progress_for(s), defer_verify=defer_verify)
    except Exception as exc:
        _apply_finish(s, error=f"apply failed: {plain_error(exc)}")
        raise
    if res["status"] == "NOT_APPLICABLE":
        _apply_finish(s, error=res.get("message", "nothing to apply"))
        raise HTTPException(status_code=422, detail=res.get("message", "nothing to apply"))
    _apply_finish(s, res)
    return res


def _status_counts(findings: list[dict[str, Any]]) -> dict[str, int]:
    out: dict[str, int] = {}
    for f in findings:
        st = str(f.get("status") or "UNKNOWN")
        out[st] = out.get(st, 0) + 1
    return out


def _sync_recalc_finding(s: Session) -> None:
    """READY-001 ("successful recalculation") is NOT_SUPPORTED in every
    analysis pass -- the pass never runs Excel -- and, worst status winning,
    it would hold the whole report at NOT_SUPPORTED after a clean Recalculate.
    So when the current version has a real recalculation, the finding says what
    it found (PASS when no new error, ERROR otherwise) and the report's status
    and summary are recomputed."""
    if not s.result:
        return
    rc = recalc_state(s.recalc, s.current_version_id)
    if not rc["ran"]:
        return
    report = s.result["validation_report"]
    finding = next((f for f in report["findings"] if f.get("rule_id") == "READY-001"), None)
    if finding is None:
        return
    short = next((v["label"].split(" — ")[0] for v in s.versions if v["id"] == s.current_version_id), s.current_version_id)
    finding["status"] = "PASS" if rc["clean"] else "ERROR"
    finding["message"] = f"Recalculated {short} in Excel: {rc['message']}."
    _recount_report(report)


def _recount_report(report: dict[str, Any]) -> None:
    """The report's status and summary again, after a finding was changed in place."""
    statuses = [Status(f["status"]) for f in report["findings"]]
    report["status"] = aggregate_status(statuses).value
    summary = report.setdefault("summary", {})
    summary["status_counts"] = _status_counts(report["findings"])
    summary["not_supported_rule_ids"] = [f["rule_id"] for f in report["findings"] if f["status"] == "NOT_SUPPORTED"]
    summary["blocking_rule_ids"] = [f["rule_id"] for f in report["findings"] if f["status"] in ("ERROR", "REQUIRES_USER_INPUT", "NOT_SUPPORTED")]


def _analysis_config(s: Session) -> dict[str, Any]:
    """The configuration every rule reads, with what the session knows about the outside input files
    (INP-010: what the macros open, the files picked, whether this line has them brought in)."""
    config = load_config()
    config["outside_inputs"] = _outside_public(s)
    return config


def _patch_outside_finding(s: Session, result: dict[str, Any] | None) -> None:
    """INP-010 again with what the session knows now: the outside files are found after the rules ran
    (the import table needs the scan), and the picked files can change without a new scan."""
    if not result:
        return
    report = result["validation_report"]
    i = next((k for k, f in enumerate(report["findings"]) if f.get("rule_id") == OUTSIDE_RULE), None)
    rule = engine().get(OUTSIDE_RULE)
    if i is None or rule is None:
        return
    new = Finding.model_validate(run_rule(rule, result["workbook_analysis"], _analysis_config(s))).model_dump(mode="json")
    new["correction_available"] = bool(report["findings"][i].get("correction_available"))
    report["findings"][i] = new
    _recount_report(report)


def _analyze(s: Session, path: Path, progress=None) -> dict[str, Any]:
    label, module = MODE_MAP[s.mode]
    work = s.work_dir / "analysis" / s.current_version_id
    previous = s.result["validation_report"] if s.result else None
    with formula_memo():  # the rules and the planners parse each formula once between them
        result = module.run(path, work, _analysis_config(s), engine=engine(), ignore_sheets=s.ignore_sheets or None, progress=progress)
        analysis = result["workbook_analysis"]
        if Path(path) == _upload_path(s):
            if s.original_environment is None:
                s.original_environment = environment_values(analysis)
            # what the macros open, the import table and the picked files: read off the upload once
            # its scan exists (the table needs it), then INP-010 again with them
            _scan_outside(s, Path(path), analysis)
            if s.outside_inputs is not None:
                _patch_outside_finding(s, result)
        analysis["original_environment_values"] = _original_environment(s)
        if progress is not None:
            progress("plan", "Planning the fixes", None)
        s.plan = plan_actions(result["workbook_analysis"], result["validation_report"], s.grid_names or None, s.pivot_mode)
    report = result["validation_report"]
    # "Fix available" means the current prep plan has an operation for the rule.
    fixable = {rid for a in s.plan if a["count"] > 0 for rid in a["rule_ids"]} | {o["rule_id"] for a in s.plan for o in a["operations"]}
    for f in report["findings"]:
        f["correction_available"] = f["rule_id"] in fixable
    report["_version_id"] = s.current_version_id
    s.result = result
    # a baseline's #REF! formulas from its own scan: the upload's, and those of the version the
    # outside files were brought into (the baseline of its line)
    if not s.ignore_sheets and str(path) not in s.ref_formulas and Path(path) in (_upload_path(s), _original_path(s)):
        refs = {
            (f["sheet"], f["cell"]) for f in result["workbook_analysis"]["workbooks"][0]["formulas"]
            if not f.get("array_ref") and not f.get("data_table") and "#REF!" in str(f.get("formula") or "")
        }
        s.ref_formulas[str(path)] = refs
        if s.original_ref_formulas is None and Path(path) == _upload_path(s):
            s.original_ref_formulas = refs
    _sync_recalc_finding(s)  # a re-analysis of a version already recalculated keeps that result
    delta = _delta(previous, report, s.current_version_id)
    s.last_delta = delta  # shown again when the page reopens the session
    summary = _summary(result["workbook_analysis"])
    summary["size"] = s.size
    # The verdict and the Prep gauge travel with every analysis
    readiness = s.readiness()
    counts = plan_counts(s.plan)
    s.plan_history.append({
        "version_id": s.current_version_id,
        "blocking_ops": readiness["prep"]["blocking_ops"] if readiness else 0,
        "optional_ops": readiness["prep"]["optional_ops"] if readiness else 0,
        "total_ops": sum(counts.values()),
        "blocking_findings": readiness["blocking_count"] if readiness else 0,
        # the repair gauge: every finding still open, by status, so the Fix panel
        # and Findings can show the run going down even when Prep has no work left
        "open_findings": sum(1 for f in report["findings"] if f.get("status") != "PASS"),
        "status_counts": _status_counts(report["findings"]),
        "by_action": counts,
    })
    return {
        "summary": summary,
        "report": {k: v for k, v in report.items() if k != "_version_id"},
        "plan": s.plan,
        "delta": delta,
        "readiness": readiness,
        "prep_progress": prep_progress(s.plan_history, s.apply_history),
        # the pivot setting travels with every analysis too: an upload must show its choices
        "pivot_mode": s.pivot_mode,
        "pivot_modes": PIVOT_MODES,
        # what the macros open from outside and which picked file goes with each (None: nothing)
        "outsideInputs": _outside_public(s),
    }


def _convert(s: Session, path: Path, progress) -> Path:
    """An .xlsb is converted through Excel first (openpyxl cannot read it);
    the converted file becomes a version of its own."""
    target = "xlsm" if has_vba_project(path) else "xlsx"
    progress("convert", f"Converting .xlsb to .{target} through Excel (openpyxl cannot read .xlsb)", None)
    conv = convert_output_format(path, s.work_dir / "convert", target)
    if conv["status"] != "APPLIED":
        raise HTTPException(status_code=500, detail=conv.get("message", "conversion failed"))
    out = Path(conv["output_path"])
    _add_version(s, out, "convert", f"converted .xlsb to .{target} via Excel", [], None, conv["change_log_entry"]["output_sha256"])
    return out


def _run_scan(s: Session, path: Path | None = None) -> dict[str, Any]:
    """One scan of `path` (default: the current version) with the session's
    ignore list, converting an .xlsb first, reporting every stage into
    `s.scan` (GET /status) -- one at a time per session."""
    if not s.lock.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="a scan is already running for this session")
    try:
        path = Path(path or s.current_path)
        needs_convert = path.suffix.lower() == ".xlsb"
        _scan_start(s, needs_convert)
        progress = _progress_for(s)
        try:
            if needs_convert:
                excel_worker.release(s.id)
                path = _convert(s, path, progress)
            with perf_stage("scan", file=path.name):
                out = _analyze(s, path, progress)
        except HTTPException as exc:
            _scan_finish(s, error=str(exc.detail))
            raise
        except Exception as exc:
            _scan_finish(s, error=f"analysis failed: {plain_error(exc)}")
            raise HTTPException(status_code=500, detail=f"analysis failed: {plain_error(exc)}") from exc
        _scan_finish(s)
        s.pending = False
        current = next((v for v in s.versions if v["id"] == s.current_version_id), s.versions[-1])
        return {**out, "version": current, "versions": s.versions, "size": s.size, "scan": _scan_snapshot(s)}
    finally:
        s.lock.release()


def _parse_ignore(s: Session, raw: Any) -> list[str]:
    """The sheets to skip, validated against the sheet list read from the
    package (when it could be read): unknown names are refused, and at least
    one sheet must remain to scan."""
    if raw in (None, ""):
        return []
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            raw = [x.strip() for x in raw.split(",") if x.strip()]
    if not isinstance(raw, list) or not all(isinstance(x, str) for x in raw):
        raise HTTPException(status_code=422, detail="ignore_sheets must be a list of sheet names")
    names = list(dict.fromkeys(x for x in raw if x))
    known = [sh["name"] for sh in (s.size or {}).get("sheets", [])]
    if known:
        unknown = [n for n in names if n not in known]
        if unknown:
            raise HTTPException(status_code=422, detail=f"unknown sheet(s): {', '.join(unknown)}")
        if len(names) >= len(known):
            raise HTTPException(status_code=422, detail="at least one sheet must be scanned")
    return names


def _apply_result(res: dict[str, Any], output_name: str) -> dict[str, Any]:
    return {
        "status": res["status"],
        "method": res.get("method", "excel_com"),
        "output_path": str(res.get("output_path", "")),
        "output_name": output_name,
        "applied": [_public_op(o) for o in res.get("applied", [])],
        "failed": [_public_op(o) for o in res.get("failed", [])],
        "verified_opens_in_excel": res.get("verified_opens_in_excel"),
        "warnings": res.get("warnings", []),
        "message": res.get("message", ""),
    }


def _before_apply(s: Session) -> dict[str, Any]:
    """What the analysis said before an Apply -- the other half of its outcome."""
    return {"plan": s.plan, "report": s.result["validation_report"] if s.result else None, "version_id": s.current_version_id}


def _outcome(s: Session, ops: list[dict[str, Any]], res: dict[str, Any], before: dict[str, Any], reanalyzed: bool) -> dict[str, Any]:
    """The Apply, repair by repair (app.readiness.apply_outcome), against the
    analysis that preceded it and -- when one ran -- the one that followed."""
    after_report = s.result["validation_report"] if reanalyzed and s.result else None
    outcome = apply_outcome(ops, res, before["plan"], before["report"], s.plan if reanalyzed else None, after_report, s.current_version_id, before["version_id"])
    s.apply_history.append(apply_record(outcome))
    return outcome


def _validate_ops(raw_ops: Any) -> list[dict[str, Any]]:
    if not isinstance(raw_ops, list) or not raw_ops:
        raise HTTPException(status_code=422, detail="operations must be a non-empty list")
    ops = []
    for i, o in enumerate(raw_ops, start=1):
        if not isinstance(o, dict) or o.get("op") not in ASSISTANT_OPS | STRUCTURAL_OPS:
            raise HTTPException(status_code=422, detail=f"operation {i}: unknown op {o.get('op') if isinstance(o, dict) else o!r}")
        if not isinstance(o.get("sheet"), str):
            raise HTTPException(status_code=422, detail=f"operation {i}: sheet is required")
        ops.append({"action_id": "frontend", "rule_id": "UI", **o})
    return ops


# --- outside input files (app/outside_inputs.py) -------------------------------------------
# A workbook whose macro opens other files and pastes their data in: the user picks the workbook
# and those files together; they are kept beside the session (never as versions), matched to the
# workbook's import table, and brought in by the workbook's own macro on a copy (the job below).
def _main_name(s: Session) -> str:
    """The name of the workbook the user uploaded."""
    return Path(s.paths[s.versions[0]["id"]]).name if s.versions else ""


def _inputs_public(s: Session) -> list[dict[str, Any]]:
    return [{"name": i["name"], "bytes": i["bytes"], "stand_in": bool(i.get("stand_in"))} for i in s.inputs]


def _stand_in_names(s: Session) -> set[str]:
    """The input files the assistant made (stand-ins), not the user: a real file of the same name replaces one."""
    return {i["name"] for i in s.inputs if i.get("stand_in")}


def _check_input_names(names: list[str]) -> None:
    """Every picked file is a workbook or a data file a macro can open; anything else is refused (400)."""
    allowed = sorted(outside_inputs.INPUT_FILE_EXTENSIONS)
    for n in names:
        if Path(n).suffix.lower() not in outside_inputs.INPUT_FILE_EXTENSIONS:
            raise HTTPException(status_code=400, detail=f"{n}: Mind Ready takes the workbook and the files its macros open ({', '.join(allowed)}), not this kind of file.")


def _save_upload(upload: UploadFile, folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / Path(upload.filename or "file").name
    with path.open("wb") as out:
        shutil.copyfileobj(upload.file, out)
    return path


def _keep_input(s: Session, path: Path) -> dict[str, Any]:
    """Record a received input file; a file of the same name replaces the earlier one (any language:
    names compare as Excel shows them)."""
    from app.inventory import sha256_of

    entry = {"name": path.name, "path": str(path), "bytes": path.stat().st_size, "sha256": sha256_of(path)}
    key = outside_inputs.normalize_name(path.name)
    s.inputs = [i for i in s.inputs if outside_inputs.normalize_name(i["name"]) != key] + [entry]
    return entry


def _input_folder(s: Session) -> Path:
    s.input_seq += 1
    return s.work_dir / "inputs" / f"{s.input_seq:02d}"


def _input_path(s: Session, name: str | None) -> Path | None:
    return next((Path(i["path"]) for i in s.inputs if i["name"] == name), None)


def _outside_vba_of(path: Path) -> dict[str, Any]:
    """What the macros of `path` open and which macro runs the import (outside_inputs.scan_vba,
    without the per-procedure detail)."""
    modules = model_map_module.read_vba(path) if has_vba_project(path) else []
    found = outside_inputs.scan_vba(modules)
    # messages: every MsgBox line as the source writes it (run_import shows a message box's text with it)
    return {"imports": found["imports"], "import_macro": found["import_macro"], "chain": found["chain"], "messages": found.get("messages") or []}


def _pick_main(paths: list[Path]) -> tuple[Path, dict[str, Any] | None]:
    """The main workbook among files picked together: the one whose macros open outside files (an
    import macro found), else the only one with macros, else the largest workbook. Returns it and
    what its macros open when that was read."""
    books = [p for p in paths if p.suffix.lower() in ALLOWED_SUFFIXES]
    if not books:
        raise HTTPException(status_code=422, detail="Pick the workbook itself (.xlsx, .xlsm or .xlsb) together with its input files.")
    read = {p: _outside_vba_of(p) for p in books if has_vba_project(p)}
    importers = [p for p, vba in read.items() if vba["imports"] and vba["import_macro"]]
    if importers:
        main = max(importers, key=lambda p: p.stat().st_size)
    elif len(read) == 1:
        main = next(iter(read))
    else:
        main = max(books, key=lambda p: p.stat().st_size)
    return main, read.get(main)


def _scan_outside(s: Session, path: Path, analysis: dict[str, Any]) -> None:
    """After a scan of the upload: what its macros open (read once), its import table and which
    picked file goes with each row. Never fails the scan: a table that cannot be read is empty."""
    if s.outside_vba is None:
        s.outside_vba = _outside_vba_of(path)
    if not s.outside_vba["imports"]:
        s.outside_inputs = None
        return
    try:
        table = outside_inputs.import_table_from_analysis(analysis)
    except Exception as exc:  # an odd workbook: say so in the table's place, never stop the scan
        perf_record("outside_inputs_table_failed", 0.0, file=path.name, error=str(exc)[:160])
        table = []
    s.outside_inputs = {"version_id": s.current_version_id, "vba": s.outside_vba, "table": table, "match": None}
    _rematch(s)


def _rematch(s: Session) -> None:
    """Which picked file goes with each row of the import table (no rescan)."""
    if s.outside_inputs is not None:
        s.outside_inputs["match"] = outside_inputs.match_files(s.outside_inputs["table"], [i["name"] for i in s.inputs])


def _match_rows(s: Session) -> list[dict[str, Any]]:
    """The rows of the import table, each with the picked file that goes with it (none: missing)."""
    oi = s.outside_inputs or {}
    if not oi:
        return []
    match = oi.get("match") or outside_inputs.match_files(oi["table"], [i["name"] for i in s.inputs])
    return list(match.get("rows") or [])


def _inputs_fingerprint(s: Session, rows: list[dict[str, Any]] | None = None) -> str:
    """The picked files the rows of the import table read, as one sha256: each row's cell with the
    content (sha256) of the file that goes with it, or none. A bring-in records it; a file replaced
    since (same name, other content) or a row matched since changes it."""
    sha = {i["name"]: i.get("sha256") or "" for i in s.inputs}
    lines = sorted(
        f"{r.get('sheet')}!{r.get('cell')}\t" + ((sha.get(r["file"]) or "?") if r.get("file") else "-")
        for r in (_match_rows(s) if rows is None else rows)
    )
    return hashlib.sha256("\n".join(lines).encode("utf-8")).hexdigest()


# why a line that went through a bring-in does not count as brought in (INP-010 and the card say it)
_FILES_CHANGED = "the input files changed since they were brought in"
_NOT_FINISHED = "the last bring-in did not finish"


def _brought_in_state(s: Session, version_id: str | None = None) -> tuple[bool, str | None]:
    """Whether the line of `version_id` (default: the current version) has the outside files brought
    in -- through a bring-in that finished, with the files picked now -- and, when its line went
    through a bring-in that does not count, why (_FILES_CHANGED / _NOT_FINISHED; None otherwise)."""
    line = _lineage(s, version_id or s.current_version_id)
    for v in reversed(line):
        recorded = s.brought_in_files.get(v["id"])
        if recorded is not None:
            return (True, None) if recorded == _inputs_fingerprint(s) else (False, _FILES_CHANGED)
    if (s.outside_run or {}).get("state") != "running" and any(v.get("source") in ("import", "mind_inputs") for v in line):
        return False, _NOT_FINISHED  # stopped half-way, or its Mind inputs were refused (they changed values)
    return False, None


def _brought_in(s: Session, version_id: str | None = None) -> bool:
    """The version descends from one a finished bring-in made, and the picked files are still the ones it read."""
    return _brought_in_state(s, version_id)[0]


def _outside_public(s: Session) -> dict[str, Any] | None:
    """What the macros open from outside and which picked file goes with each row of the import
    table (the "outsideInputs" of every scan answer); None when the macros open no outside file."""
    oi = s.outside_inputs
    if not oi:
        return None
    match = oi.get("match") or outside_inputs.match_files(oi["table"], [])
    rows = match["rows"]
    if rows:
        summary = outside_inputs.describe_match(match)
    elif oi["vba"].get("import_macro"):
        summary = "The macros open outside files, but no cell names them."
    else:
        summary = "The macros open outside files, but which macro brings them in is not clear."
    brought_in, why_not = _brought_in_state(s) if s.versions else (False, None)
    made = _stand_in_names(s)
    return {
        "imports": True,
        "macro": oi["vba"].get("import_macro"),
        "calls": [{"procedure": c.get("procedure"), "call": c.get("call"), "line": c.get("line")} for c in oi["vba"].get("imports") or []],
        "rows": [
            {"cell": f"{r['sheet']}!{r['cell']}", "expected": r.get("expected"), "keys": list(r.get("keys") or []), "target_sheet": r.get("target_sheet"),
             "file": r.get("file"), "how": r.get("how"), "candidates": list(r.get("candidates") or []), "stand_in": r.get("file") in made}
            for r in rows
        ],
        # the matched files the assistant made (stand-ins built from the model's own numbers): best
        # results come with the actual files -- INP-010 and the card say so
        "stand_ins": [r["file"] for r in rows if r.get("file") in made],
        # the assistant can make the missing files (an AI key is saved)
        "assistant": llm_available(),
        "matched": int(match["matched"]),
        "missing": [f"{r['sheet']}!{r['cell']}" for r in match["missing"]],
        "extra": list(match["extra"]),
        "summary": summary,
        "brought_in": brought_in,
        # why the line's bring-in does not count, in words (the input files changed since, it did not
        # finish); None when brought in or when the line never went through one
        "brought_in_reason": why_not,
        # the /Input grids the bring-in made for this line (their size is the model's import area)
        "mind_inputs": [Path(n).stem for n in _mind_input_files(_mind_inputs_dir_for(s))] if s.versions else [],
    }


# --- problem reports (app/feedback.py) ---------------------------------------------------
# The step a route belongs to, in the words of the screens (the route's path, as declared)
_STEPS = {
    "/api/sessions": "Upload",
    "/api/sessions/{session_id}/analyze": "Check the workbook",
    "/api/sessions/{session_id}/reanalyze": "Check the workbook",
    "/api/sessions/{session_id}/apply": "Fix the formulas",
    "/api/sessions/{session_id}/auto-fix": "Fix all issues",
    "/api/sessions/{session_id}/mind-loop-plan": "Build the Mind loop",
    "/api/sessions/{session_id}/mind-loop-plan/apply": "Build the Mind loop",
    "/api/sessions/{session_id}/mind-instances-plan": "Build the Mind instances",
    "/api/sessions/{session_id}/mind-instances-plan/apply": "Build the Mind instances",
    "/api/sessions/{session_id}/mind-instances/download": "Build the Mind instances",
    "/api/sessions/{session_id}/instance-finder": "Find where instances could help",
    "/api/sessions/{session_id}/recalculate": "Recalculate",
    "/api/sessions/{session_id}/numbers-check": "Compare the numbers",
    "/api/sessions/{session_id}/chat": "Assistant",
    "/api/sessions/{session_id}/chat/stream": "Assistant",
    "/api/sessions/{session_id}/suggest": "Assistant",
    "/api/sessions/{session_id}/grid-names": "Grid Namer",
    "/api/sessions/{session_id}/grid-namer": "Grid Namer",
    "/api/sessions/{session_id}/mind-loop": "Run in Mind",
    "/api/sessions/{session_id}/model-map": "Model map",
    "/api/sessions/{session_id}/reports": "Download the report",
    "/api/sessions/{session_id}/files/{file_name}": "Download",
    "/api/sessions/{session_id}/pivot-mode": "Pivot table setting",
    "/api/sessions/{session_id}/sheet-cells": "Workbook view",
    "/api/sessions/{session_id}/cells": "Workbook view",
    "/api/sessions/{session_id}/estimate": "Time estimate",
    "/api/sessions/{session_id}/readiness": "Verdict",
    "/api/sessions/{session_id}/inputs": "Input files",
    "/api/sessions/{session_id}/outside-inputs": "Bring in the outside files",
    "/api/sessions/{session_id}/outside-inputs/run": "Bring in the outside files",
    "/api/assistant-key": "AI key",
}


def _workbook_names(s: Session | None) -> list[str]:
    """The workbook's own names (file, sheets, defined names): masked in every problem report."""
    if s is None:
        return []
    names: list[str] = []
    for p in list(s.paths.values())[:1]:
        names.append(Path(p).stem)
    names += [Path(i["name"]).stem for i in s.inputs]  # the outside files' names are the client's too
    names += [sh.get("name", "") for sh in (s.size or {}).get("sheets", [])]
    if s.result:
        wb = s.result["workbook_analysis"]["workbooks"][0]
        names += [sh.get("name", "") for sh in wb.get("sheets", [])]
        names += [str(d.get("name", "")).split("!")[-1] for d in wb.get("defined_names", [])]
    return [n for n in dict.fromkeys(names) if n]


def _report_problem(exc: BaseException | None, step: str, s: Session | None = None, **kw: Any) -> dict[str, Any]:
    """One incident for the Tel Aviv team, with the rule and the stage the scan or the Apply was
    in when it failed, when the session knows them."""
    if s is not None:
        if not kw.get("rule_id") and (s.scan or {}).get("state") in ("error", "running"):
            kw["rule_id"] = s.scan.get("rule_id")
        if not kw.get("stage"):
            busy = s.apply_status if (s.apply_status or {}).get("state") in ("error", "running") else s.scan if (s.scan or {}).get("state") in ("error", "running") else {}
            if busy.get("stage"):
                kw["stage"] = str(busy["stage"])
    try:
        counts = feedback.workbook_counts(_summary(s.result["workbook_analysis"]) if s and s.result else None, s.size if s else None)
    except Exception:
        counts = None
    incident = feedback.record(exc, step=step, workbook=counts, names=_workbook_names(s), **kw)
    incident["plain"] = _plain_problem(exc, step, kw.get("failed_ops"))  # for this screen only, never sent
    return incident


def _plain_problem(exc: BaseException | None, step: str, failed_ops: list[dict[str, Any]] | None) -> str:
    """What stopped the user's work, in one plain sentence, for the box on their own screen (their
    workbook's words may appear here: it is not part of the message to the team)."""
    if failed_ops:
        # the rules the refused changes served, each with its plain name (the screens' own changes, "UI", are no rule)
        rules = sorted({str(o.get("rule_id")) for o in failed_ops if rule_name(o.get("rule_id"))})
        return (
            f"{step}: Excel did not accept {len(failed_ops)} of the changes" + (f", for {cite_all(rules[:4])}" if rules else "")
            + ". The other changes were made. The ones left out are listed with Excel's reason, and the check that follows shows what still blocks Mind."
        )
    said = ""
    if exc is not None:
        said = plain_com_error(exc) if is_com_error(exc) else plain_com_errors_in(str(getattr(exc, "detail", "") or exc or type(exc).__name__))
    said = re.sub(r"\s+", " ", said).strip().rstrip(".")[:220]
    who = "The app" if step == "Other" else step
    return f"{who} stopped on an error" + (f": {said}" if said else "") + ". Your original file is not changed."


def _problem_fields(incident: dict[str, Any] | None) -> dict[str, Any]:
    """The `incident` ref (for the message to the team) and the plain sentence (for the screen)."""
    return {"incident": incident["ref"], "problem": incident.get("plain")} if incident else {}


def _problem_from_request(request: Request, exc: BaseException) -> dict[str, Any] | None:
    try:
        route = request.scope.get("route")
        step = _STEPS.get(getattr(route, "path", ""), "Other")
        s = SESSIONS.get(str(request.path_params.get("session_id") or ""))
        return _report_problem(exc, step, s)
    except Exception:
        return None  # the report must never hide the error itself


@app.exception_handler(StarletteHTTPException)
async def _http_error(request: Request, exc: StarletteHTTPException):
    """As FastAPI answers; a 5xx also gets an incident (its `ref` as `incident`), the error it
    wraps (`raise HTTPException(...) from exc`) reported rather than the wrapper."""
    if exc.status_code < 500:
        return await http_exception_handler(request, exc)
    incident = _problem_from_request(request, exc.__cause__ or exc)
    detail = plain_com_errors_in(exc.detail) if isinstance(exc.detail, str) else exc.detail
    body: dict[str, Any] = {"detail": detail, **_problem_fields(incident)}
    return JSONResponse(body, status_code=exc.status_code, headers=getattr(exc, "headers", None))


@app.exception_handler(Exception)
async def _crash(request: Request, exc: Exception):
    """An error no route caught: a 500 with its incident."""
    incident = _problem_from_request(request, exc)
    body: dict[str, Any] = {"detail": _exc_text(exc)[:400], **_problem_fields(incident)}
    return JSONResponse(body, status_code=500)


def _exc_text(exc: BaseException) -> str:
    """An error no step expected, for the screen: Excel's errors said plainly, others with their type.
    A sentence that already says what Excel did ("Excel stopped with an error and did not say why",
    "Excel closed ...") stands alone: "Excel stopped with an error: " goes only in front of one that
    does not name Excel ("the macro looked for a sheet ...", Excel's own words)."""
    if not is_com_error(exc):
        return plain_error(exc, with_type=True)
    said = plain_com_error(exc)
    return said if said.startswith("Excel ") else f"Excel stopped with an error: {said}"


# --- API ----------------------------------------------------------------------------------
@app.get("/api/health")
def health() -> dict[str, Any]:
    from app.sizing import size_threshold_bytes

    return {
        "version": app.version,
        "excel": com_available(),
        "excel_worker": excel_worker.enabled() and com_available(),
        "assistant": llm_available(),
        "rules": len(engine().active_rules()),
        "frontend": DIST_DIR.is_dir(),
        # Deferred upload + sheet selection + scan status are available
        "upload_gate": True,
        "size_threshold_mb": round(size_threshold_bytes(load_config()) / MB, 1),
        # The assistant's model and POST /chat/stream
        "model": DEFAULT_MODEL_ID,
        "streaming": True,
        # POST /auto-fix ("Fix all automatically" on the Recalculate step)
        "autofix": True,
        # several files per upload, POST .../inputs and .../outside-inputs/run (app/outside_inputs.py)
        "outside_inputs": True,
        # a shared copy's end date (app/expiry.py): "expires": "2026-12-15", "days_left", "expired" and
        # its message; all None / False for the developer version. Answered after the date too.
        **expiry.current().health(),
        # GET .../instance-finder ("Find where instances could help": the finder is not a rule)
        "finder": True,
    }


@app.get("/api/plain-codes")
def plain_codes() -> dict[str, Any]:
    """The plain words for every code the screens show: each rule's name (rules/plain_names.yaml) and
    what each of Excel's error values means -- one glossary, the same as the server's (app/plain_codes.py)."""
    return glossary()


@app.get("/api/assistant-key")
def assistant_key() -> dict[str, Any]:
    """Whether a gateway key is saved, and its last 4 characters (never the key)."""
    key = ai_key.load()
    return {"saved": key is not None, "last4": ai_key.last4(key)}


@app.post("/api/assistant-key")
def save_assistant_key(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    """Check a pasted key with the gateway, then save it encrypted for this Windows user
    (app/ai_key.py). A key the gateway refuses is not saved."""
    key = str(payload.get("key") or "").strip()
    problem = ai_key.check_format(key)
    if problem:
        raise HTTPException(status_code=422, detail=problem)
    accepted, note = check_api_key(key)
    if not accepted:
        raise HTTPException(status_code=422, detail=note)
    try:
        ai_key.save(key)
    except (OSError, RuntimeError) as exc:
        raise HTTPException(status_code=500, detail=f"The key could not be saved: {exc}") from exc
    return {"saved": True, "last4": ai_key.last4(key), "message": note}


@app.post("/api/feedback")
def compose_feedback(payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """The message for "Send feedback to the Tel Aviv team": the incident `ref` (none from the
    header), the user's `note`, `useAssistant` (the assistant writes the opening lines when a key
    is saved), `sessionId` (the workbook's counts when no incident holds them). In codes, with
    nothing from the workbook: the screen shows it, then copies it and opens an email."""
    ref = str(payload.get("ref") or "").strip()
    incident = feedback.find(ref) if ref else None
    note = str(payload.get("note") or "")[:2000]
    s = SESSIONS.get(str(payload.get("sessionId") or ""))
    ai_text = None
    if payload.get("useAssistant", True):
        rule = engine().get(str((incident or {}).get("rule_id") or "").split(",")[0].strip())
        try:
            ai_text = feedback.assistant_lines(incident, note, (rule or {}).get("title"))
        except Exception:
            ai_text = None
    counts = None
    if incident is None and s is not None and s.result:
        counts = feedback.workbook_counts(_summary(s.result["workbook_analysis"]), s.size)
    return feedback.compose(incident, note, ai_text, counts)


@app.post("/api/sessions")
def create_session(
    file: UploadFile | None = File(None), files: list[UploadFile] | None = File(None),
    mode: str = Form("plan"), defer: str = Form(""), ignore_sheets: str = Form(""),
) -> dict[str, Any]:
    """Upload a workbook. The file is stored and *inspected* (sizes, sheet
    list -- from the package alone, nothing is parsed) before anything else.

    `files` (repeated, optional): every file the user picked together -- the
    workbook and the outside files its macros open. Without `file`, the main
    workbook is picked among them (the one whose macros open outside files,
    else the only one with macros, else the largest workbook); the others are
    kept as its input files (<work>/inputs/<nn>/<name>, never versions) and
    matched to its import table after the scan ("outsideInputs").

    `defer` truthy: stop there and return `{pending: true, size}` so the
    front-end can ask about skipping sheets when the workbook is above the
    size threshold, then POST .../analyze. Otherwise (default; scripts and
    older front-ends) the file is converted/analysed at once, honouring an
    optional `ignore_sheets` JSON list."""
    if mode not in MODE_MAP:
        raise HTTPException(status_code=422, detail=f"unknown mode {mode}")
    picked = [f for f in (files or []) if f is not None and (f.filename or "").strip()]
    if file is None and not picked:
        raise HTTPException(status_code=422, detail="upload an .xlsx, .xlsm or .xlsb file")
    if file is not None:
        name = Path(file.filename or "workbook.xlsx").name
        if Path(name).suffix.lower() not in ALLOWED_SUFFIXES:
            raise HTTPException(status_code=422, detail="upload an .xlsx, .xlsm or .xlsb file")
    _check_input_names([Path(f.filename or "").name for f in picked])
    if file is None and not any(Path(f.filename or "").suffix.lower() in ALLOWED_SUFFIXES for f in picked):
        raise HTTPException(status_code=422, detail="Pick the workbook itself (.xlsx, .xlsm or .xlsb) together with its input files.")
    sid = uuid.uuid4().hex[:12]
    work_dir = Path(tempfile.mkdtemp(prefix="mind_ready_"))
    upload_dir = work_dir / "upload"
    upload_dir.mkdir(parents=True)
    s = Session(id=sid, work_dir=work_dir, mode=mode)
    received = [_save_upload(f, _input_folder(s)) for f in picked]
    if file is not None:
        raw_path = _save_upload(file, upload_dir)
        main_key = outside_inputs.normalize_name(raw_path.name)
        others = [p for p in received if outside_inputs.normalize_name(p.name) != main_key]  # the workbook sent twice is not its own input
    else:
        chosen, vba = _pick_main(received)
        raw_path = upload_dir / chosen.name
        shutil.move(str(chosen), str(raw_path))
        s.outside_vba = vba  # read already: the scan does not read the macros again
        others = [p for p in received if p != chosen]
    for p in others:
        _keep_input(s, p)
    SESSIONS[sid] = s

    from app.inventory import sha256_of

    version = _add_version(s, raw_path, "upload", "original upload", [], None, sha256_of(raw_path))
    s.size = size_gate(raw_path, load_config())
    s.ignore_sheets = _parse_ignore(s, ignore_sheets)
    picked_out = {"main": raw_path.name, "inputs": _inputs_public(s)}
    if str(defer).strip().lower() in ("1", "true", "yes", "on"):
        s.pending = True
        return {"sessionId": sid, "pending": True, "mode": mode, "size": s.size, "version": version, "versions": s.versions, "scan": _scan_snapshot(s), **picked_out}
    return {"sessionId": sid, **_run_scan(s, raw_path), **picked_out}


@app.get("/api/sessions/{session_id}")
def session_state(session_id: str) -> dict[str, Any]:
    """The state of an analysed session, in the shape an upload returns, so the web app
    can open it again (`/?session=<id>`, or `/model-map?session=<id>`) after a reload."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    current = next((v for v in s.versions if v["id"] == s.current_version_id), s.versions[-1])
    return {
        "sessionId": s.id,
        "mode": s.mode,
        "summary": _summary(s.result["workbook_analysis"]),
        "report": {k: v for k, v in s.result["validation_report"].items() if k != "_version_id"},
        "plan": s.plan,
        "delta": s.last_delta,
        "version": current,
        "versions": s.versions,
        "size": s.size,
        "scan": _scan_snapshot(s),
        "readiness": s.readiness(),
        "prep_progress": prep_progress(s.plan_history, s.apply_history),
        "pivot_mode": s.pivot_mode,
        "pivot_modes": PIVOT_MODES,
        "outsideInputs": _outside_public(s),
        "main": _main_name(s),
        "inputs": _inputs_public(s),
    }


@app.post("/api/sessions/{session_id}/pivot-mode")
def set_pivot_mode(session_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    """What Prep does with the pivot tables -- "formulas" or "values" (every one blocks Mind's
    converter) -- and the plan made again with it (Do everything applies the blocking actions of
    that plan)."""
    s = _session(session_id)
    mode = str(payload.get("mode") or "")
    if mode not in PIVOT_MODES:
        raise HTTPException(status_code=422, detail=f"mode must be one of {', '.join(PIVOT_MODES)}")
    s.pivot_mode = mode
    if s.result:
        s.plan = plan_actions(s.result["workbook_analysis"], s.result["validation_report"], s.grid_names or None, s.pivot_mode)
    return {"pivot_mode": s.pivot_mode, "plan": s.plan, "readiness": s.readiness()}


@app.post("/api/sessions/{session_id}/analyze")
def analyze(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Scan a deferred upload (or scan the current version again) with the
    given `ignore_sheets`. Long call: poll GET .../status meanwhile."""
    s = _session(session_id)
    _refuse_while_bringing_in(s)  # the job scans its own versions (and every sheet of the first)
    if "ignore_sheets" in payload:
        s.ignore_sheets = _parse_ignore(s, payload.get("ignore_sheets"))
    return {"sessionId": s.id, **_run_scan(s)}


@app.get("/api/sessions/{session_id}/status")
def scan_status(session_id: str) -> dict[str, Any]:
    """Where the current (or last) scan is: state pending|running|ready|error,
    stage + title + message, in-stage fraction, overall progress, the sheet
    or rule being worked on, elapsed seconds, per-stage timings."""
    return _scan_snapshot(_session(session_id))


@app.post("/api/sessions/{session_id}/reanalyze")
def reanalyze(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    s = _session(session_id)
    _refuse_while_fixing(session_id)
    vid = payload.get("versionId") or s.current_version_id
    if vid not in s.paths:
        raise HTTPException(status_code=404, detail=f"unknown version {vid}")
    s.current_version_id = vid
    return _run_scan(s, s.paths[vid])


def _apply_and_version(s: Session, ops: list[dict[str, Any]], source: str, who: str, reanalyze: bool, minor: bool = False) -> dict[str, Any]:
    """Write `ops` into a fresh copy of the current version, register the result as the next
    version, re-analyse it when asked, and report the outcome (the shape every Apply returns).
    One Apply at a time per session: a second one is refused while the first runs."""
    if not s.apply_lock.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="an Apply is already running for this session")
    try:
        return _apply_and_version_locked(s, ops, source, who, reanalyze, minor)
    finally:
        s.apply_lock.release()


def _apply_and_version_locked(s: Session, ops: list[dict[str, Any]], source: str, who: str, reanalyze: bool, minor: bool) -> dict[str, Any]:
    before = _before_apply(s)
    reanalyze_wanted = reanalyze
    with perf_stage("apply_operations", file=s.current_path.name, operations=len(ops), source=source):
        # With a re-analysis to run, the check that the new file opens in Excel (a second
        # Excel, ~40 s on a 30 MB model) runs alongside it instead of before it
        res = _tracked_apply(s, ops, reanalyze, defer_verify=reanalyze_wanted, source=source)
    output = Path(res["output_path"])
    finish = res.pop("finish", None)
    checker: threading.Thread | None = None
    if finish is not None:
        def check() -> None:
            try:
                with perf_stage("verify_opens_alongside", file=output.name):
                    finish()
            except Exception as exc:  # the check itself failed: the version stays unverified
                res.setdefault("warnings", []).append(f"The check that the file opens in Excel did not run: {str(exc)[:160]}")

        checker = threading.Thread(target=check, name=f"verify-{s.id}", daemon=True)
        checker.start()

    def label_of() -> str:
        verified = res.get("verified_opens_in_excel")
        return f"{who}: {len(res.get('applied', []))} change(s) via {'Excel' if res.get('method') == 'excel_com' else 'openpyxl'} · {'verified' if verified else 'unverified' if verified is None else 'FAILED TO OPEN'}"

    version = _add_version(s, output, source, label_of(), res.get("applied", []), res.get("verified_opens_in_excel"), res["change_log_entry"]["output_sha256"], minor=minor)
    reanalyzed = reanalyze and res["status"] in ("APPLIED", "PARTIAL")
    scanned: dict[str, Any] = {}
    scan_error: str | None = None
    try:
        if reanalyzed:
            scanned = _run_scan(s, output)
    except HTTPException as exc:
        # the new version is already current: the Apply still answers with it (an error here would
        # leave the page on the old version, and a second Apply would write the same changes twice)
        scan_error = str(exc.detail)
        reanalyzed = False
    finally:
        if checker is not None:
            checker.join()
            version["verified_opens_in_excel"] = res.get("verified_opens_in_excel")
            version["label"] = version["label"].split(" — ")[0] + " — " + label_of()
    out: dict[str, Any] = {"result": _apply_result(res, version["file_name"]), "version": version, "versions": s.versions}
    out.update(scanned)
    if scan_error is not None:
        out["scan_error"] = scan_error
    out["outcome"] = _outcome(s, ops, res, before, reanalyzed)
    out["prep_progress"] = prep_progress(s.plan_history, s.apply_history)  # the gauge, this Apply included
    if res.get("failed"):
        step = {"loop": "Build the Mind loop", "instances": "Build the Mind instances", "assistant": "Assistant", "formula": "Formula fix"}.get(source, "Fix the formulas")
        out.update(_problem_fields(_report_problem(None, step, s, what="not written", stage="apply_writes", failed_ops=res["failed"])))
    return out


@app.post("/api/sessions/{session_id}/apply")
def apply(session_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    s = _session(session_id)
    _refuse_while_fixing(session_id)
    ops = _validate_ops(payload.get("operations"))
    action_ids = {o.get("action_id") for o in ops}
    source = "assistant" if action_ids == {"assistant"} else "formula" if action_ids == {"formula_replacement"} else "prep"
    who = {"assistant": "Assistant", "formula": "Formula fix", "prep": "Prep"}[source]
    minor = payload.get("versionStep") == "minor"  # a fix on the Recalculate step
    return _apply_and_version(s, ops, source, who, bool(payload.get("reanalyze", True)), minor=minor)


@app.post("/api/sessions/{session_id}/suggest")
def suggest(session_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    rule_id, sheet, cell = payload.get("ruleId"), payload.get("sheet"), payload.get("cell")
    finding = next((f for f in s.result["validation_report"]["findings"] if f["rule_id"] == rule_id), None)
    if finding is None:
        raise HTTPException(status_code=404, detail=f"no finding for rule {rule_id}")
    formula = next((x["formula"] for x in s.result["workbook_analysis"]["workbooks"][0]["formulas"] if x["sheet"] == sheet and x["cell"] == cell), "")
    res = suggest_formula_fix(finding, formula)
    return {"suggestion": res.get("suggestion"), "formula": extract_formula(res.get("suggestion")), "message": res.get("message")}


@app.post("/api/sessions/{session_id}/grid-names")
def grid_names(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Ask the assistant to name the grids whose deterministic name carries no
    meaning ('Cashflows C4', 'I'), and fold the answers into the prep plan.

    Nothing is written: the names only change what `create_grid_titles` would
    *propose*, which the user still reviews and approves. `apply: false`
    returns the comparison without changing the session's plan."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    analysis = s.result["workbook_analysis"]
    labels = standalone_labels(analysis)
    targets = []
    for g in all_grids(analysis):
        if g.get("name") and not is_weak_name(g.get("name"), g["sheet"]):
            continue
        deterministic, source = deterministic_name(analysis, g, labels)
        if g.get("name") or is_weak_name(deterministic, g["sheet"]):
            targets.append((g, deterministic, source))
    if not targets:
        return {"available": True, "names": [], "message": "every grid already has a meaningful name", "applied": False}
    contexts = [build_grid_context(analysis, g) for g, _, _ in targets]
    res = suggest_names(contexts)
    if not res["available"]:
        return {"available": False, "names": [], "message": res.get("message"), "applied": False}
    proposed = res["names"]
    rows = [
        {
            "grid": f'{g["sheet"]}!{g["ref"]}',
            "sheet": g["sheet"],
            "ref": g["ref"],
            "size": f'{g["n_rows"]}x{g["n_cols"]}',
            "current": g.get("name"),
            "deterministic": deterministic,
            "deterministic_source": source,
            "suggested": proposed.get(f'{g["sheet"]}!{g["ref"]}'),
        }
        for g, deterministic, source in targets
    ]
    applied = bool(payload.get("apply", True))
    if applied:
        s.grid_names.update({k: v for k, v in proposed.items() if v})
        s.plan = plan_actions(analysis, s.result["validation_report"], s.grid_names or None, s.pivot_mode)
    return {"available": True, "names": rows, "message": res.get("message"), "applied": applied, "plan": s.plan if applied else None, "readiness": s.readiness() if applied else None}


# --- the app runs itself against real Mind ------------------------------------
MIND_LOOPS: dict[str, dict[str, Any]] = {}  # session id -> {state, events, report, started, work_dir}
_MIND_LOOP_LOCK = threading.Lock()


def _loop_running() -> str | None:
    for sid, st in MIND_LOOPS.items():
        if st.get("state") == "running":
            return sid
    return None


def _busy_with(s: Session) -> str | None:
    """What is still changing this session's workbook, in words (None = nothing): Run in Mind must
    wait for it, or it would upload the version before that change."""
    if (s.outside_run or {}).get("state") == "running":
        return "the outside files are still being brought in"
    if (AUTOFIX.get(s.id) or {}).get("state") == "running":
        return "Fix all issues is still running"
    if s.id in RECALCULATING:
        return "a recalculation is still running"
    if (s.apply_status or {}).get("state") == "running" or s.apply_lock.locked():
        return "changes are still being written"
    if (s.scan or {}).get("state") == "running" or s.lock.locked() or s.pending:
        return "the workbook has not been analysed yet" if s.pending and not s.lock.locked() else "the workbook is still being analysed"
    return None


def _current_version(s: Session) -> dict[str, Any] | None:
    return next((v for v in s.versions if v["id"] == s.current_version_id), None)


def _version_public(s: Session) -> dict[str, Any]:
    v = _current_version(s) or {}
    return {"id": s.current_version_id, "label": v.get("label"), "file_name": v.get("file_name") or s.current_path.name}


def _mind_run_estimate(s: Session, use_assistant: bool, run_model: bool, check_numbers: bool, chosen: list[dict[str, Any]]) -> dict[str, Any]:
    try:
        mb = Path(s.current_path).stat().st_size / 2**20
    except OSError:
        mb = 0.0
    scan_s = None
    if (s.scan or {}).get("state") == "ready":
        scan_s = sum(float(x.get("seconds") or 0) for x in s.scan.get("stages", []) if x.get("stage") != "convert") or None
    wbs = ((s.result or {}).get("workbook_analysis") or {}).get("workbooks") or [{}]
    formulas = len(wbs[0].get("formulas") or [])
    ops = sum(int(a.get("count") or 0) for a in chosen)
    return estimate_mind_run(mb, scan_s, formulas, ops, use_assistant=use_assistant, check_numbers=check_numbers, run_model=run_model,
                             inputs=len(_mind_input_files(_mind_inputs_dir_for(s))))


def _mind_inputs_dir_for(s: Session, version_id: str | None = None) -> str | None:
    """The CSV folder of the Mind inputs of a version's line (default: the current one): the folder the
    bring-in job made for the newest "mind_inputs" version the line goes through; None without one."""
    for v in reversed(_lineage(s, version_id or s.current_version_id)):
        folder = s.mind_inputs_dirs.get(v["id"])
        if folder:
            return folder if Path(folder).is_dir() else None
    return None


def _mind_input_files(folder: str | None) -> list[str]:
    """The Mind inputs (CSV / TAB .txt files) in a folder, by name."""
    if not folder or not Path(folder).is_dir():
        return []
    return sorted(p.name for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() in (".csv", ".txt"))


def _mind_choices(s: Session) -> list[dict[str, Any]]:
    """The changes Run in Mind can make before it uploads: the actions of the current version's plan
    that have something to do, each ticked (default_on) as the loop would by default."""
    return [
        {"id": a["id"], "title": a.get("title") or a["id"], "count": int(a.get("count") or 0), "default_on": bool(a.get("default_on")), "level": a.get("level")}
        for a in (s.plan or []) if int(a.get("count") or 0) > 0
    ]


@app.get("/api/sessions/{session_id}/mind-loop/options")
def mind_loop_options(session_id: str, use_assistant: bool = True, run_model: bool = True) -> dict[str, Any]:
    """What Run in Mind will upload (always the session's current version), the changes it can make
    first (ticked as the loop would by default), what must finish first (busy) and the time estimate."""
    s = _session(session_id)
    choices = _mind_choices(s)
    newest = max(s.versions, key=lambda v: (v.get("created_at") or "", v.get("major", 0), v.get("minor", 0)))["id"] if s.versions else None
    return {
        "version": _version_public(s),
        "newest_version_id": newest,
        "busy": _busy_with(s),
        "choices": choices,
        "estimate": _mind_run_estimate(s, use_assistant, run_model, True, [c for c in choices if c["default_on"]]),
        # the Mind inputs uploaded through Mind's Inputs window before the run (the outside files)
        "inputs": _mind_input_files(_mind_inputs_dir_for(s)),
    }


@app.post("/api/sessions/{session_id}/mind-loop")
def start_mind_loop(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Start the autonomous run-in-Mind loop on the session's current version, in
    a background thread. One loop at a time (it owns the Excel session and the
    Mind browser profile). Poll GET for progress. Refused (409) while something
    is still changing the workbook: it would upload the version before that.

    `actions` (optional): the ids of the changes to make before uploading. Every
    other change of the plan is left out, and the loop never switches it on by
    itself; an empty list uploads the version as it is. Without `actions` the
    loop chooses as before (`enable` / `disable`)."""
    s = _session(session_id)
    if s.current_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=422, detail="Mind accepts .xlsx / .xlsm only")
    why = _busy_with(s)
    if why:
        raise HTTPException(status_code=409, detail=f"Wait a moment: {why}. Run in Mind uploads the latest version, so start it once that has finished.")
    if s.outside_inputs and _brought_in_state(s)[1] == _FILES_CHANGED:
        # the version (and its Mind inputs) hold the data of the files picked before: never uploaded as current
        raise HTTPException(status_code=409, detail="The input files changed since they were brought in: Run in Mind would use the data of the old files. "
                                                    "Bring them in again first (the Outside files card on the main screen, or Do everything), then run in Mind.")
    # the outside files made Mind inputs for this line: uploaded through Mind's Inputs window after
    # the model is added and before the run (checked before the loop is registered as running)
    inputs_dir = _mind_inputs_dir_for(s)
    if inputs_dir and "inputs_dir" not in getattr(LoopConfig, "__dataclass_fields__", {}):
        raise HTTPException(status_code=409, detail="This copy of Mind Ready cannot upload the Mind inputs yet: Run in Mind would run the model without the outside files.")
    enable = [str(a) for a in payload.get("enable", [])]
    disable = [str(a) for a in payload.get("disable", [])]
    chosen_ids: list[str] | None = None
    if isinstance(payload.get("actions"), list):
        chosen_ids = [str(a) for a in payload["actions"]]
        plan_ids = {a["id"] for a in (s.plan or [])}
        enable = sorted(set(chosen_ids) & plan_ids)
        disable = sorted(plan_ids - set(chosen_ids))
    use_assistant = bool(payload.get("useAssistant", True))
    run_model = bool(payload.get("runModel", True))
    check_numbers = bool(payload.get("checkNumbers", True))
    max_iterations = int(payload.get("maxIterations", 4))
    choices = _mind_choices(s)
    chosen = [c for c in choices if (c["id"] in chosen_ids if chosen_ids is not None else c["default_on"])]
    estimate = _mind_run_estimate(s, use_assistant, run_model, check_numbers, chosen)
    with _MIND_LOOP_LOCK:
        busy = _loop_running()
        if busy:
            raise HTTPException(status_code=409, detail=f"a Mind loop is already running (session {busy})")
        if MIND_SIGN_IN["state"] in ("waiting", "checking") and not bool(payload.get("skipMind", False)):
            raise HTTPException(status_code=409, detail="Finish signing in to Mind first (the Edge window), then start Run in Mind")
        work = s.work_dir / "mind_loop" / time.strftime("%Y%m%d_%H%M%S")
        state: dict[str, Any] = {
            "state": "running", "events": [], "report": None, "started": time.time(), "work_dir": str(work), "error": None,
            "stage": None, "stage_started": time.time(), "iteration": 0, "max_iterations": max_iterations,
            "estimate": estimate, "finished": None, "version": _version_public(s), "actions": chosen_ids,
            "inputs": _mind_input_files(inputs_dir),
        }
        MIND_LOOPS[session_id] = state
    cfg = LoopConfig(
        source=s.current_path,
        work_dir=work,
        max_iterations=max_iterations,
        use_assistant=use_assistant,
        run_model=run_model,
        check_numbers=check_numbers,
        enable=enable,
        disable=disable,
        delete_projects=bool(payload.get("deleteProjects", True)),
        skip_mind=bool(payload.get("skipMind", False)),
        **({"inputs_dir": Path(inputs_dir)} if inputs_dir else {}),
    )

    def progress(e: dict[str, Any]) -> None:
        now = time.time()
        if e.get("event") == "iteration" and e.get("n"):
            state["iteration"] = int(e["n"])
        stage = e.get("stage")
        if stage and stage != state.get("stage"):
            state["stage"], state["stage_started"] = stage, now
        state["events"].append({"t": now, **e})

    def worker() -> None:
        try:
            excel_worker.release(s.id)  # one Excel per session
            state["report"] = run_loop(cfg, progress)
            state["state"] = "done"
        except Exception as exc:  # the thread must never die silently
            state["error"] = str(exc)[:300]
            state["state"] = "error"
            state.update(_problem_fields(_report_problem(exc, "Run in Mind", s)))
            progress({"event": "done", "verdict": "error", "message": f"loop crashed: {exc}"[:300]})
        finally:
            state["finished"] = time.time()

    threading.Thread(target=worker, name=f"mind-loop-{session_id}", daemon=True).start()
    return {"state": "running", "work_dir": str(work), "version": state["version"], "estimate": estimate}


# --- signing in to Mind (the browser profile the loop uses) ------------------------------
MIND_SIGN_IN: dict[str, Any] = {"state": "idle", "message": None, "started": None, "finished": None, "stop": False, "checked": None}


def _sign_in_public() -> dict[str, Any]:
    from app import mind_client as mc

    sess = mc.load_session()
    st = MIND_SIGN_IN
    return {"state": st["state"], "message": st["message"], "last_login": sess.get("last_login"), "checked": st.get("checked"),
            "elapsed_s": round((st.get("finished") or time.time()) - st["started"]) if st.get("started") else None}


@app.get("/api/mind/sign-in")
def mind_sign_in_status(check: bool = False) -> dict[str, Any]:
    """The Mind sign-in: state (idle | waiting = the Edge window is open | done | failed), its last
    message, when the last sign-in was saved. `check=true` also loads Mind headless with the saved
    sign-in (10-40 s) and says whether it still works (`checked`: {signed_in, at})."""
    if check:
        with _MIND_LOOP_LOCK:
            busy = _loop_running()
            if busy or MIND_SIGN_IN["state"] == "waiting":
                raise HTTPException(status_code=409, detail="Mind is in use (Run in Mind or the sign-in window): check again when it has finished")
            MIND_SIGN_IN["state_before_check"] = MIND_SIGN_IN["state"]
            MIND_SIGN_IN["state"] = "checking"
        try:
            from app import mind_client as mc

            res = mc.is_authenticated()
            MIND_SIGN_IN["checked"] = {"signed_in": bool(res.get("authenticated")), "at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}
        except Exception as exc:  # no Edge, no Playwright driver...
            MIND_SIGN_IN["checked"] = {"signed_in": False, "at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()), "error": f"{type(exc).__name__}: {exc}"[:200]}
        finally:
            MIND_SIGN_IN["state"] = MIND_SIGN_IN.pop("state_before_check", "idle")
    return _sign_in_public()


@app.post("/api/mind/sign-in")
def mind_sign_in_start(payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Open Edge (visible) on Mind with the app's profile and wait in the background for the user to
    sign in (app/mind_client.sign_in). Poll GET. `{"stop": true}` closes a waiting window."""
    if payload.get("stop"):
        MIND_SIGN_IN["stop"] = True
        return _sign_in_public()
    with _MIND_LOOP_LOCK:
        busy = _loop_running()
        if busy:
            raise HTTPException(status_code=409, detail="Run in Mind is using Mind right now: sign in when it has finished")
        if MIND_SIGN_IN["state"] in ("waiting", "checking"):
            raise HTTPException(status_code=409, detail="the sign-in window is already open")
        MIND_SIGN_IN.update(state="waiting", message="Opening Edge on Mind...", started=time.time(), finished=None, stop=False)

    def tell(message: str) -> None:
        MIND_SIGN_IN["message"] = message

    def worker() -> None:
        from app import mind_client as mc

        try:
            res = mc.sign_in(timeout_s=int(payload.get("timeout_s", 600)), progress=tell, should_stop=lambda: bool(MIND_SIGN_IN["stop"]))
            if res.get("ok"):
                MIND_SIGN_IN.update(state="done", message="Signed in to Mind.", checked={"signed_in": True, "at": res.get("last_login")})
            else:
                MIND_SIGN_IN.update(state="failed", message=f"Not signed in: {res.get('reason')}.")
        except Exception as exc:  # no Edge, the profile in use...
            MIND_SIGN_IN.update(state="failed", message=f"The sign-in window could not open: {type(exc).__name__}: {exc}"[:300])
        finally:
            MIND_SIGN_IN["finished"] = time.time()

    threading.Thread(target=worker, name="mind-sign-in", daemon=True).start()
    return _sign_in_public()


def _mind_loop_timing(st: dict[str, Any]) -> dict[str, Any]:
    """Elapsed time, the stage in words, the stages of this iteration and a conservative time left
    (app/estimate.py estimate_mind_run, mind_run_remaining)."""
    now = st.get("finished") or time.time()
    est = st.get("estimate") or {}
    n = max(1, int(st.get("iteration") or 1))
    stages = [x for x in (est.get("stages") or []) if n == 1 or x["id"] != "names"]  # the assistant names grids once
    stage = st.get("stage")
    stage_elapsed = max(0.0, now - float(st.get("stage_started") or now))
    running = st.get("state") == "running"
    left = mind_run_remaining(stages, stage, stage_elapsed) if running else 0.0
    elapsed = max(0.0, now - float(st["started"]))
    budget = float(est.get("first_s") or 0) + (n - 1) * float(est.get("retry_s") or 0)
    ids = [x["id"] for x in stages]
    cur = ids.index(stage) if stage in ids else -1
    events = st.get("events") or []
    return {
        "elapsed_s": round(elapsed),
        "remaining_s": round(left),
        "estimate_s": est.get("first_s"),
        "retry_s": est.get("retry_s"),
        "over_estimate": bool(running and budget and elapsed > budget),
        "stage": stage,
        "stage_title": MIND_STAGE_TITLES.get(stage or "", "Starting"),
        "stage_elapsed_s": round(stage_elapsed),
        "iteration": n,
        "max_iterations": st.get("max_iterations"),
        "stages": [{"id": x["id"], "title": x["title"], "seconds": x["seconds"], "state": "done" if (i < cur or not running) else "now" if i == cur else "todo"} for i, x in enumerate(stages)],
        "last_news_s": round(now - events[-1]["t"]) if events else round(elapsed),
    }


@app.get("/api/sessions/{session_id}/mind-loop")
def mind_loop_status(session_id: str, after: int = 0) -> dict[str, Any]:
    """Progress of the session's loop: events after index `after`, the report so
    far (written after every iteration), the final verdict, the version it
    uploads and the timing (elapsed, stage, a conservative time left)."""
    _session(session_id)
    st = MIND_LOOPS.get(session_id)
    if not st:
        return {"state": "idle", "events": [], "report": None, "next": 0}
    report = st.get("report")
    if report is None:
        candidate = Path(st["work_dir"]) / "loop_report.json"
        if candidate.is_file():
            try:
                report = json.loads(candidate.read_text(encoding="utf-8"))
            except Exception:
                report = None
    events = st["events"][after:]
    return {"state": st["state"], "events": events, "next": after + len(events), "report": report, "error": st.get("error"), "incident": st.get("incident"), "problem": st.get("problem"), "started": st["started"], "work_dir": st["work_dir"], "summary": summarize(report) if report else None,
            "version": st.get("version"), "actions": st.get("actions"), "timing": _mind_loop_timing(st)}


def _chat_turn(s: Session, payload: dict[str, Any]) -> tuple[list[dict[str, str]], str, str | None]:
    """(history, question, extra_context) for one chat turn; raises the
    409/422 HTTP errors before any model call."""
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    question = str(payload.get("question") or "").strip()
    if not question:
        raise HTTPException(status_code=422, detail="question is required")
    history = [{"role": m.get("role", "user"), "content": str(m.get("content", ""))} for m in payload.get("history", []) if isinstance(m, dict) and m.get("content")]
    focus = payload.get("focus")
    extra = recalculation_context(s.recalc) if s.recalc else None
    if isinstance(focus, dict) and (focus.get("rule_id") or focus.get("error")):
        # A focused mini chat: naming the rule id and the cell makes
        # app.chat_context.retrieve() pull the finding's observed data and the
        # cell contents into the turn; a recalculation error carries the error
        # value and the formula Excel evaluated.
        where = f" at {focus['sheet']}!{focus['cell']}" if focus.get("sheet") and focus.get("cell") else ""
        cells = [c for c in (focus.get("cells") or []) if isinstance(c, dict) and c.get("sheet") and c.get("cell")]
        if focus.get("kind") == "recalc-group" or len(cells) > 1:
            listed = "; ".join(
                f"{c['sheet']}!{c['cell']} ({c.get('error', '')}) {str(c.get('formula', ''))[:120]}" + (f" [array {c['array']}]" if c.get("array") else "")
                for c in cells[:40]
            )
            arrays = sorted({f"{c['sheet']}!{c['array']}" for c in cells if c.get("array")})
            listed_keys = {f"{c['sheet']}!{c['cell']}" for c in cells}
            bits = []
            for a in arrays:
                sh, ref = a.rsplit("!", 1)
                members = _array_cells(ref)
                quiet = [c for c in members if f"{sh}!{c}" not in listed_keys]
                formula_here = next((str(c.get("formula") or "") for c in cells if c.get("array") == ref and c.get("sheet") == sh), "")
                facts = array_size_hint(ref, formula_here)
                bits.append(
                    f"{a} is ONE array formula (Ctrl+Shift+Enter) over {ref}; all {len(members)} cells share the same formula"
                    + (f", and {', '.join(quiet[:12])} belong to it although they are not in error" if quiet else "")
                    + (f". Facts: {facts}" if facts else "")
                    + "."
                )
            array_note = (
                " IMPORTANT: " + " ".join(bits) + " Excel cannot change part of an array, so the fix must cover ALL cells of the array: "
                "either one set_array_formula on the full range, or one operation for EVERY cell of it (including the cells that are not in error). "
                "Re-entering the same formula changes nothing: the new content must remove the cause -- e.g. an array entered over more cells than its "
                "source range has values must be shrunk (set_array_formula over the cells that have values, clear_cell for the rest) or replaced by "
                "per-cell formulas."
                if arrays else ""
            )
            question = (
                f"[About {len(cells)} recalculation errors sharing one root cause: {focus.get('cause') or focus.get('error', '')}. Cells: {listed}.{array_note} "
                "If you propose a fix, propose operations that fix EVERY listed cell (one operation per cell, or one range operation when the "
                "same formula applies to a contiguous range) and say explicitly that all of them can be fixed together.] "
                f"{question}"
            )
        elif focus.get("kind") == "recalc" or focus.get("error"):
            formula = f"; formula: {str(focus.get('formula'))[:300]}" if focus.get("formula") else ""
            siblings = ""
            for g in (s.recalc or {}).get("groups", []):
                members = [c for c in g["cells"] if not (c["sheet"] == focus.get("sheet") and c["cell"] == focus.get("cell"))]
                if len(members) < len(g["cells"]) and members:
                    siblings = (
                        f" This cell shares its root cause ({g['cause']}) with {len(members)} other cell(s): "
                        + ", ".join(f"{c['sheet']}!{c['cell']}" for c in members[:20])
                        + "; tell the user they can all be fixed together and, if asked to fix, include every one of them."
                    )
                    break
            own_array = focus.get("array")
            for g in (s.recalc or {}).get("groups", []) if not own_array else []:
                own_array = next((c.get("array") for c in g["cells"] if c["sheet"] == focus.get("sheet") and c["cell"] == focus.get("cell") and c.get("array")), None)
                if own_array:
                    break
            others = [c for c in _array_cells(own_array) if c != str(focus.get("cell", "")).upper()] if own_array else []
            own_facts = array_size_hint(own_array, str(focus.get("formula") or "")) if own_array else ""
            array_note = (
                f" This cell is part of the array formula {focus.get('sheet')}!{own_array} (Ctrl+Shift+Enter) together with {', '.join(others[:12])}"
                f"{' ...' if len(others) > 12 else ''} -- all of them share this formula. Excel cannot change part of an array, so a fix must cover every "
                f"cell of {own_array} (one set_array_formula on the range, or one operation per cell, including cells that are not in error); "
                "re-entering the same formula changes nothing -- the new content must remove the cause (shrink an array that is larger than its "
                "source values, or replace it by per-cell formulas)." + (f" Facts: {own_facts}." if own_facts else "")
                if own_array else ""
            )
            question = f"[About the recalculation error {focus.get('error', '')}{where}{formula}.{siblings}{array_note}] {question}"
        else:
            # A replacement is judged by a full Excel recalculation afterwards -- say so up front
            question = (
                f"[About finding {focus['rule_id']}{where}. If you propose a replacement formula, it must behave like the original on "
                "blank cells, text and errors (e.g. wrap with IF(x=\"\",0,...) when the original tolerated blanks); every new error "
                "after recalculation counts against the fix.] " + question
            )
    return history, question, extra


def _chat_payload(reply: dict[str, Any]) -> dict[str, Any]:
    proposal = None
    if reply.get("proposal") or reply.get("proposal_errors"):
        proposal = {"summary": reply.get("proposal_summary"), "operations": [_public_op(o) for o in reply.get("proposal", [])], "errors": reply.get("proposal_errors", [])}
    return {
        "text": reply.get("text") or f"(no answer: {reply.get('message')})",
        "proposal": proposal,
        "provenance": {"context_chars": reply.get("context_chars", 0), "detail_chars": reply.get("detail_chars", 0), "lookups": reply.get("lookups", 0)},
    }


@app.post("/api/sessions/{session_id}/chat")
def chat(session_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    s = _session(session_id)
    history, question, extra = _chat_turn(s, payload)
    reply = answer_question(history, question, s.result["workbook_analysis"], s.result["validation_report"], s.plan, extra_context=extra)
    return _chat_payload(reply)


@app.post("/api/sessions/{session_id}/chat/stream")
def chat_stream(session_id: str, payload: dict[str, Any] = Body(...)) -> StreamingResponse:
    """The same turn as /chat, streamed as NDJSON (one JSON object per line):
    {"type": "phase", "phase": "answer"|"lookup"|"repair", "n"} before every
    model call, {"type": "delta", "text"} for each chunk of the reply, then
    {"type": "done", <the /chat payload>} or {"type": "error", "detail"}."""
    s = _session(session_id)
    history, question, extra = _chat_turn(s, payload)
    events: queue.Queue[dict[str, Any] | None] = queue.Queue()

    def work() -> None:
        try:
            reply = answer_question(history, question, s.result["workbook_analysis"], s.result["validation_report"], s.plan, extra_context=extra, on_event=events.put)
            events.put({"type": "done", **_chat_payload(reply)})
        except Exception as exc:  # the stream has started: report in-band
            events.put({"type": "error", "detail": f"{type(exc).__name__}: {exc}"})
        finally:
            events.put(None)

    threading.Thread(target=work, name=f"chat-{session_id}", daemon=True).start()

    def lines():
        while True:
            ev = events.get()
            if ev is None:
                return
            yield json.dumps(ev, ensure_ascii=False, default=str) + "\n"

    return StreamingResponse(lines(), media_type="application/x-ndjson", headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


def _upload_path(s: Session) -> Path | None:
    """The original upload as Excel-readable file: the first .xlsx / .xlsm version."""
    return next((s.paths[v["id"]] for v in s.versions if s.paths[v["id"]].suffix.lower() in (".xlsx", ".xlsm")), None)


def _imported_version(s: Session, version_id: str | None = None) -> dict[str, Any] | None:
    """The newest version of the line of `version_id` (default: the current one) that the outside
    files were brought into (source "import"); None when the line has none."""
    return next((v for v in reversed(_lineage(s, version_id or s.current_version_id)) if v.get("source") == "import"), None)


def _original_path(s: Session) -> Path | None:
    """The workbook the current version's numbers and errors are compared with: the version the
    outside files were brought into when the current line has one (new outside data changes values
    on purpose), else the original upload."""
    imported = _imported_version(s)
    return s.paths[imported["id"]] if imported is not None else _upload_path(s)


def _original_baseline(s: Session, original: Path) -> dict[str, Any] | None:
    """The original (`_original_path`) after the same full Excel recalculation: its error
    cells (and its formulas that already contain #REF!). Computed once per baseline
    file on a fresh copy; None when Excel could not recalculate it, in which
    case the caller falls back to the original's cached values."""
    key = str(original)
    base = s.baselines.get(key)
    if base is None:
        copy_path, _ = make_immutable_copy(original, s.work_dir / "recalc_original")
        with perf_stage("recalc_original_baseline", file=original.name):
            res = recalculate(copy_path, save=False)  # a disposable copy -- nothing reads it again
        ran = bool(res.get("ran", res.get("status") != "NOT_SUPPORTED"))
        known_refs = s.ref_formulas.get(key)
        if known_refs is None and Path(original) == _upload_path(s):
            known_refs = s.original_ref_formulas
        base = {
            "ran": ran,
            "errors": list(res.get("formula_errors") or []) + list(res.get("addin_gap_errors") or []) if ran else [],
            "ref_formulas": known_refs if known_refs is not None else formulas_with_broken_refs(original),
            "message": res.get("message"),
            "path": key,
        }
        s.baselines[key] = base
    s.original_baseline = base
    return base if base.get("ran") else None


@app.post("/api/sessions/{session_id}/recalculate")
def recalc(session_id: str) -> dict[str, Any]:
    s = _session(session_id)
    _refuse_while_fixing(session_id)
    if session_id in RECALCULATING:
        raise HTTPException(status_code=409, detail="a recalculation is already running for this workbook")
    RECALCULATING.add(session_id)
    try:
        with perf_stage("recalc_endpoint", file=s.current_path.name):
            # After a Mind loop the session's Excel worker recalculated this version while it
            # was being analysed; otherwise (or when that did not happen) as before, in its own Excel
            ew = excel_worker.WORKERS.get(s.id)
            ahead = ew.take_recalc(s.current_path) if ew is not None and not ew._stopped else None
            excel_worker.release(s.id)  # one Excel per session
            if ahead is not None:
                copy_path, res = ahead
                perf_record("recalc_ahead_used", 0.0, file=s.current_path.name)
            else:
                copy_path, _ = make_immutable_copy(s.current_path, s.work_dir / "recalc")
                res = recalculate(copy_path)
            with perf_stage("recalc_out", file=s.current_path.name):
                return _recalc_out(s, res, copy_path)
    finally:
        RECALCULATING.discard(session_id)


def _recalc_out(s: Session, res: dict[str, Any], copy_path: Path, close_round: bool = True) -> dict[str, Any]:
    """What a recalculation of the current version means for the session: the
    errors against the original's, the root-cause groups, READY-001, the verdict.
    `res` has the shape of app.recalc.recalculate -- the automatic fixer
    ends with a full recalculation of its own and hands it in the same way."""
    out = {
        "status": res["status"],
        "message": res["message"],
        "version_id": s.current_version_id,
        "ran": bool(res.get("ran", res["status"] != "NOT_SUPPORTED" and not str(res.get("message", "")).startswith("Excel COM recalculation failed"))),
        "formula_errors": [
            {"sheet": e.get("sheet", ""), "cell": e.get("cell", ""), "error": e.get("error", ""), "formula": e.get("formula", ""), **({"array": e["array"]} if e.get("array") else {})}
            for e in res.get("formula_errors", [])
        ],
        "addin_gap_errors": [{"sheet": e.get("sheet", ""), "cell": e.get("cell", ""), "formula": e.get("formula", "")} for e in res.get("addin_gap_errors", [])],
    }
    out["groups"] = group_errors(out["formula_errors"], out["addin_gap_errors"])
    # Errors the original already had are the model's own, not the preparation's.
    # The baseline is the ORIGINAL after the same full recalculation (once per
    # session), never its cached values -- those can be stale.
    original = _original_path(s)
    if original is not None and out["formula_errors"]:
        base = _original_baseline(s, original)
        # the original's cells where they are now (title rows inserted, sheets renamed), as the numbers check sees them
        found = _original_and_moves(s)
        forward = found[1].forward if found is not None and found[0] == original else None
        if base is not None:
            cls = classify_errors(out["formula_errors"], base["errors"], base["ref_formulas"], forward)
        else:
            cls = classify_against_original(original, out["formula_errors"], forward)
            cls["message"] = (cls.get("message") or "") + " -- compared with the original's cached values, its own recalculation did not run"
        out["formula_errors"] = cls["errors"]
        out["preexisting_errors"], out["new_errors"] = cls["preexisting"], cls["new"]
        out["compared_with_original"] = cls["compared"]
        if cls.get("message"):
            out["message"] = f"{out['message']} ({cls['message'].strip(' -')})"
    else:
        out["preexisting_errors"], out["new_errors"] = 0, len(out["formula_errors"])
        out["compared_with_original"] = original is not None
    # A recalculation after fixes (v2.1, v2.2, ...) closes the round: the
    # same file becomes the next major version (v3), and the recalculation is its own
    current = next((v for v in s.versions if v["id"] == s.current_version_id), None)
    if close_round and out["ran"] and current is not None and current.get("minor"):
        short = current["label"].split(" — ")[0]
        version = _add_version(s, s.current_path, "recalculate", f"recalculated {short}", [], current.get("verified_opens_in_excel"), current["sha256"])
        out["version"] = version
        out["version_id"] = version["id"]
    s.recalc = out
    s.recalc_version_id = s.current_version_id
    _sync_recalc_finding(s)
    if s.result:
        out["report"] = {k: v for k, v in s.result["validation_report"].items() if k != "_version_id"}
        out["plan"] = s.plan
    s.recalc_path = copy_path
    out["readiness"] = s.readiness()  # a clean recalculation is what turns the verdict green
    return out


# --- did the preparation change what the model computes? ---------------------------------
def _lineage(s: Session, version_id: str) -> list[dict[str, Any]]:
    """The versions that lead to `version_id`, oldest first (a version restored
    from History branches: the list order alone is not the path)."""
    by_id = {v["id"]: v for v in s.versions}
    path: list[dict[str, Any]] = []
    seen: set[str] = set()
    vid: str | None = version_id
    while vid and vid in by_id and vid not in seen:
        seen.add(vid)
        path.append(by_id[vid])
        vid = by_id[vid].get("parent")
    return path[::-1]


def _original_environment(s: Session) -> dict[str, Any] | None:
    """What each CELL / INFO formula shows in the original upload, keyed "Sheet!A1" by where that
    cell sits in the current version (rows / columns inserted, sheets renamed since): the values
    FRM-012 freezes (app/prep.py). None until the original has been scanned."""
    if s.original_environment is None:
        return None
    found = _upload_and_moves(s)  # always the upload's values: the user's own Excel saved them
    moves = found[1] if found else Moves()
    out: dict[str, Any] = {}
    for (sheet, cell), value in s.original_environment.items():
        m = re.match(r"([A-Z]{1,3})(\d+)$", cell)
        if m:
            sheet_now, row, col = moves.forward(sheet, int(m.group(2)), col_to_num(m.group(1)))
            out[f"{sheet_now}!{a1(row, col)}"] = value
    return out


def _line_base_and_moves(s: Session, after_import: bool) -> tuple[Path, Moves] | None:
    """A base version of the current version's lineage and how its cells moved since (rows /
    columns inserted, sheets renamed) -- None when the current version *is* that base. The base is
    the first Excel-readable version (the upload), or with `after_import` the newest version the
    outside files were brought into when the line has one."""
    path = _lineage(s, s.current_version_id)
    start = next((i for i in range(len(path) - 1, -1, -1) if path[i].get("source") == "import"), None) if after_import else None
    if start is None:
        start = next((i for i, v in enumerate(path) if s.paths[v["id"]].suffix.lower() in (".xlsx", ".xlsm")), None)
    if start is None or start == len(path) - 1:
        return None
    return s.paths[path[start]["id"]], Moves(v.get("change_log") or [] for v in path[start + 1:])


def _original_and_moves(s: Session) -> tuple[Path, Moves] | None:
    """The version the current one is compared with -- the numbers check, and the fixer's "a good
    value never changes" gate -- and how its cells moved since: the version the outside files were
    brought into when the line has one (new outside data changes values on purpose), else the
    upload. None when the current version *is* that original."""
    return _line_base_and_moves(s, after_import=True)


def _upload_and_moves(s: Session) -> tuple[Path, Moves] | None:
    """The upload (the first Excel-readable version of the current line) and how its cells moved
    since: where the CELL / INFO values the user's own Excel saved sit now."""
    return _line_base_and_moves(s, after_import=False)


def _numbers_public(res: dict[str, Any], version_id: str) -> dict[str, Any]:
    return {**{k: v for k, v in res.items() if k != "regressions"}, "regressions": len(res.get("regressions") or []), "version_id": version_id}


@app.post("/api/sessions/{session_id}/numbers-check")
def numbers_check(session_id: str) -> dict[str, Any]:
    """Recalculate the original and the current version in Excel and compare every
    formula cell (app/numbers_check.py): did Prep change a computed value?"""
    s = _session(session_id)
    _refuse_while_fixing(session_id)
    found = _original_and_moves(s)
    if found is None:
        imported = _imported_version(s)
        verdict = ("This is the version the outside files were brought into: nothing to compare." if imported is not None and imported["id"] == s.current_version_id
                   else "This is the original workbook: nothing to compare.")
        return {"ran": False, "clean": True, "version_id": s.current_version_id, "verdict": verdict, "counts": {}}
    original, moves = found
    excel_worker.release(s.id)
    return _numbers_public(compare_workbooks(original, s.current_path, moves), s.current_version_id)


# --- bringing the outside files in (app/outside_inputs.py) --------------------------------------
_OUTSIDE_TITLES = {
    "start": "Starting", "stand_ins": "Making stand-in input files with the assistant",
    "import": "Running the workbook's import macro", "scan": "Checking the workbook with the outside data",
    "mind_inputs": "Making the Mind inputs", "numbers": "Checking that no computed value changed", "rescan": "Checking the new version", "done": "Done",
}


class _BringInStopped(Exception):
    """The job stopped for a reason the user can act on: said in plain words, no problem report."""


class _MindInputsRefused(_BringInStopped):
    """The Mind inputs would change a computed value: not kept. Its text says the whole story (which
    version holds the outside data), so it is shown as it is."""


def _bring_in_base(s: Session) -> tuple[dict[str, Any], str | None]:
    """The version the workbook's import macro runs on, and why when it is not the current one.

    The current version -- unless its line already has the outside files brought in (they would come
    in twice: the macro runs again on the version before the first bring-in), or rows, columns or
    sheet names moved since the upload (the macro writes to fixed places: it runs on the newest
    version before the first move). The changes made after that version are not in the new line."""
    line = _lineage(s, s.current_version_id)
    start = next((i for i, v in enumerate(line) if s.paths[v["id"]].suffix.lower() in (".xlsx", ".xlsm")), None)
    if start is None:
        raise HTTPException(status_code=422, detail="The workbook must be .xlsx or .xlsm for its macro to run.")
    base, why = line[start], None
    for v in line[start + 1:]:
        if v.get("source") in ("import", "mind_inputs"):
            why = "again"
            break
        if Moves([v.get("change_log") or []]).batches:
            why = "moved"
            break
        base = v
    if base["id"] == s.current_version_id:
        return base, None
    short = base["label"].split(" — ")[0]
    after = line[next(i for i, v in enumerate(line) if v["id"] == base["id"]) + 1:]
    if why == "again" and all(v.get("source") in ("import", "mind_inputs") for v in after):
        # only the versions of the earlier bring-in follow it (one that changed the files, or one that
        # did not finish): the new line leaves nothing else out
        return base, f"Brought in on {short}: the outside files were already brought in after it."
    reason = ("the outside files were already brought in after it" if why == "again"
              else "rows, columns or sheet names changed after it, and the workbook's macro writes to fixed places")
    return base, f"Brought in on {short}: {reason}. The changes made after {short} are not in this new line: run Do everything again."


def _import_ok(report: dict[str, Any]) -> bool:
    """The import macro's run can be kept: the core's own verdict ("ok": no error, no early stop, every
    picked file opened and written), or -- from a core without it -- saved and not stopped early."""
    if "ok" in report:
        return bool(report["ok"])
    return bool(report.get("saved")) and not report.get("stopped_early")


def _import_stopped(report: dict[str, Any]) -> str:
    """Why the import macro's run is not kept, in plain words (its own message boxes, any language)."""
    said = [str(d.get("message_repaired") or d.get("message") or "").strip() for d in report.get("dialogs") or []]
    said = [x for x in said if x]
    not_opened = [str(e["file"]) for e in report.get("files") or [] if e.get("file") and not e.get("opened")]
    not_written = [str(e["file"]) for e in report.get("files") or [] if e.get("file") and e.get("opened") and not e.get("areas")]
    # a VBA / COM error in plain words, Excel's own in brackets ('Subscript out of range'), else the text as it is
    err = plain_com_error(re.sub(r"\s+", " ", str(report.get("error") or "")).strip())

    def names(xs: list[str]) -> str:
        return ", ".join(xs[:6]) + (" ..." if len(xs) > 6 else "")

    if not report.get("saved") and "longer than" in err:
        text = "The workbook's import macro " + re.sub(r"^the import macro\s+", "", err).rstrip(".") + "."
    elif report.get("stopped_early") or not_opened:
        text = "The workbook's import macro stopped early" + (f': "{said[0][:200]}"' if said else f": {err[:200]}" if err else "") + "."
        if not_opened:
            text += f" Not opened: {names(not_opened)}."
    elif not_written:
        text = f"The workbook's import macro opened {names(not_written)} but wrote nothing from {'it' if len(not_written) == 1 else 'them'}" + (f': "{said[0][:200]}"' if said else f": {err[:200]}" if err else "") + "."
    else:
        text = "The workbook's import macro did not finish" + (f": {err[:200]}" if err else "") + "."
    return text + " Nothing was changed: check the input files, then try again."


def _import_log(report: dict[str, Any]) -> list[dict[str, Any]]:
    """The change log of the "import" version: each area a picked file filled."""
    return [
        {"action_id": "outside_inputs", "rule_id": OUTSIDE_RULE, "op": "bring_in_file", "sheet": a.get("sheet"), "cell": a.get("ref"), "after": e.get("file"),
         "note": f"{a.get('sheet')}!{a.get('ref')} filled from {e.get('file')} by the workbook's own macro"}
        for e in report.get("files") or [] if e.get("opened") for a in e.get("areas") or []
    ]


def _mind_inputs_log(made: dict[str, Any]) -> list[dict[str, Any]]:
    """The change log of the "mind_inputs" version: each area moved to its Mind input grid."""
    return [
        {"action_id": "outside_inputs", "rule_id": OUTSIDE_RULE, "op": "make_mind_input", "sheet": i.get("sheet"), "cell": i.get("ref"), "after": f"#{i.get('name')} /Input",
         "note": f"{i.get('sheet')}!{i.get('ref')} moved to the Mind input {i.get('name')} ({Path(str(i.get('csv') or '')).name}); {i.get('readers', 0)} formula(s) read it there"}
        for i in made.get("inputs") or []
    ]


def _import_public(report: dict[str, Any], version: dict[str, Any] | None) -> dict[str, Any]:
    return {
        "macro": report.get("macro"),
        "seconds": report.get("seconds"),
        "files": [
            {"cell": e.get("cell"), "file": e.get("file"), "expected": e.get("expected"), "exact": bool(e.get("exact")), "opened": bool(e.get("opened")),
             "areas": [{"sheet": a.get("sheet"), "ref": a.get("ref"), "rows": a.get("rows"), "cols": a.get("cols")} for a in e.get("areas") or []],
             "cleared": list(e.get("cleared") or [])}
            for e in report.get("files") or []
        ],
        "kept": list(report.get("kept") or []),
        "dialogs": list(report.get("dialogs") or []),
        "stopped_early": bool(report.get("stopped_early")),
        "error": report.get("error"),
        "ok": _import_ok(report),
        "saved": version["file_name"] if version else None,
    }


def _mind_inputs_public(made: dict[str, Any], version: dict[str, Any] | None) -> dict[str, Any]:
    return {
        "inputs": [
            {"name": i.get("name"), "sheet": i.get("sheet"), "ref": i.get("ref"), "file": i.get("file"), "rows": i.get("rows"), "cols": i.get("cols"),
             "csv": Path(str(i.get("csv") or "")).name, "separator": i.get("separator"), "readers": i.get("readers")}
            for i in made.get("inputs") or []
        ],
        "kept": list(made.get("kept") or []),
        "formulas_changed": made.get("formulas_changed", 0),
        "names_changed": made.get("names_changed", 0),
        "sums_split": made.get("sums_split", 0),
        "saved": version["file_name"] if version else None,
    }


def _register_mind_inputs(s: Session, made: dict[str, Any], run_dir: Path) -> dict[str, Any]:
    """Every CSV / TAB .txt for download under its own name (the name Mind's Input Manager matches:
    never renamed -- the newest bring-in wins), and one zip with all of them at its root."""
    names = []
    for i in made.get("inputs") or []:
        p = Path(i["csv"])
        s.files[p.name] = p
        names.append(p.name)
    zip_path = run_dir / "Mind inputs.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for i in made.get("inputs") or []:
            zf.write(i["csv"], arcname=Path(i["csv"]).name)
    s.files[zip_path.name] = zip_path
    return {"zip": zip_path.name, "files": names}


def _changed_values(numbers: dict[str, Any] | None) -> int:
    """The computed values a numbers check found changed: now an error, another value, or no longer a formula."""
    c = (numbers or {}).get("counts") or {}
    return int(c.get("good_to_error", 0)) + int(c.get("good_to_other", 0)) + int(c.get("missing_good", 0))


def _mind_inputs_refused(compared: dict[str, Any], v_import: dict[str, Any]) -> str:
    """Why the Mind inputs are not kept (they would change computed values), in plain words, with the
    first value that would change -- in the version that stays, where the user can look."""
    n = _changed_values(compared)
    what = f"{n:,} computed value{'' if n == 1 else 's'}" if n else "computed values"
    short = v_import["label"].split(" — ")[0]
    samples = compared.get("samples") or {}
    first = next((x for kind in ("good_to_error", "good_to_other", "missing") for x in samples.get(kind) or []), None)
    example = ""
    if first and first.get("original"):
        example = (f" For example {first['original']} would show {first['is']} instead of {first.get('was')}." if "is" in first
                   else f" For example {first['original']} ({first.get('was')}) would no longer be a formula.")
    return f"Making the Mind inputs would change {what}; the outside data stays inside the model ({short}); nothing else was changed.{example}"


def _bring_in_summary(report: dict[str, Any], made: dict[str, Any] | None, numbers: dict[str, Any] | None, note: str | None) -> str:
    """One plain line: '6 outside files brought in; 6 Mind inputs made (IN_BBA_01, ...); no computed value changed.'"""
    opened = sum(1 for e in report.get("files") or [] if e.get("opened"))
    parts = [f"{opened} outside file{'' if opened == 1 else 's'} brought in"]
    names = [str(i.get("name")) for i in (made or {}).get("inputs") or []]
    parts.append(f"{len(names)} Mind input{'' if len(names) == 1 else 's'} made ({', '.join(names[:8])}{', ...' if len(names) > 8 else ''})" if names else "no Mind input made")
    if numbers and numbers.get("ran") and numbers.get("clean"):
        parts.append("no computed value changed")
    elif numbers and numbers.get("ran"):
        changed = _changed_values(numbers)
        parts.append(f"{changed:,} computed value{'' if changed == 1 else 's'} changed: see Compare the numbers")
    elif numbers:  # the comparison was asked for and could not run (Excel failed it): said plainly
        why = str(numbers.get("verdict") or "").strip().rstrip(".")
        parts.append("the computed values were NOT checked" + (f" ({why})" if why else ""))
    else:
        parts.append("the computed values are not checked yet")
    return "; ".join(parts) + "." + (f" {note}" if note else "")


_STAND_IN_NOTE = ("Made with stand-in input files from the assistant, built from the numbers the model already holds: "
                  "best results come with the actual input files (add them, then bring them in again).")


def _make_stand_ins(s: Session, base_path: Path, macro: str, stage, say) -> str | None:
    """The missing files of the import table, made by the assistant from the model's own numbers and
    proven with the model's own macro before anything is kept (app/stand_in_inputs.py); then kept as
    picked files marked stand_in (a real file of the same name replaces one). None: nothing missing."""
    rows = _match_rows(s)
    missing = [r for r in rows if not r.get("file")]
    if not missing:
        return None
    ask = stand_in_inputs.assistant_ask()
    if ask is None:
        raise _BringInStopped("The assistant is not set up, so it cannot make the missing input files: "
                              "add the actual files, or paste an AI key under Advanced › Assistant.")
    need = [f"{r['sheet']}!{r['cell']}" for r in missing]
    real: dict[str, Path] = {}
    for r in rows:
        p = _input_path(s, r.get("file"))
        if p is not None and p.is_file():
            real[f"{r['sheet']}!{r['cell']}"] = p
    stage("stand_ins", f"Making {len(need)} stand-in input file{'' if len(need) == 1 else 's'} with the assistant")
    work = s.work_dir / "stand_ins" / f"{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:4]}"
    excel_worker.release(s.id)  # the stand-in steps open the model in Excels of their own
    with perf_stage("outside_inputs_stand_ins", file=base_path.name, files=len(need)):
        made = stand_in_inputs.make_stand_ins(base_path, rows, need, macro, model_map_module.read_vba(base_path), work, ask,
                                              real_files=real, messages=((s.outside_inputs or {}).get("vba") or {}).get("messages") or [],
                                              progress=say)
    s.outside_run["stand_ins"] = {"ok": bool(made.get("ok")), "explain": made.get("explain"), "files": sorted(p.name for p in (made.get("files") or {}).values()),
                                  "left_out": list(made.get("left_out") or []),
                                  "attempts": [{"n": a.get("n"), "problem": a.get("problem"), "details": a.get("details")} for a in made.get("attempts") or []]}
    if not made.get("ok"):
        n = len(made.get("attempts") or [])
        raise _BringInStopped(f"The assistant could not make input files the workbook's macro accepts ({n} attempt{'' if n == 1 else 's'}): "
                              f"{made.get('problem')}. Nothing was changed: add the actual input files.")
    if not made["files"] and not real:
        raise _BringInStopped("The model holds no data for its input files, so the assistant had nothing to make stand-ins from. "
                              "Nothing was changed: add the actual input files.")
    for path in made["files"].values():
        dest = _input_folder(s) / Path(path).name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, dest)
        _keep_input(s, dest)["stand_in"] = True
    _rematch(s)
    left = [next((str(r.get("expected") or (r.get("keys") or [None])[0] or k) for r in missing if f"{r['sheet']}!{r['cell']}" == k), k)
            for k in made.get("left_out") or []]
    if left:
        return (_STAND_IN_NOTE + f" No stand-in for {', '.join(left[:6])}{' ...' if len(left) > 6 else ''}: the model holds no data for "
                f"{'it' if len(left) == 1 else 'them'} and zeros would change a result, so {'its block stays' if len(left) == 1 else 'their blocks stay'} as the model holds {'it' if len(left) == 1 else 'them'}.")
    return _STAND_IN_NOTE


def _bring_in(s: Session, base: dict[str, Any], note: str | None, check_numbers: bool, stage, say, stand_ins: bool = False) -> dict[str, Any]:
    """The job: the workbook's own import macro on a copy of `base` with the matched files -> the
    "import" version -> its scan (every sheet: a formula that reads an import area may be anywhere)
    -> each area a Mind input, on a copy -> every computed value of that copy compared with the
    import version's (`check_numbers`) -> only when none changed (or the comparison could not run:
    said plainly, result["numbers"]["ran"] False) the "mind_inputs" version, its CSVs and zip for
    download, and its scan. Mind inputs that would change a value are not kept (_MindInputsRefused):
    the import version stays current and its line is not brought in. A bring-in counts once it got
    to the end (s.brought_in_files, with the fingerprint of the files it read)."""
    from app.inventory import sha256_of

    oi = s.outside_inputs or {}
    macro = oi["vba"]["import_macro"]
    made_note = _make_stand_ins(s, s.paths[base["id"]], macro, stage, say) if stand_ins else None
    rows = (oi.get("match") or {}).get("rows") or []
    files: dict[str, Path] = {}
    for r in rows:
        p = _input_path(s, r.get("file"))
        if p is not None and p.is_file():
            files[f"{r['sheet']}!{r['cell']}"] = p
    fingerprint = _inputs_fingerprint(s, rows)  # the files this bring-in reads (added files are refused meanwhile)
    base_path = s.paths[base["id"]]
    run_dir = s.work_dir / "outside_inputs" / f"{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:4]}"
    imported = run_dir / "import" / base_path.name
    excel_worker.release(s.id)  # one Excel per session: the macro runs in an Excel of its own
    stage("import", f"Running the workbook's import macro {macro} with {len(files)} file{'' if len(files) == 1 else 's'}")
    # a stand-in is not a real import: its row of the import table keeps the model's own status
    restore = {f"{r['sheet']}!{r['cell']}" for r in rows if r.get("file") in _stand_in_names(s)}
    with perf_stage("outside_inputs_import", file=base_path.name, files=len(files)):
        report = outside_inputs.run_import(base_path, rows, files, macro, imported, run_dir, progress=say, messages=oi["vba"].get("messages") or [],
                                           **({"restore_rows": restore} if restore else {}))
    if not _import_ok(report) or not imported.is_file():
        raise _BringInStopped(_import_stopped(report))
    s.current_version_id = base["id"]  # the new line starts from the version the macro ran on
    v_import = _add_version(s, imported, "import", "Outside files brought in", _import_log(report), None, sha256_of(imported))
    s.outside_run.setdefault("versions", []).append(v_import["id"])
    result: dict[str, Any] = {
        "import": _import_public(report, v_import), "mindInputs": None, "versions": [v_import["id"]], "download": None, "numbers": None,
        "base_version_id": base["id"], "notes": [note] if note else [],
    }
    if made_note:
        result["notes"].append(made_note)
        result["standIns"] = s.outside_run.get("stand_ins")
    not_picked = [str(e.get("expected") or e.get("cell")) for e in report.get("files") or [] if not e.get("file")]
    if not_picked:
        result["notes"].append(f"Not picked, so their data stays in the workbook as it was: {', '.join(not_picked[:6])}" + (" ..." if len(not_picked) > 6 else "") + ".")
    if report.get("error"):
        result["notes"].append(f"The macro ended with a message from Excel: {str(report['error'])[:200]}")
    stage("scan", "Checking the workbook with the outside data")
    ignored, s.ignore_sheets = s.ignore_sheets, []
    try:
        _run_scan(s, imported)
    finally:
        s.ignore_sheets = ignored
    wb0 = s.result["workbook_analysis"]["workbooks"][0]
    sheet_names = [sh["name"] for sh in wb0.get("sheets", [])] + [sh["name"] for sh in wb0.get("ignored_sheets", [])]
    plan = outside_inputs.plan_mind_inputs(report, sheet_names)
    made: dict[str, Any] | None = None
    if not plan:
        result["notes"].append("The macro wrote nothing the app could see: no Mind input made.")
    else:
        stage("mind_inputs", f"Making {len(plan)} Mind input{'' if len(plan) == 1 else 's'}")
        csv_dir = run_dir / "csv"  # the CSVs alone: Run in Mind uploads the whole folder
        made_path = run_dir / "mind_inputs" / base_path.name
        with perf_stage("outside_inputs_mind_inputs", file=base_path.name, inputs=len(plan)):
            made = outside_inputs.apply_mind_inputs(imported, plan, wb0.get("formulas") or [], made_path, csv_dir, progress=say)
        if not made.get("inputs") or not made.get("saved") or not made_path.is_file():
            why = (made.get("kept") or [{}])[0].get("why")
            result["notes"].append("No Mind input could be made" + (f": {why}" if why else "") + ". The outside data stays inside the model.")
            result["mindInputs"] = _mind_inputs_public(made, None)
        else:
            log = _mind_inputs_log(made)
            compared: dict[str, Any] | None = None
            if check_numbers:
                # before anything is kept: Mind inputs that change a computed value never become a
                # version, and their CSVs are never offered for download or to Run in Mind
                stage("numbers", "Comparing every computed value with the version the files were brought into")
                excel_worker.release(s.id)
                compared = compare_workbooks(imported, made_path, Moves([log]), lambda _stage, message, *_a, **_k: say(message))
                if compared.get("ran") and not compared.get("clean"):
                    raise _MindInputsRefused(_mind_inputs_refused(compared, v_import))
            v_made = _add_version(s, made_path, "mind_inputs", "Outside files made Mind inputs", log, None, sha256_of(made_path))
            s.outside_run.setdefault("versions", []).append(v_made["id"])
            result["versions"].append(v_made["id"])
            s.mind_inputs_dirs[v_made["id"]] = str(csv_dir)
            s.mind_inputs_dir = str(csv_dir)
            result["mindInputs"] = _mind_inputs_public(made, v_made)
            result["download"] = _register_mind_inputs(s, made, run_dir)
            if compared is not None:  # clean, or the comparison could not run: "ran" False, said in the summary
                result["numbers"] = _numbers_public(compared, v_made["id"])
            if made.get("kept"):
                result["notes"].append(f"{len(made['kept'])} area(s) stay inside the model (Mind cannot swap them): {made['kept'][0].get('why')}")
            stage("rescan", "Checking the new version")
            _run_scan(s, made_path)
    # got to the end: the line counts as brought in, while the picked files are the ones read here
    s.brought_in_files[result["versions"][-1]] = fingerprint
    with contextlib.suppress(Exception):  # INP-010 of the last scan ran before this record
        _patch_outside_finding(s, s.result)
    result["summary"] = _bring_in_summary(report, made, result["numbers"], note)
    if _stand_in_names(s) & {str(r.get("file")) for r in rows}:
        result["summary"] += " " + _STAND_IN_NOTE
    return result


def _bring_in_worker(s: Session, base: dict[str, Any], note: str | None, check_numbers: bool, stand_ins: bool = False) -> None:
    st = s.outside_run

    def stage(name: str, message: str) -> None:
        st.update(stage=name, title=_OUTSIDE_TITLES.get(name, name), message=message, updated=time.time())

    def say(message: Any) -> None:
        st.update(message=str(message)[:300], updated=time.time())

    def after_versions(text: str) -> str:
        made = [v for v in s.versions if v["id"] in (st.get("versions") or [])]
        if not made:
            return text
        short = made[-1]["label"].split(" — ")[0]
        return f"{short} holds the outside data, but the rest did not finish: {text}"

    # the end is said only once the workbook is free again: a client that sees "done" or "error"
    # and starts the next step at once must not find the lock still held (a few ms otherwise);
    # meanwhile "running" keeps every other job away (_busy_with)
    end: dict[str, Any] = {}
    try:
        result = _bring_in(s, base, note, check_numbers, stage, say, stand_ins)
        end = {"result": result, "state": "done", "stage": "done", "title": "Done", "message": result["summary"]}
    except _MindInputsRefused as exc:  # says which version holds the outside data itself
        end = {"state": "error", "error": str(exc)}
    except _BringInStopped as exc:
        end = {"state": "error", "error": after_versions(str(exc))}
    except HTTPException as exc:  # a scan refused or failed
        end = {"state": "error", "error": after_versions(str(exc.detail))[:400]}
        if exc.status_code >= 500:
            end.update(_problem_fields(_report_problem(exc.__cause__ or exc, "Bring in the outside files", s, stage=str(st.get("stage") or ""))))
    except Exception as exc:  # the thread must never die silently
        end = {"state": "error", "error": after_versions(_exc_text(exc))[:400]}
        with contextlib.suppress(Exception):
            end.update(_problem_fields(_report_problem(exc, "Bring in the outside files", s, stage=str(st.get("stage") or ""))))
    finally:
        s.apply_lock.release()
        st.update(end or {"state": "error", "error": "the job stopped without saying why"}, finished=time.time(), updated=time.time())
        if st.get("state") == "error":
            # a job that stopped half-way leaves its line not brought in: INP-010 says why (its last
            # scan ran while the job was still going)
            with contextlib.suppress(Exception):
                _patch_outside_finding(s, s.result)


@app.post("/api/sessions/{session_id}/outside-inputs/run", status_code=202)
def run_outside_inputs(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Bring the outside files in, in the background (poll GET .../outside-inputs): the workbook's own
    import macro runs once on a copy with the picked files (a new version, "Outside files brought
    in"); then each area it filled becomes a Mind input -- an IN_<sheet> /Input grid the formulas
    read, and its CSV for Mind's Input Manager (a new version, "Outside files made Mind inputs");
    then (`checkNumbers`, default true) every computed value is compared. Refused (409) while
    another job runs on the session."""
    s = _session(session_id)
    oi = s.outside_inputs
    if not oi:
        raise HTTPException(status_code=422, detail="Nothing to bring in: the workbook's macros open no outside file.")
    if not oi["vba"].get("import_macro"):
        raise HTTPException(status_code=422, detail="Which macro brings the outside files in is not clear: they cannot be brought in automatically.")
    if not oi["table"]:
        raise HTTPException(status_code=422, detail="No cell of the workbook names the outside files: they cannot be brought in automatically.")
    # standIns: the assistant makes the missing files from the model's own numbers first (the user has
    # only the workbook, or some of the files); best results come with the actual files
    missing = [r for r in _match_rows(s) if not r.get("file")]
    stand_ins = bool(payload.get("standIns")) and bool(missing)
    if stand_ins and not llm_available():
        raise HTTPException(status_code=409, detail="The assistant is not set up, so it cannot make the missing input files: add the actual files, or paste an AI key under Advanced › Assistant.")
    if not stand_ins and not (oi.get("match") or {}).get("matched"):
        raise HTTPException(status_code=422, detail="None of the picked files goes with the workbook: add the outside files first.")
    if not com_available():
        raise HTTPException(status_code=409, detail="Bringing in the outside files needs Excel on this machine.")
    why = _busy_with(s)
    if why:
        raise HTTPException(status_code=409, detail=f"Wait a moment: {why}.")
    if (MIND_LOOPS.get(session_id) or {}).get("state") == "running":
        raise HTTPException(status_code=409, detail="Wait a moment: Run in Mind is using this workbook.")
    base, note = _bring_in_base(s)
    if not has_vba_project(s.paths[base["id"]]):
        raise HTTPException(status_code=422, detail=f"{base['label'].split(' — ')[0]} has no macros left to run: bring the outside files in on the uploaded workbook.")
    if not s.apply_lock.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="Wait a moment: changes are still being written.")
    now = time.time()
    s.outside_run = {"state": "running", "stage": "start", "title": "Starting", "message": "Starting", "started": now, "updated": now, "finished": None,
                     "result": None, "error": None, "versions": [], "base_version_id": base["id"]}
    try:
        threading.Thread(target=_bring_in_worker, args=(s, base, note, bool(payload.get("checkNumbers", True)), stand_ins), name=f"outside-inputs-{session_id}", daemon=True).start()
    except Exception:
        s.outside_run.update(state="error", error="the job could not start", finished=time.time())
        s.apply_lock.release()
        raise
    return {"running": True, "state": "running", "base_version_id": base["id"], "note": note}


@app.get("/api/sessions/{session_id}/outside-inputs")
def outside_inputs_state(session_id: str) -> dict[str, Any]:
    """The outside files (what the macros open, which picked file goes with each row, whether this
    line has them brought in) and the bring-in job: running, its stage, its result or error."""
    s = _session(session_id)
    st = s.outside_run or {}
    started = st.get("started")
    return {
        "outsideInputs": _outside_public(s),
        "inputs": _inputs_public(s),
        "running": st.get("state") == "running",
        "state": st.get("state") or "idle",
        "stage": st.get("title"),  # in words, for the screen ("Running the workbook's import macro")
        "stage_id": st.get("stage"),  # start | import | scan | mind_inputs | numbers | rescan | done
        "message": st.get("message"),
        "elapsed_s": round((st.get("finished") or time.time()) - started, 1) if started else 0.0,
        "result": st.get("result"),
        "error": st.get("error"),
        "versions": list(st.get("versions") or []),
        "incident": st.get("incident"),
        "problem": st.get("problem"),
        # the stand-in files the assistant made (or tried to): each try's outcome, also when it failed
        "standIns": st.get("stand_ins"),
    }


@app.post("/api/sessions/{session_id}/inputs")
def add_inputs(session_id: str, files: list[UploadFile] = File(...)) -> dict[str, Any]:
    """Add or replace input files (by name) and match them again to the workbook's import table
    (no rescan; INP-010 says the new counts)."""
    s = _session(session_id)
    _refuse_while_bringing_in(s)
    picked = [f for f in files if (f.filename or "").strip()]
    if not picked:
        raise HTTPException(status_code=422, detail="Pick the input files to add.")
    _check_input_names([Path(f.filename or "").name for f in picked])
    main_key = outside_inputs.normalize_name(_main_name(s))
    for f in picked:
        if outside_inputs.normalize_name(Path(f.filename or "").name) != main_key:  # the workbook is not its own input
            _keep_input(s, _save_upload(f, _input_folder(s)))
    _rematch(s)
    _patch_outside_finding(s, s.result)
    return {"inputs": _inputs_public(s), "outsideInputs": _outside_public(s), "readiness": s.readiness()}


# --- fix all automatically ----------------------------------------------------------------
AUTOFIX: dict[str, dict[str, Any]] = {}  # session id -> live state of the current / last run
RECALCULATING: set[str] = set()          # sessions with a manual recalculation in progress


def _refuse_while_fixing(session_id: str) -> None:
    """The fixer ends by adding a version, a recalculation and an analysis to the
    session: a Recalculate or an Apply landing in the middle of that would be
    attached to the wrong version (seen in the first click-through: the verdict
    said "not verified" on a version that had just recalculated clean)."""
    state = AUTOFIX.get(session_id)
    if state and state.get("state") == "running":
        raise HTTPException(status_code=409, detail="the automatic fixer is running on this workbook: wait for it to finish, or stop it")
    _refuse_while_bringing_in(SESSIONS.get(session_id))


def _refuse_while_bringing_in(s: Session | None) -> None:
    """The bring-in job adds its versions and scans at the end of a long Excel run: a change landing
    in the middle would be attached to the wrong version."""
    if s is not None and (s.outside_run or {}).get("state") == "running":
        raise HTTPException(status_code=409, detail="the outside files are being brought in on this workbook: wait for it to finish")

_AUTOFIX_PUBLIC = ("state", "stage", "title", "message", "fraction", "pass_no", "errors", "errors_before", "kept", "started", "updated", "finished", "error", "version_id", "use_assistant", "keep_good_values", "incident", "problem")
_AUTOFIX_TITLES = {
    "start": "Starting", "compare": "Comparing with the original", "read": "Reading the workbook", "find": "Finding where the errors start", "advise": "Asking the assistant",
    "fix": "Fixing", "pass": "Fixing", "confirm": "Confirming", "save": "Saving", "verify": "Checking the file opens", "reanalyze": "Re-analyzing the fixed version", "done": "Done",
}


@app.post("/api/sessions/{session_id}/auto-fix")
def start_autofix(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Fix every formula error of the current version that can be fixed without
    changing a value that is good today (app/autofix.py), in a background
    thread: one Excel session that finds the cells where errors start, fixes
    them, recalculates and rolls back whatever moved a good value, pass after
    pass. Poll GET for progress; the result is a new version, already
    recalculated and re-analysed.

    Body (all optional): use_assistant (default true: the assistant says which
    value an error cell should fall back to, and why), keep_good_values (default
    true -- the owner's rule; false also fixes the errors a formula is hiding,
    and lists every value that moved), time_budget_min (default 120)."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    if s.current_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=422, detail="the automatic fixer works on .xlsx / .xlsm")
    if not com_available():
        raise HTTPException(status_code=409, detail="the automatic fixer needs Excel on this machine")
    running = AUTOFIX.get(session_id)
    if running and running.get("state") == "running":
        raise HTTPException(status_code=409, detail="the automatic fixer is already running for this session")
    _refuse_while_bringing_in(s)
    if session_id in RECALCULATING:
        raise HTTPException(status_code=409, detail="a recalculation is running for this workbook: wait for it to finish")
    if s.lock.locked():
        raise HTTPException(status_code=409, detail="the workbook is being analysed: wait for it to finish")
    use_assistant = bool(payload.get("use_assistant", True))
    keep = bool(payload.get("keep_good_values", True))
    budget = max(1.0, float(payload.get("time_budget_min", 120))) * 60
    advisor = make_advisor() if use_assistant else None
    state: dict[str, Any] = {
        "state": "running", "stage": "start", "title": "Starting", "message": "Opening the workbook in Excel", "fraction": None,
        "pass_no": 0, "errors": None, "errors_before": None, "kept": 0, "started": time.time(), "updated": time.time(), "finished": None,
        "error": None, "stop": False, "version_id": s.current_version_id, "use_assistant": advisor is not None, "keep_good_values": keep,
        "result": None, "version": None, "recalc": None, "analysis": None, "numbers": None,
    }
    AUTOFIX[session_id] = state
    source = s.current_path

    def progress(stage: str, message: str, fraction: float | None = None, **facts: Any) -> None:
        state.update(stage=stage, title=_AUTOFIX_TITLES.get(stage, stage), message=message, fraction=fraction, updated=time.time())
        if facts.get("pass_no") is not None:
            state["pass_no"] = facts["pass_no"]
        if facts.get("errors") is not None:
            state["errors"] = facts["errors"]
            if state["errors_before"] is None:
                state["errors_before"] = facts["errors"]
        if stage == "pass" and facts.get("kept") is not None:
            state["kept"] += int(facts["kept"])

    def worker() -> None:
        try:
            excel_worker.release(s.id)  # one Excel per session -- the fixer has its own
            # On a prepared version: what did the preparation itself change? An error that was a
            # value in the original is Prep's doing -- the fixer must not put a 0 over it.
            regressions = None
            numbers = None
            found = _original_and_moves(s)
            if found is not None:
                progress("compare", "Comparing this version with the original workbook", None)
                compared = compare_workbooks(found[0], source, found[1], progress)
                numbers = _numbers_public(compared, state["version_id"])
                regressions = compared.get("regressions") or None
            state["numbers"] = numbers
            cfg = AutoFixConfig(time_budget_s=budget, keep_good_values=keep, advisor=advisor, progress=progress, should_stop=lambda: bool(state["stop"]), regressions=regressions)
            res = run_autofix(source, s.work_dir / "autofix", cfg)
            res["numbers_check"] = numbers
            if res.get("status") == "error":
                raise RuntimeError(res.get("message") or "the automatic fixer failed")
            if not res.get("output_path"):
                ew = excel_worker.worker_for(s.id, s.work_dir / "apply")
                if ew is not None:
                    ew.preopen_async(source)  # nothing changed -- Prep's first round finds this version open
            if res.get("output_path"):
                output = Path(res["output_path"])
                left = int(res.get("errors_after") or 0)
                verified = res.get("verified_opens_in_excel")
                label = (
                    f"Auto-fix: {res.get('cells_rewritten', 0)} cell(s) fixed · "
                    + ("recalculated clean" if left == 0 else f"{left} error(s) left for a person")
                    + ("" if verified else " · NOT VERIFIED")
                )
                state["version"] = _add_version(s, output, "autofix", label, res.get("applied", []), verified, res["change_log_entry"]["output_sha256"])
                rec = _recalc_out(s, res["recalc"], output, close_round=False)
                ew = excel_worker.worker_for(s.id, s.work_dir / "apply")
                if ew is not None:
                    ew.preopen_async(output)  # Prep's first round finds it open
                progress("reanalyze", "Checking the fixed version against the rules", None)
                state["analysis"] = _run_scan(s, output)
                rec["readiness"] = s.readiness()
                rec.pop("report", None)
                rec.pop("plan", None)
                state["recalc"] = rec
            state["result"] = {k: v for k, v in res.items() if k not in ("applied", "change_log_entry", "recalc", "log")}
            state["result"]["log"] = res.get("log", [])[-200:]
            state.update(state="done", stage="done", title="Done", message=res.get("summary") or "", finished=time.time(), updated=time.time())
        except HTTPException as exc:  # the re-analysis refused (a scan already running)
            state.update(state="error", error=str(exc.detail)[:400], finished=time.time(), updated=time.time())
            if exc.status_code >= 500:
                state.update(_problem_fields(_report_problem(exc.__cause__ or exc, "Fix all issues", s)))
        except Exception as exc:  # the thread must never die silently
            state.update(state="error", error=_exc_text(exc)[:400], finished=time.time(), updated=time.time())
            state.update(_problem_fields(_report_problem(exc, "Fix all issues", s, stage=str(state.get("stage") or ""))))

    threading.Thread(target=worker, name=f"autofix-{session_id}", daemon=True).start()
    return _autofix_public(s, state, full=False)


def _autofix_public(s: Session, state: dict[str, Any], full: bool) -> dict[str, Any]:
    out = {k: state.get(k) for k in _AUTOFIX_PUBLIC}
    out["elapsed_s"] = round((state.get("finished") or time.time()) - state["started"], 1)
    if state.get("stage") == "reanalyze" and state.get("state") == "running":
        scan = _scan_snapshot(s)
        out["message"] = f"{scan.get('title') or 'Re-analyzing'}: {scan.get('message') or ''}".strip(": ")
        out["fraction"] = scan.get("overall")
    out["version"] = state.get("version")
    result = state.get("result")
    if result is not None:
        out["result"] = result if full else {k: v for k, v in result.items() if k not in ("fixes", "left", "good_values_changed_list", "log")}
    if full:
        out["recalc"] = state.get("recalc")
        out["analysis"] = state.get("analysis")
    return out


@app.get("/api/sessions/{session_id}/auto-fix")
def autofix_status(session_id: str, full: int = 0) -> dict[str, Any]:
    """Where the automatic fixer is (state running | done | error | idle, the pass,
    the errors left). `full=1` once it is done adds the whole result (every fix
    with its reason, what is left and why), the recalculation and the new analysis."""
    s = _session(session_id)
    state = AUTOFIX.get(session_id)
    if not state:
        return {"state": "idle"}
    return _autofix_public(s, state, full=bool(full))


@app.post("/api/sessions/{session_id}/auto-fix/stop")
def stop_autofix(session_id: str) -> dict[str, Any]:
    """Ask a running fixer to stop after the round it is in. What it has fixed so far is kept."""
    s = _session(session_id)
    state = AUTOFIX.get(session_id)
    if not state or state.get("state") != "running":
        return {"state": (state or {}).get("state", "idle")}
    state["stop"] = True
    return _autofix_public(s, state, full=False)


@app.get("/api/sessions/{session_id}/readiness")
def readiness(session_id: str) -> dict[str, Any]:
    """The one verdict -- blocked / unverified / ready -- with what stands
    in the way and the next step, plus the Prep gauge history."""
    s = _session(session_id)
    r = s.readiness()
    if r is None:
        raise HTTPException(status_code=409, detail="no analysis yet")
    return {"readiness": r, "prep_progress": prep_progress(s.plan_history, s.apply_history)}


def _stored_errors(s: Session) -> int | None:
    """Formula cells whose value Excel last stored is an error, from the scan's own load (still in the
    workbook cache); None when that load is gone or kept no values, or when the fixer has work the
    count cannot see: a sheet the scan skipped (the fixer reads every sheet), and -- with no error
    counted -- an error typed as a constant or a formula holding #REF! (the fixer settles both)."""
    from openpyxl.cell.cell import ERROR_CODES

    from app.inventory import _WORKBOOK_CACHE

    try:
        wb = _WORKBOOK_CACHE.get(str(s.result["workbook_analysis"]["workbooks"][0]["copy_path"]))
    except (KeyError, IndexError, TypeError):
        return None
    if wb is None:
        return None
    codes = set(ERROR_CODES)
    n = 0
    for ws in wb.worksheets:
        if getattr(ws, "mind_ready_ignored", False):
            return None  # its errors are not known
        cached = getattr(ws, "mind_ready_cached", None)
        if cached is None:
            if not hasattr(ws, "_cells"):
                continue  # a chart sheet
            return None
        n += sum(1 for v in cached.values() if isinstance(v, str) and v in codes)
    if n == 0:
        for ws in wb.worksheets:
            for c in getattr(ws, "_cells", {}).values():
                if c.data_type == "e" or (c.data_type == "f" and "#REF!" in str(getattr(c.value, "text", c.value))):
                    return None
    return n


@app.get("/api/sessions/{session_id}/estimate")
def chain_estimate(session_id: str, loop: bool = True) -> dict[str, Any]:
    """How long "Do everything" may take, per step (low / likely / high seconds) with the
    caveat; `loop=false` when the chain will skip the loop. See app/estimate.py."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    try:
        mb = Path(s.current_path).stat().st_size / 2**20
    except OSError:
        mb = 0.0
    scan_s = None
    if (s.scan or {}).get("state") == "ready":
        scan_s = sum(float(x.get("seconds") or 0) for x in s.scan.get("stages", []) if x.get("stage") != "convert") or None
    wb0 = s.result["workbook_analysis"]["workbooks"][0]
    blocking = sum(int(a.get("count") or 0) for a in s.plan if a.get("level") == "blocking" and a.get("default_on"))
    # the pivot tables Prep turns into cells cost far more than a plain write: priced by their size
    from openpyxl.utils.cell import range_boundaries

    pivots: list[tuple[str, int]] = []
    for a in s.plan:
        if a.get("id") == "pivot_tables" and a.get("level") == "blocking" and a.get("default_on"):
            for o in a.get("operations", []):
                try:
                    c1, r1, c2, r2 = range_boundaries(str(o.get("range") or o.get("cell")))
                    cells = (r2 - r1 + 1) * (c2 - c1 + 1)
                except (ValueError, TypeError):
                    cells = 0
                pivots.append((str(o.get("op")), cells))
    stored = _stored_errors(s)
    # with errors left, the original's own recalculation runs once per session (the first time)
    baseline_mb = 0.0
    original = _original_path(s)
    if original is not None and str(original) not in s.baselines and stored != 0:
        try:
            baseline_mb = original.stat().st_size / 2**20
        except OSError:
            baseline_mb = 0.0
    # the worker holds the next copy open only with enough free memory, and works out the
    # recalculation ahead only when that is on: otherwise every step pays its own Excel
    worker = excel_worker.enabled() and com_available() and excel_worker.free_mb() >= excel_worker.HOLD_MIN_FREE_MB
    return estimate_chain(mb, scan_s, len(wb0.get("formulas", [])), stored, blocking, loop, worker=worker,
                          recalc_ahead=excel_worker.speculative_recalc(), baseline_mb=baseline_mb, pivots=pivots)


def _model_map_of(s: Session, refresh: bool = False) -> dict[str, Any]:
    """The model map of the current analysed version, built once per version."""
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    if s.model_map is None or refresh or s.model_map_version_id != s.current_version_id:
        try:
            s.model_map = json.loads(json.dumps(build_model_map(s.result["workbook_analysis"]), default=str))  # dates and openpyxl values as text
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"model map failed: {type(exc).__name__}: {str(exc)[:200]}") from exc
        s.model_map_version_id = s.current_version_id
    return s.model_map


@app.get("/api/sessions/{session_id}/model-map")
def model_map(session_id: str, refresh: bool = False) -> dict[str, Any]:
    """Where the model's inputs, loop driver, run settings and outputs are: read off the
    analysed copy's macros, names, dropdowns, sheet names and formulas. Cached per analysed version."""
    s = _session(session_id)
    _model_map_of(s, refresh)
    return {"version_id": s.model_map_version_id, "map": s.model_map}


def _truthy(v: Any) -> bool:
    return str(v).strip().lower() in ("1", "true", "yes", "on")


def _loop_plan_of(s: Session, payload: dict[str, Any]) -> dict[str, Any]:
    """The Mind loop plan for the current version with the caller's choices:
    name / names, first, last, count_cell, clear_old, loops; repairs (default on): the formula
    repairs Prep still has for this version are written with the loop."""
    mp = _model_map_of(s)
    names = payload.get("names") if isinstance(payload.get("names"), dict) else None
    loops = payload.get("loops")
    if isinstance(loops, str):  # the GET sends each loop's choices as JSON
        try:
            loops = json.loads(loops) if loops.strip() else None
        except ValueError as exc:
            raise HTTPException(status_code=422, detail="loops: not JSON") from exc
    if loops is not None and not isinstance(loops, dict):
        raise HTTPException(status_code=422, detail="loops: an object by driver cell")
    try:
        plan = build_loop_plan(s.result["workbook_analysis"], mp, names=names, name=payload.get("name") or None, first=payload.get("first"), last=payload.get("last"),
                               count_cell=payload.get("count_cell") or None, dynamic=not _truthy(payload.get("fixed_ranges")), clear_old=bool(payload.get("clear_old")), loops=loops,
                               prep_plan=s.plan, repairs=payload.get("repairs") is None or _truthy(payload.get("repairs")))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"loop plan failed: {type(exc).__name__}: {str(exc)[:200]}") from exc
    return json.loads(json.dumps(plan, default=str))


@app.get("/api/sessions/{session_id}/mind-loop-plan")
def mind_loop_plan(session_id: str, name: str = "", first: str = "", last: str = "", count_cell: str = "", clear_old: bool = False, fixed_ranges: bool = False, loops: str = "", repairs: bool = True) -> dict[str, Any]:
    """The macro's run loop as Mind formulas -- MM_LOOP in the driver cells, MM_RESULT in
    the block the macro writes per pass -- as a plan (nothing written): the loops, the outputs,
    the cleared areas, the operations, and what cannot be built and why."""
    s = _session(session_id)
    plan = _loop_plan_of(s, {"name": name, "first": first, "last": last, "count_cell": count_cell, "clear_old": clear_old, "fixed_ranges": fixed_ranges, "loops": loops, "repairs": repairs})
    public = {k: v for k, v in plan.items() if k != "operations"}
    public["operations"] = [_public_op(o) for o in plan["operations"][:200]]
    public["operation_count"] = len(plan["operations"])
    return {"version_id": s.current_version_id, "plan": public}


@app.post("/api/sessions/{session_id}/mind-loop-plan/apply")
def mind_loop_apply(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Write the Mind loop plan (built again from the same choices, so what is written is
    what was shown) into a new version and re-analyse it."""
    s = _session(session_id)
    _refuse_while_fixing(session_id)
    plan = _loop_plan_of(s, payload)
    ops = plan["operations"]
    if not ops:
        raise HTTPException(status_code=422, detail="nothing to write: " + ("; ".join(plan["notes"]) or "no loop could be built"))
    out = _apply_and_version(s, ops, "loop", "Mind loop", bool(payload.get("reanalyze", True)))
    out["plan"] = {k: v for k, v in plan.items() if k != "operations"}
    return out


# --- Mind instances (one copy of the model per key of the model-point table) ----------------
def _instances_stem(s: Session) -> str:
    """The uploaded workbook's file stem with only letters, digits and underscores: the model's
    name in Mind and the start of every file this feature writes."""
    original = _original_path(s) or s.current_path
    stem = re.sub(r"[^A-Za-z0-9_]+", "_", Path(original).stem).strip("_")
    return stem or "model"


def _instance_plan_of(s: Session, key: Any, keys: Any) -> dict[str, Any]:
    """The Mind instances plan for the current version with the caller's choices: the key column
    (a header text or a column letter; empty = the plan's own choice) and which keys the main
    workbook lists ("present" = only keys that have rows, "all" = every key of the macro's list).
    The plan with no key and keys=present is kept per version: the panel's preview and the instance
    finder both ask for exactly that one. A key or keys=all is built fresh every time.

    As with the Mind loop, the formula repairs Prep still has for this version (CELL / INFO frozen,
    OFFSET ranges written out, ...) are written with a ready plan: a version that goes to Mind before
    Prep must compute what Excel does (Tokyo: CELL("filename") made every calculation sheet NaN)."""
    keys = str(keys or "present").strip().lower()
    if keys not in ("present", "all"):
        raise HTTPException(status_code=422, detail="keys must be 'present' or 'all'")
    key = str(key).strip() if key is not None else ""
    mp = _model_map_of(s)
    default = key == "" and keys == "present"
    if default and s.instance_plan_default is not None and s.instance_plan_default_version_id == s.current_version_id:
        return s.instance_plan_default
    analysis = s.result["workbook_analysis"]
    try:
        loop_plan = build_loop_plan(analysis, mp)
        plan = build_instance_plan(analysis, mp, loop_plan=loop_plan, key=key or None, keys=keys, model_name=f"{_instances_stem(s)}_instances")
        if plan.get("status") == "ready" and plan.get("operations"):
            plan.setdefault("counts", {})
            _add_repairs(plan, plan["operations"], s.plan, True, what="the instances")
            plan["operation_count"] = len(plan["operations"])
            plan["counts"]["operations"] = len(plan["operations"])
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"instances plan failed: {type(exc).__name__}: {str(exc)[:200]}") from exc
    plan = json.loads(json.dumps(plan, default=str))
    if default:
        s.instance_plan_default, s.instance_plan_default_version_id = plan, s.current_version_id
    return plan


@app.get("/api/sessions/{session_id}/mind-instances-plan")
def mind_instances_plan(session_id: str, key: str = "", keys: str = "present") -> dict[str, Any]:
    """The macro's run loop as Mind instances -- one copy of the model per value of a key column of
    the model-point table, a loop inside each copy over its own rows, and a second small main
    workbook with the keys and the results -- as a plan (nothing written): the loop, the table, the
    key and its candidates, the output block, the operations, and what cannot be built and why."""
    s = _session(session_id)
    plan = _instance_plan_of(s, key, keys)
    public = {k: v for k, v in plan.items() if k != "operations"}
    operations = list(plan.get("operations") or [])
    public["operations"] = [_public_op(o) for o in operations[:200]]
    public["operation_count"] = len(operations)
    return {"version_id": s.current_version_id, "plan": public}


@app.post("/api/sessions/{session_id}/mind-instances-plan/apply")
def mind_instances_apply(session_id: str, payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    """Write the Mind instances plan (built again from the same choices, so what is written is what
    was shown) into a new version and re-analyse it; then copy that version under the model's Mind
    name, write the main workbook beside it and zip both for the download. Once the version exists
    a failure of the main-workbook step is reported in `warnings`, not raised: the version stays."""
    s = _session(session_id)
    _refuse_while_fixing(session_id)
    plan = _instance_plan_of(s, payload.get("key"), payload.get("keys"))
    ops = list(plan.get("operations") or [])
    if not ops:
        raise HTTPException(status_code=422, detail="nothing to write: " + ("; ".join(plan.get("notes") or []) or "no instances could be built"))
    out = _apply_and_version(s, ops, "instances", "Mind instances", bool(payload.get("reanalyze", True)))
    stem = _instances_stem(s)
    main = plan.get("main") if isinstance(plan.get("main"), dict) else {}
    model_file = str(main.get("model_file") or f"{stem}_instances{s.current_path.suffix}")
    main_file = str(main.get("file") or f"{stem}_Main.xlsx")
    s.instances_zip = None  # an older zip belongs to an older version
    warnings: list[str] = []
    download: str | None = None
    status = out["result"]["status"]
    # a refused change (the key column, the driver, a result cell) leaves a model the plan does not
    # describe: no main workbook and no zip for it. A paste-area cell that stayed is not that.
    refused = [f for f in out["result"].get("failed") or [] if f.get("op") != "clear_cell"]
    if status == "APPLIED" or (status == "PARTIAL" and not refused):
        try:
            folder = s.work_dir / "instances" / str(out["version"]["id"])
            folder.mkdir(parents=True, exist_ok=True)
            model_path = folder / model_file  # the external references of the main workbook use this name
            shutil.copyfile(s.current_path, model_path)
            main_path = folder / main_file
            with perf_stage("instances_main_workbook", file=main_path.name, keys=len(main.get("keys") or [])):
                out["main_workbook"] = write_main_workbook(plan, model_path, main_path)
            s.instances_zip = Path(zip_for_mind(model_path, main_path, folder / f"{stem}_instances.zip"))
            download = f"/api/sessions/{s.id}/mind-instances/download"
        except Exception as exc:
            warnings.append(f"The main workbook could not be written: {type(exc).__name__}: {str(exc)[:200]}. The new version of the model exists; build the instances again to retry.")
    elif status == "PARTIAL":
        where = ", ".join(f"{f.get('sheet')}!{f.get('cell') or f.get('ref') or f.get('row') or f.get('column')}" for f in refused[:4])
        warnings.append(f"{len(refused)} change(s) were refused ({where}{', …' if len(refused) > 4 else ''}): the new version is not the model the plan describes, "
                        "so the main workbook and the zip were not made. Look at the refused changes, then build the instances again.")
    else:
        warnings.append("The changes were not written into the model, so the main workbook and the zip were not made.")
    out["plan"] = {k: v for k, v in plan.items() if k != "operations"}
    out["download"] = download
    out["files"] = [model_file, main_file]
    out["warnings"] = warnings
    return json.loads(json.dumps(out, default=str))


@app.get("/api/sessions/{session_id}/mind-instances/download")
def mind_instances_download(session_id: str):
    """Both files of the last "Build the Mind instances" -- the model and its main workbook -- as one
    zip, to upload together to one Mind project."""
    s = _session(session_id)
    path = s.instances_zip
    if path is None or not Path(path).is_file():
        raise HTTPException(status_code=404, detail="nothing built yet")
    return FileResponse(str(path), media_type="application/zip", filename=f"{_instances_stem(s)}_instances.zip")


# --- Instance finder (where Mind instances or a loop would help; nothing written) -------------
def _builder_block(s: Session, result: dict[str, Any]) -> dict[str, Any]:
    """The builder's real choice of key for this version -- its plan with no key chosen, built once
    per version and kept (the panel's preview computes exactly this plan) -- compared with the
    finder's top dimension that has a key. Falls back to the finder's own rule block (`source`
    "rule") when the plan cannot be built."""
    if s.instance_plan_default is None or s.instance_plan_default_version_id != s.current_version_id:
        try:
            _instance_plan_of(s, "", "present")  # fills the cache: the builder's sheet reads happen once per version
        except HTTPException:
            return dict(result.get("builder") or {})
    plan_key = (s.instance_plan_default or {}).get("key") or {}
    top = next((d for d in result.get("dimensions") or [] if d.get("key")), None)
    # None when there is nothing to compare -- no column dimension, or a plan without a key (a looped table
    # with no text column) -- so the panel shows no builder line instead of "another key: column null"
    same = (top["key"].get("column") == plan_key.get("column")) if (top and plan_key.get("column")) else None
    return {"column": plan_key.get("column"), "header": plan_key.get("header"), "chosen_by": plan_key.get("chosen_by"),
            "candidates": [c.get("column") for c in plan_key.get("candidates") or [] if isinstance(c, dict)],
            "matches_top": same, "source": "plan"}


def _instance_finder_of(s: Session, refresh: bool = False) -> dict[str, Any]:
    """The instance finder's result for the current analysed version, built once per version. A new
    version or `refresh` builds it again and forgets the assistant's block."""
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    if s.instance_finder is None or refresh or s.instance_finder_version_id != s.current_version_id:
        mp = _model_map_of(s)
        analysis = s.result["workbook_analysis"]
        try:
            with perf_stage("instance_finder", file=s.current_path.name, formulas=len(analysis["workbooks"][0].get("formulas") or [])):
                loop_plan = build_loop_plan(analysis, mp)
                result = json.loads(json.dumps(find_instances(analysis, mp, loop_plan=loop_plan), default=str))
                result["builder"] = _builder_block(s, result)
                s.instance_finder = result
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"instance finder failed: {type(exc).__name__}: {str(exc)[:200]}") from exc
        s.instance_finder_version_id = s.current_version_id
        s.instance_finder_assistant = None
    return s.instance_finder


@app.get("/api/sessions/{session_id}/instance-finder")
def instance_finder(session_id: str, assistant: str = "0", refresh: str = "0") -> dict[str, Any]:
    """Where Mind instances or a loop would help: grouping columns of the looped table, repeated
    blocks and sheets, near copies -- ranked, with evidence; nothing written. `assistant=1` adds the
    assistant's explanation and order on top (the deterministic result alone when there is no key or
    no answer: the assistant never turns into an error). `refresh=1` builds the result again."""
    s = _session(session_id)
    result = _instance_finder_of(s, _truthy(refresh))
    if not _truthy(assistant):
        return {"version_id": s.current_version_id, "finder": result}
    if s.instance_finder_assistant is None:
        advise = make_finder_advisor()
        if advise is None:
            block = {**EMPTY_ASSISTANT, "asked": True, "available": False, "message": "No AI key is saved yet: paste one under Advanced > Assistant."}
        else:
            try:  # last line of defence: the assistant never turns into an error
                with perf_stage("instance_finder_assistant", file=s.current_path.name):
                    block = advise(result)
            except Exception as exc:
                block = {**EMPTY_ASSISTANT, "asked": True, "available": True, "message": f"Request failed: {type(exc).__name__}"}
        if block.get("available") and not block.get("message"):
            s.instance_finder_assistant = block  # an answer is kept per version; no key or a failed call is asked again next time
    else:
        block = s.instance_finder_assistant
    return {"version_id": s.current_version_id, "finder": merge_assistant(result, block)}


@app.get("/api/sessions/{session_id}/cells")
def cells(session_id: str, sheet: str, cell: str, rows: int = 3, cols: int = 3) -> dict[str, Any]:
    """What is in the workbook around a cell: contents (values) and formulas,
    for the per-error workbook view. Values come from the last recalculation
    of the current version when there is one, else from the analysis copy."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    fresh = s.recalc_path if (s.recalc_path and s.recalc_version_id == s.current_version_id and Path(s.recalc_path).is_file()) else None
    return cell_window(s.result["workbook_analysis"], sheet, cell, rows=rows, cols=cols, values_path=fresh)


# --- Grid Namer ---------------------------------------------------------------
@app.get("/api/mind-flags")
def mind_flags() -> dict[str, Any]:
    """The documented Mind title flags (references/mind-flags.yaml), for the
    Grid Namer's flag picker. Header-cell markers and alias spellings are
    left out -- they are not written into a grid title."""
    flags = [
        {"name": spec["name"], "meaning": spec.get("meaning", "")}
        for spec in documented_flags().values()
        if spec.get("where", "title") == "title" and not spec.get("alias_of")
    ]
    return {"flags": sorted(flags, key=lambda f: f["name"].lower())}


@app.get("/api/sessions/{session_id}/sheet-cells")
def sheet_cells(session_id: str, sheet: str, max_rows: int = 400, max_cols: int = 60) -> dict[str, Any]:
    """The whole (used) area of one sheet of the current version, for the Grid
    Namer's workbook view: per cell the calculated value when the file carries
    one, else the formula text. Values come from the last recalculation of the
    current version when there is one, else from the analysis copy."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    analysis = s.result["workbook_analysis"]
    if sheet not in {sh["name"] for sh in analysis["workbooks"][0]["sheets"]}:
        raise HTTPException(status_code=404, detail=f"no sheet {sheet!r}")
    fresh = s.recalc_path if (s.recalc_path and s.recalc_version_id == s.current_version_id and Path(s.recalc_path).is_file()) else None
    path = Path(fresh or analysis["source"]["copy_path"])
    max_rows, max_cols = max(1, min(int(max_rows), 2000)), max(1, min(int(max_cols), 200))

    import openpyxl

    raw_wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    val_wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws_raw, ws_val = raw_wb[sheet], val_wb[sheet]
        if not hasattr(ws_raw, "iter_rows"):
            raise HTTPException(status_code=422, detail=f"{sheet} is a chart sheet -- there are no cells to show")
        total_rows, total_cols = int(ws_raw.max_row or 0), int(ws_raw.max_column or 0)
        rows: list[list[Any]] = []
        for raw_row, val_row in zip(
            ws_raw.iter_rows(min_row=1, max_row=min(total_rows, max_rows) or 1, max_col=min(total_cols, max_cols) or 1, values_only=True),
            ws_val.iter_rows(min_row=1, max_row=min(total_rows, max_rows) or 1, max_col=min(total_cols, max_cols) or 1, values_only=True),
        ):
            rows.append([json_safe(v if v is not None else raw) for raw, v in zip(raw_row, val_row)])
        while rows and all(v is None or v == "" for v in rows[-1]):
            rows.pop()
        width = 0
        for row in rows:
            for i in range(len(row) - 1, width - 1, -1):
                if row[i] is not None and row[i] != "":
                    width = i + 1
                    break
        rows = [row[:width] for row in rows]
    finally:
        raw_wb.close()
        val_wb.close()
    return {
        "sheet": sheet,
        "rows": rows,
        "n_rows": len(rows),
        "n_cols": width if rows else 0,
        "total_rows": total_rows,
        "total_cols": total_cols,
        "truncated": total_rows > max_rows or total_cols > max_cols,
        "values_from": "recalculation" if fresh else "analysis copy",
    }


@app.post("/api/sessions/{session_id}/grid-namer")
def grid_namer(session_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    """The Grid Namer's Submit: each user-selected area becomes a '#Name /Flags'
    grid title (written with the reference-safety rules -- a refused area comes
    back in `skipped` with the reason); with nameRest (default) every other
    untitled grid is titled by the existing conventions in the same apply. The
    result is a new verified version of the workbook."""
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    _refuse_while_bringing_in(s)
    raw_areas = payload.get("areas")
    if not isinstance(raw_areas, list) or not raw_areas:
        raise HTTPException(status_code=422, detail="areas must be a non-empty list")
    areas = []
    for i, a in enumerate(raw_areas, start=1):
        if not isinstance(a, dict) or not isinstance(a.get("sheet"), str) or not isinstance(a.get("ref"), str) or not str(a.get("name") or "").strip():
            raise HTTPException(status_code=422, detail=f"area {i}: sheet, ref and name are required")
        flags = a.get("flags") or []
        if not isinstance(flags, list):
            raise HTTPException(status_code=422, detail=f"area {i}: flags must be a list of flag names")
        areas.append({"sheet": a["sheet"], "ref": a["ref"], "name": str(a["name"]), "flags": [str(f) for f in flags]})

    analysis, report = s.result["workbook_analysis"], s.result["validation_report"]
    manual = plan_named_areas(analysis, areas)
    ops = list(manual["ops"])
    skipped = list(manual["skipped"])
    if not ops:
        raise HTTPException(status_code=422, detail="no area could be named: " + ("; ".join(skipped[:8]) or "nothing resolved"))
    auto_count = 0
    if bool(payload.get("nameRest", True)):
        auto_ops, auto_skipped = plan_create_grid_titles(
            analysis, report, s.grid_names or None,
            exclude=manual["exclude"], reserved=manual["reserved"],
            pre_titled=manual["titled_cells"], pre_inserted=manual["inserted_rows"], pre_banded=manual.get("banded_rows"),
        )
        ops += auto_ops
        skipped += auto_skipped
        auto_count = len(auto_ops)

    before = _before_apply(s)
    res = _tracked_apply(s, ops, bool(payload.get("reanalyze", True)))
    output = Path(res["output_path"])
    verified = res.get("verified_opens_in_excel")
    label = f"Grid Namer: {len(manual['named'])} named area(s) + {auto_count} automatic title op(s) · {'verified' if verified else 'unverified' if verified is None else 'FAILED TO OPEN'}"
    version = _add_version(s, output, "grid_namer", label, res.get("applied", []), verified, res["change_log_entry"]["output_sha256"])
    out: dict[str, Any] = {
        "result": _apply_result(res, version["file_name"]),
        "version": version,
        "named": manual["named"],
        "skipped": skipped,
        "manual_ops": len(manual["ops"]),
        "auto_ops": auto_count,
    }
    reanalyzed = bool(payload.get("reanalyze", True)) and res["status"] in ("APPLIED", "PARTIAL")
    if reanalyzed:
        out.update(_run_scan(s, output))
    out["outcome"] = _outcome(s, ops, res, before, reanalyzed)
    out["prep_progress"] = prep_progress(s.plan_history, s.apply_history)  # the gauge, this Apply included
    if res.get("failed"):
        out.update(_problem_fields(_report_problem(None, "Grid Namer", s, what="not written", stage="apply_writes", failed_ops=res["failed"])))
    return out


@app.post("/api/sessions/{session_id}/reports")
def reports(session_id: str) -> dict[str, Any]:
    s = _session(session_id)
    if not s.result:
        raise HTTPException(status_code=409, detail="no analysis yet")
    analysis = s.result["workbook_analysis"]
    report = s.result["validation_report"]
    copy_path = Path(analysis["source"]["copy_path"])
    stem = s.current_path.stem
    out_dir = s.work_dir / "reports" / s.current_version_id
    excel_worker.release(s.id)
    standalone = build_standalone_report(report, out_dir / f"{stem}_mind_readiness_report.xlsx", rules_by_id=engine().rules, source_name=s.current_path.name)
    built = build_report_workbook(copy_path, report, out_dir / f"{stem}_with_report{s.current_path.suffix}", rules_by_id=engine().rules, source_name=s.current_path.name)
    return {
        "standalone_name": _register_file(s, standalone),
        "workbook_name": _register_file(s, Path(built.path)),
        "method": built.method,
        "verified_opens_in_excel": built.verified_opens_in_excel,
        "warnings": built.warnings,
    }


@app.get("/api/sessions/{session_id}/versions")
def versions(session_id: str) -> list[dict[str, Any]]:
    return _session(session_id).versions


_MEDIA_TYPES = {
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    # the Mind inputs (UTF-8, no BOM: what Mind's Input Manager reads) and their zip
    ".csv": "text/csv; charset=utf-8",
    ".txt": "text/plain; charset=utf-8",
    ".zip": "application/zip",
}


@app.get("/api/sessions/{session_id}/files/{file_name}")
def download(session_id: str, file_name: str):
    s = _session(session_id)
    path = s.files.get(file_name)
    if path is None or not Path(path).is_file():
        raise HTTPException(status_code=404, detail=f"no file {file_name} in this session")
    media = _MEDIA_TYPES.get(Path(path).suffix.lower(), _MEDIA_TYPES[".xlsx"])
    return FileResponse(str(path), media_type=media, filename=file_name)


# --- built front-end ------------------------------------------------------------------------
if DIST_DIR.is_dir():
    from fastapi.staticfiles import StaticFiles

    if (DIST_DIR / "assets").is_dir():
        app.mount("/assets", StaticFiles(directory=str(DIST_DIR / "assets")), name="assets")

    @app.get("/{full_path:path}", include_in_schema=False)
    def spa(full_path: str):
        if full_path.startswith("api/"):
            raise HTTPException(status_code=404)
        candidate = DIST_DIR / full_path
        if full_path and candidate.is_file():
            return FileResponse(str(candidate))
        return FileResponse(str(DIST_DIR / "index.html"))

else:

    @app.get("/", include_in_schema=False)
    def no_frontend():
        return JSONResponse({"detail": f"front-end build not found at {DIST_DIR}; run `npx ng build --configuration production` in MindReadyUI or set MIND_READY_DIST"}, status_code=404)


def _answers_as_mind_ready(port: int, timeout: float = 20) -> bool:
    """A Mind Ready server already answers on `port` (slowly, when it is in
    the middle of analysing a large workbook -- hence the long timeout)."""
    import urllib.request

    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/health", timeout=timeout) as res:
            body = json.loads(res.read().decode("utf-8"))
    except Exception:
        return False
    return isinstance(body, dict) and "version" in body and "rules" in body


def _port_is_taken(port: int) -> bool:
    import socket

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(1)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def main() -> None:
    import uvicorn

    port = int(os.environ.get("MIND_READY_PORT") or SETTINGS.port)
    url = f"http://localhost:{port}"
    # Started a second time (F5 while the .bat window is open, or the reverse): say so and
    # open the app, instead of failing on "only one usage of each socket address".
    if _port_is_taken(port):
        print(f"Port {port} is already in use. Checking whether it is Mind Ready (up to 20 s) ...")
        if _answers_as_mind_ready(port):
            import webbrowser

            print(f"Mind Ready is already running: {url}")
            print("Opening it in the browser. To restart it (after a code change), stop the running one first:")
            print("close its window, or press the red square / Shift+F5 in VS Code, then start again.")
            webbrowser.open(url)
            return
        print(f"Whatever uses port {port} did not answer as Mind Ready.")
        print(f"- If Mind Ready is already running and busy with a large workbook, it is at {url}: wait, then reload the page.")
        print("- Otherwise close the program that uses the port, or choose another one: change 'port' in settings.yaml and start again.")
        raise SystemExit(1)
    uvicorn.run("app.web.server:app", host="127.0.0.1", port=port, reload=False)


if __name__ == "__main__":
    main()
