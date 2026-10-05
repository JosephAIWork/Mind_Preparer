import type { ApplyReport, ApplyResult, CellWindow, ChatMessage, ChatReply, ChatStreamEvent, Delta, FixTarget, Mode, Operation, PrepAction, PrepProgress, Readiness, RecalcResult, ReportBuild, ScanStatus, SizeInfo, ValidationReport, Version, WorkbookSummary } from "../types";

import workbookMock from "../mocks/workbook.json";
import reportMock from "../mocks/report.json";
import planMock from "../mocks/plan.json";
import versionsMock from "../mocks/versions.json";
import chatMock from "../mocks/chat.json";

/**
 * Live data layer. Every function talks to the FastAPI backend
 * (excel-upload-preparation/app/web/server.py) under `/api`; in development
 * Vite proxies `/api` to http://localhost:8600 (see vite.config.ts).
 * Set VITE_USE_MOCKS=true to fall back to the design-time fixtures.
 */
const USE_MOCKS = import.meta.env.VITE_USE_MOCKS === "true";
const API = "/api";

function delay(ms: number) {
  return new Promise((r) => setTimeout(r, ms));
}

function cast<T>(v: unknown): T {
  return v as T;
}

async function unwrap<T>(res: Response): Promise<T> {
  if (!res.ok) {
    let detail = `${res.status} ${res.statusText}`;
    try {
      const body = await res.json();
      if (body && typeof body.detail === "string") detail = body.detail;
    } catch {
      /* not JSON */
    }
    throw new Error(detail);
  }
  return (await res.json()) as T;
}

async function post<T>(path: string, body?: unknown): Promise<T> {
  const res = await fetch(`${API}${path}`, {
    method: "POST",
    headers: body === undefined ? undefined : { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  return unwrap<T>(res);
}

export interface SessionStart {
  sessionId: string;
  summary: WorkbookSummary;
  report: ValidationReport;
  plan: PrepAction[];
  version: Version;
  /** 1.6.6 backends also return the whole lineage, the size facts and the finished scan status. */
  versions?: Version[];
  size?: SizeInfo | null;
  scan?: ScanStatus;
  pending?: false;
  /** 1.7.2: the verdict and the Prep gauge travel with every analysis */
  readiness?: Readiness | null;
  prep_progress?: PrepProgress | null;
}

/**
 * A deferred upload (1.6.6): the file is stored and inspected -- sizes and the
 * sheet list, read from the package alone -- but not scanned yet. Call
 * analyzeSession() with the sheets to skip (if any) to run the scan.
 */
export interface PendingUpload {
  sessionId: string;
  pending: true;
  mode: Mode;
  size: SizeInfo;
  version: Version;
  versions: Version[];
  scan: ScanStatus;
}

export function isPending(r: SessionStart | PendingUpload): r is PendingUpload {
  return (r as PendingUpload).pending === true;
}

/** A fresh analysis of the current version, with what changed since the previous one. */
export interface Reanalysis {
  summary: WorkbookSummary;
  report: ValidationReport;
  plan: PrepAction[];
  delta?: Delta;
  versions?: Version[];
  readiness?: Readiness | null;
  prep_progress?: PrepProgress | null;
}

/** 1.7.2: the verdict for the current version, on demand (it also comes with every analysis). */
export async function getReadiness(sessionId: string): Promise<{ readiness: Readiness; prep_progress: PrepProgress }> {
  return unwrap(await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/readiness`));
}

/** applyOperations also returns the re-analysis of the new file when the backend ran one. */
export interface ApplyOutcome extends Partial<Reanalysis> {
  result: ApplyResult;
  version: Version;
  /** 1.7.3: the Apply, repair by repair (absent from an older backend) */
  outcome?: ApplyReport | null;
}

export interface UploadOptions {
  /** Store + inspect only; the scan runs on analyzeSession(). An older backend ignores the flag and scans at once. */
  defer?: boolean;
  /** Sheets not to scan (non-deferred upload). */
  ignoreSheets?: string[];
}

export async function uploadWorkbook(file: File, mode: Mode, options: UploadOptions = {}): Promise<SessionStart | PendingUpload> {
  if (USE_MOCKS) {
    await delay(3500);
    return {
      sessionId: "session-mock-001",
      summary: cast<WorkbookSummary>(workbookMock),
      report: cast<ValidationReport>(reportMock),
      plan: cast<PrepAction[]>(planMock),
      version: cast<Version[]>(versionsMock)[0],
    };
  }
  const form = new FormData();
  form.append("file", file, file.name);
  form.append("mode", mode);
  if (options.defer) form.append("defer", "1");
  if (options.ignoreSheets && options.ignoreSheets.length) form.append("ignore_sheets", JSON.stringify(options.ignoreSheets));
  const res = await fetch(`${API}/sessions`, { method: "POST", body: form });
  return unwrap<SessionStart | PendingUpload>(res);
}

/**
 * Scan a deferred upload (or scan the current version again) skipping
 * `ignoreSheets`. A long call: poll scanStatus() meanwhile for the indicator.
 */
export async function analyzeSession(sessionId: string, ignoreSheets: string[] = []): Promise<SessionStart> {
  if (USE_MOCKS) {
    await delay(3000);
    return {
      sessionId,
      summary: cast<WorkbookSummary>(workbookMock),
      report: cast<ValidationReport>(reportMock),
      plan: cast<PrepAction[]>(planMock),
      version: cast<Version[]>(versionsMock)[0],
    };
  }
  return post<SessionStart>(`/sessions/${encodeURIComponent(sessionId)}/analyze`, { ignore_sheets: ignoreSheets });
}

/** Where the current (or last) scan of a session is. Older backends answer 404 -- callers treat that as "no status". */
export async function scanStatus(sessionId: string): Promise<ScanStatus> {
  if (USE_MOCKS) {
    await delay(150);
    return { state: "running", stage: "load", title: "Reading the workbook", message: "Reading sheet 2/8: Cashflows", fraction: 0.25, overall: 0.3, elapsed_s: 4.2, stages: [], error: null };
  }
  return unwrap<ScanStatus>(await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/status`));
}

export async function reanalyze(sessionId: string, versionId: string): Promise<Reanalysis> {
  if (USE_MOCKS) {
    await delay(2500);
    return {
      summary: cast<WorkbookSummary>(workbookMock),
      report: cast<ValidationReport>(reportMock),
      plan: cast<PrepAction[]>(planMock),
    };
  }
  return post<Reanalysis>(`/sessions/${encodeURIComponent(sessionId)}/reanalyze`, { versionId });
}

export async function applyOperations(
  sessionId: string,
  operations: Operation[],
  options: { reanalyze?: boolean } = {}
): Promise<ApplyOutcome> {
  if (USE_MOCKS) {
    await delay(3000);
    const result: ApplyResult = {
      status: "APPLIED",
      method: "excel_com",
      output_path: "/tmp/IFRS_Risk_Model_prep.xlsm",
      output_name: "IFRS_Risk_Model_prep.xlsm",
      applied: operations,
      failed: [],
      verified_opens_in_excel: true,
      warnings: [],
      message: `${operations.length} operation(s) applied successfully. Written by Excel and verified to open.`,
    };
    const newVersion: Version = {
      id: `ver-${Date.now()}`,
      label: `v2 — Prep: ${operations.length} changes via Excel · verified`,
      file_name: "IFRS_Risk_Model_prep.xlsm",
      sha256: "mock",
      created_at: new Date().toISOString(),
      source: "prep",
      change_log: operations,
      verified_opens_in_excel: true,
    };
    return { result, version: newVersion };
  }
  return post<ApplyOutcome>(`/sessions/${encodeURIComponent(sessionId)}/apply`, {
    operations,
    reanalyze: options.reanalyze ?? true,
  });
}

export async function suggestFormulaFix(
  sessionId: string,
  ruleId: string,
  sheet: string,
  cell: string
): Promise<{ suggestion: string | null; formula: string | null; message: string | null }> {
  if (USE_MOCKS) {
    await delay(2000);
    return {
      suggestion: "Replace AGGREGATE with SUBTOTAL (which is on the supported list).",
      formula: "=SUBTOTAL(9,L8:L200)",
      message: null,
    };
  }
  return post(`/sessions/${encodeURIComponent(sessionId)}/suggest`, { ruleId, sheet, cell });
}

/**
 * One conversation turn. `focus` (optional) ties the question to one finding --
 * the backend then retrieves that finding's details and the cell contents.
 */
export async function chat(sessionId: string, history: ChatMessage[], question: string, focus?: FixTarget): Promise<ChatReply> {
  if (USE_MOCKS) {
    await delay(1800);
    const last = cast<Array<{ role: string; content: string; provenance?: unknown }>>(chatMock).find(
      (m) => m.role === "assistant" && m.provenance
    );
    return {
      text: last?.content ?? "I don't have more information on that topic in the current context.",
      proposal: null,
      provenance: { context_chars: 14200, detail_chars: 2100, lookups: 1 },
    };
  }
  return post<ChatReply>(`/sessions/${encodeURIComponent(sessionId)}/chat`, chatBody(history, question, focus));
}

function chatBody(history: ChatMessage[], question: string, focus?: FixTarget) {
  return {
    history: history.map((m) => ({ role: m.role, content: m.content, note: m.note ?? false })),
    question,
    focus: focus
      ? {
          kind: focus.kind ?? "finding",
          rule_id: focus.rule_id,
          sheet: focus.sheet,
          cell: focus.cell,
          error: focus.error,
          formula: focus.formula,
          addin_gap: focus.addin_gap ?? false,
          cells: focus.cells,
          cause: focus.cause,
        }
      : undefined,
  };
}

/**
 * 1.7.2: the same turn as chat(), streamed. The backend sends NDJSON lines --
 * `phase` before every model call, `delta` for each chunk of its text -- and
 * ends with `done` (the chat() payload) or `error`. Falls back to chat() with
 * mocks or on a backend without the route (404).
 */
export async function chatStream(
  sessionId: string,
  history: ChatMessage[],
  question: string,
  focus: FixTarget | undefined,
  onEvent: (ev: ChatStreamEvent) => void
): Promise<ChatReply> {
  if (USE_MOCKS) return chat(sessionId, history, question, focus);
  const res = await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/chat/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(chatBody(history, question, focus)),
  });
  if (res.status === 404 || res.status === 405) {
    // A pre-1.7.2 backend has no such route: FastAPI answers "Not Found", or 405 via the GET-only
    // SPA catch-all. Anything else (e.g. an unknown session) is a real error.
    const body = await res.json().catch(() => null);
    if (!body || body.detail === "Not Found" || body.detail === "Method Not Allowed") return chat(sessionId, history, question, focus);
    throw new Error(typeof body.detail === "string" ? body.detail : `${res.status} ${res.statusText}`);
  }
  if (!res.ok || !res.body) return unwrap<ChatReply>(res);
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  const out: { reply: ChatReply | null } = { reply: null };
  const handle = (line: string) => {
    if (!line.trim()) return;
    const ev = JSON.parse(line);
    if (ev.type === "done") {
      out.reply = { text: ev.text, proposal: ev.proposal ?? null, provenance: ev.provenance };
    } else if (ev.type === "error") {
      throw new Error(ev.detail || "the assistant failed");
    } else if (ev.type === "phase" || ev.type === "delta") {
      onEvent(ev as ChatStreamEvent);
    }
  };
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    let nl: number;
    while ((nl = buffer.indexOf("\n")) >= 0) {
      handle(buffer.slice(0, nl));
      buffer = buffer.slice(nl + 1);
    }
  }
  handle(buffer + decoder.decode());
  if (!out.reply) throw new Error("the reply stream ended early");
  return out.reply;
}

export async function recalculate(sessionId: string): Promise<RecalcResult> {
  if (USE_MOCKS) {
    await delay(6000);
    return {
      status: "ERROR",
      message: "Recalculation complete. Formula errors detected.",
      formula_errors: [{ sheet: "MM_Change_Log", cell: "L7", error: "#NAME?", formula: "=AGGREGATE(9,6,D7:D200)" }],
      addin_gap_errors: [{ sheet: "BEL", cell: "C5", formula: '=MM_RESULT(CohortLoop,"BEL")' }],
    };
  }
  return post<RecalcResult>(`/sessions/${encodeURIComponent(sessionId)}/recalculate`);
}

export async function generateReports(sessionId: string): Promise<ReportBuild> {
  if (USE_MOCKS) {
    await delay(4000);
    return {
      standalone_name: "IFRS_Risk_Model_report.xlsx",
      workbook_name: "IFRS_Risk_Model_with_report.xlsm",
      method: "excel_com",
      verified_opens_in_excel: true,
      warnings: [],
    };
  }
  return post<ReportBuild>(`/sessions/${encodeURIComponent(sessionId)}/reports`);
}

export async function cellWindow(sessionId: string, sheet: string, cell: string, rows = 3, cols = 3): Promise<CellWindow> {
  if (USE_MOCKS) {
    await delay(200);
    return { sheet, focus: cell, columns: ["A", "B"], rows: [{ row: 1, cells: [{ ref: "A1", value: "mock", formula: null, array: null, error: false, focus: true }, { ref: "B1", value: 1, formula: "=1", array: null, error: false, focus: false }] }], values_from: "mock" };
  }
  const q = new URLSearchParams({ sheet, cell, rows: String(rows), cols: String(cols) });
  const res = await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/cells?${q}`);
  return unwrap<CellWindow>(res);
}

/** One grid whose name carries no meaning, with the alternative the assistant proposes. */
export interface GridNameSuggestion {
  grid: string;
  sheet: string;
  ref: string;
  size: string;
  current: string | null;
  deterministic: string;
  deterministic_source: string;
  suggested: string | null;
}

export interface GridNamesResult {
  available: boolean;
  names: GridNameSuggestion[];
  message: string | null;
  applied: boolean;
  plan?: PrepAction[] | null;
  readiness?: Readiness | null;
}

/**
 * Ask the assistant (Milliman APIM) to name the grids the heuristics could only
 * call "<Sheet> <Anchor>". Nothing is written: with `apply` the names are folded
 * into the prep plan, which the user still reviews and approves.
 */
export async function suggestGridNames(sessionId: string, apply = true): Promise<GridNamesResult> {
  if (USE_MOCKS) {
    await delay(1500);
    return {
      available: true,
      applied: apply,
      message: null,
      names: [
        { grid: "Cashflows!C27:E39", sheet: "Cashflows", ref: "C27:E39", size: "13x3", current: null, deterministic: "Cashflows C27", deterministic_source: "sheet name + position", suggested: "Demographic Assumption Inputs" },
      ],
    };
  }
  return post<GridNamesResult>(`/sessions/${encodeURIComponent(sessionId)}/grid-names`, { apply });
}

// --- Grid Namer (1.7.0) ---------------------------------------------------------------

/** One sheet of the current version, cell by cell (calculated value where the file carries one, else the formula text). */
export interface SheetCells {
  sheet: string;
  rows: unknown[][];
  n_rows: number;
  n_cols: number;
  total_rows: number;
  total_cols: number;
  truncated: boolean;
  values_from: string;
}

export async function sheetCells(sessionId: string, sheet: string, maxRows = 400, maxCols = 60): Promise<SheetCells> {
  if (USE_MOCKS) {
    await delay(300);
    return { sheet, rows: [["Header", 1], [null, 2]], n_rows: 2, n_cols: 2, total_rows: 2, total_cols: 2, truncated: false, values_from: "mock" };
  }
  const q = new URLSearchParams({ sheet, max_rows: String(maxRows), max_cols: String(maxCols) });
  return unwrap<SheetCells>(await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/sheet-cells?${q}`));
}

/** A documented Mind title flag (references/mind-flags.yaml). */
export interface MindFlag { name: string; meaning: string }

export async function mindFlags(): Promise<MindFlag[]> {
  if (USE_MOCKS) {
    await delay(100);
    return [{ name: "Input", meaning: "Assumptions-only input grid." }];
  }
  return (await unwrap<{ flags: MindFlag[] }>(await fetch(`${API}/mind-flags`))).flags;
}

/** One user-named area: the selection, the chosen name, the chosen flags. */
export interface NamedArea { sheet: string; ref: string; name: string; flags: string[] }

export interface GridNamerOutcome extends Partial<Reanalysis> {
  result: ApplyResult;
  version: Version;
  outcome?: ApplyReport | null;
  /** Grid keys ("Sheet!Ref") a title was written for. */
  named: string[];
  /** Refused areas and automatic-titler skips, each with its reason. */
  skipped: string[];
  manual_ops: number;
  auto_ops: number;
}

/**
 * Submit of the Grid Namer screen: writes a '#Name /Flags' title for each
 * named area (reference-safety rules apply -- refusals come back in `skipped`),
 * optionally titles every remaining untitled grid by the existing conventions,
 * and returns the new verified version plus its re-analysis.
 */
export async function applyNamedAreas(sessionId: string, areas: NamedArea[], nameRest: boolean): Promise<GridNamerOutcome> {
  if (USE_MOCKS) {
    await delay(1500);
    throw new Error("not available in mock mode");
  }
  return post<GridNamerOutcome>(`/sessions/${encodeURIComponent(sessionId)}/grid-namer`, { areas, nameRest, reanalyze: true });
}

/** Options for the autonomous run-in-Mind loop (app/mind_loop.py). */
export interface MindLoopOptions {
  maxIterations?: number;
  useAssistant?: boolean;
  runModel?: boolean;
  checkNumbers?: boolean;
  skipMind?: boolean;
  enable?: string[];
  disable?: string[];
}

export interface MindLoopEvent { t: number; event: string; message?: string; n?: number; verdict?: string }

export interface MindLoopIteration {
  n: number;
  enabled: string[];
  disabled: string[];
  actions: { id: string; count: number }[];
  apply?: { status: string; applied: number; failed: number; verified: boolean | null; message: string; output: string };
  names?: { grids: number; baseline_grids: number; fragmented: boolean; unnamed: number; unnamed_grids: string[]; still_titleable: number; blocked: string[]; ok: boolean };
  numbers?: { ran: boolean; match?: boolean; compared?: number; differences?: number; listed?: { sheet: string; cell: string | null; prepared_cell?: string; source: unknown; prepared: unknown }[]; volatile_skipped?: number; message?: string };
  mind?: {
    skipped: boolean;
    reason?: string;
    error?: string;
    convert?: { success: boolean; steps: Record<string, string>; errors: string[] };
    templates?: { stem: string; entries: number; untitled: { label: string; cell: string }[] };
    counts?: { mind_untitled: number | null; app_untitled: number };
    run?: { completed: boolean; audit_consistent: boolean; seconds: number | null };
    exports?: { opened: boolean; elements: number | null };
  };
  decision?: { verdict: "converged" | "retry" | "stuck"; reason: string; enable: string[]; disable: string[] };
}

export interface MindLoopReport {
  source: string;
  started: string;
  finished?: string;
  iterations: MindLoopIteration[];
  projects: { name: string; status: string; ran: boolean; delete?: string }[];
  verdict: "converged" | "stuck" | "exhausted" | null;
  reason: string | null;
  final_workbook?: string;
  leftover_projects?: string[];
}

export interface MindLoopStatus {
  state: "idle" | "running" | "done" | "error";
  events: MindLoopEvent[];
  next: number;
  report: MindLoopReport | null;
  error?: string | null;
  summary?: string | null;
  work_dir?: string;
}

export async function startMindLoop(sessionId: string, options: MindLoopOptions = {}): Promise<{ state: string; work_dir: string }> {
  if (USE_MOCKS) {
    await delay(300);
    return { state: "running", work_dir: "/tmp/mind_loop" };
  }
  return post(`/sessions/${encodeURIComponent(sessionId)}/mind-loop`, options);
}

export async function mindLoopStatus(sessionId: string, after = 0): Promise<MindLoopStatus> {
  if (USE_MOCKS) {
    await delay(200);
    return { state: "idle", events: [], next: 0, report: null };
  }
  return unwrap<MindLoopStatus>(await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/mind-loop?after=${after}`));
}

export interface AppHealth {
  version: string;
  rules: number;
  excel: boolean;
  assistant: boolean;
  frontend: boolean;
  /** 1.6.6: deferred upload + sheet selection + scan status are available */
  upload_gate?: boolean;
  size_threshold_mb?: number;
}

/** Live build info from the running backend (version, active rule count, capabilities). */
export async function health(): Promise<AppHealth> {
  return unwrap<AppHealth>(await fetch(`${API}/health`));
}

export function downloadUrl(sessionId: string, fileName: string): string {
  if (USE_MOCKS) return `#download/${encodeURIComponent(fileName)}`;
  return `${API}/sessions/${encodeURIComponent(sessionId)}/files/${encodeURIComponent(fileName)}`;
}

export async function listVersions(sessionId: string): Promise<Version[]> {
  if (USE_MOCKS) {
    await delay(400);
    return cast<Version[]>(versionsMock);
  }
  const res = await fetch(`${API}/sessions/${encodeURIComponent(sessionId)}/versions`);
  return unwrap<Version[]>(res);
}
