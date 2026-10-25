import { HttpClient, HttpParams } from "@angular/common/http";
import { Injectable, inject } from "@angular/core";
import { firstValueFrom } from "rxjs";
import {
  ApplyOutcome, AppHealth, AssistantKey, AutoFixStatus, CellWindow, ChainEstimate, PivotMode, PrepAction, ChatMessage, ChatReply, ChatStreamEvent, FixTarget, GridNamerOutcome, GridNamesResult, LoopPlan, MindFlag, NamedArea, SheetCells, MindLoopOptions, MindLoopStatus, Mode, ModelMap, NumbersCheck, Operation, PendingUpload, PrepProgress, Readiness, RecalcResult, ReportBuild, ScanStatus, SessionStart, Version,
} from "./models";

/** Calls to the Mind Ready backend (FastAPI, app/web/server.py). Same contracts as FigmaOutput/src/services/api.ts. */
@Injectable({ providedIn: "root" })
export class ApiService {
  private http = inject(HttpClient);
  private base = "/api";

  private s(id: string): string {
    return `${this.base}/sessions/${encodeURIComponent(id)}`;
  }

  health(): Promise<AppHealth> {
    return firstValueFrom(this.http.get<AppHealth>(`${this.base}/health`));
  }

  /** Whether an AI gateway key is saved, and its last 4 characters (the key itself never comes back). */
  assistantKey(): Promise<AssistantKey> {
    return firstValueFrom(this.http.get<AssistantKey>(`${this.base}/assistant-key`));
  }

  /** Checked with the gateway, then saved encrypted for this Windows user; a refused key is not saved. */
  saveAssistantKey(key: string): Promise<AssistantKey> {
    return firstValueFrom(this.http.post<AssistantKey>(`${this.base}/assistant-key`, { key }));
  }

  upload(file: File, mode: Mode = "plan", defer = true): Promise<SessionStart | PendingUpload> {
    const form = new FormData();
    form.append("file", file, file.name);
    form.append("mode", mode);
    if (defer) form.append("defer", "1");
    return firstValueFrom(this.http.post<SessionStart | PendingUpload>(`${this.base}/sessions`, form));
  }

  analyze(id: string, ignoreSheets: string[] = []): Promise<SessionStart> {
    return firstValueFrom(this.http.post<SessionStart>(`${this.s(id)}/analyze`, { ignore_sheets: ignoreSheets }));
  }

  resume(id: string): Promise<SessionStart> {
    return firstValueFrom(this.http.get<SessionStart>(this.s(id)));
  }

  scanStatus(id: string): Promise<ScanStatus> {
    return firstValueFrom(this.http.get<ScanStatus>(`${this.s(id)}/status`));
  }

  /** 1.11.7: what Prep does with the pivot tables; the plan comes back made again with it. */
  setPivotMode(id: string, mode: PivotMode): Promise<{ pivot_mode: PivotMode; plan: PrepAction[]; readiness: Readiness | null }> {
    return firstValueFrom(this.http.post<{ pivot_mode: PivotMode; plan: PrepAction[]; readiness: Readiness | null }>(`${this.s(id)}/pivot-mode`, { mode }));
  }

  /** 1.11.6: how long "Do everything" may take (loop=false when the chain will skip the loop). */
  estimate(id: string, loop: boolean): Promise<ChainEstimate> {
    return firstValueFrom(this.http.get<ChainEstimate>(`${this.s(id)}/estimate`, { params: new HttpParams().set("loop", String(loop)) }));
  }

  readiness(id: string): Promise<{ readiness: Readiness; prep_progress: PrepProgress }> {
    return firstValueFrom(this.http.get<{ readiness: Readiness; prep_progress: PrepProgress }>(`${this.s(id)}/readiness`));
  }

  apply(id: string, operations: Operation[], versionStep: "major" | "minor" = "major"): Promise<ApplyOutcome> {
    return firstValueFrom(this.http.post<ApplyOutcome>(`${this.s(id)}/apply`, { operations, reanalyze: true, versionStep }));
  }

  recalculate(id: string): Promise<RecalcResult> {
    return firstValueFrom(this.http.post<RecalcResult>(`${this.s(id)}/recalculate`, {}));
  }

  startAutoFix(id: string, opts: { use_assistant?: boolean; keep_good_values?: boolean; time_budget_min?: number } = {}): Promise<AutoFixStatus> {
    return firstValueFrom(this.http.post<AutoFixStatus>(`${this.s(id)}/auto-fix`, opts));
  }

  autoFixStatus(id: string, full = false): Promise<AutoFixStatus> {
    const params = new HttpParams().set("full", full ? "1" : "0");
    return firstValueFrom(this.http.get<AutoFixStatus>(`${this.s(id)}/auto-fix`, { params }));
  }

  stopAutoFix(id: string): Promise<AutoFixStatus> {
    return firstValueFrom(this.http.post<AutoFixStatus>(`${this.s(id)}/auto-fix/stop`, {}));
  }

  mindLoopPlan(id: string, o: { name?: string; first?: string; last?: string; count_cell?: string; clear_old?: boolean; fixed_ranges?: boolean } = {}): Promise<{ version_id: string; plan: LoopPlan }> {
    let params = new HttpParams();
    for (const [k, v] of Object.entries(o)) if (v !== undefined && v !== "" && v !== false) params = params.set(k, String(v));
    return firstValueFrom(this.http.get<{ version_id: string; plan: LoopPlan }>(`${this.s(id)}/mind-loop-plan`, { params }));
  }

  applyMindLoopPlan(id: string, o: { name?: string; first?: string; last?: string; count_cell?: string; clear_old?: boolean; fixed_ranges?: boolean } = {}): Promise<ApplyOutcome & { plan: LoopPlan }> {
    return firstValueFrom(this.http.post<ApplyOutcome & { plan: LoopPlan }>(`${this.s(id)}/mind-loop-plan/apply`, { ...o, reanalyze: true }));
  }

  versions(id: string): Promise<Version[]> {
    return firstValueFrom(this.http.get<Version[]>(`${this.s(id)}/versions`));
  }

  reanalyze(id: string, versionId: string): Promise<Partial<SessionStart>> {
    return firstValueFrom(this.http.post<Partial<SessionStart>>(`${this.s(id)}/reanalyze`, { versionId }));
  }

  modelMap(id: string, refresh = false): Promise<{ version_id: string; map: ModelMap }> {
    const params = refresh ? new HttpParams().set("refresh", "true") : new HttpParams();
    return firstValueFrom(this.http.get<{ version_id: string; map: ModelMap }>(`${this.s(id)}/model-map`, { params }));
  }

  cellWindow(id: string, sheet: string, cell: string, rows = 3, cols = 3): Promise<CellWindow> {
    const params = new HttpParams().set("sheet", sheet).set("cell", cell).set("rows", String(rows)).set("cols", String(cols));
    return firstValueFrom(this.http.get<CellWindow>(`${this.s(id)}/cells`, { params }));
  }

  numbersCheck(id: string): Promise<NumbersCheck> {
    return firstValueFrom(this.http.post<NumbersCheck>(`${this.s(id)}/numbers-check`, {}));
  }

  generateReports(id: string): Promise<ReportBuild> {
    return firstValueFrom(this.http.post<ReportBuild>(`${this.s(id)}/reports`, {}));
  }

  startMindLoop(id: string, opts: MindLoopOptions): Promise<{ state: string; work_dir?: string }> {
    return firstValueFrom(this.http.post<{ state: string; work_dir?: string }>(`${this.s(id)}/mind-loop`, opts));
  }

  mindLoopStatus(id: string, after = 0): Promise<MindLoopStatus> {
    const params = new HttpParams().set("after", String(after));
    return firstValueFrom(this.http.get<MindLoopStatus>(`${this.s(id)}/mind-loop`, { params }));
  }

  /** The assistant's streamed answer (NDJSON over a POST body): `onEvent` gets phases and text deltas, the promise the whole reply. */
  async chatStream(id: string, history: ChatMessage[], question: string, focus: FixTarget | undefined, onEvent: (ev: ChatStreamEvent) => void): Promise<ChatReply> {
    const body = {
      history: history.map((m) => ({ role: m.role, content: m.content, note: m.note ?? false })),
      question,
      focus: focus
        ? { kind: focus.kind ?? "finding", rule_id: focus.rule_id, sheet: focus.sheet, cell: focus.cell, error: focus.error, formula: focus.formula, addin_gap: focus.addin_gap ?? false, cells: focus.cells, cause: focus.cause }
        : undefined,
    };
    const res = await fetch(`${this.s(id)}/chat/stream`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!res.ok || !res.body) {
      const err = await res.json().catch(() => null);
      throw new Error(typeof err?.detail === "string" ? err.detail : `${res.status} ${res.statusText}`);
    }
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let reply: ChatReply | null = null;
    const handle = (line: string) => {
      if (!line.trim()) return;
      const ev = JSON.parse(line);
      if (ev.type === "done") reply = { text: ev.text, proposal: ev.proposal ?? null, provenance: ev.provenance };
      else if (ev.type === "error") throw new Error(ev.detail || "the assistant failed");
      else if (ev.type === "phase" || ev.type === "delta") onEvent(ev as ChatStreamEvent);
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
    if (buffer.trim()) handle(buffer);
    if (!reply) throw new Error("the assistant sent no answer");
    return reply;
  }

  sheetCells(id: string, sheet: string, maxRows = 400, maxCols = 60): Promise<SheetCells> {
    const params = new HttpParams().set("sheet", sheet).set("max_rows", String(maxRows)).set("max_cols", String(maxCols));
    return firstValueFrom(this.http.get<SheetCells>(`${this.s(id)}/sheet-cells`, { params }));
  }

  async mindFlags(): Promise<MindFlag[]> {
    const out = await firstValueFrom(this.http.get<{ flags: MindFlag[] }>(`${this.base}/mind-flags`));
    return out.flags ?? [];
  }

  suggestGridNames(id: string, apply = true): Promise<GridNamesResult> {
    return firstValueFrom(this.http.post<GridNamesResult>(`${this.s(id)}/grid-names`, { apply }));
  }

  applyNamedAreas(id: string, areas: NamedArea[], nameRest: boolean): Promise<GridNamerOutcome> {
    return firstValueFrom(this.http.post<GridNamerOutcome>(`${this.s(id)}/grid-namer`, { areas, nameRest, reanalyze: true }));
  }

  downloadUrl(id: string, fileName: string): string {
    return `${this.s(id)}/files/${encodeURIComponent(fileName)}`;
  }
}

/** The text of a failed call, as the backend's `detail` when it sent one. */
export function errorText(e: unknown): string {
  const err = e as { error?: { detail?: unknown }; message?: string; status?: number };
  const detail = err?.error?.detail;
  if (typeof detail === "string") return detail;
  if (detail) return JSON.stringify(detail);
  return err?.message || "request failed";
}
