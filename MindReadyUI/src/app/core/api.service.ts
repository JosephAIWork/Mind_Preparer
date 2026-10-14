import { HttpClient, HttpParams } from "@angular/common/http";
import { Injectable, inject } from "@angular/core";
import { firstValueFrom } from "rxjs";
import {
  ApplyOutcome, AppHealth, AutoFixStatus, LoopPlan, Mode, Operation, PendingUpload, ScanStatus, SessionStart, Version, Readiness, PrepProgress, RecalcResult,
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

  mindLoopPlan(id: string, o: { name?: string; first?: string; last?: string; count_cell?: string; clear_old?: boolean } = {}): Promise<{ version_id: string; plan: LoopPlan }> {
    let params = new HttpParams();
    for (const [k, v] of Object.entries(o)) if (v !== undefined && v !== "" && v !== false) params = params.set(k, String(v));
    return firstValueFrom(this.http.get<{ version_id: string; plan: LoopPlan }>(`${this.s(id)}/mind-loop-plan`, { params }));
  }

  applyMindLoopPlan(id: string, o: { name?: string; first?: string; last?: string; count_cell?: string; clear_old?: boolean } = {}): Promise<ApplyOutcome & { plan: LoopPlan }> {
    return firstValueFrom(this.http.post<ApplyOutcome & { plan: LoopPlan }>(`${this.s(id)}/mind-loop-plan/apply`, { ...o, reanalyze: true }));
  }

  versions(id: string): Promise<Version[]> {
    return firstValueFrom(this.http.get<Version[]>(`${this.s(id)}/versions`));
  }

  reanalyze(id: string, versionId: string): Promise<Partial<SessionStart>> {
    return firstValueFrom(this.http.post<Partial<SessionStart>>(`${this.s(id)}/reanalyze`, { versionId }));
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
