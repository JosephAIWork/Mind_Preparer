import { Injectable, computed, signal } from "@angular/core";
import { PrepAction, PrepProgress, Readiness, SessionStart, ValidationReport, Version, WorkbookSummary } from "./models";

/** The loaded session, as signals. The backend keeps sessions in memory; `?session=<id>` reopens one. */
@Injectable({ providedIn: "root" })
export class SessionStore {
  readonly sessionId = signal<string | null>(null);
  readonly summary = signal<WorkbookSummary | null>(null);
  readonly report = signal<ValidationReport | null>(null);
  readonly plan = signal<PrepAction[]>([]);
  readonly versions = signal<Version[]>([]);
  readonly currentVersion = signal<Version | null>(null);
  readonly readiness = signal<Readiness | null>(null);
  readonly prepProgress = signal<PrepProgress | null>(null);

  readonly loaded = computed(() => this.sessionId() !== null && this.report() !== null);
  readonly findings = computed(() => this.report()?.findings ?? []);
  readonly statusCounts = computed(() => {
    const out: Record<string, number> = { ERROR: 0, REQUIRES_USER_INPUT: 0, WARNING: 0, NOT_SUPPORTED: 0, PASS: 0 };
    for (const f of this.findings()) out[f.status] = (out[f.status] ?? 0) + 1;
    return out;
  });

  setSession(start: SessionStart): void {
    this.sessionId.set(start.sessionId);
    this.summary.set(start.summary);
    this.report.set(start.report);
    this.plan.set(start.plan ?? []);
    this.versions.set(start.versions ?? [start.version]);
    this.currentVersion.set(start.version);
    this.readiness.set(start.readiness ?? null);
    this.prepProgress.set(start.prep_progress ?? null);
  }

  /** Merge what a re-analysis, an Apply or an auto-fix sent back (only the fields present). */
  applyAnalysis(a: Partial<SessionStart> & { version?: Version | null }): void {
    if (a.summary) this.summary.set(a.summary);
    if (a.report) this.report.set(a.report);
    if (a.plan) this.plan.set(a.plan);
    if (a.versions) this.versions.set(a.versions);
    if (a.version) this.addVersion(a.version);
    if (a.readiness) this.readiness.set(a.readiness);
    if (a.prep_progress) this.prepProgress.set(a.prep_progress);
  }

  addVersion(v: Version): void {
    const list = this.versions();
    if (!list.some((x) => x.id === v.id)) this.versions.set([...list, v]);
    this.currentVersion.set(v);
  }

  reset(): void {
    this.sessionId.set(null);
    this.summary.set(null);
    this.report.set(null);
    this.plan.set([]);
    this.versions.set([]);
    this.currentVersion.set(null);
    this.readiness.set(null);
    this.prepProgress.set(null);
  }
}
