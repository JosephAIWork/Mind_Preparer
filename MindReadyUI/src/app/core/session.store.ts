import { Injectable, computed, signal } from "@angular/core";

export type AdvancedTab = "findings" | "prep" | "recalc" | "loop" | "map" | "namer" | "assistant" | "mind" | "deliver" | "versions";
import { ChatMessage, Delta, FixTarget, PivotMode, NumbersCheck, PrepAction, PrepProgress, Readiness, RecalcResult, ReportBuild, SessionStart, ValidationReport, Version, WorkbookSummary } from "./models";

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
  readonly delta = signal<Delta | null>(null); // what changed since the previous analysis
  readonly recalc = signal<RecalcResult | null>(null); // the last recalculation shown on the Advanced page
  readonly chatHistory = signal<ChatMessage[]>([]);
  readonly fixTarget = signal<FixTarget | null>(null); // what the Assistant tab is asked about
  readonly numbers = signal<NumbersCheck | null>(null);
  readonly reports = signal<ReportBuild | null>(null);
  readonly advancedTab = signal<AdvancedTab>("findings");
  /** 1.11.7: what Prep (and Do everything) does with the pivot tables */
  readonly pivotMode = signal<PivotMode>("formulas");
  readonly pivotModes = signal<Record<string, string>>({});
  /** STR-009's pivots: [{sheet, name, ref, blocking}] (empty when the workbook has none) */
  readonly pivots = computed(() => {
    const f = this.findings().find((x) => x.rule_id === "STR-009");
    const list = (f?.observed as { pivots?: { sheet: string; name: string; ref: string; blocking: string | null }[] } | undefined)?.pivots;
    return list ?? [];
  });

  /** Point the Assistant at a finding or an error and open its tab. */
  ask(target: FixTarget): void {
    this.fixTarget.set(target);
    this.advancedTab.set("assistant");
  }

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
    this.delta.set(start.delta ?? null);
    this.pivotMode.set(start.pivot_mode ?? "formulas");
    if (start.pivot_modes) this.pivotModes.set(start.pivot_modes);
    this.recalc.set(null);
    this.chatHistory.set([]);
    this.fixTarget.set(null);
    this.numbers.set(null);
    this.reports.set(null);
  }

  /** Merge what a re-analysis, an Apply or an auto-fix sent back (only the fields present). */
  applyAnalysis(a: Partial<SessionStart> & { version?: Version | null }): void {
    if (a.summary) this.summary.set(a.summary);
    if (a.report) this.report.set(a.report);
    if (Array.isArray(a.plan)) this.plan.set(a.plan); // the loop-apply response names ITS plan `plan` too: never that one
    if (a.versions) this.versions.set(a.versions);
    if (a.version) this.addVersion(a.version);
    if (a.readiness) this.readiness.set(a.readiness);
    if (a.prep_progress) this.prepProgress.set(a.prep_progress);
    if (a.delta !== undefined) this.delta.set(a.delta ?? null);
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
    this.delta.set(null);
    this.recalc.set(null);
    this.chatHistory.set([]);
    this.fixTarget.set(null);
    this.numbers.set(null);
    this.reports.set(null);
  }
}
