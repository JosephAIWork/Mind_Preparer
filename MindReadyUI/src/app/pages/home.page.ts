import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, effect, inject, signal } from "@angular/core";
import { Router, RouterLink } from "@angular/router";
import { ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule } from "carbon-components-angular";
import { ApiService, errorText } from "../core/api.service";
import { FORMULA_ACTION_IDS, LoopPlan, Operation, PendingUpload, appliedCount, failedCount, failedText, isPending } from "../core/models";
import { SessionStore } from "../core/session.store";

type Phase = "idle" | "uploading" | "choose" | "scanning" | "loaded";
interface ActionState {
  running: boolean;
  fraction: number | null;
  line: string;
  detail: string;
  tone: "idle" | "ok" | "warn" | "error";
  /** the result text is shown only on "View details" */
  open?: boolean;
}

const idle = (): ActionState => ({ running: false, fraction: null, line: "", detail: "", tone: "idle" });

/** The main screen: one workbook, three buttons that do the work on their own, one download. */
@Component({
  selector: "mr-home",
  standalone: true,
  imports: [CommonModule, RouterLink, ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule],
  templateUrl: "./home.page.html",
  styleUrl: "./home.page.scss",
})
export class HomePage implements OnDestroy {
  private api = inject(ApiService);
  private router = inject(Router);
  readonly store = inject(SessionStore);

  readonly phase = signal<Phase>("idle");
  readonly scanLine = signal("");
  readonly uploadError = signal("");
  readonly dragOver = signal(false);
  readonly pending = signal<PendingUpload | null>(null);
  readonly skip = signal<Set<string>>(new Set());
  readonly fixAll = signal<ActionState>(idle());
  readonly fixFormulas = signal<ActionState>(idle());
  readonly loop = signal<ActionState>(idle());
  readonly recalc = signal<ActionState>(idle());
  readonly loopPlan = signal<LoopPlan | null>(null);
  readonly busy = computed(() => this.fixAll().running || this.fixFormulas().running || this.loop().running || this.recalc().running);
  readonly recalcFact = computed(() => {
    const rc = this.store.readiness()?.recalc;
    if (!rc?.ran) return "Not recalculated yet";
    const current = this.store.currentVersion()?.id;
    const stale = rc.version_id && current && rc.version_id !== current ? " (an earlier version)" : "";
    return (rc.clean ? "Last recalculation clean" : `${rc.formula_errors ?? "?"} error cells at the last recalculation`) + stale;
  });

  readonly readinessKind = computed(() => {
    const r = this.store.readiness();
    if (!r) return "info";
    return r.state === "ready" ? "success" : r.state === "unverified" ? "warning" : "error";
  });
  readonly readinessTitle = computed(() => this.store.readiness()?.headline ?? "");
  readonly readinessDetail = computed(() => {
    const r = this.store.readiness();
    if (!r) return "";
    const n = r.next_step;
    const step = n ? [n.label, n.detail].filter(Boolean).join(": ") : "";
    const counts = [r.blocking_count ? `${r.blocking_count} blocking` : "", r.manual_blocking_count ? `${r.manual_blocking_count} by hand` : "", r.optional_count ? `${r.optional_count} optional` : ""].filter(Boolean).join(", ");
    return [step, counts].filter(Boolean).join(" · ");
  });
  readonly formulaOps = computed<Operation[]>(() => this.store.plan().filter((a) => FORMULA_ACTION_IDS.has(a.id)).flatMap((a) => a.operations));
  readonly formulaCells = computed(() => this.store.plan().filter((a) => FORMULA_ACTION_IDS.has(a.id)).reduce((n, a) => n + a.count, 0));
  readonly errorCount = computed(() => this.store.statusCounts()["ERROR"] ?? 0);
  readonly downloadUrl = computed(() => {
    const id = this.store.sessionId();
    const v = this.store.currentVersion();
    return id && v ? this.api.downloadUrl(id, v.file_name) : null;
  });

  private timers: ReturnType<typeof setTimeout>[] = [];

  constructor() {
    // The session may arrive after this page is built (a `?session=<id>` link is resumed by the
    // shell asynchronously): follow the store rather than reading it once.
    effect(() => {
      const loaded = this.store.loaded();
      const phase = this.phase();
      if (loaded && phase === "idle") {
        this.phase.set("loaded");
        void this.loadLoopPlan();
      } else if (!loaded && phase === "loaded") {
        this.phase.set("idle");
        this.loopPlan.set(null);
      }
    });
  }

  ngOnDestroy(): void {
    this.timers.forEach(clearTimeout);
  }

  // ---- upload -----------------------------------------------------------------------------
  onFileInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) void this.uploadFile(file);
    input.value = "";
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.dragOver.set(false);
    const file = event.dataTransfer?.files?.[0];
    if (file) void this.uploadFile(file);
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.dragOver.set(true);
  }

  async uploadFile(file: File): Promise<void> {
    if (!/\.(xlsx|xlsm|xlsb)$/i.test(file.name)) {
      this.uploadError.set("Choose an .xlsx, .xlsm or .xlsb workbook.");
      return;
    }
    this.uploadError.set("");
    this.phase.set("uploading");
    this.scanLine.set("Uploading…");
    try {
      const started = await this.api.upload(file, "plan", true);
      if (isPending(started)) {
        // A big workbook with several sheets: let the user leave sheets out before the scan.
        if (started.size?.above_threshold && started.size.sheets?.length > 1) {
          this.pending.set(started);
          this.skip.set(new Set());
          this.phase.set("choose");
          return;
        }
        await this.scan(started.sessionId, []);
      } else {
        this.store.setSession(started);
        this.afterLoad();
      }
    } catch (e) {
      this.uploadError.set(errorText(e));
      this.phase.set("idle");
    }
  }

  async scan(id: string, ignoreSheets: string[]): Promise<void> {
    this.phase.set("scanning");
    this.scanLine.set("Reading the workbook…");
    const poll = this.pollScan(id);
    try {
      const full = await this.api.analyze(id, ignoreSheets);
      this.store.setSession(full);
      this.afterLoad();
    } catch (e) {
      this.uploadError.set(errorText(e));
      this.phase.set("idle");
    } finally {
      poll.stop();
    }
  }

  private afterLoad(): void {
    this.phase.set("loaded");
    this.pending.set(null);
    this.fixAll.set(idle());
    this.fixFormulas.set(idle());
    this.loop.set(idle());
    this.recalc.set(idle());
    void this.router.navigate([], { queryParams: { session: this.store.sessionId() }, replaceUrl: true });
    void this.loadLoopPlan();
  }

  toggleSkip(name: string): void {
    const next = new Set(this.skip());
    if (next.has(name)) next.delete(name);
    else next.add(name);
    this.skip.set(next);
  }

  mb(bytes: number): string {
    return (bytes / 1048576).toFixed(1) + " MB";
  }

  cancelChoose(): void {
    this.pending.set(null);
    this.phase.set("idle");
  }

  private pollScan(id: string): { stop: () => void } {
    let live = true;
    const tick = async () => {
      if (!live) return;
      try {
        const s = await this.api.scanStatus(id);
        const stage = s.stages?.find((x) => x.state === "running" || x.state === "current");
        if (live) this.scanLine.set(stage ? `Reading the workbook: ${stage.label}` : s.overall || "Reading the workbook…");
      } catch {
        /* keep the last line */
      }
      if (live) this.timers.push(setTimeout(tick, 700));
    };
    void tick();
    return { stop: () => (live = false) };
  }

  loadAnother(): void {
    if (this.busy()) return;
    this.store.reset();
    this.loopPlan.set(null);
    this.phase.set("idle");
    void this.router.navigate([], { queryParams: {}, replaceUrl: true });
  }

  // ---- fix all issues -------------------------------------------------------------------------
  async runFixAll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.busy()) return;
    this.fixAll.set({ running: true, fraction: 0, line: "Starting…", detail: "", tone: "idle" });
    try {
      const status = await this.api.startAutoFix(id, { use_assistant: true, keep_good_values: true });
      this.showAutoFix(status);
      await this.followAutoFix(id);
    } catch (e) {
      this.fixAll.set({ running: false, fraction: null, line: "Could not start", detail: errorText(e), tone: "error" });
    }
  }

  async stopFixAll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      await this.api.stopAutoFix(id);
      this.fixAll.update((s) => ({ ...s, line: "Stopping after the current round…" }));
    } catch (e) {
      this.fixAll.update((s) => ({ ...s, detail: errorText(e) }));
    }
  }

  private showAutoFix(s: { stage?: string; title?: string; message?: string; fraction?: number; pass_no?: number; errors?: number; errors_before?: number }): void {
    const line = s.title || s.stage || "Working…";
    const detail = [s.pass_no ? `pass ${s.pass_no}` : "", s.errors_before != null && s.errors != null ? `${s.errors_before} → ${s.errors} error cells` : "", s.message || ""]
      .filter(Boolean)
      .join(" · ");
    this.fixAll.set({ running: true, fraction: s.fraction ?? null, line, detail, tone: "idle" });
  }

  private async followAutoFix(id: string): Promise<void> {
    for (;;) {
      await this.sleep(1500);
      let s;
      try {
        s = await this.api.autoFixStatus(id, false);
      } catch (e) {
        this.fixAll.set({ running: false, fraction: null, line: "Lost the fixer", detail: errorText(e), tone: "error" });
        return;
      }
      if (s.state === "running") {
        this.showAutoFix(s);
        continue;
      }
      const full = await this.api.autoFixStatus(id, true).catch(() => s);
      if (full.analysis) this.store.applyAnalysis({ ...full.analysis, version: full.version ?? undefined });
      else if (full.version) this.store.addVersion(full.version);
      const r = full.result;
      if (s.state === "error" || r?.status === "error") {
        this.fixAll.set({ running: false, fraction: null, line: "Stopped with an error", detail: full.error || r?.summary || "", tone: "error" });
      } else {
        const before = r?.errors_before ?? full.errors_before;
        const after = r?.errors_after ?? full.errors;
        const tone = r?.status === "clean" ? "ok" : "warn";
        const line = r?.status === "clean" ? "Every formula error is fixed" : r?.status === "unchanged" ? "Nothing could be changed" : `Fixed what it could (${r?.status ?? "partial"})`;
        const left = r?.left ?? [];
        const leftText = left.length ? `${left.length} left for a person: ` + left.slice(0, 3).map((l) => `${l.sheet}!${l.cell} ${l.error}`).join(", ") + (left.length > 3 ? ", …" : "") + " (Advanced › Recalculate)" : "";
        const detail = [before != null && after != null ? `${before} → ${after} error cells` : "", r?.cells_fixed != null ? `${r.cells_fixed} cells rewritten` : "", full.version ? full.version.label : "", leftText].filter(Boolean).join(" · ");
        this.fixAll.set({ running: false, fraction: null, line, detail, tone });
      }
      return;
    }
  }

  // ---- fix the formulas ----------------------------------------------------------------------
  async runFixFormulas(): Promise<void> {
    const id = this.store.sessionId();
    const ops = this.formulaOps();
    if (!id || this.busy()) return;
    if (!ops.length) {
      this.fixFormulas.set({ running: false, fraction: null, line: "No formula to change", detail: "The current version has no formula the Mind rules would rewrite.", tone: "ok" });
      return;
    }
    this.fixFormulas.set({ running: true, fraction: null, line: `Writing ${ops.length} changes through Excel…`, detail: "", tone: "idle" });
    const stop = this.pollApply(id, this.fixFormulas);
    try {
      const out = await this.api.apply(id, ops, "major");
      stop();
      this.store.applyAnalysis({ ...out, version: out.version });
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      this.fixFormulas.set({
        running: false,
        fraction: null,
        line: refused ? `${done} changed, ${refused} refused` : `${done} formulas changed`,
        detail: [out.version?.label ?? "", r.verified_opens_in_excel ? "opens in Excel" : "", failedText(r)].filter(Boolean).join(" · "),
        tone: refused ? "warn" : "ok",
      });
    } catch (e) {
      stop();
      this.fixFormulas.set({ running: false, fraction: null, line: "Could not write the changes", detail: errorText(e), tone: "error" });
    }
  }

  private pollApply(id: string, target: typeof this.fixFormulas): () => void {
    let live = true;
    const tick = async () => {
      if (!live) return;
      try {
        const s = await this.api.scanStatus(id);
        const a = s.apply;
        if (live && a && a.state === "running") target.update((x) => ({ ...x, fraction: a.fraction ?? null, detail: [a.title, a.message].filter(Boolean).join(" · ") }));
      } catch {
        /* ignore */
      }
      if (live) this.timers.push(setTimeout(tick, 600));
    };
    void tick();
    return () => (live = false);
  }

  // ---- the Mind loop -------------------------------------------------------------------------
  private async loadLoopPlan(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      const { plan } = await this.api.mindLoopPlan(id);
      this.loopPlan.set(plan);
    } catch {
      this.loopPlan.set(null);
    }
  }

  loopSummary(): string {
    const p = this.loopPlan();
    if (!p) return "Reading the model map…";
    if (!p.found) return "No run loop in this workbook: nothing to build.";
    const ready = p.loops.filter((l) => l.status === "ready");
    const parts = ready.map((l) => `${l.name} (${l.first} to ${l.last})`);
    const open = p.loops.filter((l) => l.status !== "ready").length;
    return [parts.length ? `Loop ${parts.join(", ")}: ${p.operation_count} formulas to write` : "", open ? `${open} loop(s) need a decision (see Advanced)` : ""].filter(Boolean).join(" · ");
  }

  canBuildLoop(): boolean {
    const p = this.loopPlan();
    return !!p && p.found && p.operation_count > 0 && p.loops.some((l) => l.status === "ready");
  }

  async runLoop(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.busy() || !this.canBuildLoop()) return;
    this.loop.set({ running: true, fraction: null, line: "Writing the loop through Excel…", detail: "", tone: "idle" });
    const stop = this.pollApply(id, this.loop);
    try {
      const out = await this.api.applyMindLoopPlan(id, {});
      stop();
      const { plan: loopPlan, ...analysis } = out; // `plan` here is the LOOP plan, not the Prep actions
      this.store.applyAnalysis({ ...analysis, version: out.version });
      this.loopPlan.set(loopPlan);
      void this.api.resume(id).then((s) => this.store.applyAnalysis({ summary: s.summary, report: s.report, plan: s.plan, readiness: s.readiness, prep_progress: s.prep_progress, versions: s.versions })).catch(() => {});
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      this.loop.set({
        running: false,
        fraction: null,
        line: refused ? `${done} written, ${refused} refused` : `${done} loop formulas written`,
        detail: [out.version?.label ?? "", out.plan.how_it_runs ?? "", failedText(r)].filter(Boolean).join(" · "),
        tone: refused ? "warn" : "ok",
      });
    } catch (e) {
      stop();
      this.loop.set({ running: false, fraction: null, line: "Could not write the loop", detail: errorText(e), tone: "error" });
    }
  }

  // ---- recalculate ------------------------------------------------------------------------------
  async runRecalc(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.busy()) return;
    this.recalc.set({ running: true, fraction: null, line: "Recalculating in Excel…", detail: "Tens of seconds to a few minutes.", tone: "idle" });
    try {
      const out = await this.api.recalculate(id);
      this.store.recalc.set(out);
      this.store.applyAnalysis({ report: out.report, plan: out.plan, readiness: out.readiness, version: out.version ?? undefined });
      const tone = out.status === "PASS" ? "ok" : out.status === "ERROR" ? "warn" : "warn";
      const line = out.status === "PASS" ? "Recalculated, no formula errors" : `Recalculated, ${out.formula_errors.length} formula errors`;
      const detail = [out.message, out.new_errors != null ? `${out.new_errors} new since the original` : "", out.preexisting_errors != null ? `${out.preexisting_errors} already in the original` : "", out.addin_gap_errors.length ? `${out.addin_gap_errors.length} add-in gap` : "", "The list is under Advanced › Recalculate."].filter(Boolean).join(" · ");
      this.recalc.set({ running: false, fraction: null, line, detail, tone });
    } catch (e) {
      this.recalc.set({ running: false, fraction: null, line: "Could not recalculate", detail: errorText(e), tone: "error" });
    }
  }

  toggle(which: "fixAll" | "fixFormulas" | "loop" | "recalc"): void {
    this[which].update((s) => ({ ...s, open: !s.open }));
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((res) => this.timers.push(setTimeout(res, ms)));
  }
}
