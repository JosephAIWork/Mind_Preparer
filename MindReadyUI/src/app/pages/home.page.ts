import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, effect, inject, signal } from "@angular/core";
import { Router, RouterLink } from "@angular/router";
import { ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule } from "carbon-components-angular";
import { ApiService, errorText } from "../core/api.service";
import { FORMULA_ACTION_IDS, LoopPlan, Operation, PendingUpload, ScanStatus, appliedCount, failedCount, failedText, isPending } from "../core/models";
import { SessionStore } from "../core/session.store";

type Phase = "idle" | "uploading" | "choose" | "scanning" | "loaded";
type Tone = "idle" | "ok" | "warn" | "error";
interface ActionState {
  running: boolean;
  fraction: number | null;
  line: string;
  detail: string;
  tone: Tone;
  /** the result text is shown only on "View details" */
  open?: boolean;
  /** set after Stop: the fixer finishes the round it is in */
  stopping?: boolean;
}

const idle = (): ActionState => ({ running: false, fraction: null, line: "", detail: "", tone: "idle" });

/** The scan's stages, in order, as the backend names them (app/progress.py). */
const SCAN_STAGES: { id: string; title: string }[] = [
  { id: "convert", title: "Converting .xlsb through Excel" },
  { id: "copy", title: "Copying the workbook" },
  { id: "load", title: "Reading the workbook" },
  { id: "inventory", title: "Scanning cells and grids" },
  { id: "names", title: "Reading defined names" },
  { id: "rules", title: "Checking the rules" },
  { id: "report", title: "Building the report" },
  { id: "plan", title: "Planning the fixes" },
];

interface ScanRow {
  id: string;
  title: string;
  state: "done" | "current" | "waiting";
  seconds?: number;
  message?: string;
}

type StepId = "fixAll" | "prepare" | "loop" | "recalc";
interface Step {
  id: StepId;
  label: string;
  check: string;
  state: "waiting" | "running" | "done" | "failed" | "skipped";
  line: string;
  detail: string;
  open?: boolean;
}

const STEPS: { id: StepId; label: string; check: string }[] = [
  { id: "fixAll", label: "Fix all issues", check: "the fixer ends clean and nothing blocks the file" },
  { id: "prepare", label: "Fix the formulas and the structure", check: "every blocking action applied, nothing refused, none left" },
  { id: "loop", label: "Build the Mind loop", check: "nothing refused (skipped when the model has no loop)" },
  { id: "recalc", label: "Recalculate", check: "no formula error that was not already in the original" },
];

/** The main screen: one workbook, the buttons that do the work on their own, one download. */
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
  readonly scan = signal<ScanStatus | null>(null);
  readonly uploadError = signal("");
  readonly dragOver = signal(false);
  readonly pending = signal<PendingUpload | null>(null);
  readonly skip = signal<Set<string>>(new Set());
  readonly fixAll = signal<ActionState>(idle());
  readonly fixFormulas = signal<ActionState>(idle());
  readonly loop = signal<ActionState>(idle());
  readonly recalc = signal<ActionState>(idle());
  readonly prepare = signal<ActionState>(idle());
  readonly loopPlan = signal<LoopPlan | null>(null);
  readonly steps = signal<Step[]>([]);
  readonly chainRunning = signal(false);
  readonly chainLine = signal("");
  readonly busy = computed(() => this.chainRunning() || this.fixAll().running || this.fixFormulas().running || this.prepare().running || this.loop().running || this.recalc().running);
  /** Which step holds the buttons, for the hint on a disabled button. */
  readonly busyHint = computed(() => {
    if (this.chainRunning()) return "Do everything is running";
    if (this.fixAll().running) return this.fixAll().stopping ? "Fix all issues is stopping after its round" : "Fix all issues is running";
    if (this.fixFormulas().running) return "Fix the formulas is running";
    if (this.prepare().running) return "Fix the formulas and the structure is running";
    if (this.loop().running) return "Build the Mind loop is running";
    if (this.recalc().running) return "Recalculate is running";
    return "";
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
  readonly recalcFact = computed(() => {
    const rc = this.store.readiness()?.recalc;
    if (!rc?.ran) return "Not recalculated yet";
    const current = this.store.currentVersion()?.id;
    const stale = rc.version_id && current && rc.version_id !== current ? " (an earlier version)" : "";
    return (rc.clean ? "Last recalculation clean" : `${rc.formula_errors ?? "?"} error cells at the last recalculation`) + stale;
  });
  readonly downloadUrl = computed(() => {
    const id = this.store.sessionId();
    const v = this.store.currentVersion();
    return id && v ? this.api.downloadUrl(id, v.file_name) : null;
  });

  // ---- the scan, as rows -------------------------------------------------------------------------
  readonly scanRows = computed<ScanRow[]>(() => {
    const s = this.scan();
    const done = new Map((s?.stages ?? []).map((x) => [x.stage, x.seconds]));
    const current = s?.stage ?? null;
    const needsConvert = (s?.stages ?? []).some((x) => x.stage === "convert") || current === "convert";
    return SCAN_STAGES.filter((st) => st.id !== "convert" || needsConvert).map((st) => ({
      id: st.id,
      title: st.title,
      state: done.has(st.id) ? "done" : st.id === current ? "current" : "waiting",
      seconds: done.get(st.id),
      message: st.id === current ? s?.message || undefined : undefined,
    }));
  });
  readonly scanPercent = computed(() => Math.round(Math.max(0, Math.min(1, this.scan()?.overall ?? 0)) * 100));
  readonly scanElapsed = computed(() => {
    const e = this.scan()?.elapsed_s ?? 0;
    return e >= 60 ? `${Math.floor(e / 60)} min ${Math.round(e % 60)} s` : `${Math.round(e)} s`;
  });

  private timers: ReturnType<typeof setTimeout>[] = [];
  private fixAllFollowing = false;

  constructor() {
    // The session may arrive after this page is built (a `?session=<id>` link is resumed by the
    // shell asynchronously): follow the store rather than reading it once.
    effect(() => {
      const loaded = this.store.loaded();
      const phase = this.phase();
      if (loaded && phase === "idle") {
        this.phase.set("loaded");
        void this.loadLoopPlan();
        void this.resumeFixer();
      } else if (!loaded && phase === "loaded") {
        this.phase.set("idle");
        this.loopPlan.set(null);
      }
    });
  }

  ngOnDestroy(): void {
    this.timers.forEach(clearTimeout);
    this.fixAllFollowing = false;
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
    this.scan.set(null);
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
        await this.scanSession(started.sessionId, []);
      } else {
        this.store.setSession(started);
        this.afterLoad();
      }
    } catch (e) {
      this.uploadError.set(errorText(e));
      this.phase.set("idle");
    }
  }

  async scanSession(id: string, ignoreSheets: string[]): Promise<void> {
    this.phase.set("scanning");
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
    this.prepare.set(idle());
    this.steps.set([]);
    this.chainLine.set("");
    void this.router.navigate([], { queryParams: { session: this.store.sessionId() }, replaceUrl: true });
    void this.loadLoopPlan();
  }

  private pollScan(id: string): { stop: () => void } {
    let live = true;
    const tick = async () => {
      if (!live) return;
      try {
        const s = await this.api.scanStatus(id);
        if (live) this.scan.set(s);
      } catch {
        /* keep the last status */
      }
      if (live) this.timers.push(setTimeout(tick, 700));
    };
    void tick();
    return { stop: () => (live = false) };
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

  loadAnother(): void {
    if (this.busy()) return;
    this.store.reset();
    this.loopPlan.set(null);
    this.phase.set("idle");
    void this.router.navigate([], { queryParams: {}, replaceUrl: true });
  }

  // ---- fix all issues -------------------------------------------------------------------------
  /** A fixer may still be running from before a reload or a visit to Advanced: pick it up. */
  private async resumeFixer(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.fixAllFollowing) return;
    try {
      const s = await this.api.autoFixStatus(id, false);
      if (s.state === "running") {
        this.showAutoFix(s);
        await this.followAutoFix(id);
      }
    } catch {
      /* no fixer, nothing to resume */
    }
  }

  async runFixAll(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.fixAll().running) return false;
    this.fixAll.set({ running: true, fraction: 0, line: "Starting…", detail: "", tone: "idle" });
    try {
      const status = await this.api.startAutoFix(id, { use_assistant: true, keep_good_values: true });
      this.showAutoFix(status);
      return await this.followAutoFix(id);
    } catch (e) {
      // The backend says why in one sentence (already running, a recalculation runs, being analysed).
      const why = errorText(e);
      if (/already running/i.test(why)) {
        this.fixAll.set({ running: true, fraction: null, line: "Still running from before", detail: why, tone: "idle" });
        return await this.followAutoFix(id);
      }
      this.fixAll.set({ running: false, fraction: null, line: "Not started", detail: why, tone: "error" });
      return false;
    }
  }

  async stopFixAll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      await this.api.stopAutoFix(id);
      this.fixAll.update((s) => ({ ...s, stopping: true, line: "Stopping after this round…" }));
    } catch (e) {
      this.fixAll.update((s) => ({ ...s, detail: errorText(e) }));
    }
  }

  private showAutoFix(s: { stage?: string; title?: string; message?: string; fraction?: number; pass_no?: number; errors?: number; errors_before?: number }): void {
    const stopping = this.fixAll().stopping;
    const line = stopping ? "Stopping after this round…" : s.title || s.stage || "Working…";
    const detail = [s.pass_no ? `pass ${s.pass_no}` : "", s.errors_before != null && s.errors != null ? `${s.errors_before} → ${s.errors} error cells` : "", s.message || ""]
      .filter(Boolean)
      .join(" · ");
    this.fixAll.set({ running: true, fraction: s.fraction ?? null, line, detail, tone: "idle", stopping });
  }

  /** Follows the fixer to its end. Returns true when it ended clean. */
  private async followAutoFix(id: string): Promise<boolean> {
    if (this.fixAllFollowing) return false;
    this.fixAllFollowing = true;
    try {
      for (;;) {
        await this.sleep(1500);
        let s;
        try {
          s = await this.api.autoFixStatus(id, false);
        } catch (e) {
          this.fixAll.set({ running: false, fraction: null, line: "Lost the fixer", detail: errorText(e), tone: "error" });
          return false;
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
          return false;
        }
        const before = r?.errors_before ?? full.errors_before;
        const after = r?.errors_after ?? full.errors;
        const left = r?.left ?? [];
        // "unchanged" with no error left: the workbook already recalculated clean, nothing to fix
        const clean = r?.status === "clean" || ((r?.status === "partial" || r?.status === "unchanged") && left.length === 0 && (after ?? 0) === 0);
        const line = r?.status === "clean" ? "Every formula error is fixed" : r?.status === "unchanged" ? (clean ? "No formula error to fix" : "Nothing could be changed") : `Fixed what it could (${r?.status ?? "partial"})`;
        const leftText = left.length ? `${left.length} left for a person: ` + left.slice(0, 3).map((l) => `${l.sheet}!${l.cell} ${l.error}`).join(", ") + (left.length > 3 ? ", …" : "") + " (Advanced › Recalculate)" : "";
        const detail = [before != null && after != null ? `${before} → ${after} error cells` : "", r?.cells_fixed != null ? `${r.cells_fixed} cells rewritten` : "", full.version ? full.version.label : "", leftText].filter(Boolean).join(" · ");
        this.fixAll.set({ running: false, fraction: null, line, detail, tone: clean ? "ok" : "warn" });
        return clean;
      }
    } finally {
      this.fixAllFollowing = false;
    }
  }

  // ---- fix the formulas ----------------------------------------------------------------------
  async runFixFormulas(): Promise<boolean> {
    const id = this.store.sessionId();
    const ops = this.formulaOps();
    if (!id || this.fixFormulas().running) return false;
    if (!ops.length) {
      this.fixFormulas.set({ running: false, fraction: null, line: "No formula to change", detail: "The current version has no formula the Mind rules would rewrite.", tone: "ok" });
      return true;
    }
    this.fixFormulas.set({ running: true, fraction: null, line: `Writing ${ops.length} changes through Excel…`, detail: "", tone: "idle" });
    const stop = this.pollApply(id, this.fixFormulas);
    const errorsBefore = this.errorCount();
    try {
      const out = await this.api.apply(id, ops, "major");
      stop();
      this.store.applyAnalysis({ ...out, version: out.version });
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      const newErrors = Math.max(0, this.errorCount() - errorsBefore);
      const ok = refused === 0 && newErrors === 0;
      this.fixFormulas.set({
        running: false,
        fraction: null,
        line: refused ? `${done} changed, ${refused} refused` : `${done} formulas changed`,
        detail: [out.version?.label ?? "", r.verified_opens_in_excel ? "opens in Excel" : "", failedText(r), newErrors ? `${newErrors} new error findings` : ""].filter(Boolean).join(" · "),
        tone: ok ? "ok" : "warn",
      });
      return ok;
    } catch (e) {
      stop();
      this.fixFormulas.set({ running: false, fraction: null, line: "Could not write the changes", detail: errorText(e), tone: "error" });
      return false;
    }
  }

  // ---- fix the formulas and the structure (the chain's second step) --------------------------------
  /** Every blocking Prep action that is on by default, in rounds until none is left (3 at most). */
  async runPrepare(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.prepare().running) return false;
    this.prepare.set({ running: true, fraction: null, line: "Looking at what blocks the file…", detail: "", tone: "idle" });
    let applied = 0;
    let refused = 0;
    const names: string[] = [];
    try {
      for (let round = 1; round <= 3; round++) {
        const acts = this.store.plan().filter((a) => a.level === "blocking" && a.default_on && a.count > 0);
        const ops = acts.flatMap((a) => a.operations);
        if (!ops.length) break;
        names.push(...acts.map((a) => a.title));
        this.prepare.update((s) => ({ ...s, line: `Round ${round}: writing ${ops.length} changes through Excel…`, detail: acts.map((a) => a.title).join(", ") }));
        const stop = this.pollApply(id, this.prepare);
        let out;
        try {
          out = await this.api.apply(id, ops, "major");
        } finally {
          stop();
        }
        this.store.applyAnalysis({ ...out, version: out.version });
        applied += appliedCount(out.result);
        refused += failedCount(out.result);
        if (failedCount(out.result)) {
          this.prepare.set({ running: false, fraction: null, line: `${applied} changed, ${refused} refused`, detail: [out.version?.label ?? "", failedText(out.result)].filter(Boolean).join(" · "), tone: "warn" });
          return false;
        }
      }
      const left = this.store.plan().filter((a) => a.level === "blocking" && a.default_on && a.count > 0);
      const ok = left.length === 0;
      this.prepare.set({
        running: false,
        fraction: null,
        line: applied ? `${applied} changes applied` : "Nothing blocked the file",
        detail: [this.store.currentVersion()?.label ?? "", names.length ? [...new Set(names)].join(", ") : "", left.length ? `still blocking: ${left.map((a) => a.title).join(", ")}` : ""].filter(Boolean).join(" · "),
        tone: ok ? "ok" : "warn",
      });
      return ok;
    } catch (e) {
      this.prepare.set({ running: false, fraction: null, line: "Could not write the changes", detail: errorText(e), tone: "error" });
      return false;
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
    return [parts.length ? `Loop ${parts.join(", ")}: ${p.operation_count} formulas to write` : "", open ? `${open} loop(s) need a decision (see Advanced)` : "", this.headerRow() ? "row 1 is a header row: skipped" : ""].filter(Boolean).join(" · ");
  }

  /** The plan says the first row of the table is a header row: the loop then starts at 2. */
  private headerRow(): boolean {
    const p = this.loopPlan();
    return !!p && p.loops.some((l) => l.first === 1 && (l.notes ?? []).some((n) => /header/i.test(n)));
  }

  canBuildLoop(): boolean {
    const p = this.loopPlan();
    return !!p && p.found && p.operation_count > 0 && p.loops.some((l) => l.status === "ready");
  }

  async runLoop(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.loop().running || !this.canBuildLoop()) return false;
    this.loop.set({ running: true, fraction: null, line: "Writing the loop through Excel…", detail: "", tone: "idle" });
    const stop = this.pollApply(id, this.loop);
    try {
      const out = await this.api.applyMindLoopPlan(id, this.headerRow() ? { first: "2" } : {});
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
        detail: [out.version?.label ?? "", loopPlan.how_it_runs ?? "", failedText(r)].filter(Boolean).join(" · "),
        tone: refused ? "warn" : "ok",
      });
      return refused === 0;
    } catch (e) {
      stop();
      this.loop.set({ running: false, fraction: null, line: "Could not write the loop", detail: errorText(e), tone: "error" });
      return false;
    }
  }

  // ---- recalculate ------------------------------------------------------------------------------
  async runRecalc(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.recalc().running) return false;
    this.recalc.set({ running: true, fraction: null, line: "Recalculating in Excel…", detail: "Tens of seconds to a few minutes.", tone: "idle" });
    try {
      const out = await this.api.recalculate(id);
      this.store.recalc.set(out);
      this.store.applyAnalysis({ report: out.report, plan: out.plan, readiness: out.readiness, version: out.version ?? undefined });
      const newErrors = out.new_errors ?? out.formula_errors.length;
      const ok = out.status === "PASS" || newErrors === 0;
      const line = out.status === "PASS" ? "Recalculated, no formula errors" : `Recalculated, ${out.formula_errors.length} formula errors`;
      const detail = [out.message, out.new_errors != null ? `${out.new_errors} new since the original` : "", out.preexisting_errors != null ? `${out.preexisting_errors} already in the original` : "", out.addin_gap_errors.length ? `${out.addin_gap_errors.length} add-in gap` : "", "The list is under Advanced › Recalculate."].filter(Boolean).join(" · ");
      this.recalc.set({ running: false, fraction: null, line, detail, tone: ok ? "ok" : "warn" });
      return ok;
    } catch (e) {
      this.recalc.set({ running: false, fraction: null, line: "Could not recalculate", detail: errorText(e), tone: "error" });
      return false;
    }
  }

  // ---- do everything --------------------------------------------------------------------------
  /** The whole chain, each step followed by its check, stopping at the first failed check. */
  async runAll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.busy()) return;
    this.chainRunning.set(true);
    this.chainLine.set("");
    this.steps.set(STEPS.map((s) => ({ ...s, state: "waiting", line: "", detail: "" })));
    const runners: Record<StepId, () => Promise<boolean>> = {
      fixAll: () => this.runFixAll(),
      prepare: () => this.runPrepare(),
      loop: () => this.runLoop(),
      recalc: () => this.runRecalc(),
    };
    const states: Record<StepId, typeof this.fixAll> = { fixAll: this.fixAll, prepare: this.prepare, loop: this.loop, recalc: this.recalc };
    try {
      for (const step of STEPS) {
        if (step.id === "loop" && !this.canBuildLoop()) {
          this.setStep(step.id, { state: "skipped", line: this.loopPlan()?.found ? "The loop needs a decision first (Advanced › Mind loop)" : "No run loop in this workbook" });
          continue;
        }
        this.setStep(step.id, { state: "running" });
        let ok = false;
        try {
          ok = await runners[step.id]();
        } catch (e) {
          states[step.id].set({ running: false, fraction: null, line: "Failed", detail: errorText(e), tone: "error" });
        }
        const r = states[step.id]();
        this.setStep(step.id, { state: ok ? "done" : "failed", line: r.line, detail: r.detail });
        if (!ok) {
          this.chainLine.set(`Stopped at "${step.label}": ${r.line}. Fix that under Advanced, then press Do everything again.`);
          return;
        }
      }
      this.chainLine.set("Everything ran and every check passed.");
    } finally {
      this.chainRunning.set(false);
    }
  }

  private setStep(id: StepId, patch: Partial<Step>): void {
    this.steps.update((list) => list.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  }

  toggleStep(id: StepId): void {
    this.setStep(id, { open: !this.steps().find((s) => s.id === id)?.open });
  }

  chainStep(): Step | null {
    return this.steps().find((s) => s.state === "running") ?? null;
  }

  stateOf(id: StepId): ActionState {
    return this[id]();
  }

  toggle(which: StepId | "fixFormulas"): void {
    this[which].update((s) => ({ ...s, open: !s.open }));
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((res) => this.timers.push(setTimeout(res, ms)));
  }
}
