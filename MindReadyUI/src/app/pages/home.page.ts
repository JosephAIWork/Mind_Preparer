import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, effect, inject, signal } from "@angular/core";
import { Router, RouterLink } from "@angular/router";
import { ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule } from "carbon-components-angular";
import { ApiService, errorText } from "../core/api.service";
import {
  ChainEstimate, ChainStepId, FORMULA_ACTION_IDS, InputFile, LoopPlan, Operation, OutsideInputRow, OutsideInputs, OutsideRunResult, OutsideRunStatus, PICK_ACCEPT, PendingUpload, ScanStatus,
  appliedCount, failedCount, failedText, firstClause, inputFilesProblem, isPending, nameList, outsideButtonLabel, outsideOutcome, outsideReasonLine, outsideRowLabel, outsideSkipReason,
  pickedFilesProblem, withOutsideStep,
} from "../core/models";
import { SessionStore } from "../core/session.store";
import { KeyBox } from "./advanced/key-box";

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

/** The "Do everything right after the upload" box, remembered in this browser. */
const AUTO_RUN_KEY = "mindReady.autoRun";

function readAutoRun(): boolean {
  try {
    return localStorage.getItem(AUTO_RUN_KEY) === "1";
  } catch {
    return false; // storage blocked: the box starts unticked
  }
}

/** "under a minute", "12 min", "1 h 25 min" */
function minutes(seconds: number): string {
  if (seconds < 60) return "under a minute";
  const m = Math.round(seconds / 60);
  return m < 90 ? `${m} min` : `${Math.floor(m / 60)} h ${m % 60} min`;
}

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

type StepId = ChainStepId;
interface Step {
  id: StepId;
  label: string;
  check: string;
  /** stopped: amber -- the step did its work, but a check could not run (the user goes on on purpose) */
  state: "waiting" | "running" | "done" | "failed" | "stopped" | "skipped";
  line: string;
  detail: string;
  open?: boolean;
}

/** The chain's steps in order; "outside" only when the model's macros open outside files. */
const STEPS: { id: StepId; label: string; check: string }[] = [
  { id: "outside", label: "Bring in the outside files", check: "every input file matched and read by the model's own import macro" },
  { id: "fixAll", label: "Fix all issues", check: "the fixer ends clean, or every error left was already in the workbook (none caused by the preparation)" },
  { id: "prepare", label: "Fix the formulas and the structure", check: "every blocking action applied, nothing refused, none left" },
  { id: "loop", label: "Build the Mind loop", check: "nothing refused (skipped when the model has no loop)" },
  { id: "recalc", label: "Recalculate", check: "no formula error that was not already in the original" },
];

/** The main screen: one workbook, the buttons that do the work on their own, one download. */
@Component({
  selector: "mr-home",
  standalone: true,
  imports: [CommonModule, RouterLink, ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule, KeyBox],
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
  /** Words for the uploading screen: one workbook, or the model with its input files. */
  readonly sending = signal("");
  readonly dragOver = signal(false);
  readonly pending = signal<PendingUpload | null>(null);
  readonly skip = signal<Set<string>>(new Set());
  readonly fixAll = signal<ActionState>(idle());
  readonly fixFormulas = signal<ActionState>(idle());
  readonly loop = signal<ActionState>(idle());
  readonly recalc = signal<ActionState>(idle());
  readonly prepare = signal<ActionState>(idle());
  /** Bring in the outside files: the model's own import macro, then one Mind input per file */
  readonly outside = signal<ActionState>(idle());
  readonly outsideResult = signal<OutsideRunResult | null>(null);
  /** the input files sent with the workbook (the upload's answer), for the note when its macros open none */
  readonly pickedInputs = signal<InputFile[]>([]);
  readonly addingInputs = signal(false);
  readonly inputsError = signal("");
  readonly cardDragOver = signal(false);
  /** the last run on this page failed (the job stopped, its check failed, it could not start): the card offers "Try again" */
  readonly outsideFailed = signal(false);
  /** the last run brought the files in but could not compare the computed values: amber, and Do everything stops there */
  readonly outsideUnchecked = signal(false);
  readonly pickAccept = PICK_ACCEPT;
  readonly loopPlan = signal<LoopPlan | null>(null);
  readonly steps = signal<Step[]>([]);
  /** Ticked: once an uploaded workbook is read, Do everything starts by itself -- and a big workbook is
   *  read whole, without stopping to ask which sheets to leave out. */
  readonly autoRun = signal(readAutoRun());
  readonly chainRunning = signal(false);
  readonly chainLine = signal("");
  /** 1.11.6: the estimate for "Do everything" (shown before the click) and the countdown while it runs */
  readonly estimate = signal<ChainEstimate | null>(null);
  private readonly clock = signal(Date.now());
  private chainT0 = 0;
  private stepT0 = 0;
  private stepSeconds: Partial<Record<StepId, number>> = {};
  private ticker: ReturnType<typeof setInterval> | null = null;
  /** Do everything starts by bringing in the outside files: the macros open some, the app knows which
   *  cells name them, and they are not in the current version yet. */
  readonly outsideWillRun = computed(() => {
    const o = this.store.outsideInputs();
    return !!o && !o.brought_in && this.canBringIn(o);
  });
  /** The server's estimate, with the outside step in it when that step runs and the server left it out. */
  readonly fullEstimate = computed(() => {
    const include = this.chainRunning()
      ? this.steps().some((s) => s.id === "outside" && (s.state === "waiting" ? this.outsideWillRun() : s.state !== "skipped"))
      : this.outsideWillRun();
    return withOutsideStep(this.estimate(), include);
  });
  readonly estimateLine = computed(() => {
    const e = this.fullEstimate();
    if (!e) return "";
    const t = e.total;
    return `Rough estimate: about ${minutes(t.mid_s)} (between ${minutes(t.low_s)} and ${minutes(t.high_s)}).`;
  });
  /** Elapsed and what is left, recomputed every 10 s: each step's likely time, scaled by how the
   *  finished steps compared with their own estimate, minus what the running step has already taken. */
  readonly timerLine = computed(() => {
    const now = this.clock();
    const e = this.fullEstimate();
    if (!this.chainRunning() || !this.chainT0) return "";
    const elapsed = `${minutes((now - this.chainT0) / 1000)} so far`;
    if (!e) return elapsed;
    const est = new Map(e.steps.map((s) => [s.id, s]));
    let planned = 0;
    let took = 0;
    for (const s of this.steps()) {
      const actual = this.stepSeconds[s.id];
      const x = est.get(s.id);
      // a step without an estimate would only push the ratio up
      if (s.state === "done" && actual !== undefined && x) {
        planned += x.mid_s;
        took += actual;
      }
    }
    const k = planned >= 60 ? Math.min(3, Math.max(0.5, took / planned)) : 1;
    let left = 0;
    let slow = false;
    for (const s of this.steps()) {
      const x = est.get(s.id);
      if (!x) continue;
      if (s.state === "waiting") left += x.mid_s * k;
      else if (s.state === "running") {
        const inStep = (now - this.stepT0) / 1000;
        left += Math.max(x.mid_s * k - inStep, 0.1 * x.mid_s * k, 30);
        slow = inStep > x.high_s * k;
      }
    }
    return `${elapsed} · about ${minutes(left)} left${slow ? " · this step is taking longer than usual" : ""}`;
  });
  readonly busy = computed(
    () => this.chainRunning() || this.outside().running || this.addingInputs() || this.fixAll().running || this.fixFormulas().running || this.prepare().running || this.loop().running || this.recalc().running,
  );
  /** Which step holds the buttons, for the hint on a disabled button. */
  readonly busyHint = computed(() => {
    if (this.chainRunning()) return "Do everything is running";
    if (this.outside().running) return "Bring in the outside files is running";
    if (this.addingInputs()) return "The input files are being added";
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
  /** What the Outside files card shows: the job's own state -- or, once the current version no longer
   *  has the files brought in by that job (the server says why: files replaced or added since, a
   *  bring-in that did not finish; or another version was chosen), that, not the old result. The
   *  server is the one source: the same after a reload. */
  readonly outsideCard = computed<ActionState>(() => {
    const s = this.outside();
    const o = this.store.outsideInputs();
    if (!o || o.brought_in || s.running || s.tone === "error") return s;
    const why = outsideReasonLine(o);
    if (why) return { running: false, fraction: null, line: why, detail: "", tone: "warn" };
    return this.outsideResult() ? idle() : s;
  });
  /** "Mind inputs.zip": one CSV per Mind input, made by the last Bring in the outside files -- offered
   *  only while the current version has them brought in (not the CSVs of files replaced since) */
  readonly zipUrl = computed(() => {
    const id = this.store.sessionId();
    const r = this.outsideResult();
    return id && this.store.outsideInputs()?.brought_in && r?.download?.zip && r.mindInputs?.inputs?.length ? this.api.downloadUrl(id, r.download.zip) : null;
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
  private outsideFollowing = false;
  private estimateSeq = 0;
  /** Set when the page is left (Advanced): a running Do everything starts no further step unseen. */
  private destroyed = false;

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
        void this.resumeOutside();
      } else if (!loaded && phase === "loaded") {
        this.phase.set("idle");
        this.loopPlan.set(null);
      }
    });
    // the estimate follows the analysed version and the loop plan (refreshed once a step is done)
    effect(() => {
      const id = this.store.sessionId();
      const version = this.store.currentVersion()?.id;
      this.loopPlan(); // whether the loop step runs changes the estimate
      if (id && version && !this.chainRunning()) void this.loadEstimate(id);
    });
  }

  private async loadEstimate(id: string): Promise<void> {
    // only the latest request counts: an older answer (made before the loop plan arrived) may come last
    const seq = ++this.estimateSeq;
    try {
      const e = await this.api.estimate(id, this.canBuildLoop());
      if (seq === this.estimateSeq) this.estimate.set(e);
    } catch {
      if (seq === this.estimateSeq) this.estimate.set(null); // an older server without the estimate: nothing shown
    }
  }

  ngOnDestroy(): void {
    this.destroyed = true;
    this.timers.forEach(clearTimeout);
    if (this.ticker) clearInterval(this.ticker);
    this.fixAllFollowing = false;
  }

  // ---- upload -----------------------------------------------------------------------------
  onFileInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    const files = Array.from(input.files ?? []);
    input.value = "";
    if (files.length) void this.uploadFiles(files);
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.dragOver.set(false);
    const files = Array.from(event.dataTransfer?.files ?? []);
    if (files.length) void this.uploadFiles(files);
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.dragOver.set(true);
  }

  /** One workbook, or the model workbook together with its input files (the server finds the model). */
  async uploadFiles(files: File[]): Promise<void> {
    const problem = pickedFilesProblem(files.map((f) => f.name));
    if (problem) {
      this.uploadError.set(problem);
      return;
    }
    this.uploadError.set("");
    this.sending.set(files.length === 1 ? "Sending the file to the app…" : `Sending the ${files.length} files to the app…`);
    this.phase.set("uploading");
    this.scan.set(null);
    this.pickedInputs.set([]);
    try {
      const started = await this.api.upload(files.length === 1 ? files[0] : files, "plan", true);
      this.pickedInputs.set(started.inputs ?? []);
      if (isPending(started)) {
        // A big workbook with several sheets: let the user leave sheets out before the scan
        // (not when Do everything is to follow on its own: nothing stops the run).
        if (started.size?.above_threshold && started.size.sheets?.length > 1 && !this.autoRun()) {
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
    this.resetOutside();
    this.fixAll.set(idle());
    this.fixFormulas.set(idle());
    this.loop.set(idle());
    this.recalc.set(idle());
    this.prepare.set(idle());
    this.steps.set([]);
    this.chainLine.set("");
    void this.router.navigate([], { queryParams: { session: this.store.sessionId() }, replaceUrl: true });
    const loop = this.loadLoopPlan();
    // the chain needs the loop plan first: it decides whether the loop step runs
    if (this.autoRun()) void loop.then(() => this.runAll());
  }

  setAutoRun(on: boolean): void {
    this.autoRun.set(on);
    try {
      localStorage.setItem(AUTO_RUN_KEY, on ? "1" : "0");
    } catch {
      /* storage blocked: the box holds for this visit only */
    }
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
    this.resetOutside();
    this.pickedInputs.set([]);
    this.phase.set("idle");
    void this.router.navigate([], { queryParams: {}, replaceUrl: true });
  }

  // ---- bring in the outside files (the model's macros open other files) ---------------------------
  private resetOutside(): void {
    this.outside.set(idle());
    this.outsideResult.set(null);
    this.inputsError.set("");
    this.cardDragOver.set(false);
    this.outsideFailed.set(false);
    this.outsideUnchecked.set(false);
  }

  rowLabel(r: OutsideInputRow): string {
    return outsideRowLabel(r);
  }

  /** The app can run the model's own import: it knows the macro and the cells that name the files. */
  canBringIn(o: OutsideInputs): boolean {
    return o.rows.length > 0 && !!o.macro;
  }

  /** The card's sentence: how the picked files match, or why the app cannot bring them in. */
  outsideLead(o: OutsideInputs): string {
    const cannot = "the app cannot bring them in, so their data stays inside the model.";
    if (!o.rows.length) return `${o.macro ? `The macro ${o.macro} opens` : "The workbook's macros open"} outside files, but no cell of the workbook names them: ${cannot}`;
    if (!o.macro) return `The workbook's macros open outside files, but which macro brings them in is not clear: ${cannot}`;
    return o.summary;
  }

  /** "Try again" after a failed run; with files missing, it says it brings in only the matched ones;
   *  "Bring in again" when the current line was brought in before (files changed since). */
  outsideButton(o: OutsideInputs): string {
    return outsideButtonLabel(o, this.outsideFailed());
  }

  /** Why Do everything skips the outside step ("" when it runs it): only when the server says the
   *  current line has the files brought in (a bring-in that finished, with the files picked now). It
   *  runs again after a failed or half-way bring-in, and after input files were replaced or added. */
  private outsideSkip(): string {
    return outsideSkipReason(this.store.outsideInputs(), this.outsideUnchecked());
  }

  /** Bring in the outside files: the model's own import macro runs on a copy with the matched files, then
   *  each file's area becomes a Mind input (a background job, followed every second). From Do everything
   *  (`strict`) a missing file stops the chain: add it, or bring in the others with the card's button. */
  async runOutside(strict = false): Promise<boolean> {
    const id = this.store.sessionId();
    const o = this.store.outsideInputs();
    if (!id || !o || this.outside().running) return false;
    this.outsideFailed.set(false);
    this.outsideUnchecked.set(false);
    if (!this.canBringIn(o)) {
      this.outside.set({ running: false, fraction: null, line: "The app cannot bring these files in", detail: this.outsideLead(o), tone: "error" });
      return false;
    }
    if (!o.matched) {
      this.outside.set({ running: false, fraction: null, line: "No input file matched yet", detail: "Add the model's input files on the Outside files card (Add files).", tone: "error" });
      return false;
    }
    const missing = o.rows.filter((r) => !r.file).map((r) => outsideRowLabel(r));
    if (strict && missing.length) {
      this.outside.set({
        running: false,
        fraction: null,
        line: `${missing.length} of ${o.rows.length} input file${o.rows.length === 1 ? "" : "s"} missing: ${nameList(missing)}`,
        detail: `Add ${missing.length === 1 ? "it" : "them"} on the Outside files card (Add files), or press "${this.outsideButton(o)}" there: a missing file keeps the data the model already holds.`,
        tone: "error",
      });
      return false;
    }
    this.inputsError.set("");
    this.outside.set({ running: true, fraction: null, line: "Starting…", detail: "", tone: "idle" });
    try {
      await this.api.startOutsideInputs(id);
    } catch (e) {
      // refused: it may run already (started before a reload, or in another tab): then follow that run
      const st = await this.api.outsideInputs(id).catch(() => null);
      if (!st?.running) {
        this.outside.set({ running: false, fraction: null, line: "Not started", detail: errorText(e), tone: "error" });
        this.outsideFailed.set(true);
        return false;
      }
      return await this.followOutside(id, true);
    }
    return await this.followOutside(id);
  }

  /** Follows the job to its end, a look every second. True when the files came in and the check passed.
   *  `seenRunning`: the caller saw the job running (else a status without it waits a few looks). */
  private async followOutside(id: string, seenRunning = false): Promise<boolean> {
    if (this.outsideFollowing) return false;
    this.outsideFollowing = true;
    const t0 = Date.now();
    let looks = 0;
    let failures = 0;
    try {
      for (;;) {
        if (this.destroyed) return false; // the page was left: the next one picks the job up
        let st: OutsideRunStatus;
        try {
          st = await this.api.outsideInputs(id);
          failures = 0;
        } catch (e) {
          if (++failures < 5) {
            await this.sleep(1000);
            continue;
          }
          this.outside.set({ running: false, fraction: null, line: "Lost the job", detail: errorText(e), tone: "error" });
          this.outsideFailed.set(true);
          return false;
        }
        looks++;
        if (st.outsideInputs) this.store.outsideInputs.set(st.outsideInputs);
        if (st.running) {
          seenRunning = true;
          this.showOutsideRunning(st, t0);
          await this.sleep(1000);
          continue;
        }
        // a server that shows the job a moment after it answered: a few more looks first
        if (!seenRunning && looks < 3) {
          await this.sleep(1000);
          continue;
        }
        await this.reloadSession(id); // the job's new versions, also when it stopped half-way
        if (st.error) {
          this.showOutsideError(st.error);
          return false;
        }
        if (!st.result) {
          this.outside.set({ running: false, fraction: null, line: "The job ended without a result", detail: "", tone: "error" });
          this.outsideFailed.set(true);
          return false;
        }
        return this.showOutsideResult(st.result);
      }
    } finally {
      this.outsideFollowing = false;
    }
  }

  /** The running job on the card: its stage (or its latest line, when that says the stage in full),
   *  then the latest line and the time it has taken. */
  private showOutsideRunning(st: OutsideRunStatus, t0: number): void {
    const secs = Math.round(st.elapsed_s ?? (Date.now() - t0) / 1000);
    const took = secs < 60 ? `${secs} s so far` : `${Math.floor(secs / 60)} min ${secs % 60} s so far`;
    const stage = st.stage || "Working…";
    const said = st.message && st.message !== stage ? st.message : "";
    const full = !!said && said.startsWith(stage);
    this.outside.set({ running: true, fraction: null, line: full ? said : stage, detail: [full ? "" : said, took].filter(Boolean).join(" · "), tone: "idle" });
  }

  /** The result on the card (and in the chain): the summary, then the details and the new versions. */
  private showOutsideResult(r: OutsideRunResult): boolean {
    this.outsideResult.set(r);
    const out = outsideOutcome(r);
    const labels = (r.versions ?? []).map((v) => this.store.versions().find((x) => x.id === v)?.label ?? "");
    this.outside.set({ running: false, fraction: null, line: out.line, detail: [out.detail, ...labels].filter(Boolean).join(" · "), tone: out.tone });
    this.outsideFailed.set(out.tone === "error"); // the macro stopped early, or a value changed: "Try again"
    this.outsideUnchecked.set(out.unchecked);
    return out.ok;
  }

  /** A job that stopped (the server's words: why, and which version holds the outside data): its first
   *  clause as the card's line, all of it in the details (open), and "Try again" on the card. */
  private showOutsideError(error: string): void {
    this.outside.set({ running: false, fraction: null, line: firstClause(error) || "Could not bring in the outside files", detail: error, tone: "error", open: true });
    this.outsideFailed.set(true);
    this.outsideUnchecked.set(false);
  }

  /** The job may still run from before a reload or a visit to Advanced (follow it), or its result is there to show. */
  private async resumeOutside(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.store.outsideInputs() || this.outsideFollowing) return;
    let st: OutsideRunStatus;
    try {
      st = await this.api.outsideInputs(id);
    } catch {
      return; // nothing to pick up
    }
    if (st.outsideInputs) this.store.outsideInputs.set(st.outsideInputs);
    // the card then follows the server: a result that no longer describes the current version (files
    // replaced since, another version chosen) gives way to why (outsideCard)
    if (st.running) {
      this.showOutsideRunning(st, Date.now());
      await this.followOutside(id, true);
    } else if (st.error) this.showOutsideError(st.error);
    else if (st.result) this.showOutsideResult(st.result);
  }

  /** The match and `brought_in` as the server has them now (a version restored under Advanced changes them). */
  private async refreshOutside(id: string): Promise<void> {
    try {
      const st = await this.api.outsideInputs(id);
      if (st.outsideInputs) this.store.outsideInputs.set(st.outsideInputs);
    } catch {
      /* keep what the scan said */
    }
  }

  /** After a background job made new versions: the session as the server holds it now, and the loop
   *  plan of the new current version. */
  private async reloadSession(id: string): Promise<void> {
    try {
      const s = await this.api.resume(id);
      this.store.applyAnalysis({ ...s, version: s.version });
    } catch {
      /* the next step's answer brings the new version */
    }
    await this.loadLoopPlan();
  }

  onAddInputs(event: Event): void {
    const input = event.target as HTMLInputElement;
    const files = Array.from(input.files ?? []);
    input.value = "";
    if (files.length) void this.addInputs(files);
  }

  onCardDragOver(event: DragEvent): void {
    event.preventDefault(); // always: a file dropped on the card must not open in the browser
    const o = this.store.outsideInputs();
    this.cardDragOver.set(!!o && this.canBringIn(o));
  }

  onCardDrop(event: DragEvent): void {
    event.preventDefault();
    this.cardDragOver.set(false);
    const files = Array.from(event.dataTransfer?.files ?? []);
    if (files.length) void this.addInputs(files);
  }

  /** Add files: more input files (one of the same name replaces the old one); the server matches them
   *  again. After a bring-in, a file that changes what the model reads (a corrected file, a missing one)
   *  makes the server say "not brought in" (with why): the card offers "Bring in again", Do everything
   *  runs the step again. */
  async addInputs(files: File[]): Promise<void> {
    const id = this.store.sessionId();
    const o = this.store.outsideInputs();
    if (!id || !o || !this.canBringIn(o)) return;
    if (this.busy()) {
      this.inputsError.set(`Wait: ${this.busyHint()}.`);
      return;
    }
    const problem = inputFilesProblem(files.map((f) => f.name));
    if (problem) {
      this.inputsError.set(problem);
      return;
    }
    this.inputsError.set("");
    this.addingInputs.set(true);
    try {
      const out = await this.api.addInputs(id, files);
      this.store.applyAnalysis({ outsideInputs: out.outsideInputs, readiness: out.readiness ?? undefined });
      if (this.outside().tone === "error") {
        // the last stop was about the files before
        this.outside.set(idle());
        this.outsideFailed.set(false);
      }
    } catch (e) {
      this.inputsError.set(errorText(e));
    } finally {
      this.addingInputs.set(false);
    }
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
        // 1.11.6: the errors left could not be fixed without changing a good value, and the
        // preparation caused none of them (no regression): the workbook was already living with
        // them, so the chain goes on (Recalculate then compares with the original). A Stop or a
        // used-up time budget still stops it.
        const acceptable = !clean && !!r && ["partial", "stuck", "unchanged"].includes(r.status) && !r.stopped && (r.regression_cells ?? 0) === 0;
        // On a prepared version "no regression" means something only when the comparison with the
        // original ran (none is made on the original itself): when Excel failed it, nothing was checked.
        const compared = r?.numbers_check == null || r.numbers_check.ran !== false;
        const accepted = acceptable && compared;
        const kept = after ?? left.length;
        const line = r?.status === "clean" ? "Every formula error is fixed" : accepted ? `Fixed what it could; ${kept} error${kept === 1 ? "" : "s"} the workbook already had stay, moving on` : acceptable ? "The errors left could not be checked against the original (the comparison did not run)" : r?.status === "unchanged" ? (clean ? "No formula error to fix" : "Nothing could be changed") : `Fixed what it could (${r?.status ?? "partial"})`;
        const leftText = left.length ? `${left.length} left for a person: ` + left.slice(0, 3).map((l) => `${l.sheet}!${l.cell} ${l.error}`).join(", ") + (left.length > 3 ? ", …" : "") + " (Advanced › Recalculate)" : "";
        const detail = [before != null && after != null ? `${before} → ${after} error cells` : "", r?.cells_fixed != null ? `${r.cells_fixed} cells rewritten` : "", full.version ? full.version.label : "", leftText].filter(Boolean).join(" · ");
        this.fixAll.set({ running: false, fraction: null, line, detail, tone: clean ? "ok" : "warn" });
        return clean || accepted;
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
        if (this.destroyed) return false; // the page was left: no further round unseen
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
    // the outside step only for a model whose macros open outside files
    const plan = STEPS.filter((s) => s.id !== "outside" || this.store.outsideInputs() !== null);
    this.steps.set(plan.map((s) => ({ ...s, state: "waiting", line: "", detail: "" })));
    // the clock starts before the estimate is fetched: until then the timer would show the last run's time
    this.stepSeconds = {};
    this.chainT0 = this.stepT0 = Date.now();
    this.clock.set(Date.now());
    await this.loadEstimate(id);
    if (this.ticker) clearInterval(this.ticker);
    this.ticker = setInterval(() => this.clock.set(Date.now()), 10_000);
    const runners: Record<StepId, () => Promise<boolean>> = {
      outside: () => this.runOutside(true),
      fixAll: () => this.runFixAll(),
      prepare: () => this.runPrepare(),
      loop: () => this.runLoop(),
      recalc: () => this.runRecalc(),
    };
    const states: Record<StepId, typeof this.fixAll> = { outside: this.outside, fixAll: this.fixAll, prepare: this.prepare, loop: this.loop, recalc: this.recalc };
    try {
      for (const step of plan) {
        if (this.destroyed) return; // the page was left: the next step would run unseen
        if (step.id === "outside") {
          await this.refreshOutside(id);
          if (this.destroyed) return;
          const why = this.outsideSkip();
          if (why) {
            this.setStep(step.id, { state: "skipped", line: why });
            continue;
          }
        }
        if (step.id === "loop" && !this.canBuildLoop()) {
          this.setStep(step.id, { state: "skipped", line: this.loopPlan()?.found ? "The loop needs a decision first (Advanced › Mind loop)" : "No run loop in this workbook" });
          continue;
        }
        this.stepT0 = Date.now();
        this.clock.set(this.stepT0);
        this.setStep(step.id, { state: "running" });
        let ok = false;
        try {
          ok = await runners[step.id]();
        } catch (e) {
          states[step.id].set({ running: false, fraction: null, line: "Failed", detail: errorText(e), tone: "error" });
        }
        const r = states[step.id]();
        this.stepSeconds[step.id] = (Date.now() - this.stepT0) / 1000;
        // the files are in, but no computed value could be compared: amber, and the chain stops -- going
        // on is the user's choice (Do everything again skips the step: the server says brought in)
        const unchecked = step.id === "outside" && !ok && this.outsideUnchecked();
        this.setStep(step.id, { state: ok ? "done" : unchecked ? "stopped" : "failed", line: r.line, detail: r.detail });
        if (unchecked) {
          this.chainLine.set(
            `Stopped at "${step.label}": the computed values were not checked. Check them under Advanced › Deliver (Check the numbers), or press Do everything again to go on without the check.`,
          );
          return;
        }
        if (!ok) {
          const where = step.id === "outside" ? "See the Outside files card above" : "Fix that under Advanced";
          this.chainLine.set(`Stopped at "${step.label}": ${r.line.replace(/\.$/, "")}. ${where}, then press Do everything again.`);
          return;
        }
      }
      this.chainLine.set(`Everything ran and every check passed (${minutes((Date.now() - this.chainT0) / 1000)}).`);
    } finally {
      if (this.ticker) clearInterval(this.ticker);
      this.ticker = null;
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
