import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, effect, inject, signal, untracked } from "@angular/core";
import { ButtonModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { MindLoopEvent, MindLoopIteration, MindLoopReport, MindRunOptionsInfo, MindRunTiming, MindRunVersion, MindSignIn } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** The app running itself against real Mind: prepare, gates, upload, convert, run, decide; one run at a time.
 *  Before the click: which version goes up (always the latest), the changes to make first (ticks) and a
 *  conservative time estimate. While it runs: the stages, the one going on now, time spent and time left. */
@Component({
  selector: "mr-mind-run-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule],
  templateUrl: "./mind-run-panel.html",
  styleUrl: "./panels.scss",
})
export class MindRunPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  readonly maxIterations = signal(4);
  readonly runModel = signal(true);
  readonly useAssistant = signal(true);
  readonly state = signal<"idle" | "running" | "done" | "error">("idle");
  readonly events = signal<MindLoopEvent[]>([]);
  readonly report = signal<MindLoopReport | null>(null);
  readonly summary = signal("");
  readonly error = signal("");
  readonly options = signal<MindRunOptionsInfo | null>(null);
  readonly ticked = signal<Set<string>>(new Set());
  readonly timing = signal<MindRunTiming | null>(null);
  readonly runVersion = signal<MindRunVersion | null>(null);
  readonly showLog = signal(false);
  readonly signIn = signal<MindSignIn | null>(null);
  readonly signInError = signal("");
  private signInTimer: ReturnType<typeof setTimeout> | null = null;
  /** seconds since the last status answer, so the clock moves every second between polls */
  readonly drift = signal(0);
  private next = 0;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private clock: ReturnType<typeof setInterval> | null = null;
  private live = true;

  readonly iterations = computed<MindLoopIteration[]>(() => this.report()?.iterations ?? []);
  readonly verdictClass = computed(() => {
    const v = this.report()?.verdict;
    return v === "converged" ? "chip chip--ok" : v === "stuck" || v === "exhausted" ? "chip chip--warn" : "chip";
  });
  readonly running = computed(() => this.state() === "running");
  readonly tickedCount = computed(() => this.ticked().size);
  /** the estimate before the click (the ticks change only the short 'making the changes' stage) */
  readonly estimateText = computed(() => {
    const o = this.options();
    if (!o) return "";
    return `about ${this.minutes(o.estimate.first_s)}`;
  });
  readonly elapsed = computed(() => (this.timing()?.elapsed_s ?? 0) + (this.running() ? this.drift() : 0));
  readonly remaining = computed(() => Math.max(0, (this.timing()?.remaining_s ?? 0) - (this.running() ? this.drift() : 0)));

  constructor() {
    // a new version (Fix all issues, Prep...) changes what goes up and what can be ticked
    effect(() => {
      this.store.currentVersion();
      this.store.plan();
      untracked(() => {
        if (!this.running()) void this.loadOptions();
      });
    });
  }

  ngOnInit(): void {
    void this.poll();
    void this.loadSignIn(false);
    this.clock = setInterval(() => this.drift.update((d) => d + 1), 1000);
  }

  ngOnDestroy(): void {
    this.live = false;
    if (this.timer) clearTimeout(this.timer);
    if (this.clock) clearInterval(this.clock);
    if (this.signInTimer) clearTimeout(this.signInTimer);
  }

  // ---- the Mind sign-in ------------------------------------------------------------------
  /** One line on the sign-in: signed in / not / never, and when. */
  signInLine(): string {
    const s = this.signIn();
    if (!s) return "";
    if (s.state === "waiting") return s.message || "An Edge window is open on Mind: sign in there.";
    if (s.state === "checking") return "Checking the Mind sign-in…";
    if (s.state === "failed" && s.message) return s.message;
    const when = (t: string | null | undefined) => (t ? new Date(t).toLocaleString([], { dateStyle: "medium", timeStyle: "short" }) : "");
    if (s.checked) return s.checked.signed_in ? `Signed in to Mind (checked ${when(s.checked.at)}).` : "Not signed in to Mind, or the sign-in has expired: sign in again.";
    if (s.last_login) return `Last signed in ${when(s.last_login)}. A sign-in lasts about a week: check it, or sign in again.`;
    return "Not signed in to Mind yet: Run in Mind needs it once (about a week).";
  }

  signedIn(): boolean {
    return !!this.signIn()?.checked?.signed_in;
  }

  async loadSignIn(check: boolean): Promise<void> {
    this.signInError.set("");
    try {
      if (check) this.signIn.update((s) => (s ? { ...s, state: "checking" } : s));
      this.signIn.set(await this.api.mindSignIn(check));
    } catch (e) {
      this.signInError.set(errorText(e));
      void this.loadSignIn(false).catch(() => undefined);
    }
    this.watchSignIn();
  }

  async openSignIn(): Promise<void> {
    this.signInError.set("");
    try {
      this.signIn.set(await this.api.startMindSignIn());
    } catch (e) {
      this.signInError.set(errorText(e));
    }
    this.watchSignIn();
  }

  async stopSignIn(): Promise<void> {
    try {
      this.signIn.set(await this.api.startMindSignIn(true));
    } catch {
      /* the next poll shows where it is */
    }
  }

  /** while the Edge window is open, look every 2 s */
  private watchSignIn(): void {
    if (this.signInTimer) clearTimeout(this.signInTimer);
    if (!this.live || this.signIn()?.state !== "waiting") return;
    this.signInTimer = setTimeout(async () => {
      try {
        this.signIn.set(await this.api.mindSignIn(false));
      } catch {
        /* try again */
      }
      this.watchSignIn();
    }, 2000);
  }

  async loadOptions(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      const o = await this.api.mindLoopOptions(id, { useAssistant: this.useAssistant(), runModel: this.runModel() });
      this.options.set(o);
      this.ticked.set(new Set(o.choices.filter((c) => c.default_on).map((c) => c.id)));
    } catch {
      /* the panel still works without it: the loop then chooses by itself */
    }
  }

  toggle(id: string): void {
    this.ticked.update((s) => {
      const n = new Set(s);
      if (n.has(id)) n.delete(id);
      else n.add(id);
      return n;
    });
  }

  tickAll(on: boolean): void {
    this.ticked.set(new Set(on ? (this.options()?.choices ?? []).map((c) => c.id) : []));
  }

  async start(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.running()) return;
    this.error.set("");
    try {
      const opts = { maxIterations: this.maxIterations(), runModel: this.runModel(), useAssistant: this.useAssistant() };
      const out = await this.api.startMindLoop(id, this.options() ? { ...opts, actions: [...this.ticked()] } : opts);
      this.events.set([]);
      this.report.set(null);
      this.timing.set(null);
      this.runVersion.set((out as { version?: MindRunVersion }).version ?? null);
      this.next = 0;
      this.drift.set(0);
      this.state.set("running");
      void this.poll();
    } catch (e) {
      this.error.set(errorText(e));
      void this.loadOptions();
    }
  }

  private async poll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.live) return;
    try {
      const s = await this.api.mindLoopStatus(id, this.next);
      if (s.events?.length) this.events.update((e) => [...e, ...s.events].slice(-300));
      this.next = s.next ?? this.next;
      this.state.set(s.state);
      if (s.report) this.report.set(s.report);
      if (s.summary) this.summary.set(s.summary);
      if (s.error) this.error.set(s.error);
      if (s.timing) {
        this.timing.set(s.timing);
        this.drift.set(0);
      }
      if (s.version) this.runVersion.set(s.version);
    } catch {
      /* the next tick tries again */
    }
    if (this.live && this.state() === "running") this.timer = setTimeout(() => void this.poll(), 2000);
  }

  /** the newest line of the run, in plain words (no "iteration n:" prefix) */
  latestLine(): string {
    const e = this.events();
    const last = e.length ? e[e.length - 1] : null;
    return (last?.message ?? "").replace(/^iteration \d+:\s*/, "");
  }

  gate(it: MindLoopIteration, which: "apply" | "numbers" | "names" | "convert" | "run"): { label: string; cls: string } {
    const ok = (b: boolean | null | undefined) => (b === true ? "chip chip--ok" : b === false ? "chip chip--error" : "chip");
    switch (which) {
      case "apply":
        return { label: it.apply ? `${it.apply.applied} applied` : "not prepared", cls: ok(it.apply ? it.apply.failed === 0 : undefined) };
      case "numbers":
        return { label: it.numbers?.ran ? (it.numbers.match ? "numbers match" : `${it.numbers.differences ?? "?"} differ`) : "numbers skipped", cls: ok(it.numbers?.ran ? it.numbers.match : undefined) };
      case "names":
        return { label: it.names ? `${it.names.unnamed} untitled` : "names skipped", cls: ok(it.names?.ok) };
      case "convert":
        return { label: it.mind?.skipped ? "Mind skipped" : it.mind?.convert ? (it.mind.convert.success ? "converted" : "convert failed") : "no convert", cls: ok(it.mind?.convert?.success) };
      case "run":
        return { label: it.mind?.run ? (it.mind.run.completed ? (it.mind.run.audit_consistent ? "run ok" : "run inconsistent") : "run failed") : "not run", cls: ok(it.mind?.run ? it.mind.run.completed && it.mind.run.audit_consistent : undefined) };
    }
  }

  num(event: Event): number {
    return Number((event.target as HTMLInputElement).value) || 1;
  }

  when(t: number): string {
    const d = new Date(t * 1000);
    return isNaN(d.getTime()) ? "" : d.toLocaleTimeString();
  }

  /** 754 -> "12:34" */
  clockText(s: number): string {
    const t = Math.max(0, Math.round(s));
    const h = Math.floor(t / 3600);
    const m = Math.floor((t % 3600) / 60);
    const sec = t % 60;
    return (h ? `${h}:${String(m).padStart(2, "0")}` : String(m)) + ":" + String(sec).padStart(2, "0");
  }

  /** 754 -> "13 min" (rounded up: the estimate stays on the long side) */
  minutes(s: number): string {
    const m = Math.max(1, Math.ceil(s / 60));
    return m >= 90 ? `${Math.floor(m / 60)} h ${m % 60} min` : `${m} min`;
  }

  /** Clock time when it should be done, e.g. "14:35" */
  doneBy(): string {
    const d = new Date(Date.now() + this.remaining() * 1000);
    return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  }
}
