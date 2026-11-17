import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, effect, inject, signal, untracked } from "@angular/core";
import { ButtonModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { CancelButton } from "../../core/cancel-button";
import { MIND_KEPT_LEAD, MIND_NOTHING_UPLOADED, MindLoopEvent, MindLoopIteration, MindLoopReport, MindLoopStatus, MindProject, MindRunOptionsInfo, MindRunTiming, MindRunVersion, MindSignIn, mindDeleteByHand, mindDeleteGone, mindEarlierLine, mindStoppedLine } from "../../core/models";
import { SessionStore } from "../../core/session.store";
import { ShowMeLink } from "../../core/show-me-link";

/** One  line under  a stopped run: a project it left in Mind, with its "Delete it from Mind" button. */
export interface KeptRow {
  name: string;
  /** still  in Mind  as far as the server knows: the button is drawn */
  button: boolean;
  /** "Deleting…", then what Mind did (the delete route's sentence); "" before the first click */
  note: string;
  /** Mind's address, when the person has to delete the  project there by hand (the link "Open Mind") */
  openMind: string | null;
}

/** The  app running  itself against real Mind: prepare, gates, upload, convert, run, decide; one run at a time.
 *  Before the click: which version goes up (always the latest), the changes to make first (ticks) and  a
 *  conservative time estimate. While it runs: the stages, the one going  on now, time spent and time  left.
 *  17 Nov 2026: Cancel beside the progress.  The run stops on the server; the project it was making  stays
 *  in Mind and gets a "Delete it  from Mind" button (nothing is deleted without a click). */
@Component({
  selector: "mr-mind-run-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, ShowMeLink, CancelButton],
  templateUrl: "./mind-run-panel.html",
  styleUrl: "./panels.scss",
})
export class MindRunPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  readonly maxIterations = signal(4);
  readonly runModel = signal(true);
  readonly useAssistant = signal(true);
  /** 17  Nov 2026:  "cancelling" from the click on Cancel until the run has really stopped, then "cancelled" */
  readonly state = signal<MindLoopStatus["state"]>("idle");
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
  /** the newest look at the  run; an older look that answers late is dropped (it would  start a second loop of looks) */
  private pollSeq = 0;

  // ----  Cancel (17 Nov 2026) ----------------------------------------------------------------
  /** Between the click  on Run in Mind and the server's job: the tab's own  sign-in check or its options
   *  are still on  their way, and the run waits for them. */
  readonly starting = signal(false);
  /** this tab's own "Check the sign-in" is on its way (10 to 40 s; the server has no job for it) */
  readonly checking = signal(false);
  /** Cancel was pressed (set at the button's `stopping`, taken back at the top of `start()` and of a new
   *  sign-in check). While the server has no job yet nothing can be stopped there, so it  is read after
   *  the sign-in check, after the options and before the run is started: when it is set, nothing more starts. */
  readonly cancelled = signal(false);
  /** "Stopped. Nothing was uploaded to Mind.": a Cancel that came before anything reached the server */
  readonly nothingUploaded = signal(false);
  /** the server's sentence under a stopped run ("Stopped. The project … is still in Mind.") */
  readonly message = signal("");
  /** the newest project of the run on screen (null: none was made) */
  readonly project = signal<MindProject | null>(null);
  /** this workbook's projects from stopped runs that are still in Mind: each gets "Delete it from Mind" */
  readonly kept = signal<MindProject[]>([]);
  /** other projects the run left in  Mind (earlier rounds that were  run): names only, no button */
  readonly earlier = signal<string[]>([]);
  /** the project being  deleted from Mind  right now  (about a minute, no Cancel); null: none */
  readonly deleting = signal<string | null>(null);
  /** what Mind did with each project "Delete it from Mind" was pressed for, by name */
  readonly deleteNotes = signal<Record<string, { text: string; openMind: string | null }>>({});
  /** the projects shown under the line, in the order first seen  (one that was deleted keeps its line) */
  private readonly seen = signal<string[]>([]);
  private checkWork: Promise<void> | null = null;
  private optionsWork: Promise<void> | null = null;

  readonly iterations = computed<MindLoopIteration[]>(() => this.report()?.iterations ?? []);
  readonly verdictClass = computed(() => {
    const v = this.report()?.verdict;
    return v === "converged" ? "chip chip--ok" : v === "stuck" || v === "exhausted" ? "chip chip--warn" : "chip";
  });
  /** the run is at work  on the server: also while it  is on its way out after a  Cancel */
  readonly running = computed(() => this.state() === "running" || this.state() === "cancelling");
  /** from the click on Cancel until the run has really stopped:  the line says "Stopping…" */
  readonly stoppingNow = computed(() => this.state() === "cancelling" || (this.running() && this.cancelled()));
  /** The line after a  Cancel. Before anything reached the server: "Stopped. Nothing was uploaded to
   *  Mind."; after a stopped run: the server's sentence, which names the project. */
  readonly stopLine = computed(() => {
    if (this.running() || this.starting()) return "";
    if (this.nothingUploaded()) return MIND_NOTHING_UPLOADED;
    return this.state() === "cancelled" ? mindStoppedLine({ message: this.message(), project: this.project() }) : "";
  });
  /** Above the projects of stopped runs when the tab opens later (after a restart, after another run):
   *  the stopped run's own sentence is not there then, so the list says what it is (the banner's words). */
  readonly keptLead = computed(() => (this.kept().length && this.state() !== "cancelled" ? MIND_KEPT_LEAD : ""));
  /** one line per project of a  stopped run: its name, the button while it is still in Mind, what a delete answered */
  readonly keptRows = computed<KeptRow[]>(() => {
    const inMind = new Set(this.kept().map((p) => p.name));
    const notes = this.deleteNotes();
    return this.seen()
      .filter((name) => inMind.has(name) || !!notes[name])
      .map((name) => ({ name, button: inMind.has(name), note: this.deleting() === name ? "Deleting…" : (notes[name]?.text ?? ""), openMind: notes[name]?.openMind ?? null }));
  });
  /** under a stopped run: "Earlier rounds of this run are still in Mind: …" (no button: they were run), or "" */
  readonly earlierLine = computed(() => (this.state() === "cancelled" ? mindEarlierLine(this.earlier()) : ""));
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
    // 17  Nov 2026: Cancel during this  tab's own check: the check cannot be  stopped on the  server, the  line says the screen is on its way out
    if (s.state === "checking") return this.checking() && this.cancelled() ? "Stopping…" : "Checking the Mind sign-in…";
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
    if (!check) {
      await this.readSignIn(false);
      this.watchSignIn();
      return;
    }
    // 17 Nov 2026: the check loads Mind (10 to 40 s) and the server has no job for it, so Cancel cannot
    // stop it  there. The Cancel beside the line sets `cancelled`; it is read  when the check answers, and
    // a Run  in Mind pressed meanwhile waits for this  check (`start`) and then does not start.
    if (this.checking()) return;
    this.cancelled.set(false);
    this.nothingUploaded.set(false);
    this.checking.set(true);
    const work = this.readSignIn(true);
    this.checkWork = work;
    await work;
    if (this.checkWork === work) this.checkWork = null;
    this.checking.set(false);
    if (this.cancelled() && !this.starting()) this.nothingUploaded.set(true); // a waiting start() says it  itself
    this.watchSignIn();
  }

  /** One read of the sign-in (`check`: Mind  itself is loaded); never throws. */
  private async readSignIn(check: boolean): Promise<void> {
    try {
      if (check) this.signIn.update((s) => (s ? { ...s, state: "checking" } : s));
      this.signIn.set(await this.api.mindSignIn(check));
    } catch (e) {
      this.signInError.set(errorText(e));
      try {
        this.signIn.set(await this.api.mindSignIn(false));
      } catch {
        /* the line keeps what it had */
      }
    }
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
    // 17 Nov 2026: kept while it is on its way, so a Run in Mind pressed meanwhile waits for it (`start`)
    const work = this.readOptions();
    this.optionsWork = work;
    await work;
    if (this.optionsWork === work) this.optionsWork = null;
  }

  /** One read of what Run in Mind will upload and change; never throws. */
  private async readOptions(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      const o = await this.api.mindLoopOptions(id, { useAssistant: this.useAssistant(), runModel: this.runModel() });
      this.options.set(o);
      this.ticked.set(new Set(o.choices.filter((c) => c.default_on).map((c) => c.id)));
    } catch {
      /* the panel still works  without it: the loop then chooses by itself */
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
    if (!id || this.running() || this.starting()) return;
    this.error.set("");
    // 17 Nov 2026, "before the server has a job": the Cancel is on screen from this click on, and the
    // server has nothing to stop until the run is started. So `cancelled` is taken back here and read
    // after the tab's own sign-in check and after its options, when one is still on its way (both open
    // or feed the run: the check holds Mind's sign-in, the options are the ticks this run sends).
    this.cancelled.set(false);
    this.nothingUploaded.set(false);
    this.starting.set(true);
    try {
      if (this.checkWork) await this.checkWork;
      if (this.cancelled()) return this.stoppedBefore();
      if (this.optionsWork) await this.optionsWork;
      if (this.cancelled()) return this.stoppedBefore();
      const opts = { maxIterations: this.maxIterations(), runModel: this.runModel(), useAssistant: this.useAssistant() };
      const out = await this.api.startMindLoop(id, this.options() ? { ...opts, actions: [...this.ticked()] } : opts);
      this.events.set([]);
      this.report.set(null);
      this.timing.set(null);
      this.message.set("");
      this.project.set(null);
      this.earlier.set([]);
      this.runVersion.set((out as { version?: MindRunVersion }).version ?? null);
      this.next = 0;
      this.drift.set(0);
      this.state.set("running");
      // the click on Cancel  came while the start itself  was on its way: the server had  no run to  stop
      // then, and  has one now
      if (this.cancelled()) void this.api.cancel(id, "mind").catch(() => undefined);
      void this.poll();
    } catch (e) {
      this.error.set(errorText(e));
      void this.loadOptions();
    } finally {
      this.starting.set(false);
    }
  }

  /** Cancel came before the run was started: nothing more starts, and the line says so. */
  private stoppedBefore(): void {
    this.nothingUploaded.set(true);
  }

  /** The Cancel button's `stopping`: the click, before the server is asked. */
  onStopping(): void {
    this.cancelled.set(true);
  }

  /** The Cancel call itself could  not be made (the server's words): nothing was stopped, the work  goes on. */
  cancelFailed(why: string): void {
    this.cancelled.set(false);
    this.error.set(why);
  }

  /** The Cancel button's `stopped`: the  server has finished; look at the run  now, not at the next tick. */
  onStopped(): void {
    if (this.timer) clearTimeout(this.timer);
    void this.poll();
  }

  private async poll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.live) return;
    const seq = ++this.pollSeq;
    try {
      const s = await this.api.mindLoopStatus(id, this.next);
      if (seq !== this.pollSeq) return; // a newer look is on its  way: it reads the same events again
      this.take(s);
    } catch {
      /* the  next tick tries again */
    }
    // 17 Nov 2026: the looks never end at the click on Cancel: they go on through "cancelling" until
    // the server says cancelled, done or error
    if (this.live && seq === this.pollSeq && this.running()) this.timer = setTimeout(() => void this.poll(), 2000);
  }

  /** One answer of GET .../mind-loop onto the screen. */
  private take(s: MindLoopStatus): void {
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
    this.message.set(s.message ?? "");
    this.project.set(s.project ?? null);
    this.earlier.set(s.earlier ?? []);
    this.takeKept(s.kept ?? []);
  }

  /** The projects of stopped runs that are still in Mind; each one seen keeps its place in the list. */
  private takeKept(kept: MindProject[]): void {
    this.kept.set(kept);
    const seen = this.seen();
    const fresh = kept.map((p) => p.name).filter((n) => !seen.includes(n));
    if (fresh.length) this.seen.set([...seen, ...fresh]);
  }

  /** "Delete it from Mind", for a project a stopped run  left there. Blocking, about a minute, no Cancel.
   *  Nothing is deleted without this click. Afterwards the button goes or stays as the server's own list
   *  says (GET .../mind-loop `kept`); when that cannot be read, as the answer itself says. */
  async deleteProject(name: string): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.deleting() || this.running() || this.starting()) return;
    this.deleting.set(name);
    try {
      const d = await this.api.deleteMindProject(id, name);
      this.deleteNotes.update((n) => ({ ...n, [name]: { text: d.message, openMind: mindDeleteByHand(d) ? d.url || null : null } }));
      try {
        const s = await this.api.mindLoopStatus(id, this.next);
        ++this.pollSeq;
        this.take(s);
      } catch {
        if (mindDeleteGone(d)) {
          this.takeKept(this.kept().filter((p) => p.name !== name));
          // the sentence above must not go on saying "is still in Mind" for it (the server's words for both)
          if (this.project()?.name === name) this.message.set(d.result === "deleted" ? `Stopped. The project ${name} was deleted from Mind.` : `Stopped. The project ${name} is not in Mind.`);
        }
      }
    } catch (e) {
      // refused before Mind was asked ("Mind is busy. Try again in a moment."): the server's words, the button stays
      this.deleteNotes.update((n) => ({ ...n, [name]: { text: errorText(e), openMind: null } }));
    } finally {
      this.deleting.set(null);
    }
  }

  /** the newest line of the run, in plain words  (no "iteration n:" prefix) */
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

  /** 754 ->  "12:34" */
  clockText(s: number): string {
    const t = Math.max(0, Math.round(s));
    const h = Math.floor(t / 3600);
    const m = Math.floor((t % 3600) / 60);
    const sec = t % 60;
    return (h ? `${h}:${String(m).padStart(2, "0")}` : String(m)) + ":" + String(sec).padStart(2, "0");
  }

  /** 754 -> "13 min" (rounded up: the  estimate stays on  the long  side) */
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
