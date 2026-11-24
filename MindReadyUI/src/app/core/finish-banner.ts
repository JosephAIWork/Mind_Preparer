import { Component, OnDestroy, OnInit, computed, inject, input, signal } from "@angular/core";
import { ApiService, errorText } from "./api.service";
import { CancelButton } from "./cancel-button";
import { FeedbackService } from "./feedback.service";
import { MIND_KEPT_LEAD, MIND_NOTHING_UPLOADED, MindDelete, MindLoopOptions, MindLoopStatus, MindProject, MindSignIn, Readiness, mindDeleteByHand, mindDeleteGone, mindEarlierLine, mindStoppedLine } from "./models";
import { RecalcNow } from "./recalc-now";
import { NOT_SIGNED_IN, SIGN_IN_WAITING } from "./sign-in-text";
import { Thinking } from "./thinking";
import { SessionStore } from "./session.store";

/** What  the end-of-check  banner says and allows, from the readiness verdict (1.15.0). */
export interface FinishState {
  /** the  workbook can  go to Mind */
  acceptable: boolean;
  /** the one line */
  text: string;
  /** why Upload to Mind is greyed ("" when it is not) */
  reason: string;
  /** blocking problems are left (a warning colour); false for a step still to do (calm colour) */
  problem?: boolean;
  /** 17 Nov  2026: a line under an acceptable verdict: the cells that show  an error only for the Mind  loop's first value */
  note?: string;
}

/** The  banner's words  for a verdict: acceptable, blocked (how many blocking problems are left), or
 *  not yet verified by a recalculation. */
export function finishState(r: Readiness | null | undefined): FinishState | null {
  if (!r) return null;
  if (r.state === "ready") {
    // 17  Nov 2026  (Model A): errors that show only for the Mind loop's first value do not stop it -- said in one line
    const note = r.recalc?.loop_note;
    return { acceptable: true, text: "The check is finished. The workbook can be uploaded to Mind.", reason: "", ...(note ? { note } : {}) };
  }
  if (r.state === "unverified") {
    if (r.recalc?.ran) {
      const n = r.recalc.new_errors ?? r.recalc.formula_errors ?? 0;
      // 16 Nov 2026: after the Mind loop the errors usually come  from the loop's first value -- said in plain words
      // (17 Nov  2026: counted. `n` is only the  errors that also show without the loop;  the line names the others,
      // or says that it could not be  checked)
      return { acceptable: false, text: `The recalculation found ${n} new formula error${n === 1 ? "" : "s"}.`, reason: r.recalc.loop_note || "The list is under Advanced › Recalculate." };
    }
    return { acceptable: false, text: "One step left: the recalculation has not run yet.", reason: "Press Recalculate: it finishes the check." };
  }
  const n = r.blocking_count;
  return {
    acceptable: false,
    text: `${n} blocking problem${n === 1 ? " is" : "s are"} left.`,
    reason: "Not possible while blocking problems are left.",
    problem: true,
  };
}

/** One  line under  the banner for a project in Mind: one a stopped run left there (`kept`: it has the
 *  button "Delete it from Mind"), or one that was just deleted (the sentence only). */
export interface ProjectLine {
  name: string;
  kept: boolean;
  /** what  the last  "Delete it from Mind" said about it */
  note: string;
  /** Mind did not delete it: the link "Open Mind" (a new tab) stands beside the sentence */
  openMind: string | null;
}

/** The  end-of-check banner  of the Basic screen: the verdict in one line, with two big buttons.
 *  "Upload to Mind" does what Run in Mind (Advanced) does: upload, convert, run, read back; when the
 *  Mind sign-in is  missing it asks for it first. "Download" gives the current version,  as
 *  "Download this version" does. Not acceptable:  Download stays, Upload is greyed with its  reason.
 *
 *  17 Nov 2026: Cancel stands beside the progress line (and beside  the sign-in line while the sign-in
 *  window is  open). A stopped upload keeps the project it was making in Mind: the banner names it and
 *  offers "Delete it from Mind"; nothing is deleted without that click. */
@Component({
  selector: "mr-finish-banner",
  standalone: true,
  imports: [RecalcNow, Thinking, CancelButton],
  template: `
    @if (state(); as s) {
      <section class="finish" [class.finish--stop]="!s.acceptable" [class.finish--problem]="s.problem" data-show-me="finish-banner">
        <p class="finish__text" role="status">{{ s.text }}</p>
        <div class="finish__buttons">
          <mr-recalc-now [disabled]="running()"></mr-recalc-now>
          <button type="button" class="cds--btn cds--btn--primary cds--btn--lg finish__upload" data-show-me="finish-upload" (click)="upload()" [disabled]="!s.acceptable || running() || deleting() !== null" [title]="s.reason">
            {{ running() ? "Uploading…" : "Upload to Mind" }}
          </button>
          @if (downloadUrl(); as url) {
            <a class="cds--btn cds--btn--secondary cds--btn--lg finish__download" data-show-me="finish-download" [href]="url">Download</a>
          }
        </div>
        @if (!s.acceptable && s.reason) {
          <p class="finish__reason">{{ s.reason }}</p>
        }
        @if (s.acceptable && s.note) {
          <p class="finish__note">{{ s.note }}</p>
        }
        @if (askSignIn()) {
          <div class="finish__signin">
            <p>{{ signInLine() }}</p>
            <div class="finish__row">
              <button type="button" class="cds--btn cds--btn--tertiary cds--btn--md finish__signin-open" (click)="openSignIn()" [disabled]="signIn()?.state === 'waiting'">Sign in to Mind</button>
              @if (signIn()?.state === "failed" && signIn()?.feedback_ref) {
                <!-- 19 Nov 2026: a sign-in that could not start: the program's own error goes with Send feedback, never on the screen -->
                <button type="button" class="cds--btn cds--btn--ghost cds--btn--md" data-test="finish-sign-in-feedback" (click)="signInFeedback()">Send feedback</button>
              }
              @if (signIn()?.state === "waiting") {
                <!-- 17 Nov 2026: the upload waits for the sign-in window: Cancel closes it, nothing is uploaded -->
                <mr-cancel kind="mind" size="md" (stopping)="onStopping()"></mr-cancel>
              }
            </div>
          </div>
        }
        @if (running()) {
          <div class="finish__row finish__working">
            @if (progress()) {
              <p class="finish__progress" role="status">@if (aiAtWork()) {<mr-thinking></mr-thinking> }{{ progress() }}</p>
            }
            <!-- 17 Nov 2026: Cancel beside the progress line, from the sign-in check to the end of the run -->
            <mr-cancel kind="mind" size="md" [busy]="stoppingNow()" (stopping)="onStopping()"></mr-cancel>
          </div>
        }
        @if (outcome()) {
          <p class="finish__outcome" [class.finish__outcome--error]="failed()" role="status">{{ outcome() }}</p>
        }
        @if (projectLines().length) {
          <!-- 17 Nov 2026: the projects stopped runs left in Mind: nothing is deleted without a click -->
          @if (keptLead()) {
            <p class="finish__kept-lead">{{ keptLead() }}</p>
          }
          <ul class="finish__projects">
            @for (p of projectLines(); track p.name) {
              <li class="finish__project">
                <bdi class="finish__project-name">{{ p.name }}</bdi>
                @if (p.kept) {
                  <button type="button" class="cds--btn cds--btn--danger--tertiary cds--btn--sm finish__delete" (click)="deleteProject(p.name)" [disabled]="deleting() !== null || running()">
                    {{ deleting() === p.name ? "Deleting…" : "Delete it from Mind" }}
                  </button>
                }
                @if (p.note) {
                  <span class="finish__delete-note">{{ p.note }}</span>
                }
                @if (p.openMind; as url) {
                  <a class="finish__open-mind" [href]="url" target="_blank" rel="noopener">Open Mind</a>
                }
              </li>
            }
          </ul>
        }
        @if (earlierLine()) {
          <p class="finish__earlier">{{ earlierLine() }}</p>
        }
      </section>
    }
  `,
  styles: [
    `
      :host {
        display: block;
      }
      .finish {
        padding: 1rem;
        border-left: 4px solid var(--cds-support-success, #24a148);
        background: var(--cds-layer-01, #f4f4f4);
        margin: 1rem 0;
      }
      .finish--stop {
        border-left-color: var(--cds-support-info, #0043ce);
      }
      .finish--problem {
        border-left-color: var(--cds-support-warning, #f1c21b);
      }
      .finish__text {
        font-weight: 600;
        margin: 0 0 0.75rem;
      }
      .finish__buttons {
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
      }
      .finish__reason,
      .finish__note,
      .finish__progress,
      .finish__signin p,
      .finish__kept-lead,
      .finish__earlier {
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
        margin: 0.5rem 0 0;
      }
      .finish__row {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        flex-wrap: wrap;
        margin-top: 0.5rem;
      }
      .finish__working .finish__progress {
        margin: 0;
      }
      .finish__outcome {
        margin: 0.5rem 0 0;
        font-weight: 600;
      }
      .finish__outcome--error {
        color: var(--cds-text-error, #da1e28);
      }
      .finish__projects {
        list-style: none;
        margin: 0.25rem 0 0;
        padding: 0;
      }
      .finish__project {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        flex-wrap: wrap;
        margin-top: 0.25rem;
        font-size: 0.875rem;
      }
      .finish__project-name {
        font-weight: 600;
        overflow-wrap: anywhere;
      }
      .finish__delete-note {
        color: var(--cds-text-secondary, #525252);
      }
    `,
  ],
})
export class FinishBanner implements OnInit, OnDestroy {
  private api = inject(ApiService);
  private feedback = inject(FeedbackService);
  private store = inject(SessionStore);

  /** how  often the  run in Mind (and the sign-in window) is looked at, in ms (the tests shorten it) */
  static everyMs = 2000;

  /** the verdict, from the page */
  readonly readiness = input<Readiness | null | undefined>(null);
  /** the current version's download address */
  readonly downloadUrl = input<string | null>(null);

  readonly state = computed(() => finishState(this.readiness()));
  /** the stage running now is the assistant's (naming  the grids, explaining, making stand-ins) */
  readonly aiAtWork = computed(() => /\b(nam(e|ing)|assistant|suggest|explain|stand-in)/i.test(this.progress()));
  readonly running = signal(false);
  readonly progress = signal("");
  readonly outcome = signal("");
  readonly failed = signal(false);
  readonly signIn = signal<MindSignIn | null>(null);
  readonly askSignIn = signal(false);
  /** 17 Nov 2026: Cancel was clicked and  the upload has not ended yet:  the line and the button say "Stopping…" */
  readonly stoppingNow = signal(false);
  /** the  projects stopped runs of this workbook left in Mind (the server's  list): each has "Delete it from Mind" */
  readonly kept = signal<MindProject[]>([]);
  /** other  projects the last run left in Mind (earlier rounds that were run): named, no button */
  readonly earlier = signal<string[]>([]);
  /** the project being deleted now (null: none) */
  readonly deleting = signal<string | null>(null);
  /** what "Delete it from Mind" last said, by project */
  private readonly answers = signal<Record<string, { message: string; openMind: string | null }>>({});
  /** the line above the list is the stopped run's own sentence (the server's), and it names a project of the
   *  list: the list then needs no lead of its own */
  private readonly stoppedShown = signal(false);
  /** the outcome line is a  stopped run's sentence: it follows the server after a delete ("... was deleted from Mind.") */
  private stoppedRun = false;
  readonly projectLines = computed<ProjectLine[]>(() => {
    const answers = this.answers();
    const kept = this.kept();
    const lines: ProjectLine[] = kept.map((p) => ({ name: p.name, kept: true, note: answers[p.name]?.message ?? "", openMind: answers[p.name]?.openMind ?? null }));
    // a project that was just deleted keeps its sentence ("Deleted from Mind."); its button is gone
    for (const [name, a] of Object.entries(answers)) if (!kept.some((p) => p.name === name)) lines.push({ name, kept: false, note: a.message, openMind: a.openMind });
    return lines;
  });
  readonly keptLead = computed(() => (this.kept().length && !this.stoppedShown() ? MIND_KEPT_LEAD : ""));
  readonly earlierLine = computed(() => mindEarlierLine(this.earlier()));
  /** 17 Nov 2026: Cancel was clicked. Set at the click, reset at the top of `upload()`; read after every call
   *  that comes before the server has a job to stop (the sign-in check, the options, the wait for the
   *  sign-in window): once it is set nothing more starts. */
  private cancelled = false;
  private live = true;
  private timer: ReturnType<typeof setTimeout> | null = null;

  /** 17 Nov 2026: when the banner opens  (it is drawn again after  every Start): an upload that still runs  is
   *  followed again, and  the projects  stopped runs left in Mind are listed again with their button. */
  ngOnInit(): void {
    void this.look();
  }

  ngOnDestroy(): void {
    this.live = false;
    if (this.timer) clearTimeout(this.timer);
  }

  signInLine(): string {
    const s = this.signIn();
    if (s?.state === "waiting") return s.message || SIGN_IN_WAITING;
    if (s?.state === "failed" && s.message) return s.message;
    return NOT_SIGNED_IN;
  }

  /** "Send feedback" beside a sign-in that could not start: the incident the server wrote, the sentence as the note. */
  signInFeedback(): void {
    const s = this.signIn();
    if (s?.state === "failed" && s.feedback_ref) this.feedback.open(s.feedback_ref, s.message || "");
  }

  /** Upload to Mind:  the sign-in first (asked for when missing), then the same run  as Run in Mind. */
  async upload(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.running() || !this.state()?.acceptable) return;
    this.cancelled = false;
    this.stoppingNow.set(false);
    this.stoppedShown.set(false);
    this.stoppedRun = false;
    this.outcome.set("");
    this.failed.set(false);
    this.running.set(true);
    this.progress.set("Checking the Mind sign-in…");
    try {
      const s = await this.api.mindSignIn(true);
      if (this.cancelled) return this.stoppedBefore(); // Cancel during the  check (10 to 40 s): nothing starts
      this.signIn.set(s);
      if (!s.checked?.signed_in) {
        this.running.set(false);
        this.progress.set("");
        this.askSignIn.set(true);
        return;
      }
      this.askSignIn.set(false);
      await this.start(id);
    } catch (e) {
      if (this.cancelled) return this.stoppedBefore();
      this.fail(errorText(e));
    }
  }

  /** Opens  the Edge window; when the sign-in is done the upload starts  by itself. */
  async openSignIn(): Promise<void> {
    this.cancelled = false; // the person's own click: a Cancel  before it is over
    try {
      this.signIn.set(await this.api.startMindSignIn());
    } catch (e) {
      this.fail(errorText(e));
      return;
    }
    this.watchSignIn();
  }

  private watchSignIn(): void {
    if (this.timer) clearTimeout(this.timer);
    if (!this.live) return;
    this.timer = setTimeout(async () => {
      try {
        const s = await this.api.mindSignIn(false);
        if (this.cancelled) return; // 17 Nov 2026: Cancel while the sign-in window waited: no upload starts
        this.signIn.set(s);
        if (s.state === "waiting") return this.watchSignIn();
        if (s.checked?.signed_in || s.state === "done") {
          this.askSignIn.set(false);
          void this.upload();
        }
      } catch {
        if (!this.cancelled) this.watchSignIn();
      }
    }, FinishBanner.everyMs);
  }

  /** 17 Nov 2026: the click on Cancel (the button's `stopping`, before the server is asked). The button
   *  itself tells the server (kind "mind": a run is stopped, the sign-in window is closed). Here: nothing
   *  more starts. A run the server already has is followed until the server says "cancelled";  before
   *  that (the sign-in check cannot be stopped half-way) the line says "Stopping…" until the check is back. */
  onStopping(): void {
    this.cancelled = true;
    if (!this.running()) {
      // the wait for the sign-in window: it ends here, and nothing was uploaded
      if (this.timer) clearTimeout(this.timer);
      this.signIn.set(null);
      this.askSignIn.set(false);
      this.stoppedBefore();
      return;
    }
    this.stoppingNow.set(true);
    this.progress.set("Stopping…");
  }

  /** Cancel came before the server had a job (the sign-in check, the sign-in window, the options). */
  private stoppedBefore(): void {
    this.running.set(false);
    this.stoppingNow.set(false);
    this.progress.set("");
    this.failed.set(false);
    this.stoppedShown.set(false);
    this.stoppedRun = false;
    this.outcome.set(MIND_NOTHING_UPLOADED);
  }

  /** The same call as Run in Mind: the changes it ticks by default, 4 rounds at most, the model run, the assistant naming grids. */
  private async start(id: string): Promise<void> {
    this.progress.set("Starting the upload…");
    try {
      const opts: MindLoopOptions = { maxIterations: 4, runModel: true, useAssistant: true };
      try {
        const o = await this.api.mindLoopOptions(id, { useAssistant: true, runModel: true });
        opts.actions = o.choices.filter((c) => c.default_on).map((c) => c.id);
      } catch {
        /* the run then chooses by itself */
      }
      if (this.cancelled) return this.stoppedBefore(); // Cancel while the options were read: no run starts
      await this.api.startMindLoop(id, opts);
    } catch (e) {
      if (this.cancelled) return this.stoppedBefore();
      this.fail(errorText(e));
      return;
    }
    // Cancel was  clicked while the start was  on its way: the server had no  run to stop  then, and  has one
    // now. Ask again; the run is followed to its "cancelled" like any other.
    if (this.cancelled) await this.cancelAgain(id);
    await this.poll(id, 0);
  }

  private async cancelAgain(id: string): Promise<void> {
    try {
      await this.api.cancel(id, "mind");
    } catch {
      /* the run then goes on and its result is shown: nothing was promised */
    }
  }

  private async poll(id: string, after: number): Promise<void> {
    if (!this.live) return;
    let next = after;
    try {
      const s: MindLoopStatus = await this.api.mindLoopStatus(id, after);
      next = s.next ?? after;
      // (the look does not end at  the click: it goes on until the server says cancelled, done  or error)
      if (s.state === "cancelling" || this.stoppingNow()) this.progress.set("Stopping…");
      else if (s.timing?.stage_title) this.progress.set(`${s.timing.stage_title}…`);
      if (s.state === "cancelled") {
        // 17 Nov 2026: stopped with  Cancel: the server's sentence names the project  it left in Mind (the
        // project is kept: nothing is deleted without  a click); under it, each kept project  with its button
        this.running.set(false);
        this.stoppingNow.set(false);
        this.progress.set("");
        this.showProjects(s);
        this.showStopped(s);
        return;
      }
      if (s.state === "done" || s.state === "error") {
        // (also after a Cancel that came too late: the run's result stands and is shown)
        this.running.set(false);
        this.stoppingNow.set(false);
        this.progress.set("");
        const verdict = s.report?.verdict;
        this.failed.set(s.state === "error" || (!!verdict && verdict !== "converged"));
        this.outcome.set(s.state === "error" ? (s.error ?? "The run in Mind stopped.") : (s.report?.reason ?? s.summary ?? (this.failed() ? "The run in Mind did not finish cleanly." : "The workbook ran in Mind.")));
        this.stoppedShown.set(false);
        this.stoppedRun = false;
        this.showProjects(s);
        return;
      }
    } catch {
      /* the next look tries again */
    }
    this.timer = setTimeout(() => void this.poll(id, next), FinishBanner.everyMs);
  }

  /** The banner opens: one look at the run in Mind. */
  private async look(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    let s: MindLoopStatus;
    try {
      s = await this.api.mindLoopStatus(id, 0);
    } catch {
      return; // nothing to pick up
    }
    // (an upload pressed meanwhile reads the list itself when it ends)
    if (!s || !this.live || this.running() || this.store.sessionId() !== id) return;
    this.showProjects(s);
    // the last run was stopped and left something in Mind: its sentence stands above the list again
    if (s.state === "cancelled" && this.kept().length && !this.outcome()) this.showStopped(s);
    if (s.state !== "running" && s.state !== "cancelling") return;
    this.cancelled = false;
    this.outcome.set("");
    this.failed.set(false);
    this.running.set(true);
    this.stoppingNow.set(s.state === "cancelling");
    this.progress.set(s.state === "cancelling" ? "Stopping…" : s.timing?.stage_title ? `${s.timing.stage_title}…` : "Uploading to Mind…");
    this.timer = setTimeout(() => void this.poll(id, s.next ?? 0), FinishBanner.everyMs);
  }

  /** The sentence of  a run that was stopped with Cancel (the server's: it names the project the run left in
   *  Mind, or says that nothing was uploaded). Not an error. */
  private showStopped(s: MindLoopStatus): void {
    this.failed.set(false);
    this.outcome.set(mindStoppedLine(s));
    this.stoppedRun = true;
    const own = s.project?.name;
    this.stoppedShown.set(!!own && this.kept().some((p) => p.name === own));
  }

  /** The projects stopped runs left in Mind, and the earlier rounds' projects, as the server lists them now. */
  private showProjects(s: MindLoopStatus): void {
    this.kept.set(Array.isArray(s.kept) ? s.kept : []);
    this.earlier.set(Array.isArray(s.earlier) ? s.earlier : []);
  }

  /** "Delete it from Mind": about a minute, no Cancel. The server says in one sentence what Mind did; the
   *  button goes when Mind no longer has the project, and stays (with the link "Open Mind" when Mind
   *  refused) when it still  does. */
  async deleteProject(name: string): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.deleting() !== null || this.running()) return;
    this.deleting.set(name);
    this.answers.update(({ [name]: _old, ...rest }) => rest);
    let answer: MindDelete | null = null;
    let note: string;
    try {
      answer = await this.api.deleteMindProject(id, name);
      note = answer.message;
    } catch (e) {
      note = errorText(e); // "Mind is busy. Try again  in a moment.": the button stays
    }
    const url = answer?.url || this.kept().find((p) => p.name === name)?.url || null;
    this.answers.update((all) => ({ ...all, [name]: { message: note, openMind: mindDeleteByHand(answer) ? url : null } }));
    const gone = mindDeleteGone(answer);
    let seen = false;
    try {
      // the list  as the server  has it  now says whether the button stays
      const s = await this.api.mindLoopStatus(id, 0);
      if (this.store.sessionId() === id) {
        this.showProjects(s);
        // the stopped run's sentence follows what was done to its project ("... was deleted from Mind.")
        if (this.stoppedRun && s.state === "cancelled" && !this.running()) this.showStopped(s);
        seen = true;
      }
    } catch {
      /* keep the list; the answer itself decides below */
    }
    if (gone) {
      this.kept.update((list) => list.filter((p) => p.name !== name));
      // the list could not  be read again: a sentence that still says "is still in  Mind" would be wrong now
      if (!seen && this.stoppedRun && this.outcome().includes(name)) {
        this.outcome.set("");
        this.stoppedRun = false;
        this.stoppedShown.set(false);
      }
    }
    this.deleting.set(null);
  }

  private fail(message: string): void {
    this.running.set(false);
    this.stoppingNow.set(false);
    this.progress.set("");
    this.failed.set(true);
    this.outcome.set(message);
  }
}
