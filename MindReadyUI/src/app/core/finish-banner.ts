import { Component, OnDestroy, computed, inject, input, signal } from "@angular/core";
import { ApiService, errorText } from "./api.service";
import { MindLoopOptions, MindLoopStatus, MindSignIn, Readiness } from "./models";
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
}

/** The  banner's words  for a verdict: acceptable, blocked (how many blocking problems are left), or
 *  not yet verified by a recalculation. */
export function finishState(r: Readiness | null | undefined): FinishState | null {
  if (!r) return null;
  if (r.state === "ready") return { acceptable: true, text: "The check is finished. The workbook can be uploaded to Mind.", reason: "" };
  if (r.state === "unverified") {
    return { acceptable: false, text: "One step left: the recalculation has not run yet.", reason: "Press Go first: the recalculation finishes the check." };
  }
  const n = r.blocking_count;
  return {
    acceptable: false,
    text: `${n} blocking problem${n === 1 ? " is" : "s are"} left.`,
    reason: "Not possible while blocking problems are left.",
    problem: true,
  };
}

/** The  end-of-check banner  of the Basic screen: the verdict in one line, with two big buttons.
 *  "Upload to Mind" does what Run in Mind (Advanced) does: upload, convert, run, read back; when the
 *  Mind sign-in is  missing it asks for it first. "Download" gives the current version,  as
 *  "Download this version" does. Not acceptable:  Download stays, Upload is greyed with its  reason. */
@Component({
  selector: "mr-finish-banner",
  standalone: true,
  template: `
    @if (state(); as s) {
      <section class="finish" [class.finish--stop]="!s.acceptable" [class.finish--problem]="s.problem" data-show-me="finish-banner">
        <p class="finish__text" role="status">{{ s.text }}</p>
        <div class="finish__buttons">
          <button type="button" class="cds--btn cds--btn--primary cds--btn--lg finish__upload" data-show-me="finish-upload" (click)="upload()" [disabled]="!s.acceptable || running()" [title]="s.reason">
            {{ running() ? "Uploading…" : "Upload to Mind" }}
          </button>
          @if (downloadUrl(); as url) {
            <a class="cds--btn cds--btn--secondary cds--btn--lg finish__download" data-show-me="finish-download" [href]="url">Download</a>
          }
        </div>
        @if (!s.acceptable && s.reason) {
          <p class="finish__reason">{{ s.reason }}</p>
        }
        @if (askSignIn()) {
          <div class="finish__signin">
            <p>{{ signInLine() }}</p>
            <button type="button" class="cds--btn cds--btn--tertiary cds--btn--md finish__signin-open" (click)="openSignIn()" [disabled]="signIn()?.state === 'waiting'">Sign in to Mind</button>
          </div>
        }
        @if (progress()) {
          <p class="finish__progress">{{ progress() }}</p>
        }
        @if (outcome()) {
          <p class="finish__outcome" [class.finish__outcome--error]="failed()">{{ outcome() }}</p>
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
      .finish__progress,
      .finish__signin p {
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
        margin: 0.5rem 0 0;
      }
      .finish__outcome {
        margin: 0.5rem 0 0;
        font-weight: 600;
      }
      .finish__outcome--error {
        color: var(--cds-text-error, #da1e28);
      }
    `,
  ],
})
export class FinishBanner implements OnDestroy {
  private api = inject(ApiService);
  private store = inject(SessionStore);

  /** the  verdict, from  the page */
  readonly readiness = input<Readiness | null | undefined>(null);
  /** the current version's download address */
  readonly downloadUrl = input<string | null>(null);

  readonly state = computed(() => finishState(this.readiness()));
  readonly running = signal(false);
  readonly progress = signal("");
  readonly outcome = signal("");
  readonly failed = signal(false);
  readonly signIn = signal<MindSignIn | null>(null);
  readonly askSignIn = signal(false);
  private live = true;
  private timer: ReturnType<typeof setTimeout> | null = null;

  ngOnDestroy(): void {
    this.live = false;
    if (this.timer) clearTimeout(this.timer);
  }

  signInLine(): string {
    const s = this.signIn();
    if (s?.state === "waiting") return s.message || "An Edge window is open on Mind: sign in there, then the upload goes on by itself.";
    if (s?.state === "failed" && s.message) return s.message;
    return "Not signed in to Mind, or the sign-in has expired. Sign in once (it lasts about a week); the upload then goes on by itself.";
  }

  /** Upload to Mind: the sign-in first (asked for when missing), then the same run as Run in Mind. */
  async upload(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.running() || !this.state()?.acceptable) return;
    this.outcome.set("");
    this.failed.set(false);
    this.running.set(true);
    this.progress.set("Checking the Mind sign-in…");
    try {
      const s = await this.api.mindSignIn(true);
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
      this.fail(errorText(e));
    }
  }

  /** Opens the Edge window; when the sign-in is done  the upload starts by itself. */
  async openSignIn(): Promise<void> {
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
        this.signIn.set(s);
        if (s.state === "waiting") return this.watchSignIn();
        if (s.checked?.signed_in || s.state === "done") {
          this.askSignIn.set(false);
          void this.upload();
        }
      } catch {
        this.watchSignIn();
      }
    }, 2000);
  }

  /** The same call as Run in Mind:  the changes it ticks by default,  4 rounds at most, the model run,  the assistant naming grids. */
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
      await this.api.startMindLoop(id, opts);
    } catch (e) {
      this.fail(errorText(e));
      return;
    }
    await this.poll(id, 0);
  }

  private async poll(id: string, after: number): Promise<void> {
    if (!this.live) return;
    let next = after;
    try {
      const s: MindLoopStatus = await this.api.mindLoopStatus(id, after);
      next = s.next ?? after;
      if (s.timing?.stage_title) this.progress.set(`${s.timing.stage_title}…`);
      if (s.state === "done" || s.state === "error") {
        this.running.set(false);
        this.progress.set("");
        const verdict = s.report?.verdict;
        this.failed.set(s.state === "error" || (!!verdict && verdict !== "converged"));
        this.outcome.set(s.state === "error" ? (s.error ?? "The run in Mind stopped.") : (s.report?.reason ?? s.summary ?? (this.failed() ? "The run in Mind did not finish cleanly." : "The workbook ran in Mind.")));
        return;
      }
    } catch {
      /* the next look  tries again */
    }
    this.timer = setTimeout(() => void this.poll(id, next), 2000);
  }

  private fail(message: string): void {
    this.running.set(false);
    this.progress.set("");
    this.failed.set(true);
    this.outcome.set(message);
  }
}
