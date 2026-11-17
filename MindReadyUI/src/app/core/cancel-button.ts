import { Component, OnDestroy, OnInit, inject, input, output, signal } from "@angular/core";
import { ApiService, errorText } from "./api.service";
import { CancelResult, SessionStart } from "./models";
import { SessionStore } from "./session.store";
import { WorkService } from "./work.service";

/** The  Cancel button,  the same one wherever the app works (17 Nov 2026, Joseph: "VERY IMPORTANT").
 *  One click, no question: it stops the work on the server -- one `kind` of work, or everything that
 *  runs  for this workbook when `kind` is null -- and the workbook  goes back to what it was  before that
 *  work began.
 *
 *  While it works it is  greyed and says "Stopping…". It gives two signs to the page  around it:
 *  `stopping` at the click (before the  server is asked: the page starts nothing new from then on) and
 *  `stopped` once the server has really finished, with the workbook as it is now already in the store.
 *  The page does not have to abandon anything: the stopped work's own request answers by itself
 *  (`wasCancelled`), and a job that is followed says "cancelled".
 *
 *  Drawn with Carbon's own classes (as the Recalculate button is),  not through the cdsButton directive:
 *  the look is the same and every screen test can draw it without loading Carbon. */
@Component({
  selector: "mr-cancel",
  standalone: true,
  template: `<button type="button" class="cds--btn cds--btn--danger--tertiary cds--btn--{{ size() }} mr-cancel" data-show-me="cancel" (click)="cancel()" [disabled]="working() || busy()">{{ working() || busy() ? "Stopping…" : "Cancel" }}</button>`,
  styles: [
    `
      :host {
        display: inline-block;
        vertical-align: middle;
      }
    `,
  ],
})
export class CancelButton implements OnInit, OnDestroy {
  private api = inject(ApiService);
  private store = inject(SessionStore);
  private work = inject(WorkService);

  /** how  often a  job still on its way out is looked at, in ms (the tests shorten it) */
  static everyMs = 1000;

  /** which work to stop ("apply", "recalc", "mind" ...); null: everything that runs for this workbook */
  readonly kind = input<string | null>(null);
  /** Carbon's button size */
  readonly size = input<"sm" | "md" | "lg">("md");
  /** the  workbook, when the store does not hold it yet (the first  check after an upload); else the  open one */
  readonly session = input<string | null>(null);
  /** false for the shell's strip: its  button is not one "on screen for the work" (it shows  when none is) */
  readonly counted = input(true);
  /** the page knows the stopped  work is still on its way out (its own request has not answered yet, a
   *  check that cannot be stopped still runs): the button stays greyed and keeps saying "Stopping…" */
  readonly busy = input(false);

  /** at the click, before the server is asked */
  readonly stopping = output<void>();
  /** the server has really finished; the store already holds the workbook as it is now */
  readonly stopped = output<CancelResult>();
  /** the server could not be asked (its words): nothing was stopped, the  button is back */
  readonly failed = output<string>();

  /** the click was made and the server has not finished yet */
  readonly working = signal(false);

  private gone = false;
  private leave: (() => void) | null = null;

  ngOnInit(): void {
    if (this.counted()) this.leave = this.work.show(() => this.kind());
  }

  ngOnDestroy(): void {
    this.gone = true;
    this.leave?.();
  }

  /** The click. A second click while it works does nothing. */
  async cancel(): Promise<void> {
    const id = this.session() ?? this.store.sessionId();
    if (!id || this.working() || this.busy()) return;
    this.working.set(true);
    this.stopping.emit();
    let result: CancelResult;
    try {
      result = await this.api.cancel(id, this.kind());
      const waitFor = result.stopping ?? [];
      if (waitFor.length) {
        // the work is still on its way out (a page of Mind that hangs, one very large sheet): the server
        // has not put the workbook back yet. Look every second until those kinds are gone, then read
        // the workbook as it is.
        await this.untilGone(id, waitFor);
        // (a workbook the store does not hold -- the first check after an upload -- has nothing to read back)
        if (this.store.sessionId() === id) this.keep(id, await this.api.resume(id));
      } else if (result.session) {
        this.keep(id, result.session);
      }
    } catch (e) {
      this.working.set(false);
      if (!this.gone) this.failed.emit(errorText(e));
      return;
    }
    this.working.set(false);
    void this.work.refresh(); // the strip  goes now, not at its  next look
    // the page may have taken the  button away already  (its own  request answered first): the store is
    // right either way, and an output that is gone cannot be given anything
    if (!this.gone) this.stopped.emit(result);
  }

  /** The workbook as the server holds it after the stop, into the store (never  onto another workbook). */
  private keep(id: string, session: SessionStart | null | undefined): void {
    if (!session || this.store.sessionId() !== id) return;
    this.store.applyAnalysis(session);
    this.store.recalc.set(session.recalc ?? null);
  }

  /** Waits until none of `kinds` is in GET .../jobs  any more (a look every second;  a look that fails is tried again). */
  private async untilGone(id: string, kinds: string[]): Promise<void> {
    let failures = 0;
    for (;;) {
      await new Promise<void>((res) => setTimeout(res, CancelButton.everyMs));
      try {
        const jobs = await this.api.jobs(id);
        failures = 0;
        if (!jobs.some((j) => kinds.includes(j.kind))) return;
      } catch (e) {
        if (++failures >= 5) throw e; // the  server is not there any more: say so instead of waiting  for ever
      }
    }
  }
}
