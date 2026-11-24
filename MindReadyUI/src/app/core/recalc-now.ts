import { Component, computed, inject, input, signal } from "@angular/core";
import { ApiService, errorText, wasCancelled } from "./api.service";
import { CancelButton } from "./cancel-button";
import { Readiness, STOPPED_NOTHING_CHANGED } from "./models";
import { SessionStore } from "./session.store";

/** The  recalculation is  the only thing left: nothing blocks, and the current version was not recalculated. */
export function onlyRecalcLeft(r: Readiness | null | undefined): boolean {
  return !!r && r.state === "unverified" && !r.recalc?.ran && !r.blocking_count;
}

/** "Recalculate"  wherever the  verdict is shown (16 Nov 2026): it appears any time the recalculation is the
 *  only thing left, and does what the Recalculate tab's button does, so nobody has to look for that  tab.
 *  17 Nov 2026: while it recalculates, Cancel stands beside it; a  stopped recalculation changes nothing
 *  (the last good  one stays) and the button is back  to "Recalculate". */
@Component({
  selector: "mr-recalc-now",
  standalone: true,
  imports: [CancelButton],
  template: `
    @if (show()) {
      <div class="recalc-now">
        <button type="button" class="cds--btn cds--btn--{{ kind() }} cds--btn--{{ size() }} recalc-now__button" data-show-me="recalc-now" (click)="run()" [disabled]="running() || disabled()">
          {{ running() ? "Recalculating…" : "Recalculate" }}
        </button>
        @if (running()) {
          <mr-cancel kind="recalc" [size]="size()"></mr-cancel>
          <span class="recalc-now__note">Excel recalculates the whole workbook: tens of seconds to a few minutes.</span>
        } @else if (error()) {
          <span class="recalc-now__note recalc-now__note--error">{{ error() }}</span>
        } @else if (note()) {
          <span class="recalc-now__note" role="status">{{ note() }}</span>
        }
      </div>
    }
  `,
  styles: [
    `
      :host {
        display: contents;
      }
      .recalc-now {
        display: inline-flex;
        align-items: center;
        gap: 0.75rem;
        flex-wrap: wrap;
      }
      .recalc-now__note {
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
      }
      .recalc-now__note--error {
        color: var(--cds-text-error, #da1e28);
      }
    `,
  ],
})
export class RecalcNow {
  private api = inject(ApiService);
  private store = inject(SessionStore);

  /** Carbon  button kind  and size */
  readonly kind = input<"primary" | "secondary" | "tertiary">("primary");
  readonly size = input<"md" | "lg">("lg");
  /** another job is running on the page */
  readonly disabled = input(false);

  readonly running = signal(false);
  readonly error = signal("");
  /** 17 Nov 2026: after a Cancel: "Stopped. Nothing was changed." (calm: it is not an error) */
  readonly note = signal("");
  readonly show = computed(() => this.running() || onlyRecalcLeft(this.store.readiness()));

  async run(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.running()) return;
    this.running.set(true);
    this.error.set("");
    this.note.set("");
    try {
      const out = await this.api.recalculate(id);
      this.store.recalc.set(out);
      this.store.applyAnalysis({ report: out.report, plan: out.plan, readiness: out.readiness, version: out.version ?? undefined });
    } catch (e) {
      // 17 Nov 2026: stopped with Cancel: no error, and  nothing of this call reaches the store (the
      // workbook as it is  comes from the Cancel's own answer)
      if (wasCancelled(e)) this.note.set(STOPPED_NOTHING_CHANGED);
      else this.error.set(errorText(e));
    } finally {
      this.running.set(false);
    }
  }
}
