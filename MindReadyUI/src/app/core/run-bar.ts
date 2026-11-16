import { Component, ViewEncapsulation, computed, input } from "@angular/core";
import { ProgressBarModule } from "carbon-components-angular";

/** How  far a  long run is: Carbon's progress bar with the percentage beside it (night of 16 Nov 2026, Joseph:
 *  "some sort of loading bar or animation", so it is clear the app is working  and has not frozen).
 *  With no percentage (`percent` null: there is no  time estimate yet) it is Carbon's  moving bar, and no
 *  number is shown. A  screen reader gets Carbon's own role="progressbar" with the value, named by  `label`.
 *  Motion stops for people who ask their  system for less motion (Carbon's bar alone keeps moving). */
@Component({
  selector: "mr-run-bar",
  standalone: true,
  imports: [ProgressBarModule],
  // the  rules below  reach inside Carbon's bar, so they are not scoped: every one starts with .run-bar
  encapsulation: ViewEncapsulation.None,
  template: `<div class="run-bar" [class.run-bar--moving]="shown() === null">
    <cds-progress-bar class="run-bar__bar" [label]="label()" [hideLabel]="true" [value]="shown() ?? undefined" [max]="100"></cds-progress-bar>
    @if (shown() !== null) {<span class="run-bar__percent" aria-hidden="true">{{ shown() }}%</span>}
  </div>`,
  styles: [
    `
      mr-run-bar {
        display: block;
      }
      .run-bar {
        display: flex;
        align-items: center;
        gap: 0.75rem;
      }
      .run-bar__bar {
        display: block;
        flex: 1 1 auto;
        min-width: 0;
      }
      .run-bar__percent {
        flex: 0 0 2.5rem;
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
        font-variant-numeric: tabular-nums;
      }
      /* no number yet: its room is kept, so the bar is as long as it will be */
      .run-bar--moving .run-bar__bar {
        margin-inline-end: 3.25rem;
      }
      /* the value arrives every 10 seconds: the bar glides to it rather than jumping */
      .run-bar .cds--progress-bar__bar {
        transition: transform 10s linear;
      }
      @media (prefers-reduced-motion: reduce) {
        .run-bar .cds--progress-bar__bar {
          transition: none;
        }
        /* the moving bar stands still: a pale full-width band, so it is not read as "a quarter done" */
        .run-bar .cds--progress-bar--indeterminate .cds--progress-bar__track::after {
          animation: none;
          background-image: none;
          background-color: var(--cds-interactive, #0f62fe);
          inset: 0;
          inline-size: 100%;
          opacity: 0.35;
          transform: none;
        }
      }
    `,
  ],
})
export class RunBar {
  /** 0  to 100;  null while it is not known (the bar only moves) */
  readonly percent = input<number | null>(null);
  /** what a screen reader calls the bar */
  readonly label = input("Work done so far");
  /** the whole number shown and given to the bar */
  readonly shown = computed(() => {
    const p = this.percent();
    return p === null || !Number.isFinite(p) ? null : Math.max(0, Math.min(100, Math.round(p)));
  });
}
