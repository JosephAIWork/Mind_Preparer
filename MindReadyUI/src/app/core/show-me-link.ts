import { Component, computed, inject, input } from "@angular/core";
import { ShowMeId, ShowMeNextStep, ShowMeRequest, showMeFor, showMeForStep } from "./show-me";
import { ShowMeService } from "./show-me.service";

/** The quiet "Show me" link next to an instruction. It appears only when there is a place to show:
 *  - `[text]`: a text the app shows (a result, an error, the server's words): when it names a place;
 *  - `to`: a place named in the template itself ("Change it in Advanced › Prep actions");
 *  - `[step]`: the readiness verdict's next step (by its screen).
 *  `[unless]`: another text shown next to this one; when it leads to the same place, its own link is
 *  enough (one link per place). The click starts the pointer (ShowMeService). */
@Component({
  selector: "mr-show-me",
  standalone: true,
  template: `@if (request(); as r) {<button type="button" class="show-me-link" title="Shows where to click, one step at a time" (click)="go(r)">Show me</button>}`,
  styleUrl: "./show-me-link.scss",
})
export class ShowMeLink {
  private readonly showMe = inject(ShowMeService);
  readonly text = input<string | null | undefined>(null);
  readonly to = input<ShowMeId | null | undefined>(null);
  readonly step = input<ShowMeNextStep | null | undefined>(null);
  readonly unless = input<string | null | undefined>(null);

  readonly request = computed<ShowMeRequest | null>(() => {
    const to = this.to();
    const step = this.step();
    const r = to ? { id: to } : step ? showMeForStep(step) : showMeFor(this.text());
    if (!r) return null;
    const other = showMeFor(this.unless());
    return other && other.id === r.id ? null : r;
  });

  go(r: ShowMeRequest): void {
    this.showMe.start(r);
  }
}
