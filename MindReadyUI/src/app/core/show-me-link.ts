import { Component, computed, inject, input } from "@angular/core";
import { ShowMeId, ShowMeNextStep, ShowMeRequest, showMeFor, showMeForStep } from "./show-me";
import { ShowMeService } from "./show-me.service";

/** The quiet "Show me" link next to an instruction. It appears only when there is a place to show:
 *  - `[text]`: a text the app shows (a result, an error, the server's words): when it names a place;
 *  - `to`: a place named in the template itself ("Change it in Advanced › Prep actions");
 *  - `[step]`: the readiness verdict's next step (by its screen).
 *  `[rule]`: the rule the text is about (a finding's message): the pointer goes to that rule's own row or
 *  Prep action. `except`: a place already next to the text (a Findings row's own "Ask the assistant"):
 *  no link to it, the next place the text names counts. `[unless]`: another text shown next to this
 *  one; when it leads to the same place, its own link is enough (one link per place). The click starts
 *  the pointer (ShowMeService). */
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
  readonly rule = input<string | null | undefined>(null);
  readonly except = input<ShowMeId | null | undefined>(null);
  readonly unless = input<string | null | undefined>(null);

  readonly request = computed<ShowMeRequest | null>(() => {
    const to = this.to();
    const step = this.step();
    const found = to ? { id: to } : step ? showMeForStep(step) : showMeFor(this.text(), this.except());
    if (!found) return null;
    const rule = this.rule()?.trim();
    const r: ShowMeRequest = rule && !found.ruleId ? { ...found, ruleId: rule } : found;
    const other = showMeFor(this.unless());
    return other && other.id === r.id ? null : r;
  });

  go(r: ShowMeRequest): void {
    this.showMe.start(r);
  }
}
