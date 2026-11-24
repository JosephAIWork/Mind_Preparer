import { Component, input } from "@angular/core";
import { NotificationModule, NotificationType } from "carbon-components-angular";
import { ShowMeNextStep } from "./show-me";
import { ShowMeLink } from "./show-me-link";

/** The  readiness verdict  as Carbon's inline notification, with "Show me" after the next step: the
 *  same banner on the main screen and on Advanced.
 *
 *  Carbon's own template is replaced to put the link inside the words,  so the title carries the
 *  notification's id itself, as Carbon's does: the  notification is labelled by it (aria-labelledby),
 *  and  a screen reader names the banner by  the verdict ("Acceptable by Mind…"). */
@Component({
  selector: "mr-readiness-banner",
  standalone: true,
  imports: [NotificationModule, ShowMeLink],
  template: `
    <cds-inline-notification #banner [notificationObj]="{ type: kind(), title: headline(), message: detail(), lowContrast: true, showClose: false, template: words }"></cds-inline-notification>
    <ng-template #words let-n>
      <div class="cds--inline-notification__title" [id]="banner.notificationID">{{ n.title }}</div>
      <div class="cds--inline-notification__subtitle">
        <span>{{ n.message }}</span><mr-show-me [step]="step()"></mr-show-me>
      </div>
    </ng-template>
  `,
  styles: [":host { display: block; }"],
})
export class ReadinessBanner {
  /** Carbon's  colour: success  (ready), warning (not verified), error (blocked) */
  readonly kind = input<NotificationType>("info");
  /** the verdict itself ("Not acceptable by Mind yet: 3 blocking problem(s)") */
  readonly headline = input("");
  /** the next step and the counts, in words */
  readonly detail = input("");
  /** the readiness verdict's next step: its "Show me" */
  readonly step = input<ShowMeNextStep | null | undefined>(null);
}
