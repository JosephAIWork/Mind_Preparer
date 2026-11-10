import { NgTemplateOutlet } from "@angular/common";
import { Component, Input, TemplateRef } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import { provideRouter } from "@angular/router";
import type { ReadinessBanner } from "./readiness-banner";

/** Carbon's inline notification as far as the banner relies on it (carbon-components-angular 5,
 *  `Notification`): the host is labelled by `notificationID` (aria-labelledby="notification-N"), and a
 *  custom template is drawn with the notification object as its $implicit. Carbon itself cannot load
 *  under Node here (its icons are imported without file extensions). */
@Component({
  selector: "cds-inline-notification",
  standalone: true,
  imports: [NgTemplateOutlet],
  host: { "[attr.aria-labelledby]": "notificationID" },
  template: `<ng-container *ngTemplateOutlet="notificationObj.template ?? null; context: { $implicit: notificationObj }"></ng-container>`,
})
class CarbonNotification {
  @Input() notificationObj!: { type: string; title: string; message?: string; template?: TemplateRef<unknown> };
  notificationID = "notification-7";
}

describe("the readiness banner", () => {
  let Banner: typeof ReadinessBanner;

  beforeAll(async () => {
    // The spec files share one module registry here (the builder runs them without isolation): a page
    // spec that ran before this one in the same worker may have loaded the banner under its own inert
    // Carbon. Let those modules go, so the banner is compiled against the stand-in below.
    vi.resetModules();
    // The banner is compiled ahead of time against Carbon: its template names Carbon's `Notification`
    // component directly, and the module it imports only matters at build time.
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : name === "Notification" ? CarbonNotification : Inert), has: () => true });
    });
    Banner = (await import("./readiness-banner")).ReadinessBanner;
  });

  beforeEach(() => TestBed.configureTestingModule({ providers: [provideRouter([])] }));

  function draw(step: { screen: string; label: string } | null): HTMLElement {
    const f = TestBed.createComponent(Banner);
    f.componentRef.setInput("kind", "warning");
    f.componentRef.setInput("headline", "Acceptable by Mind, not verified: run the Excel recalculation");
    f.componentRef.setInput("detail", "Recalculate in Excel: the only way the app can call the file ready");
    f.componentRef.setInput("step", step);
    f.detectChanges();
    return f.nativeElement as HTMLElement;
  }

  it("is named by its verdict: the title carries the id the notification is labelled by", () => {
    const host = draw({ screen: "recalculate", label: "Recalculate in Excel" });
    const label = host.querySelector("cds-inline-notification")?.getAttribute("aria-labelledby");
    expect(label).toBe("notification-7");
    const title = host.querySelector(`[id="${label}"]`);
    expect(title?.classList.contains("cds--inline-notification__title")).toBe(true);
    expect(title?.textContent).toBe("Acceptable by Mind, not verified: run the Excel recalculation");
  });

  it("says the next step with its Show me, and no link without one", () => {
    const host = draw({ screen: "recalculate", label: "Recalculate in Excel" });
    const words = host.querySelector(".cds--inline-notification__subtitle");
    expect(words?.textContent).toContain("Recalculate in Excel: the only way the app can call the file ready");
    expect(words?.querySelector("button")?.textContent).toBe("Show me");
    expect(draw(null).querySelector(".cds--inline-notification__subtitle button")).toBeNull();
  });
});
