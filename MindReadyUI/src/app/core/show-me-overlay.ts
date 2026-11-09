import { Component, DOCUMENT, ElementRef, NgZone, effect, inject, signal, viewChild } from "@angular/core";
import { PLACE_SIDES, STEP_SIDES, ShowMeBox, ShowMeLayout, ShowMeSide, placeShowMe } from "./show-me";
import { ShowMeService } from "./show-me.service";

/** The arrow's shapes (a stroked arrow, as Carbon's arrow icons), by the side the bubble is on: it
 *  points from the bubble at the element. */
const ARROWS: Record<ShowMeLayout["side"], string> = {
  below: "M10 24 V3 M3 10 L10 3 L17 10",
  inside: "M10 24 V3 M3 10 L10 3 L17 10",
  above: "M10 2 V23 M3 16 L10 23 L17 16",
  right: "M24 10 H3 M10 3 L3 10 L10 17",
  left: "M2 10 H23 M16 3 L23 10 L16 17",
};

/** What a bubble beside the place must not cover: anything the user may click there. */
const CLICKABLE = "a[href], button, input, select, textarea, label, [role=button], [role=tab], [role=link], [tabindex]:not([tabindex='-1'])";

/** The pointer on screen, placed once in the shell: one arrow with a short bubble at the element the
 *  ShowMeService points at, following it (scrolling, resizing, the page moving under it); at the place
 *  itself a soft outline and "Got it", on a side where the bubble covers no other button. It sits above
 *  everything and takes no click (pointer-events: none) except its own button. */
@Component({
  selector: "mr-show-me-overlay",
  standalone: true,
  templateUrl: "./show-me-overlay.html",
  styleUrl: "./show-me-overlay.scss",
})
export class ShowMeOverlay {
  readonly showMe = inject(ShowMeService);
  private readonly zone = inject(NgZone);
  private readonly doc = inject(DOCUMENT);
  private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
  private readonly bubble = viewChild<ElementRef<HTMLElement>>("bubble");
  /** where the arrow, the bubble and the outline go now (null: nothing to draw yet, or the element has no size) */
  readonly layout = signal<ShowMeLayout | null>(null);

  constructor() {
    effect((onCleanup) => {
      const p = this.showMe.pointer();
      this.layout.set(null); // each new step fades in at its own place
      const view = this.doc.defaultView;
      if (!p || !view?.requestAnimationFrame) return;
      // Looked at every frame, outside Angular; placed again only when the element, the bubble or the
      // window changed size or place (the layout signal changes only then).
      let seen = "";
      let side: ShowMeSide | null = null;
      let frame = 0;
      const free = (box: ShowMeBox) => this.clear(box, p.el);
      const tick = () => {
        const r = p.el.getBoundingClientRect();
        const b = this.bubble()?.nativeElement;
        const size = { width: b?.offsetWidth || 240, height: b?.offsetHeight || 40 };
        const now = [r.top, r.left, r.width, r.height, size.width, size.height, view.innerWidth, view.innerHeight].join(",");
        if (now !== seen) {
          seen = now;
          const next = r.width > 0 || r.height > 0 ? placeShowMe(r, size, { width: view.innerWidth, height: view.innerHeight }, p.last ? PLACE_SIDES : STEP_SIDES, p.last ? free : undefined, side) : null;
          side = next && next.side !== "inside" ? next.side : side;
          this.layout.set(next);
        }
        frame = view.requestAnimationFrame(tick);
      };
      this.zone.runOutsideAngular(() => (frame = view.requestAnimationFrame(tick)));
      onCleanup(() => view.cancelAnimationFrame(frame));
    });
  }

  /** The arrow's shape for the side the bubble is on. */
  arrow(side: ShowMeLayout["side"]): string {
    return ARROWS[side];
  }

  /** True when the box covers nothing to click (its corners and middle, as the page sees them; the
   *  pointer's own bubble does not count, the place itself does: never cover it). */
  private clear(box: ShowMeBox, place: HTMLElement): boolean {
    if (!this.doc.elementsFromPoint) return true;
    const own = this.host.nativeElement;
    const x = [box.left + 3, box.left + box.width / 2, box.left + box.width - 3];
    const y = [box.top + 3, box.top + box.height / 2, box.top + box.height - 3];
    for (const px of x) {
      for (const py of y) {
        const hit = this.doc.elementsFromPoint(px, py).find((e) => !own.contains(e));
        if (hit && (place.contains(hit) || hit.closest(CLICKABLE))) return false;
      }
    }
    return true;
  }
}
