import { DOCUMENT, DestroyRef, Injectable, computed, inject, signal } from "@angular/core";
import { Router } from "@angular/router";
import { SessionStore } from "./session.store";
import { PAGE_PATHS, ShowMePage, ShowMeRequest, ShowMeStep, ShowMeTarget, ShowMeWhere, planShowMe } from "./show-me";

/** What  the pointer  shows now: the element, the bubble's words, and whether it is the place itself
 *  (outlined, with "Got it"). */
export interface ShowMePointer {
  el: HTMLElement;
  text: string;
  last: boolean;
}

/** How  long a  step's element is waited for (a screen or a tab still being drawn) before the pointer
 *  ends quietly. */
export const SHOW_ME_WAIT_MS = 2500;
/** How  often a  missing element is looked for. */
export const SHOW_ME_LOOK_MS = 100;
/** After  the click  the arrow asked for: the screen or the tab changes first, then the next step is looked for. */
export const SHOW_ME_AFTER_CLICK_MS = 60;
/** How  often the  pointed-at element is checked: Angular may draw it again, and the pointer follows the new one. */
export const SHOW_ME_WATCH_MS = 300;

/** The  "Show me"  pointer: one arrow at a time, at the next thing to click, until the place itself.
 *
 *  `start` plans the steps from the screen the user is on (show-me.ts); each step is then  looked for
 *  in the page by its `data-show-me` words (the header's link,  else the page's own: in a  narrow window
 *  Carbon hides the header's). A step  the user is already past (that screen, that tab) is skipped;  when
 *  a step's element does not come, the  pointer ends quietly -- a click on the way that cannot be shown
 *  leaves nothing after it to reach, and the pointer never stops the page. A click on the element the
 *  arrow points at moves it on (at the place itself: ends it, the user acts there); Esc, "Got it" or a
 *  click anywhere else ends it, but not the second click of  a double click, nor a click again on the
 *  element just clicked (a screen slow to change). The clicks are watched on the whole document in the
 *  capture phase, so the element's own handler (a router link, a tab) still runs as usual: the overlay
 *  never takes a click.
 *
 *  17 Nov 2026: a request with `open` (a stop message on Basic) is the one case where the service moves
 *  the page itself: it goes to the screen as the header's link would, opens the Advanced tab,  and points
 *  at the place only.  A request's `where` (the cell the problem  is in) is  kept in  the signal `where` while
 *  the pointer lives: the panel that draws the line naming that cell reads it there. */
@Injectable({ providedIn: "root" })
export class ShowMeService {
  private readonly router = inject(Router);
  private readonly store = inject(SessionStore);
  private readonly doc = inject(DOCUMENT);

  /** the  steps of  the pointer going on (empty: none) */
  readonly steps = signal<ShowMeStep[]>([]);
  /** the step it is at */
  readonly at = signal(0);
  /** what it points at now (null while the element is looked for) */
  readonly pointer = signal<ShowMePointer | null>(null);
  readonly active = computed(() => this.steps().length > 0);
  /** the request of the pointer going on (null: none): its place,  its rule */
  readonly request = signal<ShowMeRequest | null>(null);
  /** the cell the pointer going on is about (null: none,  or a pointer with no cell).  Set when the pointer
   *  starts, null again when  it ends: a panel that opened a list for it keeps  that in its own state. */
  readonly where = signal<ShowMeWhere | null>(null);

  private lookTimer: ReturnType<typeof setTimeout> | null = null;
  private watchTimer: ReturnType<typeof setInterval> | null = null;
  /** the element of  the step just clicked: a click on it again is no click elsewhere */
  private passed: HTMLElement | null = null;
  /** which of the step's targets the arrow is at (0: the best one) */
  private rank = 0;
  /** counts the pointers started and ended: a screen that opens late belongs to the pointer that asked for it */
  private run = 0;

  constructor() {
    const click = (e: Event) => this.onClick(e);
    const key = (e: KeyboardEvent) => this.onKey(e);
    this.doc.addEventListener("click", click, true);
    this.doc.addEventListener("keydown", key, true);
    inject(DestroyRef).onDestroy(() => {
      this.doc.removeEventListener("click", click, true);
      this.doc.removeEventListener("keydown", key, true);
      this.stop();
    });
  }

  /** Starts a pointer to the place (a pointer going on ends first). False when there is no way there.
   *  With  `request.open` the service goes there itself and points at the place only. */
  start(request: ShowMeRequest): boolean {
    this.stop();
    const steps = planShowMe(request, this.page());
    if (!steps.length) return false;
    const where = request.where?.sheet && request.where?.cell ? { sheet: request.where.sheet, cell: request.where.cell } : null;
    this.request.set(request);
    this.where.set(where);
    this.watchTimer = setInterval(() => this.watch(), SHOW_ME_WATCH_MS);
    if (request.open) {
      // only the place is pointed at; the clicks on the way are made here
      this.steps.set(steps.filter((s) => s.kind === "place"));
      void this.openWay(steps, this.run);
    } else {
      this.steps.set(steps);
      this.go(0, 0);
    }
    return true;
  }

  /** Ends the pointer (Esc, "Got it", a click elsewhere, the place reached or not found). */
  stop(): void {
    this.run++;
    if (this.lookTimer) clearTimeout(this.lookTimer);
    if (this.watchTimer) clearInterval(this.watchTimer);
    this.lookTimer = null;
    this.watchTimer = null;
    this.passed = null;
    this.rank = 0;
    this.pointer.set(null);
    this.at.set(0);
    if (this.steps().length) this.steps.set([]);
    if (this.request()) this.request.set(null);
    if (this.where()) this.where.set(null);
  }

  /** `open`: the clicks on the way, made by the service: the screen first (as the header's link goes
   *  there, with the open workbook in the address), then the Advanced tab; then the place is looked for.
   *  A pointer ended meanwhile (Esc, another Show me) is  left ended: `run` has moved  on. A screen that
   *  did not open ends  the pointer quietly,  as a  step that never comes does. */
  private async openWay(steps: ShowMeStep[], run: number): Promise<void> {
    try {
      for (const step of steps) {
        if (step.kind === "page" && step.page && this.page() !== step.page) {
          const id = this.store.sessionId();
          await this.router.navigate([PAGE_PATHS[step.page]], { queryParams: id ? { session: id } : {} });
          if (run !== this.run) return;
          if (this.page() !== step.page) {
            this.stop();
            return;
          }
        } else if (step.kind === "tab" && step.tab) {
          this.store.advancedTab.set(step.tab);
        }
      }
    } catch {
      if (run === this.run) this.stop();
      return;
    }
    if (run === this.run) this.go(0, 0);
  }

  /** The screen the user is on, from the router's address ("/advanced?session=..." is Advanced). */
  page(): ShowMePage {
    const path = this.router.url.split(/[?#]/)[0].replace(/\/+$/, "");
    if (path === "/advanced" || path.startsWith("/advanced/")) return "advanced";
    return path === "/custom" || path.startsWith("/custom/") ? "custom" : "basic";
  }

  /** Step `i`, looked for now (`delay` 0) or after the click's effects (`delay` ms), then  every
   *  SHOW_ME_LOOK_MS until it is there, the user is past it, or  SHOW_ME_WAIT_MS have gone. */
  private go(i: number, delay: number): void {
    if (this.lookTimer) clearTimeout(this.lookTimer);
    this.lookTimer = null;
    this.pointer.set(null);
    if (i >= this.steps().length) {
      this.stop();
      return;
    }
    this.at.set(i);
    const since = Date.now();
    const look = (): void => {
      this.lookTimer = null;
      const step = this.steps()[i];
      if (!step || this.at() !== i) return; // ended, or moved on  meanwhile
      if (this.reached(step)) {
        this.go(i + 1, 0);
        return;
      }
      const found = this.find(step.targets);
      if (found) {
        this.rank = found.rank;
        this.pointer.set({ el: found.el, text: found.text, last: step.kind === "place" });
        this.reveal(found.el);
        return;
      }
      if (Date.now() - since >= SHOW_ME_WAIT_MS) {
        // never came: the pointer ends quietly. Skipping  a click on the way would only wait for steps
        // that cannot  come (the tab and the place are  on the screen that click leads to).
        this.stop();
        return;
      }
      this.lookTimer = setTimeout(look, SHOW_ME_LOOK_MS);
    };
    if (delay > 0) this.lookTimer = setTimeout(look, delay);
    else look();
  }

  /** The user is already where a click on the way leads: that screen, or that Advanced tab. */
  private reached(step: ShowMeStep): boolean {
    if (step.kind === "page") return this.page() === step.page;
    if (step.kind === "tab") return this.page() === "advanced" && this.store.advancedTab() === step.tab;
    return false;
  }

  /** The first target drawn on the page (an element with no size is hidden: not pointed at), and which
   *  of the targets it is. */
  private find(targets: ShowMeTarget[]): { el: HTMLElement; text: string; rank: number } | null {
    for (const [rank, t] of targets.entries()) {
      const word = t.el.replace(/["\\]/g, "\\$&");
      for (const el of Array.from(this.doc.querySelectorAll<HTMLElement>(`[data-show-me~="${word}"]`))) {
        const r = el.getBoundingClientRect();
        if (el.isConnected && r.width > 0 && r.height > 0) return { el, text: t.text, rank };
      }
    }
    return null;
  }

  /** Scrolls the element into the window when it is not in it (the fixed header never needs  it). */
  private reveal(el: HTMLElement): void {
    const view = this.doc.defaultView;
    if (!view || el.closest(".cds--header")) return;
    const r = el.getBoundingClientRect();
    if (r.top >= 64 && r.bottom <= view.innerHeight - 16) return;
    const calm = view.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
    el.scrollIntoView?.({ block: r.height > view.innerHeight - 200 ? "start" : "center", behavior: calm ? "auto" : "smooth" });
  }

  /** Angular drew the element again, hid it or took it away: look for it again; gone for good, the
   *  pointer ends, as when it is not found at first. At the place itself, a better target drawn a moment
   *  after the one pointed at takes the arrow (17 Nov 2026: the line of the cell's own change comes when
   *  the Prep panel has opened its list, the action's card is there before it). */
  private watch(): void {
    const p = this.pointer();
    if (!p) return;
    const r = p.el.getBoundingClientRect();
    if (!p.el.isConnected || (r.width === 0 && r.height === 0)) {
      this.go(this.at(), 0);
      return;
    }
    const step = this.steps()[this.at()];
    if (!p.last || !step || this.rank === 0) return;
    const better = this.find(step.targets.slice(0, this.rank));
    if (!better) return;
    this.rank = better.rank;
    this.pointer.set({ el: better.el, text: better.text, last: true });
    this.reveal(better.el);
  }

  private onClick(e: Event): void {
    if (!this.active()) return;
    const target = e.target as (Node & { closest?: (s: string) => Element | null }) | null;
    if (target?.closest?.("[data-show-me-ui]")) return; // the bubble's own "Got it": its handler ends the pointer
    // The second click  of a double click (the  browser counts them in `detail`): the first  one already
    // counted. It  may land  where the next element is not drawn yet, or on the next screen, which came
    // under the mouse: neither a click elsewhere nor the next click asked for.
    if ((e as MouseEvent).detail > 1) return;
    const p = this.pointer();
    if (p && target && p.el.contains(target)) {
      // the click the arrow asked  for: on to the next step once the screen or the  tab has changed; at
      // the place itself  the user acts there, and the pointer  goes
      if (p.last) this.stop();
      else {
        this.passed = p.el;
        this.go(this.at() + 1, SHOW_ME_AFTER_CLICK_MS);
      }
      return;
    }
    // the element just clicked, clicked again (the screen is slow to  change): nothing new happened
    if (target && this.passed?.contains(target)) return;
    this.stop(); // a click anywhere else  ends it
  }

  private onKey(e: KeyboardEvent): void {
    if (e.key === "Escape" && this.active()) this.stop();
  }
}
