import { DOCUMENT, DestroyRef, Injectable, computed, inject, signal } from "@angular/core";
import { Router } from "@angular/router";
import { SessionStore } from "./session.store";
import { ShowMePage, ShowMeRequest, ShowMeStep, ShowMeTarget, planShowMe } from "./show-me";

/** What the pointer shows now: the element, the bubble's words, and whether it is the place itself
 *  (outlined, with "Got it"). */
export interface ShowMePointer {
  el: HTMLElement;
  text: string;
  last: boolean;
}

/** How long a step's element is waited for (a screen or a tab still being drawn) before the step is
 *  skipped -- or, for the place itself, before the pointer ends quietly. */
export const SHOW_ME_WAIT_MS = 2500;
/** How often a missing element is looked for. */
export const SHOW_ME_LOOK_MS = 100;
/** After the click the arrow asked for: the screen or the tab changes first, then the next step is looked for. */
export const SHOW_ME_AFTER_CLICK_MS = 60;
/** How often the pointed-at element is checked: Angular may draw it again, and the pointer follows the new one. */
export const SHOW_ME_WATCH_MS = 300;

/** The "Show me" pointer: one arrow at a time, at the next thing to click, until the place itself.
 *
 *  `start` plans the steps from the screen the user is on (show-me.ts); each step is then looked for
 *  in the page by its `data-show-me` word. A step the user is already past (that screen, that tab) is
 *  skipped; one whose element does not come is skipped too, and at the place itself the pointer ends
 *  quietly: it never stops the page. A click on the element the arrow points at moves it on (at the
 *  place itself: ends it, the user acts there); Esc, "Got it" or a click anywhere else ends it. The
 *  clicks are watched on the whole document in the capture phase, so the element's own handler (a
 *  router link, a tab) still runs as usual: the overlay never takes a click. */
@Injectable({ providedIn: "root" })
export class ShowMeService {
  private readonly router = inject(Router);
  private readonly store = inject(SessionStore);
  private readonly doc = inject(DOCUMENT);

  /** the steps of the pointer going on (empty: none) */
  readonly steps = signal<ShowMeStep[]>([]);
  /** the step it is at */
  readonly at = signal(0);
  /** what it points at now (null while the element is looked for) */
  readonly pointer = signal<ShowMePointer | null>(null);
  readonly active = computed(() => this.steps().length > 0);

  private lookTimer: ReturnType<typeof setTimeout> | null = null;
  private watchTimer: ReturnType<typeof setInterval> | null = null;

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

  /** Starts a pointer to the place (a pointer going on ends first). False when there is no way there. */
  start(request: ShowMeRequest): boolean {
    this.stop();
    const steps = planShowMe(request, this.page());
    if (!steps.length) return false;
    this.steps.set(steps);
    this.watchTimer = setInterval(() => this.watch(), SHOW_ME_WATCH_MS);
    this.go(0, 0);
    return true;
  }

  /** Ends the pointer (Esc, "Got it", a click elsewhere, the place reached or not found). */
  stop(): void {
    if (this.lookTimer) clearTimeout(this.lookTimer);
    if (this.watchTimer) clearInterval(this.watchTimer);
    this.lookTimer = null;
    this.watchTimer = null;
    this.pointer.set(null);
    this.at.set(0);
    if (this.steps().length) this.steps.set([]);
  }

  /** The screen the user is on, from the router's address ("/advanced?session=..." is Advanced). */
  page(): ShowMePage {
    const path = this.router.url.split(/[?#]/)[0].replace(/\/+$/, "");
    return path === "/advanced" || path.startsWith("/advanced/") ? "advanced" : "main";
  }

  /** Step `i`, looked for now (`delay` 0) or after the click's effects (`delay` ms), then every
   *  SHOW_ME_LOOK_MS until it is there, the user is past it, or SHOW_ME_WAIT_MS have gone. */
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
      if (!step || this.at() !== i) return; // ended, or moved on meanwhile
      if (this.reached(step)) {
        this.go(i + 1, 0);
        return;
      }
      const found = this.find(step.targets);
      if (found) {
        this.pointer.set({ el: found.el, text: found.text, last: step.kind === "place" });
        this.reveal(found.el);
        return;
      }
      if (Date.now() - since >= SHOW_ME_WAIT_MS) {
        // never came: a click on the way is skipped; the place itself missing ends the pointer quietly
        if (step.kind === "place") this.stop();
        else this.go(i + 1, 0);
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

  /** The first target drawn on the page (an element with no size is hidden: not pointed at). */
  private find(targets: ShowMeTarget[]): { el: HTMLElement; text: string } | null {
    for (const t of targets) {
      const word = t.el.replace(/["\\]/g, "\\$&");
      for (const el of Array.from(this.doc.querySelectorAll<HTMLElement>(`[data-show-me~="${word}"]`))) {
        const r = el.getBoundingClientRect();
        if (el.isConnected && r.width > 0 && r.height > 0) return { el, text: t.text };
      }
    }
    return null;
  }

  /** Scrolls the element into the window when it is not in it (the fixed header never needs it). */
  private reveal(el: HTMLElement): void {
    const view = this.doc.defaultView;
    if (!view || el.closest(".cds--header")) return;
    const r = el.getBoundingClientRect();
    if (r.top >= 64 && r.bottom <= view.innerHeight - 16) return;
    const calm = view.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
    el.scrollIntoView?.({ block: r.height > view.innerHeight - 200 ? "start" : "center", behavior: calm ? "auto" : "smooth" });
  }

  /** Angular drew the element again, hid it or took it away: look for it again; gone for good, the
   *  step is skipped or the pointer ends, as when it is not found at first. */
  private watch(): void {
    const p = this.pointer();
    if (!p) return;
    const r = p.el.getBoundingClientRect();
    if (!p.el.isConnected || (r.width === 0 && r.height === 0)) this.go(this.at(), 0);
  }

  private onClick(e: Event): void {
    if (!this.active()) return;
    const target = e.target as (Node & { closest?: (s: string) => Element | null }) | null;
    if (target?.closest?.("[data-show-me-ui]")) return; // the bubble's own "Got it": its handler ends the pointer
    const p = this.pointer();
    if (p && target && p.el.contains(target)) {
      // the click the arrow asked for: on to the next step once the screen or the tab has changed; at
      // the place itself the user acts there, and the pointer goes
      if (p.last) this.stop();
      else this.go(this.at() + 1, SHOW_ME_AFTER_CLICK_MS);
      return;
    }
    this.stop(); // a click anywhere else ends it
  }

  private onKey(e: KeyboardEvent): void {
    if (e.key === "Escape" && this.active()) this.stop();
  }
}
