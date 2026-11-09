import { HttpErrorResponse, HttpInterceptorFn } from "@angular/common/http";
import { Injectable, computed, inject, signal } from "@angular/core";
import { tap } from "rxjs";
import { AppHealth } from "./models";

/** A shared copy's end date (app/expiry.py, docs/SHARING.md): the quiet "Works until" in the header, a
 *  gentle line in the last 7 days, and the full-page notice once the copy stopped working. The developer
 *  version has no end date: nothing is shown. */

/** The gentle line shows in the last 7 days (days left 6 ... 0; 0 is the last day). */
export const WARN_DAYS = 7;
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
/** When a 410 comes without words (it should not): still plain. */
const STOPPED = "This copy of Mind Ready stopped working. Ask the person who shared it for a new copy.";

/** "2026-12-15" -> "15 Dec 2026", read as written (a Date would move it a day in some time zones); null when not a date. */
export function shortDate(iso: string | null | undefined): string | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso ?? "");
  const month = m ? MONTHS[Number(m[2]) - 1] : undefined;
  return m && month ? `${Number(m[3])} ${month} ${m[1]}` : null;
}

/** The quiet words in the header; null without an end date. */
export function untilLine(expires: string | null | undefined): string | null {
  const day = shortDate(expires);
  return day ? `Works until ${day}` : null;
}

/** The gentle line in the last 7 days; null before them, after the date, or without one. */
export function soonLine(expires: string | null | undefined, daysLeft: number | null | undefined): string | null {
  const day = shortDate(expires);
  if (!day || daysLeft == null || daysLeft < 0 || daysLeft >= WARN_DAYS) return null;
  return `This copy stops working on ${day}.`;
}

/** The notice in two lines: what happened (the message's first sentence), then what to do. */
export function noticeParts(message: string): { title: string; text: string } {
  const cut = message.indexOf(". ");
  return cut < 0 ? { title: message, text: "" } : { title: message.slice(0, cut + 1), text: message.slice(cut + 2).trim() };
}

@Injectable({ providedIn: "root" })
export class ExpiryService {
  /** the last day this copy works, "2026-12-15"; null: no end date */
  readonly expires = signal<string | null>(null);
  readonly daysLeft = signal<number | null>(null);
  /** the server's message once this copy stopped working: the screens show only that */
  readonly stopped = signal<string | null>(null);
  readonly until = computed(() => (this.stopped() ? null : untilLine(this.expires())));
  readonly soon = computed(() => (this.stopped() ? null : soonLine(this.expires(), this.daysLeft())));

  /** GET /api/health says the date, the days left, and whether the copy already stopped. */
  fromHealth(h: AppHealth): void {
    this.expires.set(h.expires ?? null);
    this.daysLeft.set(h.days_left ?? null);
    if (h.expired) this.stop(h.expired_message);
  }

  stop(message: unknown): void {
    this.stopped.set(typeof message === "string" && message.trim() ? message.trim() : STOPPED);
  }

  /** An answer of the server, as status and body: a 410 that says `expired` turns the screens into the notice. */
  noteAnswer(status: number, body: unknown): void {
    const b = body as { expired?: unknown; detail?: unknown } | null;
    if (status === 410 && b?.expired === true) this.stop(b.detail);
  }
}

/** Any API call after the copy's end date answers 410 (app/web/server.py): the app gives way to the notice. */
export const expiryInterceptor: HttpInterceptorFn = (req, next) => {
  const expiry = inject(ExpiryService);
  return next(req).pipe(
    tap({
      error: (err: unknown) => {
        if (err instanceof HttpErrorResponse) expiry.noteAnswer(err.status, err.error);
      },
    }),
  );
};
