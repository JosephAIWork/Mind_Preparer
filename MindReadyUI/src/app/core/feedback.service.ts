import { HttpErrorResponse, HttpInterceptorFn, HttpResponse } from "@angular/common/http";
import { Injectable, inject, signal } from "@angular/core";
import { tap } from "rxjs";

/** "Send feedback to the Tel Aviv team": the problem the backend just reported (its incident ref),
 *  shown as a box under the header, and the dialog that turns it into an email. */
@Injectable({ providedIn: "root" })
export class FeedbackService {
  /** the incident waiting in the box under the header (null: no box), with what stopped, in plain words */
  readonly incident = signal<{ ref: string; text: string | null } | null>(null);
  /** the dialog is open, for this incident (null: feedback from the header) */
  readonly dialog = signal<{ ref: string | null } | null>(null);
  private seen = new Set<string>();

  /** A problem the backend recorded: shown once, however often a poll repeats it. */
  report(ref: unknown, text?: unknown): void {
    if (typeof ref !== "string" || !ref || this.seen.has(ref)) return;
    this.seen.add(ref);
    this.incident.set({ ref, text: typeof text === "string" && text ? text : null });
  }

  open(ref: string | null = null): void {
    this.dialog.set({ ref });
  }

  close(): void {
    this.dialog.set(null);
  }

  dismiss(): void {
    this.incident.set(null);
  }
}

/** Every answer of the backend that carries an `incident` (a failed step, changes Excel did not
 *  write, a background job that stopped on an error) puts the feedback box up. */
export const feedbackInterceptor: HttpInterceptorFn = (req, next) => {
  const feedback = inject(FeedbackService);
  return next(req).pipe(
    tap({
      next: (ev) => {
        if (ev instanceof HttpResponse) {
          const body = ev.body as { incident?: unknown; problem?: unknown } | null;
          feedback.report(body?.incident, body?.problem);
        }
      },
      error: (err: unknown) => {
        if (err instanceof HttpErrorResponse) {
          const body = err.error as { incident?: unknown; problem?: unknown } | null;
          feedback.report(body?.incident, body?.problem);
        }
      },
    }),
  );
};
