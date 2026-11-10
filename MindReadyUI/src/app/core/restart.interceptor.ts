import { HttpErrorResponse, HttpInterceptorFn } from "@angular/common/http";
import { inject } from "@angular/core";
import { tap } from "rxjs";
import { SessionStore } from "./session.store";

/** What the server's 404 says for a session the restart closed (app/web/server.py `_session`). */
export const RESTART_MARK = "(the app was restarted)";

/** 1.14.0: any call about a session the server no longer holds answers 404 "This workbook is not open
 *  any more (the app was restarted). Open it again from Recent workbooks on the main screen." The page
 *  then lets go of the session (the main screen goes back to its list, where the sentence stands and
 *  its Show me can point at the list) instead of showing the sentence inside a page of a workbook that
 *  is not there any more. The call still fails for its caller, as before. */
export const restartInterceptor: HttpInterceptorFn = (req, next) => {
  const store = inject(SessionStore);
  return next(req).pipe(
    tap({
      error: (err: unknown) => {
        if (!(err instanceof HttpErrorResponse) || err.status !== 404) return;
        const detail = (err.error as { detail?: unknown } | null)?.detail;
        if (typeof detail === "string" && detail.includes(RESTART_MARK)) store.closedByRestart(detail);
      },
    }),
  );
};
