/** What  the screens  say about the Mind sign-in (19 Nov 2026, Joseph's wording). The server sends the same waiting
 *  sentence in `message`; these stand in when it has sent none. No error text of the  program goes on a screen: the
 *  server keeps it in `detail`, and  "Send feedback" carries the incident. */
export const SIGN_IN_WAITING =
  "Sign in to Mind in the Edge window the app opened (also when this page is in Chrome). It closes by itself. If you do not see it, look for Edge on the taskbar.";

export const NOT_SIGNED_IN = "Not signed in to Mind. Press Sign in to Mind and sign in once in the Edge window that opens. It lasts about a week.";

/** Where  the sign-in  is kept: the app's own Edge profile folder, as the server names it ("" when it did not). */
export function keptInLine(folder: string | null | undefined): string {
  return folder ? `The sign-in is kept on this PC only, in ${folder}.` : "";
}
