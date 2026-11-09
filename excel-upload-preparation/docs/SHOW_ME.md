# "Show me" pointers

Where the app tells the user to go somewhere ("Apply the blocking repairs in Prep", "Check them
under Advanced › Deliver", "on the Outside files card"), a small quiet **Show me** link sits next
to the words. It starts a pointer: one arrow with a short bubble at the next thing to click
("Click Advanced", "Click Prep actions"), moving on as the user clicks it, until the place itself.

## What the user sees

- One arrow at a time, in Carbon blue, with a short dark bubble that says what to click.
- At the place itself: a soft outline, the bubble says what to do there ("Apply the blocking
  repairs here") and **Got it**.
- Steps the user is already past are skipped: already on Advanced, the tab already open.
- In a window narrower than 66rem (a browser snapped to half a screen, a laptop at 125 %) Carbon
  hides the header's Main and Advanced, and the app has no side menu: the arrow then points at the
  page's own link ("Advanced" at the foot of the main screen, "Back to the main screen" on Advanced).
- **Esc**, **Got it**, a click on the place itself, or a click anywhere else ends it. A double
  click does not (its second click is not "a click elsewhere"), nor a click again on what was just
  clicked while the screen changes.
- The pointer never takes a click: the arrow, the bubble and the outline let every click through
  to the page (only "Got it" is a button). The user's click works as usual.
- Something missing (a tab not drawn, a button not shown in this state): the pointer ends quietly
  after 2.5 seconds. It does not skip a click on the way: what comes after it is on the screen that
  click leads to, and would not come either.
- The bubble sits under the element; at the place itself it moves beside or above it when a
  button sits under it, so it never hides the next thing to click.

## Where the links are

| Words in the app | Show me leads to |
|---|---|
| The banner's next step (main screen and Advanced) | by its `screen` (below) |
| "Check them under Advanced › Deliver (Check the numbers)" | Deliver: Check the numbers |
| "Pivot tables (…). Change it in Advanced › Prep actions." | Prep actions: the pivot tables choice |
| "Fix that under Advanced › Prep actions" (Do everything stopped) | Prep actions: Apply |
| "(Advanced › Recalculate)", "The list is under Advanced › Recalculate" (only when there is a list) | Recalculate: the list of formula errors, else its button |
| "The loop needs a decision first (Advanced › Mind loop)", "loop(s) need a decision" | Mind loop: the loop's choices |
| "Add … on the Outside files card (Add files)" | main screen: Add files |
| "See the Outside files card above" | main screen: the Outside files card |
| "Bring them in again first (the Outside files card …)" (Run in Mind) | main screen: the card's button that brings the files in |
| "paste an AI key under Advanced › Assistant" (server; the chat's "(no answer: …)") | Assistant: the AI key box |
| "added to the Prep plan" (Grid Namer) | Prep actions: Apply |
| A finding's "Use Bring in the outside files" (INP-010) | main screen: the card's button that brings the files in |
| A finding's "The Prep action '…'" (FRM, REF, STR rules) | Prep actions: the card of the action that repairs that rule, else the list |
| A finding's "by hand or with the assistant" | on the Findings row: none (its own "Ask the assistant" is next to it); in Prep's "By hand" list: that rule's "Ask the assistant" |
| "Apply them with the others from the list at the top of this page" (Prep), "(applied with the others from the top of the page)" (the assistant's note) | Advanced, above the tabs: Apply all |

In the chat only the app's own words get a link (its notes, and the server's "(no answer: …)"):
the assistant's free text may name a place to say there is nothing there ("There is no repair in
Prep for FRM-009"), and a link would lead the other way.

The readiness verdict's next step (`app/readiness.py`, `_next_step`) is read by its stable
`screen`, not by its words:

| `screen` | Show me leads to |
|---|---|
| `outside-inputs` | main screen: the Outside files card's button |
| `prep` | Advanced › Prep actions: Apply |
| `findings` (with `rule_id`) | Advanced › Findings: "Ask the assistant" on that rule's row |
| `recalculate` | "Recalculate in Excel": the main screen's Recalculate (on the main screen) or Advanced › Recalculate; "Review the new formula errors": the list |
| `reports` | "Download or send to Mind": Download this version (main screen) or Advanced › Run in Mind |

"Fix that under Advanced", which named no tab, now says which one: Prep actions after the formulas
and the structure, Mind loop after the loop, Recalculate after Recalculate. Fix all issues has no tab
of its own, so how it ended decides: the errors it left are in Recalculate's list ("Fix that under
Advanced › Recalculate"); when its comparison with the original did not run, "Check the numbers
under Advanced › Deliver"; when it did not start, was lost or stopped with an error, nothing under
Advanced fixes that: "Press Do everything again."

The readiness banner (`core/readiness-banner.ts`, on both screens) replaces Carbon's template to put
the link inside the words; its title keeps the notification's id (`notificationID`), so the banner
is still named by its verdict for a screen reader (Carbon's `aria-labelledby`).

## How it works

All in `MindReadyUI/src/app/core/`:

- `show-me.ts` -- the map: the places and the ways to them (`SHOW_ME_PLACES`), the words that name
  them (`SHOW_ME_PHRASES`; a text naming several places leads to the one it names first; a quoted
  name such as `Stopped at "Bring in the outside files"` is not read), `showMeFor(text)`,
  `showMeForStep(next_step)`, `planShowMe` (the clicks from the screen the user is on) and
  `placeShowMe` (where the arrow and the bubble go).
- `show-me.service.ts` -- `ShowMeService` (signals): `start`, `stop`, the steps, what it points at.
  It watches clicks on the whole document in the capture phase, so the element's own handler (a
  router link, a tab) runs as usual.
- `show-me-overlay.*` -- the arrow, the bubble and the outline, drawn once in the shell (`app.html`),
  above everything, following the element on scroll and resize.
- `show-me-link.ts` -- the link: `<mr-show-me [text]="…">` (appears only when the words name a
  place), `<mr-show-me to="prep-pivots">`, `<mr-show-me [step]="readiness.next_step">`;
  `[rule]="f.rule_id"` points at that rule's own row or Prep action; `except="assistant-ask"` leaves
  out a place already next to the text (the next place it names counts); `[unless]="…"` hides it
  when a text shown next to it already leads to the same place.

Elements are found by their `data-show-me` words, never by their text: `nav-main`, `nav-advanced`
(the header), `link-main`, `link-advanced` (the pages' own links, for a narrow window), `tab-<tab>`
for each Advanced tab, and the places (`prep-apply`, `prep-for-<rule>` on each Prep action's card,
`prep-list`, `prep-pivots`, `stack-apply`, `deliver-numbers`, `assistant-key`, `ask`, `ask-<rule>`,
`findings-list`, `recalc-run`, `recalc-errors`, `main-recalc`, `loop-choices`, `mind-run`,
`outside-card`, `outside-add`, `outside-bring-in`, `do-everything`, `download`).

To add one: put `<mr-show-me [text]="…">` next to the instruction. When it names a new place, add
the place to `SHOW_ME_PLACES`, its words to `SHOW_ME_PHRASES`, `data-show-me` on the element, and
the text to `FOUND` in `show-me.spec.ts` (that test lists every text that sends the user somewhere:
the app's own, the server's, and the findings' messages in `app/validators`).
