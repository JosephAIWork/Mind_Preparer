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
- **Esc**, **Got it**, a click on the place itself, or a click anywhere else ends it.
- The pointer never takes a click: the arrow, the bubble and the outline let every click through
  to the page (only "Got it" is a button). The user's click works as usual.
- Something missing (a tab not drawn, a button not shown in this state): that click is skipped
  after 2.5 seconds; when the place itself is missing, the pointer ends quietly.
- The bubble sits under the element; at the place itself it moves beside or above it when a
  button sits under it, so it never hides the next thing to click.

## Where the links are

| Words in the app | Show me leads to |
|---|---|
| The banner's next step (main screen and Advanced) | by its `screen` (below) |
| "Check them under Advanced › Deliver (Check the numbers)" | Deliver: Check the numbers |
| "Pivot tables (…). Change it in Advanced › Prep actions." | Prep actions: the pivot tables choice |
| "Fix that under Advanced › Prep actions" (Do everything stopped) | Prep actions: Apply |
| "(Advanced › Recalculate)", "The list is under Advanced › Recalculate" | Recalculate: the list of formula errors, else its button |
| "The loop needs a decision first (Advanced › Mind loop)", "loop(s) need a decision" | Mind loop: the loop's choices |
| "Add … on the Outside files card (Add files)" | main screen: Add files |
| "See the Outside files card above" | main screen: the Outside files card |
| "Bring them in again first (the Outside files card …)" (Run in Mind) | main screen: the card's button that brings the files in |
| "paste an AI key under Advanced › Assistant" (server, assistant answers) | Assistant: the AI key box |
| "added to the Prep plan" (Grid Namer) | Prep actions: Apply |

The readiness verdict's next step (`app/readiness.py`, `_next_step`) is read by its stable
`screen`, not by its words:

| `screen` | Show me leads to |
|---|---|
| `outside-inputs` | main screen: the Outside files card's button |
| `prep` | Advanced › Prep actions: Apply |
| `findings` (with `rule_id`) | Advanced › Findings: "Ask the assistant" on that rule's row |
| `recalculate` | "Recalculate in Excel": the main screen's Recalculate (on the main screen) or Advanced › Recalculate; "Review the new formula errors": the list |
| `reports` | "Download or send to Mind": Download this version (main screen) or Advanced › Run in Mind |

"Fix that under Advanced", which named no tab, now says which one (Recalculate after Fix all
issues or Recalculate, Prep actions after the formulas and the structure, Mind loop after the loop).

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
  `[unless]="…"` hides it when a text shown next to it already leads to the same place.

Elements are found by their `data-show-me` words, never by their text: `nav-main`, `nav-advanced`,
`tab-<tab>` for each Advanced tab, and the places (`prep-apply`, `prep-pivots`, `deliver-numbers`,
`assistant-key`, `ask`, `ask-<rule>`, `findings-list`, `recalc-run`, `recalc-errors`,
`main-recalc`, `loop-choices`, `mind-run`, `outside-card`, `outside-add`, `outside-bring-in`,
`do-everything`, `download`).

To add one: put `<mr-show-me [text]="…">` next to the instruction. When it names a new place, add
the place to `SHOW_ME_PLACES`, its words to `SHOW_ME_PHRASES`, `data-show-me` on the element, and
the text to `FOUND` in `show-me.spec.ts` (that test lists every text that sends the user somewhere).
