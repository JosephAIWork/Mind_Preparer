import { AdvancedTab } from "./session.store";

/** "Show  me" pointers.  Where the app tells the user to go somewhere ("Check them under Advanced ›
 *  Deliver", "Apply the blocking repairs in Prep", "on the Outside files card"), a quiet "Show me" link
 *  starts a  pointer: one arrow with a short bubble at the next thing  to click, until the place itself
 *  (ShowMeService  moves it, ShowMeOverlay draws it). This file  is the map: the places, the words that
 *  name them, and the  clicks that lead there.
 *
 *  Elements are found by  their `data-show-me` words, never by their text: the header's links
 *  (`nav-main`, `nav-advanced`), the links on the pages themselves (`link-main`, `link-advanced`), each
 *  Advanced tab (`tab-<tab>`) and each place below. One element may carry several words, space
 *  separated ("ask ask-FRM-009"). See excel-upload-preparation/docs/SHOW_ME.md. */

/** The two screens: the main screen (route "") and Advanced (route "advanced"). */
export type ShowMePage = "main" | "advanced";

/** The  places a  "Show me" leads to. */
export type ShowMeId =
  | "prep-apply" // Advanced  › Prep  actions: the Apply button
  | "prep-action" // Advanced › Prep actions: the action that repairs a rule ("The Prep action '…'" in a finding)
  | "prep-pivots" // Advanced › Prep actions: what happens to the pivot tables
  | "stack-apply" // Advanced, above the  tabs: the list of changes waiting to be applied, its Apply  all
  | "deliver-numbers" // Advanced › Deliver: Check the numbers
  | "assistant-key" // Advanced  › Assistant: the AI key box
  | "assistant-ask" // Advanced ›  Findings: "Ask the assistant" on the rule's row, else the list
  | "recalc-run" // Recalculate:  the main screen's button, or the one  under Advanced
  | "recalc-errors" // Advanced › Recalculate: the list of formula errors, else the button that makes it
  | "loop-choices" // Advanced › Mind loop: the answer a loop still needs, else the fold "Change the choices"
  | "loop-build" // Advanced › Mind loop: the Build it button
  | "loop-result" // Advanced › Mind loop: the line that says what the build did
  | "instances-choice" // Advanced › Instances: the question "Make one copy per…" (the loops to tick, or  a column)
  | "instances-build" // Advanced › Instances: the Build it button
  | "instances-result" // Advanced › Instances: the line that says what the build did, with the download
  | "outside-card" // main screen: the Outside files card
  | "outside-add" // main screen: the card's Add files
  | "outside-bring-in" // main screen: the card's button that brings the files in
  | "outside-waiting" // main screen: the box under Do everything while it waits for the input files (1.14.0)
  | "do-everything" // main screen: Do everything
  | "recent-workbooks" // main screen: the Recent workbooks list (a saved workbook to open again)
  | "download"; // main screen: Download this version; on Advanced: Run in Mind (send it to Mind)

/** A  pointer to start: the place,  and the rule whose Findings row to  point at ("Fix  FRM-009 with  the assistant"). */
export interface ShowMeRequest {
  id: ShowMeId;
  ruleId?: string;
}

/** What  to point  at: an element's `data-show-me` word, and the bubble's words there. */
export interface ShowMeTarget {
  el: string;
  text: string;
}

/** One  way to  a place: its screen, the Advanced tab to open there, and what to point at, best first
 *  (the first one on the screen is used: the list of formula errors, else the  button that makes it). */
export interface ShowMeWay {
  page: ShowMePage;
  tab?: AdvancedTab;
  targets: ShowMeTarget[];
}

/** One  step of  a pointer: a click to another screen ("page"), to an Advanced tab ("tab"), or the place
 *  itself ("place", outlined, with "Got it"). A page or tab step is skipped once the user  is there. */
export interface ShowMeStep {
  kind: "page" | "tab" | "place";
  page?: ShowMePage;
  tab?: AdvancedTab;
  targets: ShowMeTarget[];
}

/** The  readiness verdict's  next step, as app/readiness.py sends it. */
export interface ShowMeNextStep {
  screen?: string | null;
  label?: string | null;
  rule_id?: string | null;
}

/** The  Advanced tabs  by the names on their buttons. */
export const TAB_NAMES: Record<AdvancedTab, string> = {
  findings: "Findings",
  prep: "Prep actions",
  recalc: "Recalculate",
  loop: "Mind loop",
  instances: "Instances",
  map: "Model map",
  namer: "Grid Namer",
  assistant: "Assistant",
  mind: "Run in Mind",
  deliver: "Deliver",
  versions: "Versions",
};

/** Stands  for the  rule id in a target ("ask-{rule}"); a target with it is left out when no rule is given. */
const RULE = "{rule}";

const OUTSIDE_CARD: ShowMeTarget = { el: "outside-card", text: "The Outside files card is here" };

/** Every  place and  the ways to it. With several ways, the one on the screen the user is on wins (no
 *  click to another screen), else the first. */
export const SHOW_ME_PLACES: Record<ShowMeId, ShowMeWay[]> = {
  "prep-apply": [{ page: "advanced", tab: "prep", targets: [{ el: "prep-apply", text: "Apply the blocking repairs here" }] }],
  "prep-action": [
    {
      page: "advanced",
      tab: "prep",
      // the  card of  the action that repairs the finding's rule; an action with nothing left to change
      // is not drawn, and then the list of actions is shown
      targets: [
        { el: `prep-for-${RULE}`, text: "Tick this action, then click Apply" },
        { el: "prep-list", text: "The Prep actions are listed here" },
      ],
    },
  ],
  "prep-pivots": [{ page: "advanced", tab: "prep", targets: [{ el: "prep-pivots", text: "Choose what happens to the pivot tables here" }] }],
  // above the tabs: on any tab of Advanced
  "stack-apply": [{ page: "advanced", targets: [{ el: "stack-apply", text: "Apply the whole list here" }] }],
  "deliver-numbers": [{ page: "advanced", tab: "deliver", targets: [{ el: "deliver-numbers", text: "Check the numbers here" }] }],
  "assistant-key": [{ page: "advanced", tab: "assistant", targets: [{ el: "assistant-key", text: "Paste the AI key here" }] }],
  "assistant-ask": [
    {
      page: "advanced",
      tab: "findings",
      targets: [
        { el: `ask-${RULE}`, text: `Ask the assistant about ${RULE} here` },
        { el: "ask", text: "Ask the assistant about a finding here" },
        { el: "findings-list", text: "The findings are listed here" },
      ],
    },
  ],
  "recalc-run": [
    { page: "main", targets: [{ el: "main-recalc", text: "Click Recalculate" }] },
    { page: "advanced", tab: "recalc", targets: [{ el: "recalc-run", text: "Recalculate here" }] },
  ],
  "recalc-errors": [
    {
      page: "advanced",
      tab: "recalc",
      targets: [
        { el: "recalc-errors", text: "The formula errors are listed here" },
        { el: "recalc-run", text: "Recalculate here to list the errors" },
      ],
    },
  ],
  // the  box of the one answer a loop still needs (drawn above  the fold only while it is  needed), else
  // the fold that holds every choice
  "loop-choices": [
    {
      page: "advanced",
      tab: "loop",
      targets: [
        { el: "loop-answer", text: "Type the answer here, then click Build it" },
        { el: "loop-choices", text: "Open this to change the loop's choices" },
      ],
    },
  ],
  "loop-build": [{ page: "advanced", tab: "loop", targets: [{ el: "loop-build", text: "Click Build it" }] }],
  "loop-result": [{ page: "advanced", tab: "loop", targets: [{ el: "loop-result", text: "What the build did is here" }] }],
  // the  Instances tab: the question, the Build it button, the line with  the download
  "instances-choice": [{ page: "advanced", tab: "instances", targets: [{ el: "instances-choice", text: "Tick a loop or pick a column here" }] }],
  "instances-build": [{ page: "advanced", tab: "instances", targets: [{ el: "instances-build", text: "Click Build it" }] }],
  "instances-result": [{ page: "advanced", tab: "instances", targets: [{ el: "instances-result", text: "What the build did is here, with the download" }] }],
  "outside-card": [{ page: "main", targets: [OUTSIDE_CARD] }],
  "outside-add": [{ page: "main", targets: [{ el: "outside-add", text: "Add the input files here" }, OUTSIDE_CARD] }],
  "outside-bring-in": [{ page: "main", targets: [{ el: "outside-bring-in", text: "Bring in the outside files here" }, OUTSIDE_CARD] }],
  "do-everything": [{ page: "main", targets: [{ el: "do-everything", text: "Click Do everything" }] }],
  // the box is drawn only while  Do everything waits for the input files: the picker first, else the box
  "outside-waiting": [{ page: "main", targets: [{ el: "outside-waiting-pick", text: "Pick the input files here, then click Go on" }, { el: "outside-waiting", text: "Do everything waits here for the input files" }] }],
  // the list is drawn only while a saved workbook exists and no workbook is open (the idle main screen)
  "recent-workbooks": [{ page: "main", targets: [{ el: "recent-workbooks", text: "Reopen the workbook here" }] }],
  download: [
    { page: "main", targets: [{ el: "download", text: "Download the file here" }] },
    { page: "advanced", tab: "mind", targets: [{ el: "mind-run", text: "Send it to Mind here" }] },
  ],
};

/** "Advanced  › Deliver"  as the app writes it, also "Advanced > Deliver" (the backend's plain texts). */
const UNDER = String.raw`Advanced\s*(?:›|>|&gt;)\s*`;
const under = (tab: string) => new RegExp(UNDER + tab, "i");

/** The  words in  the app's texts (and the backend's) that send the user somewhere, and the place each
 *  one names. Found with grep for "Advanced ›", "in Prep", "Prep action", "Assistant", "Deliver",
 *  "Recalculate", "Mind loop", "card",  "Use ", "press", "top of" in MindReadyUI/src and in
 *  excel-upload-preparation/app -- the  findings' own messages included (app/validators) --  see
 *  show-me.spec.ts for every text. When a text  names several places, the one it names first counts:
 *  the first thing  to do. */
export const SHOW_ME_PHRASES: { words: RegExp; id: ShowMeId }[] = [
  { words: /pivot tables?\b[^]*?Advanced\s*(?:›|>|&gt;)\s*Prep actions/i, id: "prep-pivots" },
  { words: under("Deliver"), id: "deliver-numbers" },
  { words: /\bCheck the numbers\b/i, id: "deliver-numbers" },
  { words: under("Prep actions"), id: "prep-apply" },
  { words: /\bblocking repairs in Prep\b|\bin Prep\b|\bPrep has a repair\b|\badded to the Prep plan\b/i, id: "prep-apply" },
  // a  finding's message:  "The Prep action 'Freeze external links' removes them"
  { words: /\bthe Prep action\b/i, id: "prep-action" },
  // "Apply them with the others from the list at the top of this page" (Prep, the assistant)
  { words: /\b(?:the list at|from) the top of (?:this|the) page\b/i, id: "stack-apply" },
  { words: under("Assistant"), id: "assistant-key" },
  { words: /\bpaste (?:an|the|your) AI key\b/i, id: "assistant-key" },
  { words: /\bFix [A-Z]{2,5}-\d{3} with the assistant\b|\bby hand or with the assistant\b/i, id: "assistant-ask" },
  { words: under("Recalculate"), id: "recalc-errors" },
  { words: /\bReview the new formula errors\b/i, id: "recalc-errors" },
  { words: /\bRecalculate in Excel\b/i, id: "recalc-run" },
  // "Build it under Advanced › Mind loop": the  button, named before the tab; "click Build it" is the loop's
  // unless  the same sentence names the Instances  tab ("click Build it under Advanced ›  Instances")
  { words: /\b(?:press|click) Build it\b(?![^.!?]*\bInstances\b)|\bBuild it under Advanced\s*(?:›|>|&gt;)\s*Mind loop\b/i, id: "loop-build" },
  { words: /\bthe loop's result\b/i, id: "loop-result" },
  { words: under("Mind loop"), id: "loop-choices" },
  // "1 loop needs one answer", "2 loops need one answer each";  the older "needs a decision" still read
  { words: /\bloop(?:\(s\)|s)? needs? (?:one answer|a decision)\b/i, id: "loop-choices" },
  // the  Instances tab: "Build it under Advanced › Instances" or "click Build it (Advanced › Instances)" (the
  // button, with the tab named in the same sentence), "the copies' result" / "the instances' result" (the
  // line with the download), else the question
  { words: /\bBuild it under Advanced\s*(?:›|>|&gt;)\s*Instances\b|\b(?:press|click) Build it\b(?=[^.!?]*\bInstances\b)/i, id: "instances-build" },
  { words: /\bthe (?:copies|instances)' result\b/i, id: "instances-result" },
  { words: under("Instances"), id: "instances-choice" },
  { words: /\bOutside files card \(Add files\)/i, id: "outside-add" },
  { words: /\bBring (?:them|the outside files) in again\b|\bBring in the outside files\b/i, id: "outside-bring-in" },
  { words: /\bOutside files card\b/i, id: "outside-card" },
  // "See the box under Do everything." (the Outside files card while Do everything waits for the input files)
  { words: /\bthe box under Do everything\b/i, id: "outside-waiting" },
  { words: /\bpress Do everything\b/i, id: "do-everything" },
  { words: /\bDownload or send to Mind\b/i, id: "download" },
  // "Open it again from Recent workbooks on  the main screen" (the server, for a session the restart closed)
  { words: /\bRecent workbooks\b/i, id: "recent-workbooks" },
];

/** The  place a  text sends the user to, or null when it names none (most texts: a "Show me" appears
 *  only next to an instruction). A quoted name ('Stopped at "Bring in the outside files":  …', 'press
 *  "Make them with AI"') names a step or a button,  not a place, so it is  not read. `except`: a place
 *  already next to  the text (a Findings row's own "Ask the assistant"): the next  place it names counts. */
export function showMeFor(text: string | null | undefined, except?: ShowMeId | null): ShowMeRequest | null {
  const plain = (text ?? "").replace(/"[^"]*"|“[^”]*”/g, '""');
  if (!plain.trim()) return null;
  let best: { at: number; id: ShowMeId } | null = null;
  for (const p of SHOW_ME_PHRASES) {
    if (p.id === except) continue;
    const m = p.words.exec(plain);
    if (m && (!best || m.index < best.at)) best = { at: m.index, id: p.id };
  }
  if (!best) return null;
  // "Fix  FRM-009 with  the assistant": point at that rule's row
  const rule = best.id === "assistant-ask" ? /\b[Ff]ix ([A-Z]{2,5}-\d{3}) with the assistant/.exec(plain)?.[1] : undefined;
  return rule ? { id: best.id, ruleId: rule } : { id: best.id };
}

/** A  chat message  the app wrote itself: its notes ("12 changes added to the list …") and the server's
 *  "(no answer: …)" when the assistant could not answer (app/web/server.py, `_chat_payload`). Only these
 *  get a "Show me".  The assistant's own words may name a place only to say  there is nothing to do
 *  there ("There  is no repair in Prep for FRM-009:  change the formula by hand"): a link would lead the
 *  other way. */
export function chatIsAppWords(m: { role: string; content: string; note?: boolean }): boolean {
  return m.role !== "user" && (!!m.note || m.content.startsWith("(no answer:"));
}

/** The  place the  readiness verdict's next step names, by its stable `screen` (app/readiness.py
 *  `_next_step`): outside-inputs, prep, findings (with the rule), recalculate, reports. */
export function showMeForStep(step: ShowMeNextStep | null | undefined): ShowMeRequest | null {
  if (!step) return null;
  switch (step.screen) {
    case "outside-inputs":
      return { id: "outside-bring-in" };
    case "prep":
      return { id: "prep-apply" };
    case "findings":
      return step.rule_id ? { id: "assistant-ask", ruleId: step.rule_id } : { id: "assistant-ask" };
    case "recalculate":
      // "Review  the new  formula errors" (the list) or "Recalculate in Excel" (the button)
      return { id: /error/i.test(step.label ?? "") ? "recalc-errors" : "recalc-run" };
    case "reports":
      return { id: "download" };
    default:
      return showMeFor(step.label); // a screen this front end does not know: its words may still name a place
  }
}

/** The  links to  each screen: the header's, else the one on the page itself. Below 66rem (a window
 *  snapped to half a screen, a laptop at 125 %) Carbon hides the header's links and the  app has no side
 *  menu, so the page's own link is the  way: "Advanced" at the foot of  the main screen, "Back to the
 *  main screen"  at the top of Advanced. */
export const PAGE_LINKS: Record<ShowMePage, ShowMeTarget[]> = {
  advanced: [
    { el: "nav-advanced", text: "Click Advanced" },
    { el: "link-advanced", text: "Click Advanced" },
  ],
  main: [
    { el: "nav-main", text: "Click Main" },
    { el: "link-main", text: "Click Back to the main screen" },
  ],
};

/** The  steps from  the screen the user is on (`here`) to the place: a click on Advanced or Main when the
 *  place is on the other screen, a click on the Advanced tab, then the  place itself. Whether the user is
 *  already on that screen or tab  is checked again as each step  comes up (ShowMeService). */
export function planShowMe(request: ShowMeRequest, here: ShowMePage): ShowMeStep[] {
  const ways = SHOW_ME_PLACES[request.id];
  if (!ways?.length) return [];
  const way = ways.find((w) => w.page === here) ?? ways[0];
  const steps: ShowMeStep[] = [];
  if (way.page !== here) steps.push({ kind: "page", page: way.page, targets: [...PAGE_LINKS[way.page]] });
  if (way.tab) steps.push({ kind: "tab", tab: way.tab, targets: [{ el: `tab-${way.tab}`, text: `Click ${TAB_NAMES[way.tab]}` }] });
  const rule = request.ruleId?.trim();
  const targets = way.targets.flatMap((t) => {
    if (!t.el.includes(RULE) && !t.text.includes(RULE)) return [t];
    return rule ? [{ el: t.el.split(RULE).join(rule), text: t.text.split(RULE).join(rule) }] : [];
  });
  steps.push({ kind: "place", targets });
  return steps;
}

/** The  size and  place of an element, as getBoundingClientRect gives them. */
export interface ShowMeBox {
  top: number;
  left: number;
  width: number;
  height: number;
}

/** Which  side of  the element the bubble sits on ("inside": a card or list taller than the window, the
 *  bubble just inside its top edge). The arrow points from the bubble at the element. */
export type ShowMeSide = "below" | "right" | "above" | "left";

/** Where  the overlay  draws: the arrow's box, the bubble, and the soft outline around the place itself. */
export interface ShowMeLayout {
  side: ShowMeSide | "inside";
  arrow: ShowMeBox;
  bubble: { x: number; y: number };
  ring: ShowMeBox;
}

/** The  arrow's length  and width, the gap between it and the element, the margin kept from the window's
 *  edges, and Carbon's fixed header (3rem): nothing is put under it. */
export const ARROW_LONG = 26;
export const ARROW_WIDE = 20;
const GAP = 6;
const EDGE = 8;
const HEADER = 48;

/** A  click on  the way (a header link, a tab): the bubble always under it, the same place each time. */
export const STEP_SIDES: ShowMeSide[] = ["below", "above"];
/** The  place itself,  where the user acts: under it, else beside it, else above it -- the first side
 *  that covers no other button (the user may want the one next to it). */
export const PLACE_SIDES: ShowMeSide[] = ["below", "right", "above", "left"];

const clamp = (v: number, lo: number, hi: number) => Math.min(Math.max(v, lo), Math.max(lo, hi));

/** The  arrow and  the bubble on one side of the element, and whether they fit in the window. */
function sideLayout(side: ShowMeSide, el: ShowMeBox, bubble: { width: number; height: number }, view: { width: number; height: number }): { arrow: ShowMeBox; bubble: { x: number; y: number }; fits: boolean } {
  const right = el.left + el.width;
  const bottom = el.top + el.height;
  const cx = clamp(el.left + el.width / 2, EDGE + ARROW_WIDE / 2, view.width - EDGE - ARROW_WIDE / 2);
  const cy = clamp(el.top + el.height / 2, HEADER + EDGE + ARROW_WIDE / 2, view.height - EDGE - ARROW_WIDE / 2);
  const bx = clamp(cx - bubble.width / 2, EDGE, view.width - EDGE - bubble.width);
  const by = clamp(cy - bubble.height / 2, HEADER + EDGE, view.height - EDGE - bubble.height);
  switch (side) {
    case "below": {
      const arrow = { left: cx - ARROW_WIDE / 2, top: bottom + GAP, width: ARROW_WIDE, height: ARROW_LONG };
      return { arrow, bubble: { x: bx, y: arrow.top + ARROW_LONG }, fits: arrow.top + ARROW_LONG + bubble.height <= view.height - EDGE };
    }
    case "above": {
      const arrow = { left: cx - ARROW_WIDE / 2, top: el.top - GAP - ARROW_LONG, width: ARROW_WIDE, height: ARROW_LONG };
      return { arrow, bubble: { x: bx, y: arrow.top - bubble.height }, fits: arrow.top - bubble.height >= HEADER + EDGE };
    }
    case "right": {
      const arrow = { left: right + GAP, top: cy - ARROW_WIDE / 2, width: ARROW_LONG, height: ARROW_WIDE };
      return { arrow, bubble: { x: arrow.left + ARROW_LONG, y: by }, fits: arrow.left + ARROW_LONG + bubble.width <= view.width - EDGE };
    }
    case "left": {
      const arrow = { left: el.left - GAP - ARROW_LONG, top: cy - ARROW_WIDE / 2, width: ARROW_LONG, height: ARROW_WIDE };
      return { arrow, bubble: { x: arrow.left - bubble.width, y: by }, fits: arrow.left - bubble.width >= EDGE };
    }
  }
}

/** Where  the arrow  and the bubble go: the first of `sides` that fits in the window and, when `free` is
 *  given, covers nothing to click (`free` asks the page about the bubble's and the arrow's  boxes);
 *  `keep`, the side used a moment ago, is tried first so  the bubble does not jump while  the page
 *  scrolls. When none is free, the  first that fits; when none fits (a card or list taller  than the
 *  window), just inside the element's top  edge. */
export function placeShowMe(
  el: ShowMeBox,
  bubble: { width: number; height: number },
  view: { width: number; height: number },
  sides: ShowMeSide[] = STEP_SIDES,
  free?: (box: ShowMeBox) => boolean,
  keep?: ShowMeSide | null,
): ShowMeLayout {
  const order = keep && sides.includes(keep) ? [keep, ...sides.filter((s) => s !== keep)] : sides;
  const fitting = order.map((side) => ({ side, ...sideLayout(side, el, bubble, view) })).filter((c) => c.fits);
  const clear = (c: (typeof fitting)[number]) => !free || (free({ left: c.bubble.x, top: c.bubble.y, ...bubble }) && free(c.arrow));
  const chosen = fitting.find(clear) ?? fitting[0];
  const round = (b: ShowMeBox): ShowMeBox => ({ top: Math.round(b.top), left: Math.round(b.left), width: Math.round(b.width), height: Math.round(b.height) });
  const ring = round({ top: el.top - 4, left: el.left - 4, width: el.width + 8, height: el.height + 8 });
  if (chosen) {
    return { side: chosen.side, arrow: round(chosen.arrow), bubble: { x: Math.round(chosen.bubble.x), y: Math.round(chosen.bubble.y) }, ring };
  }
  const tip = Math.max(el.top, HEADER) + GAP;
  const cx = clamp(el.left + el.width / 2, EDGE + ARROW_WIDE / 2, view.width - EDGE - ARROW_WIDE / 2);
  const bx = clamp(cx - bubble.width / 2, EDGE, view.width - EDGE - bubble.width);
  return { side: "inside", arrow: round({ left: cx - ARROW_WIDE / 2, top: tip, width: ARROW_WIDE, height: ARROW_LONG }), bubble: { x: Math.round(bx), y: Math.round(tip + ARROW_LONG) }, ring };
}
