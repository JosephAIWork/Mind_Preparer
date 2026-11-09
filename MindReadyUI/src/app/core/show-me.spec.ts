import { Component } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import { Router, provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import { ApiService } from "./api.service";
import { LoopPlan, OutsideInputs, OutsideRunResult, outsideAddLine, outsideOutcome } from "./models";
import { SessionStore } from "./session.store";
import { PLACE_SIDES, SHOW_ME_PHRASES, SHOW_ME_PLACES, ShowMeId, TAB_NAMES, placeShowMe, planShowMe, showMeFor, showMeForStep } from "./show-me";
import { ShowMeOverlay } from "./show-me-overlay";
import { SHOW_ME_AFTER_CLICK_MS, SHOW_ME_WAIT_MS, SHOW_ME_WATCH_MS, ShowMeService } from "./show-me.service";

// ---- every text that sends the user somewhere, and the place its "Show me" leads to -----------------
// Found with grep for "Advanced ›", "Advanced >", "in Prep", "Prep actions", "Assistant", "Deliver",
// "Recalculate", "Mind loop", "card" in MindReadyUI/src and excel-upload-preparation/app.
const FOUND: { where: string; text: string; id: ShowMeId; ruleId?: string }[] = [
  // the main screen
  { where: "Outside files card: numbers not checked", text: "Check them under Advanced › Deliver (Check the numbers), or go on without the check.", id: "deliver-numbers" },
  { where: "Do everything: the pivot tables", text: "Pivot tables (2): Formulas where they give the same numbers. Change it in Advanced › Prep actions.", id: "prep-pivots" },
  {
    where: "Outside files card: files missing (Do everything)",
    text: 'Add them on the Outside files card (Add files), or press "Bring in the 1 matched file" there: a missing file keeps the data the model already holds.',
    id: "outside-add",
  },
  { where: "Fix all issues: errors left", text: "3 → 1 error cells · 12 cells rewritten · v3 · 1 left for a person: Calc!B2 #DIV/0! (Advanced › Recalculate)", id: "recalc-errors" },
  { where: "Build the Mind loop tile", text: "Loop Policy (1 to 40): 120 formulas to write · 1 loop(s) need a decision (Advanced › Mind loop)", id: "loop-choices" },
  { where: "Recalculate tile", text: "Recalculated in Excel · 2 new since the original · 3 already in the original · The list is under Advanced › Recalculate.", id: "recalc-errors" },
  { where: "Do everything: the loop step skipped", text: "The loop needs a decision first (Advanced › Mind loop)", id: "loop-choices" },
  {
    where: "Do everything: stopped, numbers not checked",
    text: 'Stopped at "Bring in the outside files": the computed values were not checked. Check them under Advanced › Deliver (Check the numbers), or press Do everything again to go on without the check.',
    id: "deliver-numbers",
  },
  {
    where: "Do everything: stopped at the outside step",
    text: 'Stopped at "Bring in the outside files": Making the Mind inputs would change 3 computed values. See the Outside files card above, then press Do everything again.',
    id: "outside-card",
  },
  { where: "Do everything: stopped at Fix all issues", text: 'Stopped at "Fix all issues": Not started. Fix that under Advanced › Recalculate, then press Do everything again.', id: "recalc-errors" },
  {
    where: "Do everything: stopped at the formulas and the structure",
    text: 'Stopped at "Fix the formulas and the structure": 3 changed, 1 refused. Fix that under Advanced › Prep actions, then press Do everything again.',
    id: "prep-apply",
  },
  { where: "Do everything: stopped at the loop", text: 'Stopped at "Build the Mind loop": Could not write the loop. Fix that under Advanced › Mind loop, then press Do everything again.', id: "loop-choices" },
  { where: "Do everything: stopped at Recalculate", text: 'Stopped at "Recalculate": Recalculated, 4 formula errors. Fix that under Advanced › Recalculate, then press Do everything again.', id: "recalc-errors" },
  // Advanced
  { where: "Grid Namer: suggestions", text: "12 suggestions, added to the Prep plan", id: "prep-apply" },
  // the backend's words
  { where: "readiness next step (app/readiness.py)", text: "Apply the blocking repairs in Prep", id: "prep-apply" },
  { where: "readiness next step", text: "Fix FRM-009 with the assistant", id: "assistant-ask", ruleId: "FRM-009" },
  { where: "readiness next step", text: "Bring in the outside files", id: "outside-bring-in" },
  { where: "readiness next step", text: "Recalculate in Excel", id: "recalc-run" },
  { where: "readiness next step", text: "Review the new formula errors", id: "recalc-errors" },
  { where: "readiness next step", text: "Download or send to Mind", id: "download" },
  {
    where: "readiness: after an Apply",
    text: "12 of 14 changes written. Blocking problems: 3 before, 2 now -- left: FRM-001 (Prep has a repair: apply it); FRM-009 (no automatic repair: by hand or with the assistant).",
    id: "prep-apply",
  },
  { where: "readiness: after an Apply", text: "Still open: nothing more Prep can do -- fix the rest by hand or with the assistant.", id: "assistant-ask" },
  { where: "app/autofix.py: a #REF! that stays", text: "rewrite: the cell's own value changed, so the #REF! stays (fix it in Prep or by hand)", id: "prep-apply" },
  { where: "app/llm.py: no AI key (grid names)", text: "No AI key is saved yet: paste one under Advanced > Assistant.", id: "assistant-key" },
  { where: "app/llm.py: no AI key (the chat's answer)", text: "(no answer: No AI key is saved yet: paste one under Advanced > Assistant.)", id: "assistant-key" },
  {
    where: "server: stand-ins without an AI key",
    text: "The assistant is not set up, so it cannot make the missing input files: add the actual files, or paste an AI key under Advanced › Assistant.",
    id: "assistant-key",
  },
  {
    where: "server: Run in Mind after the input files changed",
    text: "The input files changed since they were brought in: Run in Mind would use the data of the old files. Bring them in again first (the Outside files card on the main screen, or Do everything), then run in Mind.",
    id: "outside-bring-in",
  },
  { where: "the Do everything button itself", text: "press Do everything", id: "do-everything" },
];

/** Texts next to which no "Show me" belongs: they name no place to go. */
const NOT_A_PLACE = [
  "",
  "Making the missing input files with the assistant…",
  "Recalculated, no formula errors",
  "Everything ran and every check passed (12 min).",
  "2 of 2 input files matched.",
  "Already brought in",
  'Running "Bring in the outside files"', // a quoted name is a step's name, not an instruction
  "Fixed what it could; 3 errors the workbook already had stay, moving on",
];

describe("Show me: the texts that send the user somewhere", () => {
  it.each(FOUND)("$where leads to $id", ({ text, id, ruleId }) => {
    expect(showMeFor(text)).toEqual(ruleId ? { id, ruleId } : { id });
  });

  it("gives no link to a text that names no place", () => {
    for (const t of NOT_A_PLACE) expect(showMeFor(t), t).toBeNull();
    expect(showMeFor(null)).toBeNull();
    expect(showMeFor(undefined)).toBeNull();
  });

  it("covers the texts the app builds itself", () => {
    const o = (assistant: boolean) => ({ assistant }) as OutsideInputs;
    expect(showMeFor(outsideAddLine(o(true)))).toEqual({ id: "outside-add" });
    expect(showMeFor(outsideAddLine(o(false)))).toEqual({ id: "outside-add" });
    const unchecked = {
      import: { macro: "calcCF", files: [], kept: [], dialogs: [], stopped_early: false, error: null },
      mindInputs: { inputs: [{ name: "IN_A" }], kept: [], formulas_changed: 10, names_changed: 0, sums_split: 0 },
      numbers: { ran: false, clean: null, verdict: "Excel is not available here: the numbers could not be compared." },
      summary: "1 outside file brought in.",
    } as unknown as OutsideRunResult;
    expect(showMeFor(outsideOutcome(unchecked).detail)).toEqual({ id: "deliver-numbers" });
  });

  it("names the place the text names first: the first thing to do", () => {
    expect(showMeFor("Check them under Advanced › Deliver, or press Do everything again.")?.id).toBe("deliver-numbers");
    expect(showMeFor("Press Do everything again, or check them under Advanced › Deliver.")?.id).toBe("do-everything");
  });

  it("knows every place it names, and every place is named by some text", () => {
    const named = new Set(SHOW_ME_PHRASES.map((p) => p.id));
    for (const id of Object.keys(SHOW_ME_PLACES) as ShowMeId[]) {
      expect(named.has(id), id).toBe(true);
      expect(SHOW_ME_PLACES[id].length, id).toBeGreaterThan(0);
      for (const way of SHOW_ME_PLACES[id]) {
        expect(way.targets.length, id).toBeGreaterThan(0);
        if (way.tab) expect(TAB_NAMES[way.tab], id).toBeTruthy();
      }
    }
  });
});

describe("Show me: the readiness verdict's next step", () => {
  it("follows the step's screen, with the rule for the Findings row", () => {
    expect(showMeForStep({ screen: "outside-inputs", label: "Bring in the outside files" })).toEqual({ id: "outside-bring-in" });
    expect(showMeForStep({ screen: "prep", label: "Apply the blocking repairs in Prep" })).toEqual({ id: "prep-apply" });
    expect(showMeForStep({ screen: "findings", label: "Fix FRM-009 with the assistant", rule_id: "FRM-009" })).toEqual({ id: "assistant-ask", ruleId: "FRM-009" });
    expect(showMeForStep({ screen: "recalculate", label: "Recalculate in Excel" })).toEqual({ id: "recalc-run" });
    expect(showMeForStep({ screen: "recalculate", label: "Review the new formula errors" })).toEqual({ id: "recalc-errors" });
    expect(showMeForStep({ screen: "reports", label: "Download or send to Mind" })).toEqual({ id: "download" });
  });

  it("reads the words of a screen it does not know, and nothing from no step", () => {
    expect(showMeForStep({ screen: "somewhere-new", label: "Apply the blocking repairs in Prep" })).toEqual({ id: "prep-apply" });
    expect(showMeForStep({ screen: "somewhere-new", label: "Wait" })).toBeNull();
    expect(showMeForStep(null)).toBeNull();
  });
});

describe("Show me: the clicks to a place", () => {
  it("goes through Advanced and the tab from the main screen, through the tab only on Advanced", () => {
    expect(planShowMe({ id: "prep-apply" }, "main")).toEqual([
      { kind: "page", page: "advanced", targets: [{ el: "nav-advanced", text: "Click Advanced" }] },
      { kind: "tab", tab: "prep", targets: [{ el: "tab-prep", text: "Click Prep actions" }] },
      { kind: "place", targets: [{ el: "prep-apply", text: "Apply the blocking repairs here" }] },
    ]);
    expect(planShowMe({ id: "prep-apply" }, "advanced").map((s) => s.kind)).toEqual(["tab", "place"]);
    expect(planShowMe({ id: "outside-card" }, "advanced")[0]).toEqual({ kind: "page", page: "main", targets: [{ el: "nav-main", text: "Click Main" }] });
  });

  it("takes the way on the screen the user is on when there are two", () => {
    expect(planShowMe({ id: "recalc-run" }, "main")).toEqual([{ kind: "place", targets: [{ el: "main-recalc", text: "Click Recalculate" }] }]);
    expect(planShowMe({ id: "recalc-run" }, "advanced").map((s) => s.targets[0].el)).toEqual(["tab-recalc", "recalc-run"]);
    expect(planShowMe({ id: "download" }, "main").map((s) => s.targets[0].el)).toEqual(["download"]);
    expect(planShowMe({ id: "download" }, "advanced").map((s) => s.targets[0].el)).toEqual(["tab-mind", "mind-run"]);
  });

  it("points at the rule's own row when there is a rule, else at any finding", () => {
    expect(planShowMe({ id: "assistant-ask", ruleId: "FRM-009" }, "advanced")[1].targets).toEqual([
      { el: "ask-FRM-009", text: "Ask the assistant about FRM-009 here" },
      { el: "ask", text: "Ask the assistant about a finding here" },
      { el: "findings-list", text: "The findings are listed here" },
    ]);
    expect(planShowMe({ id: "assistant-ask" }, "advanced")[1].targets.map((t) => t.el)).toEqual(["ask", "findings-list"]);
  });
});

describe("Show me: where the arrow and the bubble go", () => {
  const view = { width: 1200, height: 800 };
  const bubble = { width: 200, height: 40 };

  it("puts them under the element when there is room", () => {
    const l = placeShowMe({ top: 8, left: 200, width: 80, height: 32 }, bubble, view);
    expect(l).toMatchObject({ side: "below", arrow: { left: 230, top: 46, width: 20, height: 26 }, bubble: { x: 140, y: 72 } });
    expect(l.ring).toEqual({ top: 4, left: 196, width: 88, height: 40 });
  });

  it("puts them above an element at the bottom of the window", () => {
    expect(placeShowMe({ top: 700, left: 200, width: 80, height: 40 }, bubble, view)).toMatchObject({ side: "above", arrow: { top: 668 }, bubble: { y: 628 } });
  });

  it("keeps the bubble inside the window, and inside a card taller than the window", () => {
    expect(placeShowMe({ top: 300, left: 1180, width: 20, height: 20 }, bubble, view).bubble.x).toBe(992);
    expect(placeShowMe({ top: 60, left: 0, width: 1000, height: 900 }, bubble, view)).toMatchObject({ side: "inside", arrow: { top: 66 } });
  });

  it("puts the place's bubble beside it when a button sits under it, and keeps that side", () => {
    const apply = { top: 600, left: 120, width: 190, height: 44 };
    const notUnder = (b: { top: number; left: number }) => b.top < 600 || b.left > 310; // "Add to the list" sits under Apply
    const l = placeShowMe(apply, bubble, view, PLACE_SIDES, notUnder);
    expect(l).toMatchObject({ side: "right", arrow: { left: 316, top: 612, width: 26, height: 20 }, bubble: { x: 342, y: 602 } });
    expect(placeShowMe(apply, bubble, view, PLACE_SIDES, () => true, "right").side).toBe("right"); // no jump while the page scrolls
    expect(placeShowMe(apply, bubble, view, PLACE_SIDES, () => false).side).toBe("below"); // nothing free: the first side that fits
  });
});

// ---- the pointer itself, on a page made of plain elements ----------------------------------------------
@Component({ standalone: true, template: "" })
class Blank {}

/** An element the pointer can find: drawn (a size), with its data-show-me word. */
function draw(word: string, size = { width: 80, height: 32 }): HTMLElement {
  const el = document.createElement("button");
  el.setAttribute("data-show-me", word);
  el.getBoundingClientRect = () => ({ top: 120, left: 100, ...size, right: 100 + size.width, bottom: 120 + size.height, x: 100, y: 120, toJSON: () => ({}) }) as DOMRect;
  document.body.appendChild(el);
  return el;
}

describe("Show me: the pointer", () => {
  let showMe: ShowMeService;
  let store: SessionStore;
  let router: Router;

  beforeEach(async () => {
    TestBed.configureTestingModule({ providers: [provideRouter([{ path: "", component: Blank }, { path: "advanced", component: Blank }])] });
    showMe = TestBed.inject(ShowMeService);
    store = TestBed.inject(SessionStore);
    router = TestBed.inject(Router);
    await router.navigateByUrl("/");
  });

  afterEach(() => {
    showMe.stop();
    vi.useRealTimers();
    document.body.innerHTML = "";
  });

  it("points at Advanced, then at the tab, then at the place, as each one is clicked", async () => {
    store.advancedTab.set("findings");
    const nav = draw("nav-advanced");
    vi.useFakeTimers();
    expect(showMe.start({ id: "prep-apply" })).toBe(true);
    expect(showMe.pointer()).toMatchObject({ el: nav, text: "Click Advanced", last: false });

    nav.click(); // the router link opens Advanced, which draws its tabs
    await router.navigateByUrl("/advanced?session=s1");
    const tab = draw("tab-prep");
    vi.advanceTimersByTime(SHOW_ME_AFTER_CLICK_MS);
    expect(showMe.pointer()).toMatchObject({ el: tab, text: "Click Prep actions", last: false });

    tab.click();
    store.advancedTab.set("prep"); // the tab's own handler
    const apply = draw("prep-apply");
    vi.advanceTimersByTime(SHOW_ME_AFTER_CLICK_MS);
    expect(showMe.pointer()).toMatchObject({ el: apply, text: "Apply the blocking repairs here", last: true });

    apply.click(); // at the place itself the user acts: the pointer goes
    expect(showMe.active()).toBe(false);
    expect(showMe.pointer()).toBeNull();
  });

  it("skips the clicks the user is already past", async () => {
    await router.navigateByUrl("/advanced?session=s1");
    store.advancedTab.set("prep");
    const apply = draw("prep-apply");
    showMe.start({ id: "prep-apply" });
    expect(showMe.pointer()).toMatchObject({ el: apply, last: true });

    store.advancedTab.set("findings");
    const tab = draw("tab-deliver");
    showMe.start({ id: "deliver-numbers" });
    expect(showMe.pointer()).toMatchObject({ el: tab, text: "Click Deliver" });
  });

  it("skips a click whose element never comes, and the rule's row when it is not there", async () => {
    vi.useFakeTimers();
    const tab = draw("tab-prep"); // no Advanced link drawn
    showMe.start({ id: "prep-apply" });
    expect(showMe.pointer()).toBeNull();
    vi.advanceTimersByTime(SHOW_ME_WAIT_MS);
    expect(showMe.pointer()).toMatchObject({ el: tab, text: "Click Prep actions" });

    await router.navigateByUrl("/advanced");
    store.advancedTab.set("findings");
    draw("ask", { width: 0, height: 0 }); // hidden: not pointed at
    const ask = draw("ask");
    showMe.start({ id: "assistant-ask", ruleId: "FRM-009" });
    expect(showMe.pointer()).toMatchObject({ el: ask, text: "Ask the assistant about a finding here", last: true });
  });

  it("ends quietly when the place itself is not on the page", async () => {
    vi.useFakeTimers();
    await router.navigateByUrl("/advanced");
    store.advancedTab.set("prep");
    expect(showMe.start({ id: "prep-apply" })).toBe(true);
    expect(showMe.active()).toBe(true);
    vi.advanceTimersByTime(SHOW_ME_WAIT_MS + 200);
    expect(showMe.active()).toBe(false);
  });

  it("ends on Esc and on a click anywhere else", () => {
    draw("nav-advanced");
    showMe.start({ id: "prep-apply" });
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(showMe.active()).toBe(false);

    showMe.start({ id: "prep-apply" });
    draw("something-else").click();
    expect(showMe.active()).toBe(false);
  });

  it("follows the element when Angular draws it again", async () => {
    vi.useFakeTimers();
    await router.navigateByUrl("/advanced");
    store.advancedTab.set("deliver");
    const first = draw("deliver-numbers");
    showMe.start({ id: "deliver-numbers" });
    expect(showMe.pointer()?.el).toBe(first);
    first.remove();
    const again = draw("deliver-numbers");
    vi.advanceTimersByTime(SHOW_ME_WATCH_MS);
    expect(showMe.pointer()?.el).toBe(again);
  });

  it("ends on Got it, and never takes the click meant for the page", async () => {
    await router.navigateByUrl("/advanced");
    store.advancedTab.set("deliver");
    const check = draw("deliver-numbers");
    let clicked = 0;
    check.addEventListener("click", () => clicked++);
    const fixture = TestBed.createComponent(ShowMeOverlay);
    showMe.start({ id: "deliver-numbers" });
    fixture.detectChanges();
    const host = fixture.nativeElement as HTMLElement;
    expect(host.textContent).toContain("Check the numbers here");
    const got = host.querySelector<HTMLButtonElement>(".show-me__done");
    expect(got?.textContent).toBe("Got it");
    got!.click();
    expect(showMe.active()).toBe(false);
    fixture.detectChanges();
    expect(host.textContent?.trim()).toBe("");

    showMe.start({ id: "deliver-numbers" });
    check.click(); // the place's own handler runs: the pointer only watches
    expect(clicked).toBe(1);
    expect(showMe.active()).toBe(false);
  });
});

// ---- the main screen's own texts, built by the page -----------------------------------------------------
describe("Show me on the main screen's texts", () => {
  let Page: typeof HomePage;
  let page: HomePage;

  beforeAll(async () => {
    // The page's logic, without drawing it: Carbon (a package Node cannot load as it is) stands in as
    // inert classes, as in outside-inputs.spec.ts.
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
  });

  beforeEach(() => {
    const api = {
      startAutoFix: () => Promise.reject(new Error("the test stops here")),
      estimate: () => Promise.reject(new Error("no estimate")),
      resume: () => Promise.reject(new Error("offline")),
      mindLoopPlan: () => Promise.reject(new Error("no loop plan")),
      downloadUrl: (id: string, name: string) => `/api/sessions/${id}/files/${encodeURIComponent(name)}`,
    } as unknown as ApiService;
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: api }] });
    page = TestBed.runInInjectionContext(() => new Page());
    TestBed.inject(SessionStore).sessionId.set("s1");
  });

  afterEach(() => page?.ngOnDestroy());

  it("says under which tab to fix what stopped Do everything, and that tab has a Show me", async () => {
    await page.runAll();
    expect(page.chainLine()).toBe('Stopped at "Fix all issues": Not started. Fix that under Advanced › Recalculate, then press Do everything again.');
    expect(showMeFor(page.chainLine())).toEqual({ id: "recalc-errors" });
  });

  it("names the Mind loop tab when a loop needs a decision", () => {
    page.loopPlan.set({ found: true, loops: [{ name: "Policy", status: "needs_decision", cells: [], first: 1, last: 40 }], outputs: [], clears: [], operation_count: 0 } as LoopPlan);
    expect(page.loopSummary()).toBe("1 loop(s) need a decision (Advanced › Mind loop)");
    expect(showMeFor(page.loopSummary())).toEqual({ id: "loop-choices" });
  });
});
