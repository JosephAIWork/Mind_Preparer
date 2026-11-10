import { NO_ERRORS_SCHEMA, Type, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { AdvancedPage } from "../pages/advanced.page";
import type { LoopPanel } from "../pages/advanced/loop-panel";
import { ApiService } from "./api.service";
import { LoopChoices, LoopPlan, LoopPlanLoop } from "./models";
import { SessionStore } from "./session.store";

// The Mind loop tab after the clean-up of 10 November 2026: one sentence, one Build it button; the
// choices and what gets written behind folds (closed until opened); a loop that still needs one answer
// shows only that box, above the folds; the instances builder on its own tab.

const cohort: LoopPlanLoop = { name: "CohortRng", driver: "Ctrl_01!C22", status: "ready", cells: [{}, {}], first: 2021, last: 2026 };
const scenarioOpen: LoopPlanLoop = {
  name: "Scenario",
  driver: "Ctrl_01!C27",
  status: "needs_decision",
  cells: [],
  first: 0,
  last: undefined,
  why: "MM_LOOP needs a cell that holds the count (4); the macro's last value is Scenarios, not a cell. Name a free cell for it (count cell) and the plan writes 4 there. A free cell: Ctrl_01!N27.",
  suggested_count_cell: "Ctrl_01!N27",
};
const scenarioReady: LoopPlanLoop = { name: "Scenario", driver: "Ctrl_01!C27", status: "ready", cells: [{}], first: 0, last: 3 };

const plan = (over: Partial<LoopPlan> = {}): LoopPlan => ({
  found: true,
  loops: [cohort, scenarioOpen],
  outputs: [{ target: "Out!A1:J6", status: "ready", cells: 60 }, { target: "Other!A1", status: "not_built", why: "no loop was built" }],
  clears: [{ target: "Out!A1:J6", cells: 12, included: false }],
  operation_count: 7,
  ...over,
});

function mockCarbon(): void {
  vi.doMock("carbon-components-angular", () => {
    class Inert {}
    return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
  });
}

async function render<T>(cls: Type<T>, api: unknown, store: unknown): Promise<{ page: HTMLElement; panel: T; detect: () => Promise<void> }> {
  TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: api }, { provide: SessionStore, useValue: store }] });
  TestBed.overrideComponent(cls, { set: { imports: [], schemas: [NO_ERRORS_SCHEMA] } });
  const fixture = TestBed.createComponent(cls);
  const detect = async () => {
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();
  };
  await detect();
  return { page: fixture.nativeElement as HTMLElement, panel: fixture.componentInstance, detect };
}

describe("the Mind loop tab: one sentence, one button, the rest folded", () => {
  let PanelClass: typeof LoopPanel;
  let sent: LoopChoices[];
  let served: LoopPlan;

  beforeAll(async () => {
    mockCarbon();
    PanelClass = (await import("../pages/advanced/loop-panel")).LoopPanel;
  });

  async function draw(p: LoopPlan = plan(), api: Record<string, unknown> = {}) {
    sent = [];
    served = p;
    const store = {
      sessionId: signal("s1"),
      loopChoices: signal(null),
      advancedTab: signal("loop"),
      applyAnalysis: () => undefined,
    };
    const calls = {
      mindLoopPlan: (_id: string, o: LoopChoices) => {
        sent.push(o);
        return Promise.resolve({ version_id: "v1", plan: served });
      },
      ...api,
    };
    return { store, ...(await render(PanelClass, calls, store)) };
  }

  const folds = (page: HTMLElement) => Array.from(page.querySelectorAll("details")) as HTMLDetailsElement[];

  it("opens with one sentence from the plan, and the folds closed", async () => {
    const { page } = await draw(plan({ loops: [cohort, scenarioReady] }));
    expect(page.querySelector('[data-test="loop-sentence"]')?.textContent?.trim()).toBe("The macro runs CohortRng 2021 to 2026 and Scenario 0 to 3. Mind can do the same.");
    const build = page.querySelector('[data-show-me="loop-build"]') as HTMLButtonElement;
    expect(build.textContent?.trim()).toBe("Build it (7 formulas)");
    expect(build.disabled).toBe(false);
    const all = folds(page);
    expect(all.map((d) => d.querySelector("summary")?.textContent?.trim())).toEqual(["Change the choices", "What gets written (7)", "What is this?"]);
    expect(all.map((d) => d.open)).toEqual([false, false, false]);
    expect(page.querySelector('[data-show-me="loop-choices"]')?.tagName).toBe("DETAILS");
    // nothing above the folds but the sentence and the button: no answer box, no cards
    expect(page.querySelectorAll(".loop-answer").length).toBe(0);
    expect(page.querySelector(".card")?.closest("details")).not.toBeNull();
  });

  it("says in plain words what is still needed, and shows only that loop's missing box above the folds", async () => {
    const { page, detect } = await draw();
    expect(page.querySelector('[data-test="loop-sentence"]')?.textContent?.trim()).toBe("The macro runs CohortRng 2021 to 2026 and Scenario from 0. Mind can do the same, after one answer below.");
    const answers = page.querySelectorAll(".loop-answer");
    expect(answers.length).toBe(1);
    expect(answers[0].closest("details")).toBeNull(); // above the folds
    expect(answers[0].getAttribute("data-show-me")).toBe("loop-answer");
    const boxes = answers[0].querySelectorAll("input");
    expect(boxes.length).toBe(1);
    expect(boxes[0].getAttribute("placeholder")).toBe("Ctrl_01!N27");
    const label = answers[0].querySelector("label")?.textContent ?? "";
    expect(label).toContain("Mind needs a cell that holds the number of passes (4)");
    expect(label).toContain("Name a free cell for it and the plan writes 4 there.");
    expect(label).not.toMatch(/count cell|MM_LOOP/);
    // the offered cell in one click, then the plan is read again with it
    (answers[0].querySelector("button") as HTMLButtonElement).click();
    await detect();
    expect(sent.at(-1)!.loops).toEqual({ "Ctrl_01!C27": { count_cell: "Ctrl_01!N27" } });
    // every box is still there, behind the fold
    expect(page.querySelectorAll("details .loop-choices").length).toBe(2);
  });

  it("asks for the first or the last value when that is what the plan lacks", async () => {
    const noLast: LoopPlanLoop = { name: "Year", driver: "Ctrl!B2", status: "needs_decision", cells: [], first: 1, why: "the first or last value of the loop could not be read from the workbook (LastYear): type them in" };
    const { page, panel } = await draw(plan({ loops: [noLast] }));
    expect(panel.missingKey(noLast)).toBe("last");
    const box = page.querySelector(".loop-answer input") as HTMLInputElement;
    expect(box.getAttribute("placeholder")).toBe("a number, or a cell (Sheet!A1)");
    box.value = "5";
    box.dispatchEvent(new Event("input"));
    expect(panel.pick(noLast, "last")).toBe("5");
    expect(panel.missingKey({ ...noLast, first: undefined })).toBe("first");
    expect(panel.missingKey({ ...noLast, status: "ready" })).toBeNull();
  });

  it("uses plain words for the choices and the chips", async () => {
    const { page } = await draw(plan({ loops: [cohort, scenarioOpen, { name: "GroupName", driver: "R!B4", status: "not_built", other_macro: true, cells: [], why: "a loop of the macro EndemicCalculationRun" }] }));
    const text = page.querySelector('[data-show-me="loop-choices"]')?.textContent ?? "";
    expect(text).toContain("the cell the macro changes");
    expect(text).toContain("Runs from");
    expect(text).toContain("A cell for the number of passes");
    expect(text).toContain("Only today's rows");
    expect(text).not.toMatch(/\bdriver\b|Fixed ranges|Count cell|needs a decision|not built/);
    const chips = Array.from(page.querySelectorAll('[data-show-me="loop-choices"] .chip')).map((c) => c.textContent?.trim());
    expect(chips).toEqual(["ready", "needs one answer", "left as is"]);
    // the fold with what gets written: the output cards and the cleared areas
    const written = page.querySelector('[data-test="loop-written"]')?.textContent ?? "";
    expect(written).toContain("Output Out!A1:J6");
    expect(written).toContain("Cleared before a run");
    expect(written).toContain("Show the formulas");
    // the MM_LOOP / MM_RESULT explanation only behind "What is this?"
    const what = page.querySelector('[data-test="loop-what"]')?.textContent ?? "";
    expect(what).toContain("MM_LOOP");
    expect(page.querySelector('[data-test="loop-sentence"]')?.textContent).not.toContain("MM_LOOP");
  });

  it("offers the Instances tab as the alternative, in one line", async () => {
    const { page, store } = await draw();
    const link = page.querySelector('[data-show-me="link-instances"]') as HTMLButtonElement;
    expect(link.textContent?.trim()).toBe("Instances");
    expect(link.closest("p")?.textContent?.replace(/\s+/g, " ").trim()).toBe("One model that loops (usual). Or one copy of the model per value: Instances");
    link.click();
    expect(store.advancedTab()).toBe("instances");
  });

  it("after a build, says what to do next: Run in Mind", async () => {
    const api = {
      applyMindLoopPlan: () => Promise.resolve({ version: { id: "v2", label: "v2 loop" }, plan: plan({ loops: [cohort, scenarioReady] }), result: { applied: 7, failed: [] } }),
      resume: () => Promise.resolve({}),
      scanStatus: () => Promise.resolve({}),
    };
    const { page, panel, detect } = await draw(plan({ loops: [cohort, scenarioReady] }), api);
    expect(page.querySelector('[data-test="loop-next"]')).toBeNull();
    await panel.write();
    await detect();
    expect(panel.line()).toBe("7 loop formulas written");
    const result = page.querySelector('[data-show-me="loop-result"]');
    expect(result?.querySelector('[data-test="loop-next"]')?.textContent).toContain("Done. Next: Run in Mind");
    // a new preview (other choices) drops the line
    await panel.preview();
    await detect();
    expect(page.querySelector('[data-test="loop-next"]')).toBeNull();
  });
});

describe("the Advanced page: the loop and the instances on their own tabs", () => {
  let PageClass: typeof AdvancedPage;

  beforeAll(async () => {
    mockCarbon();
    PageClass = (await import("../pages/advanced.page")).AdvancedPage;
  });

  // The page pulls in every panel and the readiness banner under the inert Carbon: the spec files
  // share one module cache here, so the modules are let go for the files that come after (the
  // banner's own spec draws a Carbon notification of its own).
  afterAll(() => {
    vi.doUnmock("carbon-components-angular");
    vi.resetModules();
  });

  async function draw() {
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: {} }] });
    TestBed.overrideComponent(PageClass, { set: { imports: [], schemas: [NO_ERRORS_SCHEMA] } });
    const store = TestBed.inject(SessionStore);
    store.sessionId.set("s1");
    store.report.set({ findings: [] } as unknown as NonNullable<ReturnType<SessionStore["report"]>>);
    const fixture = TestBed.createComponent(PageClass);
    fixture.detectChanges();
    return { page: fixture.nativeElement as HTMLElement, store, detect: () => fixture.detectChanges() };
  }

  it("has a tab named Instances next to Mind loop, and shows one builder at a time", async () => {
    const { page, store, detect } = await draw();
    const tabs = Array.from(page.querySelectorAll(".switch__item")).map((b) => [b.getAttribute("data-show-me"), b.textContent?.trim()]);
    expect(tabs).toContainEqual(["tab-loop", "Mind loop"]);
    expect(tabs).toContainEqual(["tab-instances", "Instances"]);
    expect(tabs.findIndex((t) => t[0] === "tab-instances")).toBe(tabs.findIndex((t) => t[0] === "tab-loop") + 1);

    store.advancedTab.set("loop");
    detect();
    expect(page.querySelector("mr-loop-panel")).not.toBeNull();
    expect(page.querySelector("mr-instances-panel")).toBeNull();

    (page.querySelector('[data-show-me="tab-instances"]') as HTMLButtonElement).click();
    detect();
    expect(store.advancedTab()).toBe("instances");
    expect(page.querySelector("mr-instances-panel")).not.toBeNull();
    expect(page.querySelector("mr-loop-panel")).toBeNull();
  });
});
