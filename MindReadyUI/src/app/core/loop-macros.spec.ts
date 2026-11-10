import { NO_ERRORS_SCHEMA, Type, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { InstancesPanel } from "../pages/advanced/instances-panel";
import type { LoopPanel } from "../pages/advanced/loop-panel";
import { ApiService } from "./api.service";
import { InstanceChoices, InstancePlan, LoopChoice, LoopChoices, LoopPlan } from "./models";
import { SessionStore } from "./session.store";

// review of 10 November 2026: one macro's loops at a time, and the instances panel uses the loop panel's choices
const loopPlan = (macro: string): LoopPlan => ({
  found: true,
  macro,
  macros: [
    { name: "AllScenarioRun", loops: ["CohortRng", "Scenario"], chosen: macro === "AllScenarioRun" },
    { name: "EndemicCalculationRun", loops: ["GroupName", "CohortYear"], chosen: macro === "EndemicCalculationRun" },
  ],
  loops: [
    { name: "CohortRng", driver: "Ctrl_01!C22", status: macro === "AllScenarioRun" ? "ready" : "not_built", other_macro: macro !== "AllScenarioRun", cells: [{}], first: 2021, last: 2026 },
    { name: "GroupName", driver: "RollForward!B4", status: macro === "AllScenarioRun" ? "not_built" : "ready", other_macro: macro === "AllScenarioRun", cells: [], why: "a loop of the macro EndemicCalculationRun" },
  ],
  outputs: [],
  clears: [],
  operation_count: 1,
});

const instancePlan: InstancePlan = {
  found: true,
  status: "ready",
  mode: "loops",
  operation_count: 5,
  loop_options: [{ name: "CohortRng", driver: "Ctrl_01!C22", kind: "counter", count: 6, first: 2021, last: 2026, usable: true }],
};

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

describe("the Mind loop panel with two macros", () => {
  let PanelClass: typeof LoopPanel;
  beforeAll(async () => {
    mockCarbon();
    PanelClass = (await import("../pages/advanced/loop-panel")).LoopPanel;
  });

  it("offers the macros, builds the chosen one's loops and shares the choices", async () => {
    const sent: LoopChoices[] = [];
    const api = {
      mindLoopPlan: (_id: string, o: LoopChoices) => {
        sent.push(o);
        return Promise.resolve({ version_id: "v1", plan: loopPlan(o.macro || "AllScenarioRun") });
      },
    };
    const store = { sessionId: signal("s1"), loopChoices: signal<{ loops?: Record<string, LoopChoice>; macro?: string } | null>(null) };
    const { page, panel, detect } = await render(PanelClass, api, store);
    const select = page.querySelector(".macro-choice select") as HTMLSelectElement;
    expect(Array.from(select.options).map((o) => o.textContent)).toEqual(["AllScenarioRun (CohortRng, Scenario)", "EndemicCalculationRun (GroupName, CohortYear)"]);
    expect(page.querySelectorAll(".loop-choices").length).toBe(1); // the other macro's loop has no boxes
    select.value = "EndemicCalculationRun";
    select.dispatchEvent(new Event("change"));
    await detect();
    expect(sent.at(-1)!.macro).toBe("EndemicCalculationRun");
    expect(panel.plan()!.macro).toBe("EndemicCalculationRun");
    expect(store.loopChoices()).toEqual({ loops: undefined, macro: "EndemicCalculationRun" });
    panel.setPick(panel.plan()!.loops[1], "last", "3");
    await panel.preview();
    expect(store.loopChoices()).toEqual({ loops: { "RollForward!B4": { last: "3" } }, macro: "EndemicCalculationRun" });
  });
});

describe("the Mind instances panel and the loop panel's choices", () => {
  let PanelClass: typeof InstancesPanel;
  beforeAll(async () => {
    mockCarbon();
    PanelClass = (await import("../pages/advanced/instances-panel")).InstancesPanel;
  });

  async function draw(start: { loops?: Record<string, LoopChoice>; macro?: string } | null) {
    const sent: InstanceChoices[] = [];
    const api = {
      mindInstancesPlan: (_id: string, o: InstanceChoices) => {
        sent.push(o);
        return Promise.resolve({ version_id: "v1", plan: instancePlan });
      },
      assistantKey: () => Promise.resolve({ saved: false }),
    };
    const store = { sessionId: signal("s1"), loopChoices: signal(start) };
    return { sent, store, ...(await render(PanelClass, api, store)) };
  }

  it("sends each loop's choices and the macro, and says so", async () => {
    const { sent, page, panel, detect } = await draw({ loops: { "Ctrl_01!C22": { first: "2022", last: "2025" } }, macro: "AllScenarioRun" });
    expect(sent[0]).toEqual({ key: undefined, keys: undefined, loop_choices: { "Ctrl_01!C22": { first: "2022", last: "2025" } }, macro: "AllScenarioRun" });
    expect(page.querySelector(".loop-choices-used")?.textContent).toContain("Mind loop panel");
    panel.toggleLoop("CohortRng");
    await detect();
    expect(sent.at(-1)).toEqual({ loops: "CohortRng", loop_choices: { "Ctrl_01!C22": { first: "2022", last: "2025" } } });
  });

  it("previews again when the loop panel's choices change", async () => {
    const { sent, store, detect } = await draw(null);
    expect(sent.length).toBe(1);
    store.loopChoices.set({ loops: { "Ctrl_01!C27": { count_cell: "Ctrl_01!N27" } } });
    await detect();
    expect(sent.length).toBe(2);
    expect(sent.at(-1)!.loop_choices).toEqual({ "Ctrl_01!C27": { count_cell: "Ctrl_01!N27" } });
  });

  it("drops the earlier write's download when a new plan is previewed", async () => {
    const { panel, page, detect } = await draw(null);
    panel.downloadUrl.set("/api/sessions/s1/mind-instances/download");
    panel.files.set(["model_instances.xlsm", "model_Main.xlsx"]);
    panel.written.set(true);
    await panel.preview();
    await detect();
    expect(panel.downloadUrl()).toBeNull();
    expect(panel.files()).toEqual([]);
    expect(page.textContent).not.toContain("Download both files");
  });
});
