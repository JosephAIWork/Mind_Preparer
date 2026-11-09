import { NO_ERRORS_SCHEMA, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { LoopPanel } from "../pages/advanced/loop-panel";
import { ApiService } from "./api.service";
import { LoopChoices, LoopPlan } from "./models";
import { SessionStore } from "./session.store";

// the Mind loop panel: every loop has its own choices, sent by driver cell
const plan = (over: Partial<LoopPlan> = {}): LoopPlan => ({
  found: true,
  loops: [
    { name: "CohortRng", driver: "Ctrl_01!C22", status: "ready", cells: [{}], first: 2021, last: 2026 },
    { name: "Scenario", driver: "Ctrl_01!C27", status: "needs_decision", cells: [], first: 0, last: undefined, why: "MM_LOOP needs a cell that holds the count (count cell). A free cell: Ctrl_01!N27.", suggested_count_cell: "Ctrl_01!N27" },
  ],
  outputs: [],
  clears: [],
  operation_count: 1,
  ...over,
});

describe("the Mind loop panel's choices", () => {
  let PanelClass: typeof LoopPanel;
  let sent: LoopChoices[];

  beforeAll(async () => {
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    PanelClass = (await import("../pages/advanced/loop-panel")).LoopPanel;
  });

  async function draw(): Promise<{ page: HTMLElement; panel: LoopPanel; detect: () => Promise<void> }> {
    sent = [];
    const api = {
      mindLoopPlan: (_id: string, o: LoopChoices) => {
        sent.push(o);
        return Promise.resolve({ version_id: "v1", plan: plan() });
      },
    } as unknown as ApiService;
    const store = { sessionId: signal("s1") } as unknown as SessionStore;
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: api }, { provide: SessionStore, useValue: store }] });
    TestBed.overrideComponent(PanelClass, { set: { imports: [], schemas: [NO_ERRORS_SCHEMA] } });
    const fixture = TestBed.createComponent(PanelClass);
    const detect = async () => {
      fixture.detectChanges();
      await fixture.whenStable();
      fixture.detectChanges();
    };
    await detect();
    return { page: fixture.nativeElement as HTMLElement, panel: fixture.componentInstance, detect };
  }

  it("gives every loop its own boxes, with what the model says greyed in", async () => {
    const { page } = await draw();
    const rows = page.querySelectorAll(".loop-choices");
    expect(rows.length).toBe(2);
    const boxes = (i: number) => Array.from(rows[i].querySelectorAll("input")).map((b) => b.getAttribute("placeholder"));
    expect(boxes(0)).toEqual(["CohortRng", "2021", "2026"]); // no count box: the loop asks for none
    expect(boxes(1)).toEqual(["Scenario", "0", "type it", "Ctrl_01!N27"]);
    expect(rows[1].querySelector("button")?.textContent?.trim()).toBe("Use Ctrl_01!N27");
    expect(sent).toEqual([{ clear_old: undefined, fixed_ranges: undefined, loops: undefined }]); // nothing typed: nothing sent
  });

  it("sends what was typed for each loop, by its driver cell, and drops empty boxes", async () => {
    const { panel } = await draw();
    const [cohort, scenario] = panel.plan()!.loops;
    panel.setPick(scenario, "last", " 3 ");
    panel.setPick(cohort, "name", "");
    panel.setPick(cohort, "first", "Ctrl_01!C24");
    await panel.preview();
    expect(sent.at(-1)!.loops).toEqual({ "Ctrl_01!C27": { last: "3" }, "Ctrl_01!C22": { first: "Ctrl_01!C24" } });
  });

  it("takes the offered count cell in one click and previews again", async () => {
    const { page, panel, detect } = await draw();
    (page.querySelectorAll(".loop-choices")[1].querySelector("button") as HTMLButtonElement).click();
    await detect();
    expect(sent.at(-1)!.loops).toEqual({ "Ctrl_01!C27": { count_cell: "Ctrl_01!N27" } });
    expect(panel.pick(panel.plan()!.loops[1], "count_cell")).toBe("Ctrl_01!N27");
  });
});
