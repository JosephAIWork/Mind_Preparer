import { NO_ERRORS_SCHEMA, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { InstancesPanel } from "../pages/advanced/instances-panel";
import { ApiService } from "./api.service";
import { InstanceChoices, InstancePlan } from "./models";
import { SessionStore } from "./session.store";

// the Mind instances panel: copies made from loops (a scenario x a cohort) instead of a key column
const options = [
  { name: "CohortRng", driver: "Ctrl_01!C22", kind: "counter", count: 6, first: 2021, last: 2026, usable: true },
  { name: "Scenario", driver: "Ctrl_01!C27", kind: "list", count: 4, first: 0, last: 3, usable: true },
  { name: "Rows", driver: "Calc!D2", kind: "row_of_table", count: 9, usable: false, why: "the driver reads a row of a table" },
];

const plan = (choices: InstanceChoices): InstancePlan => {
  const loops = (choices.loops ?? "").split(",").filter(Boolean);
  if (!loops.length) {
    return { found: true, status: "needs_decision", mode: "table", notes: ["No formula reads a row of a table."], operation_count: 0, loop_options: options };
  }
  return {
    found: true,
    status: "ready",
    mode: "loops",
    operation_count: 10,
    loop_options: options,
    key: { header: "Instance key", distinct: 24, column: "B" },
    copies_from: [
      { name: "CohortRng", driver: "Ctrl_01!C22", count: 6, formula: "=INDEX(Ctrl_01!$C$49:$C$72,1)" },
      { name: "Scenario", driver: "Ctrl_01!C27", count: 4, formula: "=INDEX(Ctrl_01!$D$49:$D$72,1)" },
    ],
    split_grid: { title: "#Instance values /InstanceSplit", title_cell: "Ctrl_01!B47" },
    kept_loops: [{ name: "GroupName" }, { name: "CohortYear" }],
    results: { block: "Result_CF!A3:CI8", rows: 6, cols: 87, chosen_by: "outputs" },
  };
};

describe("the Mind instances panel's copies from loops", () => {
  let PanelClass: typeof InstancesPanel;
  let sent: InstanceChoices[];

  beforeAll(async () => {
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    PanelClass = (await import("../pages/advanced/instances-panel")).InstancesPanel;
  });

  async function draw(): Promise<{ page: HTMLElement; panel: InstancesPanel; detect: () => Promise<void> }> {
    sent = [];
    const api = {
      mindInstancesPlan: (_id: string, o: InstanceChoices) => {
        sent.push(o);
        return Promise.resolve({ version_id: "v1", plan: plan(o) });
      },
      assistantKey: () => Promise.resolve({ saved: false }),
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

  const boxes = (page: HTMLElement) => Array.from(page.querySelectorAll(".copies-from input")) as HTMLInputElement[];

  it("offers every loop as a tick box, the ones copies cannot come from greyed out", async () => {
    const { page } = await draw();
    const b = boxes(page);
    expect(b.length).toBe(3);
    expect(b.map((x) => x.disabled)).toEqual([false, false, true]);
    expect(page.querySelector(".copies-from")?.textContent).toContain("6 values, 2021 to 2026");
    expect(sent).toEqual([{ key: undefined, keys: undefined }]); // the first preview: by the key column
  });

  it("sends the ticked loops in the plan's order and shows the copies card", async () => {
    const { page, panel, detect } = await draw();
    boxes(page)[1].click(); // Scenario first ...
    await detect();
    boxes(page)[0].click(); // ... then CohortRng: still sent in the macro's order
    await detect();
    expect(sent.at(-1)).toEqual({ loops: "CohortRng,Scenario" });
    expect(panel.loopCopies()).toBe(24);
    const card = page.querySelector(".copies-card")?.textContent ?? "";
    expect(card).toContain("CohortRng (6) × Scenario (4)");
    expect(card).toContain("24 copies");
    expect(card).toContain("GroupName");
    expect(card).toContain("Result_CF!A3:CI8");
    expect(page.textContent).toContain("One copy of the model per combination of CohortRng and Scenario");
  });

  it("goes back to the key column when a key is picked", async () => {
    const { page, panel, detect } = await draw();
    boxes(page)[0].click();
    await detect();
    expect(panel.loops()).toEqual(["CohortRng"]);
    panel.useAsKey({ key: { column: "C" } } as never);
    await detect();
    expect(panel.loops()).toEqual([]);
    expect(sent.at(-1)).toEqual({ key: "C", keys: undefined });
  });
});
