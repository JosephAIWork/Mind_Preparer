import { NO_ERRORS_SCHEMA, Type, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { InstancesPanel } from "../pages/advanced/instances-panel";
import { ApiService } from "./api.service";
import { FinderDimension, FinderResult, InstanceChoices, InstancePlan } from "./models";
import { SessionStore } from "./session.store";

// The Instances tab after the clean-up of 10 November 2026, one thing at a time: (1) one question, "Make
// one copy per…", and nothing else until it has an answer; (2) one line of counts and one Build it
// button, the cards behind "Details"; (3) after the build, the download and "Next: Run in Mind". The
// finder ("Where instances could help") sits behind "Not sure? Let the app look", top suggestion first.

const options = [
  { name: "CohortRng", driver: "Ctrl_01!C22", kind: "counter", count: 6, first: 2021, last: 2026, usable: true },
  { name: "Scenario", driver: "Ctrl_01!C27", kind: "list", count: 4, first: 0, last: 3, usable: true },
  { name: "Rows", driver: "Calc!D2", kind: "row_of_table", count: 9, usable: false, why: "the driver reads a row of a table" },
];

/** The plan by the choices: nothing chosen = no table key (the question stays open); loops = 24 copies; a column = 3 copies. */
const plan = (choices: InstanceChoices): InstancePlan => {
  const loops = (choices.loops ?? "").split(",").filter(Boolean);
  if (loops.length) {
    return {
      found: true,
      status: "ready",
      mode: "loops",
      operation_count: 10,
      loop_options: options,
      counts: { key_values: 24, result_cells: 522, inserted_rows: 3, cleared: 0 },
      key: { header: "Instance key", distinct: 24, column: "B" },
      copies_from: [
        { name: "CohortRng", driver: "Ctrl_01!C22", count: 6, formula: "=INDEX(Ctrl_01!$C$49:$C$72,1)" },
        { name: "Scenario", driver: "Ctrl_01!C27", count: 4, formula: "=INDEX(Ctrl_01!$D$49:$D$72,1)" },
      ],
      split_grid: { title: "#Instance values /InstanceSplit", title_cell: "Ctrl_01!B47" },
      results: { block: "Result_CF!A3:CI8", rows: 6, cols: 87, chosen_by: "outputs" },
      outputs: [{ block: "Out!A1:B3", status: "not_built", cells: 6, why: "no formula reads it" }],
      how_it_runs: "Mind runs the 24 copies at once and the main workbook gathers their results.",
    };
  }
  if (choices.key) {
    return {
      found: true,
      status: "ready",
      mode: "table",
      operation_count: 4,
      loop_options: options,
      counts: { key_values: 3, result_cells: 30 },
      key: { header: "Product", distinct: 3, column: choices.key, chosen_by: "user", candidates: [{ column: "C", header: "Product", distinct: 3 }], counts: { A: 4, B: 2, C: 3 } },
      table: { sheet: "MP", ref: "A1:F10", rows: 9 },
      loop: { name: "Rows", driver: "Calc!D2", count_cell: "Calc!D1" },
    };
  }
  return { found: true, status: "needs_decision", mode: "table", notes: ["No formula reads a row of a table."], operation_count: 0, loop_options: options };
};

const dimension = (n: number, column: string): FinderDimension =>
  ({
    id: `d${n}`,
    rank: n,
    score: 10 - n,
    label: `Product ${n}`,
    headline: `The macro groups by column ${column}.`,
    kinds: ["macro"],
    key: { sheet: "MP", column, distinct: 3, rows: 9 },
    copies: 3,
    biggest_copy: 4,
    formulas_saved: 12,
    cells_changed: 5,
    evidence: [{ kind: "macro", text: "the macro groups by it" }],
    fit: { instances: { fits: "yes", reasons: [] }, loop: { fits: "partly", reasons: [] }, leave: { fits: "no", reasons: [] } },
  }) as FinderDimension;

const finder: FinderResult = {
  schema: "finder",
  version: "1",
  file: "model.xlsm",
  counts: { sheets: 3, formulas: 1200, dimensions: 3, chains: 0, time_axes: 0, seconds: 2.5 },
  dimensions: [dimension(1, "C"), dimension(2, "D"), dimension(3, "E")],
  chains: [],
  macro_loops: [],
  skipped: [],
  assistant: { asked: false, available: false, text: null, ranking: null, explanations: {}, edits: {}, extra: [], message: null },
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

describe("the Instances tab: one question, then one line and one button", () => {
  let PanelClass: typeof InstancesPanel;
  let sent: InstanceChoices[];
  let finds: number;

  beforeAll(async () => {
    mockCarbon();
    PanelClass = (await import("../pages/advanced/instances-panel")).InstancesPanel;
  });

  async function draw(api: Record<string, unknown> = {}) {
    sent = [];
    finds = 0;
    const store = { sessionId: signal("s1"), loopChoices: signal(null), applyAnalysis: () => undefined };
    const calls = {
      mindInstancesPlan: (_id: string, o: InstanceChoices) => {
        sent.push(o);
        return Promise.resolve({ version_id: "v1", plan: plan(o) });
      },
      assistantKey: () => Promise.resolve({ saved: false }),
      instanceFinder: () => {
        finds += 1;
        return Promise.resolve({ version_id: "v1", finder });
      },
      ...api,
    };
    return { store, ...(await render(PanelClass, calls, store)) };
  }

  const boxes = (page: HTMLElement) => Array.from(page.querySelectorAll(".copies-from input")) as HTMLInputElement[];
  const folds = (page: HTMLElement) => Array.from(page.querySelectorAll("details")) as HTMLDetailsElement[];
  const summaries = (page: HTMLElement) => folds(page).map((d) => d.querySelector("summary")?.textContent?.trim());

  it("shows nothing but the question before a choice", async () => {
    const { page, panel } = await draw();
    expect(page.querySelector('[data-test="inst-question"]')?.textContent?.trim()).toBe("Make one copy per…");
    expect(page.querySelector('[data-show-me="instances-choice"]')).not.toBeNull();
    expect(boxes(page).length).toBe(3);
    expect(panel.chosen()).toBe(false);
    // no summary, no Build it, no Details, no card, no counts
    expect(page.querySelector('[data-test="inst-summary"]')).toBeNull();
    expect(page.querySelector('[data-show-me="instances-build"]')).toBeNull();
    expect(page.querySelector('[data-test="inst-details"]')).toBeNull();
    expect(page.querySelectorAll(".card").length).toBe(0);
    expect(page.textContent).not.toContain("result cells");
    // the only other things: the finder behind its button, and "What is this?", both closed
    expect(summaries(page)).toEqual(["Not sure? Let the app look", "What is this?"]);
    expect(folds(page).map((d) => d.open)).toEqual([false, false]);
    expect(finds).toBe(0);
  });

  it("after a choice: one line of counts, one Build it button, the cards behind Details (closed)", async () => {
    const { page, panel, detect } = await draw();
    boxes(page)[0].click();
    await detect();
    boxes(page)[1].click();
    await detect();
    expect(sent.at(-1)).toEqual({ loops: "CohortRng,Scenario" });
    expect(panel.chosen()).toBe(true);
    const line = page.querySelector('[data-test="inst-summary"]');
    expect(line?.textContent?.replace(/\s+/g, " ").trim()).toBe("ready 24 copies · 522 result cells");
    const build = page.querySelector('[data-show-me="instances-build"]') as HTMLButtonElement;
    expect(build.textContent?.trim()).toBe("Build it (10 changes)");
    expect(build.disabled).toBe(false);
    expect(page.textContent).toContain("One copy of the model per combination of CohortRng and Scenario.");
    // the Details fold: closed, and every card inside it
    const details = page.querySelector('[data-test="inst-details"]') as HTMLDetailsElement;
    expect(details.open).toBe(false);
    expect(details.querySelector("summary")?.textContent?.trim()).toBe("Details");
    expect(page.querySelectorAll(".card").length).toBeGreaterThan(0);
    for (const card of Array.from(page.querySelectorAll(".card"))) expect(card.closest('[data-test="inst-details"]')).not.toBeNull();
    const text = details.textContent ?? "";
    expect(text).toContain("Copies from CohortRng (6) × Scenario (4)");
    expect(text).toContain("Result_CF!A3:CI8");
    expect(text).toContain("Mind runs the 24 copies at once");
    expect(text).toContain("Show the changes");
    expect(summaries(page)).toEqual(["Not sure? Let the app look", "Details", "What is this?"]);
  });

  it("uses plain words: no key column, no decision, no 'write it'; MM_INSTANCE only behind What is this?", async () => {
    const { page, panel, detect } = await draw();
    panel.useAsKey(dimension(1, "C"));
    await detect();
    expect(sent.at(-1)).toEqual({ key: "C", keys: undefined });
    const chip = page.querySelector('[data-test="inst-summary"] .chip');
    expect(chip?.textContent?.trim()).toBe("ready");
    expect(panel.statusText("needs_decision")).toBe("needs one answer");
    expect(panel.statusText("not_built")).toBe("left as is");
    const details = page.querySelector('[data-test="inst-details"]');
    expect(details?.textContent).toContain("One copy per Product");
    expect(details?.textContent).toContain("Also keys with no rows");
    expect(details?.textContent).toContain("the cell the macro changes");
    const all = page.textContent ?? "";
    // (the backend's own why texts, "the driver reads a row of a table", are shown as they come)
    expect(all).not.toMatch(/Key column|Every key of the list|Write it into a new version|needs a decision|not built|Where instances could help/);
    const what = page.querySelector('[data-test="inst-what"]') as HTMLDetailsElement;
    expect(what.open).toBe(false);
    expect(what.textContent).toContain("MM_INSTANCE");
    expect(all.replace(what.textContent ?? "", "")).not.toContain("MM_INSTANCE");
  });

  it("keeps the finder behind its button: top suggestion first, Use it answers the question", async () => {
    const { page, panel, detect } = await draw();
    const fold = page.querySelector('[data-test="inst-finder"]') as HTMLDetailsElement;
    expect(fold.open).toBe(false);
    expect(finds).toBe(0);
    (fold.querySelector("summary") as HTMLElement).click();
    await detect();
    expect(finds).toBe(1);
    expect(panel.finderOpen()).toBe(true);
    // one suggestion, with Use it; the other two behind Show all
    expect(page.querySelectorAll(".suggestion").length).toBe(1);
    expect(page.querySelector(".suggestion .card__title")?.textContent).toContain("#1 Product 1");
    expect(fold.textContent).toContain("1,200 formulas · 3 places found · 0 set aside · 2.5 s");
    const chips = Array.from(page.querySelectorAll(".suggestion .card__meta .chip")).map((c) => c.textContent?.trim());
    expect(chips).toEqual(["copies: fits", "a loop: fits partly", "leave as is: does not fit"]);
    const showAll = page.querySelector('[data-test="inst-show-all"]') as HTMLButtonElement;
    expect(showAll.textContent?.trim()).toBe("Show all 3");
    showAll.click();
    await detect();
    expect(page.querySelectorAll(".suggestion").length).toBe(3);
    // the assistant tick stays inside the fold (shown when a key is saved)
    expect(fold.textContent).toContain("Look again");
    (page.querySelector(".suggestion .use-it") as HTMLButtonElement).click();
    await detect();
    expect(sent.at(-1)).toEqual({ key: "C", keys: undefined });
    expect(panel.finderOpen()).toBe(false);
    expect(page.querySelector('[data-test="inst-summary"]')?.textContent).toContain("3 copies · 30 result cells");
  });

  it("after the build: Done, the download, and Next: Run in Mind", async () => {
    const api = {
      applyMindInstancesPlan: (_id: string, o: InstanceChoices) => Promise.resolve({ version: { id: "v2", label: "v2 instances" }, plan: plan(o), download: true, files: ["model_instances.xlsm", "model_Main.xlsx"], result: { applied: 10, failed: [] } }),
      mindInstancesDownloadUrl: () => "/api/sessions/s1/mind-instances/download",
      resume: () => Promise.resolve({}),
      scanStatus: () => Promise.resolve({}),
    };
    const { page, panel, detect } = await draw(api);
    boxes(page)[0].click();
    await detect();
    boxes(page)[1].click();
    await detect();
    expect(page.querySelector('[data-test="inst-next"]')).toBeNull();
    await panel.write();
    await detect();
    expect(panel.line()).toBe("10 changes written");
    const result = page.querySelector('[data-show-me="instances-result"]');
    expect(result?.querySelector('[data-test="inst-next"]')?.textContent).toContain("Done. Download both files (zip). Next: Run in Mind - upload both files as one project");
    expect((page.querySelector('[data-test="inst-download"]') as HTMLAnchorElement).getAttribute("href")).toBe("/api/sessions/s1/mind-instances/download");
    expect((page.querySelector('[data-show-me="instances-build"]') as HTMLButtonElement).disabled).toBe(true);
    // other choices: a new plan, the line and the download go
    boxes(page)[1].click();
    await detect();
    expect(page.querySelector('[data-test="inst-next"]')).toBeNull();
    expect(page.querySelector('[data-test="inst-download"]')).toBeNull();
  });
});
