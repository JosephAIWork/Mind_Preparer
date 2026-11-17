import { CommonModule } from "@angular/common";
import { NO_ERRORS_SCHEMA, Type } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { AdvancedPage } from "../pages/advanced.page";
import type { AssistantPanel } from "../pages/advanced/assistant-panel";
import type { DeliverPanel } from "../pages/advanced/deliver-panel";
import type { GridNamerPanel } from "../pages/advanced/grid-namer-panel";
import type { InstancesPanel } from "../pages/advanced/instances-panel";
import type { LoopPanel } from "../pages/advanced/loop-panel";
import type { MindRunPanel } from "../pages/advanced/mind-run-panel";
import type { PrepPanel } from "../pages/advanced/prep-panel";
import type { RecalcPanel } from "../pages/advanced/recalc-panel";
import type { StackBar } from "../pages/advanced/stack-bar";
import type { ApiService } from "./api.service";
import type { CancelButton } from "./cancel-button";
import type { ChangeStack } from "./change-stack";
import { CancelResult, FinderResult, InstancePlan, LoopPlan, MindDelete, MindLoopStatus, MindProject, MindRunTiming, MindSignIn, Operation, PrepAction, RecalcResult, SessionStart, Version } from "./models";
import type { SessionStore } from "./session.store";

// The Cancel button in the Advanced panels (17 November 2026; DESIGN.md section 4, rows 8 to 17). One
// test per row: the button is drawn while that work runs, a click sends that kind of work to the
// server, and the place is left as the table says. Then the pattern every blocking call follows (a
// stopped call is no error and never reaches the store), and Run in Mind: a Cancel before the server
// has a job, the projects a stopped run left in Mind with "Delete it from Mind", and its five answers.
//
// The panels are drawn with the real <mr-cancel> inside them (it is made of Carbon's classes, so it
// needs no Carbon), a real click on it, and a fake ApiService that records what it was asked.

const v1: Version = { id: "ver-001", major: 1, minor: 0, label: "v1 — original", file_name: "model.xlsx", source: "upload" };
const v2: Version = { id: "ver-002", major: 2, minor: 0, label: "v2 — Prep", file_name: "model_v2.xlsx", source: "prep" };
const lastRecalc: RecalcResult = { status: "PASS", message: "No formula errors.", version_id: "ver-002", ran: true, formula_errors: [], addin_gap_errors: [] };

/** The workbook as the server holds it after a stop: what the Cancel's own answer carries. */
const RESTORED = { sessionId: "s1", versions: [v1, v2], version: v2, plan: [], recalc: lastRecalc } as unknown as SessionStart;

const STOPPED = "Stopped. Nothing was changed.";

/** A blocking call that was stopped, as HttpClient hands it over: HTTP 409 with `cancelled: true`. */
const stoppedError = (kind: string) => ({ status: 409, error: { detail: STOPPED, cancelled: true, kind } });

const answer = (kind: string | null, over: Partial<CancelResult> = {}): CancelResult => ({ cancelled: kind ? [kind] : [], late: [], stopping: [], ahead: false, message: STOPPED, session: RESTORED, mind: null, ...over });

const ops: Operation[] = [
  { op: "set_value", sheet: "Calc", cell: "A1", value: "#Title" },
  { op: "set_value", sheet: "Calc", cell: "A9", value: "#Other" },
];

function deferred<T>(): { promise: Promise<T>; resolve: (v: T) => void; reject: (e: unknown) => void } {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

/** Lets every promise that is ready run (no timer: it works under fake timers too). */
async function settle(): Promise<void> {
  for (let i = 0; i < 60; i++) await Promise.resolve();
}

const text = (el: Element | null | undefined) => (el?.textContent ?? "").replace(/\s+/g, " ").trim();
const cancelButton = (within: Element | null | undefined) => (within?.querySelector("mr-cancel button") ?? null) as HTMLButtonElement | null;
const buttonSaying = (within: Element, words: string) => Array.from(within.querySelectorAll("button")).find((b) => text(b).startsWith(words)) as HTMLButtonElement | undefined;

function click(el: Element | null | undefined): void {
  expect(el).toBeTruthy();
  (el as HTMLElement).click();
}

describe("Cancel in the Advanced panels", () => {
  // every class from one module registry: the panels, the button inside them and the tokens they inject
  let Api: typeof ApiService;
  let Store: typeof SessionStore;
  let Stack: typeof ChangeStack;
  let Cancel: typeof CancelButton;
  let Recalc: typeof RecalcPanel;
  let Bar: typeof StackBar;
  let Prep: typeof PrepPanel;
  let Loop: typeof LoopPanel;
  let Instances: typeof InstancesPanel;
  let Namer: typeof GridNamerPanel;
  let Assistant: typeof AssistantPanel;
  let Deliver: typeof DeliverPanel;
  let Advanced: typeof AdvancedPage;
  let MindRun: typeof MindRunPanel;

  /** the kinds the fake server was asked to stop, in order */
  let cancels: (string | null)[];

  beforeAll(async () => {
    vi.resetModules();
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Api = (await import("./api.service")).ApiService;
    Store = (await import("./session.store")).SessionStore;
    Stack = (await import("./change-stack")).ChangeStack;
    Cancel = (await import("./cancel-button")).CancelButton;
    Recalc = (await import("../pages/advanced/recalc-panel")).RecalcPanel;
    Bar = (await import("../pages/advanced/stack-bar")).StackBar;
    Prep = (await import("../pages/advanced/prep-panel")).PrepPanel;
    Loop = (await import("../pages/advanced/loop-panel")).LoopPanel;
    Instances = (await import("../pages/advanced/instances-panel")).InstancesPanel;
    Namer = (await import("../pages/advanced/grid-namer-panel")).GridNamerPanel;
    Assistant = (await import("../pages/advanced/assistant-panel")).AssistantPanel;
    Deliver = (await import("../pages/advanced/deliver-panel")).DeliverPanel;
    Advanced = (await import("../pages/advanced.page")).AdvancedPage;
    MindRun = (await import("../pages/advanced/mind-run-panel")).MindRunPanel;
  });

  // the files that come after load their own Carbon stand-ins
  afterAll(() => {
    vi.doUnmock("carbon-components-angular");
    vi.resetModules();
  });

  interface Drawn<T> {
    page: HTMLElement;
    panel: T;
    store: SessionStore;
    /** everything that was merged into the store through `applyAnalysis`, in order */
    applied: unknown[];
    detect: () => Promise<void>;
  }

  /** Draws one panel on the workbook "s1" (two versions, v2 current) with a fake server. `calls` are
   *  the server's answers for this test; `setup` runs on the store before the panel is made. */
  async function draw<T>(cls: Type<T>, calls: Record<string, unknown> = {}, setup?: (store: SessionStore) => void): Promise<Drawn<T>> {
    cancels = [];
    const api = {
      plainCodes: () => Promise.resolve({ rules: {}, errors: {} }),
      scanStatus: () => Promise.resolve({ state: "ready" }),
      resume: () => Promise.resolve(RESTORED),
      jobs: () => Promise.resolve([]),
      assistantKey: () => Promise.resolve({ saved: false }),
      aiUsage: () => Promise.resolve({ aiUsage: null, line: "" }),
      mindFlags: () => Promise.resolve([]),
      downloadUrl: (id: string, name: string) => `/api/sessions/${id}/files/${name}`,
      cancel: (_id: string, kind: string | null) => {
        cancels.push(kind);
        return Promise.resolve(answer(kind));
      },
      ...calls,
    };
    TestBed.resetTestingModule();
    TestBed.configureTestingModule({ providers: [{ provide: Api, useValue: api }] });
    TestBed.overrideComponent(cls, { set: { imports: [CommonModule, Cancel], schemas: [NO_ERRORS_SCHEMA] } });
    const store = TestBed.inject(Store);
    store.sessionId.set("s1");
    store.versions.set([v1, v2]);
    store.currentVersion.set(v2);
    setup?.(store);
    const applied: unknown[] = [];
    const keep = store.applyAnalysis.bind(store);
    store.applyAnalysis = (a) => {
      applied.push(a);
      keep(a);
    };
    const fixture = TestBed.createComponent(cls);
    const detect = async () => {
      for (let i = 0; i < 3; i++) {
        fixture.detectChanges();
        await settle();
      }
      fixture.detectChanges();
    };
    await detect();
    return { page: fixture.nativeElement as HTMLElement, panel: fixture.componentInstance, store, applied, detect };
  }

  // ---- rows 9 to 17: one blocking call each -------------------------------------------------

  it("row 9, Recalculate: Cancel beside the loading line stops the recalculation, and the last good one stays", async () => {
    const call = deferred<RecalcResult>();
    const { page, store, applied, detect } = await draw(Recalc, { recalculate: () => call.promise }, (s) => s.recalc.set(lastRecalc));
    expect(page.querySelector("mr-cancel")).toBeNull(); // nothing runs: no Cancel
    click(page.querySelector('[data-show-me="recalc-run"]'));
    await detect();
    const button = cancelButton(page)!;
    expect(text(button)).toBe("Cancel");
    expect(button.closest(".row")?.querySelector("cds-inline-loading")).not.toBeNull(); // beside the loading line
    button.click();
    await detect();
    expect(cancels).toEqual(["recalc"]);
    call.reject(stoppedError("recalc"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="stopped"]'))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull(); // stopped is no error
    expect(store.recalc()).toBe(lastRecalc);
    expect(applied).toEqual([RESTORED]); // only the Cancel's own answer reached the store
    const run = page.querySelector('[data-show-me="recalc-run"]') as HTMLButtonElement;
    expect(run.disabled).toBe(false);
    expect(text(run)).toBe("Recalculate again");
  });

  it("row 10, Apply all: Cancel beside the bar stops the Apply, and the entries stay in the list", async () => {
    const call = deferred<never>();
    const { page, applied, detect } = await draw(Bar, { apply: () => call.promise }, () => TestBed.inject(Stack).add("prep", "Prep: title the grids", ops));
    const stack = TestBed.inject(Stack);
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(page.querySelector('[data-show-me="stack-apply"]'));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest(".row")?.querySelector("cds-progress-bar")).not.toBeNull(); // beside the bar
    button.click();
    await detect();
    expect(cancels).toEqual(["apply"]);
    call.reject(stoppedError("apply"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector(".result strong"))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull();
    expect(stack.entries().length).toBe(1); // nothing was written: nothing leaves the list
    const again = page.querySelector('[data-show-me="stack-apply"]') as HTMLButtonElement;
    expect(text(again)).toBe("Apply all 2 changes");
    expect(again.disabled).toBe(false);
    expect(applied).toEqual([RESTORED]);
  });

  const action: PrepAction = { id: "title_grids", title: "Title the grids", rule_ids: ["STR-001"], default_on: true, level: "blocking", operations: ops, count: 2 };

  it("row 11, Prep actions Apply: Cancel beside the bar stops the Apply, and no version is made", async () => {
    const call = deferred<never>();
    const { page, store, applied, detect } = await draw(Prep, { apply: () => call.promise }, (s) => s.plan.set([action]));
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(page.querySelector('[data-show-me="prep-apply"]'));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest(".row")?.querySelector("cds-progress-bar")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["apply"]);
    call.reject(stoppedError("apply"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector(".bar .result strong"))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull();
    expect(store.versions()).toEqual([v1, v2]);
    expect(store.currentVersion()).toBe(v2);
    expect(applied).toEqual([RESTORED]);
  });

  const loopPlan: LoopPlan = {
    found: true,
    loops: [{ name: "CohortRng", driver: "Ctrl_01!C22", status: "ready", cells: [{}, {}], first: 2021, last: 2026 }],
    outputs: [{ target: "Out!A1:J6", status: "ready", cells: 60 }],
    clears: [],
    operation_count: 7,
  };

  it("row 12, Mind loop Build it: Cancel beside the bar stops the build, and the plan stays ready to build", async () => {
    const call = deferred<never>();
    const { page, applied, detect } = await draw(Loop, { mindLoopPlan: () => Promise.resolve({ version_id: "ver-002", plan: loopPlan }), applyMindLoopPlan: () => call.promise });
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(page.querySelector('[data-show-me="loop-build"]'));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest(".row")?.querySelector("cds-progress-bar")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["apply"]);
    call.reject(stoppedError("apply"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    const result = page.querySelector('[data-show-me="loop-result"]')!;
    expect(text(result.querySelector("strong"))).toBe(STOPPED);
    expect(result.classList.contains("result--error")).toBe(false);
    expect(page.querySelector('[data-test="loop-next"]')).toBeNull(); // no "Done. Next: Run in Mind"
    expect((page.querySelector('[data-show-me="loop-build"]') as HTMLButtonElement).disabled).toBe(false);
    expect(applied).toEqual([RESTORED]);
  });

  const instancePlan: InstancePlan = {
    found: true,
    status: "ready",
    mode: "table",
    operation_count: 4,
    counts: { key_values: 3, result_cells: 30 },
    key: { header: "Product", distinct: 3, column: "C", chosen_by: "macro", candidates: [{ column: "C", header: "Product", distinct: 3 }] },
  } as unknown as InstancePlan;

  it("row 13, Instances Build it: Cancel beside the bar stops the build: no version, no zip", async () => {
    const call = deferred<never>();
    const { page, applied, detect } = await draw(Instances, { mindInstancesPlan: () => Promise.resolve({ version_id: "ver-002", plan: instancePlan }), applyMindInstancesPlan: () => call.promise });
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(page.querySelector('[data-show-me="instances-build"]'));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest(".row")?.querySelector("cds-progress-bar")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["apply"]);
    call.reject(stoppedError("apply"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    const result = page.querySelector('[data-show-me="instances-result"]')!;
    expect(text(result.querySelector("strong"))).toBe(STOPPED);
    expect(result.classList.contains("result--error")).toBe(false);
    expect(page.querySelector('[data-test="inst-download"]')).toBeNull();
    expect(page.querySelector('[data-test="inst-next"]')).toBeNull();
    expect((page.querySelector('[data-show-me="instances-build"]') as HTMLButtonElement).disabled).toBe(false);
    expect(applied).toEqual([RESTORED]);
  });

  const finder = {
    schema: "finder",
    version: "1",
    file: "model.xlsm",
    counts: { sheets: 3, formulas: 1200, dimensions: 0, chains: 0, time_axes: 0, seconds: 2.5 },
    dimensions: [],
    chains: [],
    macro_loops: [],
    skipped: [],
    assistant: { asked: false, available: false, text: null, ranking: null, explanations: {}, edits: {}, extra: [], message: null },
  } as unknown as FinderResult;

  it("row 13, Instances Look and Look again: Cancel beside the loading line stops the look; what was found before stays", async () => {
    let call = deferred<{ version_id: string; finder: FinderResult }>();
    const { page, panel, detect } = await draw(Instances, { mindInstancesPlan: () => Promise.resolve({ version_id: "ver-002", plan: instancePlan }), instanceFinder: () => call.promise });
    panel.toggleFinder(); // "Not sure? Let the app look": the first look starts by itself
    await detect();
    let button = cancelButton(page.querySelector('[data-test="inst-finder"]'))!;
    expect(button.closest(".row")?.querySelector("cds-inline-loading")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["finder"]);
    call.reject(stoppedError("finder"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="inst-find-stopped"]'))).toBe(STOPPED);
    expect(text(page.querySelector('[data-test="inst-finder"]'))).not.toContain("Could not read the model");
    // the Look button is back; a look that ends shows its result
    call = deferred();
    click(buttonSaying(page, "Look"));
    await detect();
    call.resolve({ version_id: "ver-002", finder });
    await detect();
    expect(page.querySelector('[data-test="inst-find-stopped"]')).toBeNull();
    expect(panel.finder()).toBe(finder);
    // Look again, stopped: the earlier result stays on screen
    call = deferred();
    click(buttonSaying(page, "Look again"));
    await detect();
    button = cancelButton(page.querySelector('[data-test="inst-finder"]'))!;
    button.click();
    await detect();
    expect(cancels).toEqual(["finder", "finder"]);
    call.reject(stoppedError("finder"));
    await detect();
    expect(panel.finder()).toBe(finder);
    expect(text(page.querySelector('[data-test="inst-find-stopped"]'))).toBe(STOPPED);
  });

  const area = { sheet: "Calc", ref: "A1:B4", gridRef: "A1:B4", name: "Rates", flags: [], wasNamed: null };

  it("row 14, Grid Namer Submit: Cancel beside the loading line stops the write, and the areas to be named stay", async () => {
    const call = deferred<never>();
    const { page, panel, applied, detect } = await draw(Namer, { applyNamedAreas: () => call.promise });
    panel.pending.set([area]);
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(buttonSaying(page, "Submit 1 area"));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest("mr-cancel")?.previousElementSibling?.tagName).toBe("CDS-INLINE-LOADING"); // beside "Writing the titles through Excel…"
    button.click();
    await detect();
    expect(cancels).toEqual(["apply"]);
    call.reject(stoppedError("apply"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="namer-submit-stopped"]'))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull(); // no "Not written"
    expect(panel.pending()).toEqual([area]);
    expect(text(buttonSaying(page, "Submit"))).toBe("Submit 1 area");
    expect(applied).toEqual([RESTORED]);
  });

  it("row 14, Grid Namer Suggest names: Cancel beside the thinking dots stops the assistant, and the plan is as it was", async () => {
    const call = deferred<never>();
    // the workbook as the server holds it after the stop: the Prep plan it had before
    const cancel = (_id: string, kind: string | null) => {
      cancels.push(kind);
      return Promise.resolve(answer(kind, { session: { ...RESTORED, plan: [action] } }));
    };
    const { page, panel, store, detect } = await draw(Namer, { suggestGridNames: () => call.promise, cancel }, (s) => s.plan.set([action]));
    click(buttonSaying(page, "Suggest names"));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest("mr-cancel")?.previousElementSibling?.tagName).toBe("MR-THINKING"); // beside the thinking dots
    button.click();
    await detect();
    expect(cancels).toEqual(["names"]);
    call.reject(stoppedError("names"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="namer-suggest-stopped"]'))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull();
    expect(panel.suggestions()).toBeNull();
    expect(store.plan()).toEqual([action]);
    expect(text(buttonSaying(page, "Suggest names"))).toBe("Suggest names");
  });

  it("row 15, the assistant's answer: Cancel beside the thinking dots gives the request up; the question stays and a note says Stopped.", async () => {
    let given: AbortSignal | undefined;
    const chatStream = (_id: string, _history: unknown, _q: string, _focus: unknown, onEvent: (ev: { type: "delta"; text: string }) => void, signal?: AbortSignal) =>
      new Promise((resolve) => {
        given = signal;
        onEvent({ type: "delta", text: "Half an ans" });
        signal?.addEventListener("abort", () => resolve(null)); // an aborted fetch answers null (api.service)
      });
    const { page, panel, store, detect } = await draw(Assistant, { chatStream });
    void panel.send("Why does Mind reject this?");
    await detect();
    const draft = page.querySelector(".bubble--draft")!;
    expect(text(draft)).toContain("Half an ans");
    expect(draft.querySelector("mr-thinking")).not.toBeNull();
    const button = cancelButton(draft)!; // beside the thinking dots
    expect(given?.aborted).toBe(false);
    button.click();
    await detect();
    expect(cancels).toEqual(["chat"]); // the server is told too
    expect(given?.aborted).toBe(true); // the fetch was aborted at the click
    expect(store.chatHistory()).toEqual([
      { role: "user", content: "Why does Mind reject this?" },
      { role: "assistant", content: "Stopped.", note: true },
    ]);
    expect(page.querySelector(".bubble--draft")).toBeNull();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page)).not.toContain("The assistant failed");
    expect(text(page)).not.toContain("Half an ans"); // the half answer is not kept
    expect(panel.streaming()).toBe(false);
  });

  it("row 15, the assistant's answer: the server's own last line 'cancelled' is the same note, not an error", async () => {
    const { page, panel, store, detect } = await draw(Assistant, { chatStream: () => Promise.resolve(null) });
    await panel.send("What blocks this workbook?");
    await detect();
    expect(store.chatHistory().at(-1)).toEqual({ role: "assistant", content: "Stopped.", note: true });
    expect(text(page)).not.toContain("The assistant failed");
  });

  it("row 15, Apply the proposal: Cancel beside the bar stops the Apply, and the proposal's card stays", async () => {
    const call = deferred<never>();
    const proposal = { summary: "Title the two grids", operations: ops, errors: [] };
    const { page, panel, store, applied, detect } = await draw(Assistant, { apply: () => call.promise });
    panel.proposal.set(proposal);
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(buttonSaying(page, "Apply 2 changes"));
    await detect();
    const button = cancelButton(page.querySelector(".card"))!;
    expect(button.closest(".row")?.querySelector("cds-progress-bar")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["apply"]);
    call.reject(stoppedError("apply"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="assistant-apply-stopped"]'))).toBe(STOPPED);
    expect(panel.proposal()).toBe(proposal); // the card stays: it can be applied again
    expect(text(page.querySelector(".card .card__title"))).toBe("Proposed: Title the two grids");
    expect((buttonSaying(page, "Apply 2 changes") as HTMLButtonElement).disabled).toBe(false);
    expect(text(page)).not.toContain("The assistant failed");
    expect(store.chatHistory()).toEqual([]); // no "changes applied" note
    expect(applied).toEqual([RESTORED]);
  });

  it("row 16, Deliver Generate reports: Cancel beside the loading line stops it, and no report is kept", async () => {
    const call = deferred<never>();
    const { page, store, detect } = await draw(Deliver, { generateReports: () => call.promise });
    expect(page.querySelector("mr-cancel")).toBeNull();
    click(buttonSaying(page, "Generate reports"));
    await detect();
    const button = cancelButton(page)!;
    expect(text(button.closest(".row")?.querySelector("button"))).toBe("Generate reports"); // in the reports' own row
    expect(button.closest(".row")?.querySelector("cds-inline-loading")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["reports"]);
    call.reject(stoppedError("reports"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="deliver-build-stopped"]'))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull(); // no "Failed"
    expect(store.reports()).toBeNull();
    expect((buttonSaying(page, "Generate reports") as HTMLButtonElement).disabled).toBe(false);
  });

  it("row 16, Deliver Check the numbers: Cancel beside the loading line stops the comparison", async () => {
    const call = deferred<never>();
    const { page, store, detect } = await draw(Deliver, { numbersCheck: () => call.promise });
    click(page.querySelector('[data-show-me="deliver-numbers"]'));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest(".row")?.querySelector('[data-show-me="deliver-numbers"]')).not.toBeNull(); // in the numbers' own row
    expect(button.closest(".row")?.querySelector("cds-inline-loading")).not.toBeNull();
    button.click();
    await detect();
    expect(cancels).toEqual(["numbers"]);
    call.reject(stoppedError("numbers"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="deliver-check-stopped"]'))).toBe(STOPPED);
    expect(page.querySelector(".result--error")).toBeNull();
    expect(store.numbers()).toBeNull();
    expect(text(page.querySelector('[data-show-me="deliver-numbers"]'))).toBe("Check the numbers");
  });

  it("row 17, Versions Make current: one at a time, Cancel beside 'Restoring…', and the version list as it was", async () => {
    const call = deferred<never>();
    let asked = 0;
    const reanalyze = () => {
      asked += 1;
      return call.promise;
    };
    const { page, panel, store, applied, detect } = await draw(Advanced, { reanalyze }, (s) => {
      s.report.set({ findings: [] } as unknown as NonNullable<ReturnType<SessionStore["report"]>>);
      s.advancedTab.set("versions");
    });
    expect(page.querySelector("mr-cancel")).toBeNull();
    const make = buttonSaying(page, "Make current")!; // v1's row: v2 is the current one
    make.click();
    await detect();
    const line = page.querySelector('[data-test="restore-line"]')!;
    expect(text(line)).toContain("Restoring ver-001…");
    const button = cancelButton(line)!; // beside "Restoring…"
    // the busy flag: the button is off while it works, and a second call does nothing
    expect((buttonSaying(page, "Make current") as HTMLButtonElement).disabled).toBe(true);
    void panel.restore("ver-001");
    await detect();
    expect(asked).toBe(1);
    button.click();
    await detect();
    expect(cancels).toEqual(["scan"]);
    call.reject(stoppedError("scan"));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="restore-line"]'))).toBe(STOPPED);
    expect(store.versions()).toEqual([v1, v2]);
    expect(store.currentVersion()).toBe(v2);
    expect(page.querySelector("tr.current")?.textContent).toContain("v2 — Prep");
    expect((buttonSaying(page, "Make current") as HTMLButtonElement).disabled).toBe(false);
    expect(applied).toEqual([RESTORED]);
  });

  // ---- the pattern at every blocking call ----------------------------------------------------

  it("a stopped call that answers before the Cancel does: the line is written, and the workbook still comes from the Cancel's answer", async () => {
    const call = deferred<never>();
    const stop = deferred<CancelResult>();
    const cancel = (_id: string, kind: string | null) => {
      cancels.push(kind);
      return stop.promise;
    };
    const { page, store, applied, detect } = await draw(Prep, { apply: () => call.promise, cancel }, (s) => s.plan.set([action]));
    click(page.querySelector('[data-show-me="prep-apply"]'));
    await detect();
    const button = cancelButton(page)!;
    button.click();
    await detect();
    expect(text(button)).toBe("Stopping…"); // while the server works on the stop
    expect(button.disabled).toBe(true);
    call.reject(stoppedError("apply")); // the Apply's own request answers 409 first
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull(); // the bar and its Cancel are gone already
    expect(text(page.querySelector(".bar .result strong"))).toBe(STOPPED);
    expect(applied).toEqual([]);
    stop.resolve(answer("apply"));
    await detect();
    expect(applied).toEqual([RESTORED]);
    expect(store.currentVersion()).toBe(v2);
  });

  it("a call that answers normally after a Cancel (too late to stop) is shown as if no Cancel was pressed", async () => {
    const call = deferred<unknown>();
    const cancel = (_id: string, kind: string | null) => {
      cancels.push(kind);
      return Promise.resolve(answer(null, { late: ["apply"], message: "It had already finished.", session: null }));
    };
    const { page, store, detect } = await draw(Prep, { apply: () => call.promise, cancel }, (s) => s.plan.set([action]));
    click(page.querySelector('[data-show-me="prep-apply"]'));
    await detect();
    cancelButton(page)!.click();
    await detect();
    const v3: Version = { id: "ver-003", major: 3, minor: 0, label: "v3 — Prep", file_name: "model_v3.xlsx", source: "prep" };
    call.resolve({ result: { status: "ok", applied: ops, failed: [] }, version: v3, versions: [v1, v2, v3], plan: [] });
    await detect();
    expect(text(page.querySelector(".bar .result strong"))).toBe("2 changes applied");
    expect(store.currentVersion()).toBe(v3);
    expect(text(page)).not.toContain("Stopped.");
  });

  it("an error that is not a Cancel is still an error", async () => {
    const { page, detect } = await draw(Recalc, { recalculate: () => Promise.reject({ status: 500, error: { detail: "Excel did not start." } }) });
    click(page.querySelector('[data-show-me="recalc-run"]'));
    await detect();
    const failed = page.querySelector(".result--error");
    expect(text(failed?.querySelector("strong"))).toBe("Could not recalculate");
    expect(text(failed?.querySelector("span"))).toBe("Excel did not start.");
    expect(page.querySelector('[data-test="stopped"]')).toBeNull();
  });

  // ---- row 8: Run in Mind ---------------------------------------------------------------------

  const timing = {
    elapsed_s: 30,
    remaining_s: 600,
    estimate_s: 630,
    retry_s: null,
    over_estimate: false,
    stage: "upload",
    stage_title: "Uploading to Mind",
    stage_elapsed_s: 5,
    iteration: 1,
    max_iterations: 4,
    stages: [{ id: "upload", title: "Uploading to Mind", seconds: 60, state: "now" }],
    last_news_s: 1,
  } as MindRunTiming;

  const run = (over: Partial<MindLoopStatus> = {}): MindLoopStatus => ({ state: "idle", events: [], next: 0, report: null, project: null, kept: [], earlier: [], message: null, ...over });
  const project = (name: string, over: Partial<MindProject> = {}): MindProject => ({ name, status: "model added", ran: false, deleted: null, url: "https://mind.example/app", ...over });
  const signedOut: MindSignIn = { state: "idle", message: null, last_login: null, checked: null, elapsed_s: null };
  const signedIn: MindSignIn = { state: "done", message: null, last_login: "2026-11-16T10:00:00", checked: { signed_in: true, at: "2026-11-17T00:10:00" }, elapsed_s: 12 };

  /** The Run in Mind tab on a fake Mind: `mind.now` is what GET .../mind-loop answers. */
  async function drawMind(now: MindLoopStatus, calls: Record<string, unknown> = {}) {
    const mind = { now, looks: 0, starts: 0, signInStops: 0 };
    const drawn = await draw(MindRun, {
      mindLoopStatus: () => {
        mind.looks += 1;
        return Promise.resolve(mind.now);
      },
      mindSignIn: () => Promise.resolve(signedOut),
      startMindSignIn: (stop = false) => {
        if (stop) mind.signInStops += 1;
        return Promise.resolve(signedOut);
      },
      mindLoopOptions: () => Promise.reject(new Error("no options in this test")), // the tab works without them
      startMindLoop: () => {
        mind.starts += 1;
        mind.now = run({ state: "running", timing });
        return Promise.resolve({ state: "running" });
      },
      ...calls,
    });
    return { ...drawn, mind };
  }

  it("row 8, Run in Mind: Cancel beside the progress sends 'mind', says Stopping…, then the server's sentence with the project", async () => {
    const stop = deferred<CancelResult>();
    const p = project("zzz_mindready_model_1");
    const { page, mind, detect } = await drawMind(run({ state: "running", timing }), {
      cancel: (_id: string, kind: string | null) => {
        cancels.push(kind);
        return stop.promise;
      },
    });
    const progress = () => page.querySelector('[data-test="mind-progress"]');
    expect(text(progress())).toContain("Uploading to Mind…");
    const button = cancelButton(progress())!; // beside the progress
    expect(text(button)).toBe("Cancel");
    expect(page.querySelectorAll("mr-cancel").length).toBe(1);
    button.click();
    await detect();
    expect(cancels).toEqual(["mind"]);
    expect(text(progress()?.querySelector("strong"))).toBe("Stopping…");
    expect(text(progress())).not.toContain("Uploading to Mind…");
    expect(page.querySelector('[data-test="mind-stopped"]')).toBeNull(); // not before the server has finished
    // the server has finished: its answer comes, and the tab looks at the run at once
    mind.now = run({ state: "cancelled", timing, project: p, kept: [p], message: "Stopped. The project zzz_mindready_model_1 is still in Mind." });
    stop.resolve(answer("mind", { message: mind.now.message!, mind: p }));
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(progress()).toBeNull(); // no "Finished in …" for a run that was stopped
    expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. The project zzz_mindready_model_1 is still in Mind.");
    const rows = page.querySelectorAll('[data-test="mind-kept-project"]');
    expect(rows.length).toBe(1);
    expect(text(rows[0].querySelector(".mono"))).toBe("zzz_mindready_model_1");
    expect(text(rows[0].querySelector("button"))).toBe("Delete it from Mind");
    expect(page.querySelector(".result--error")).toBeNull(); // stopped is no error
    expect(text(page.querySelector('[data-show-me="mind-run"]'))).toBe("Run in Mind"); // and it can be started again
  });

  it("row 8, Run in Mind: stopped before a project was made, the sentence says nothing was uploaded", async () => {
    const { page, mind, detect } = await drawMind(run({ state: "running", timing }));
    mind.now = run({ state: "cancelled", message: "Stopped. Nothing was uploaded to Mind." });
    cancelButton(page)!.click();
    await detect();
    expect(cancels).toEqual(["mind"]);
    expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. Nothing was uploaded to Mind.");
    expect(page.querySelector('[data-test="mind-kept"]')).toBeNull(); // nothing to delete
  });

  it("row 8, Run in Mind: the looks go on through 'cancelling' and end when the server says cancelled", async () => {
    vi.useFakeTimers();
    try {
      const { page, mind, detect } = await drawMind(run({ state: "running", timing }));
      expect(mind.looks).toBe(1);
      mind.now = run({ state: "cancelling", timing, message: "Stopping…" });
      await vi.advanceTimersByTimeAsync(2000);
      await detect();
      expect(mind.looks).toBe(2);
      expect(text(page.querySelector('[data-test="mind-progress"] strong'))).toBe("Stopping…");
      // a tab opened while the run stops: the button is there, greyed, and says so too
      expect(text(cancelButton(page))).toBe("Stopping…");
      expect(cancelButton(page)!.disabled).toBe(true);
      await vi.advanceTimersByTimeAsync(2000);
      expect(mind.looks).toBe(3); // still on its way out: looked at again
      mind.now = run({ state: "cancelled", message: "Stopped. Nothing was uploaded to Mind." });
      await vi.advanceTimersByTimeAsync(2000);
      await detect();
      expect(mind.looks).toBe(4);
      expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. Nothing was uploaded to Mind.");
      await vi.advanceTimersByTimeAsync(10000);
      expect(mind.looks).toBe(4); // the looks have ended
    } finally {
      vi.useRealTimers();
    }
  });

  it("row 8, Run in Mind: a Cancel that came too late leaves the finished run as it is", async () => {
    const { page, mind, detect } = await drawMind(run({ state: "running", timing }), {
      cancel: (_id: string, kind: string | null) => {
        cancels.push(kind);
        return Promise.resolve(answer(null, { late: ["mind"], message: "It had already finished.", session: null }));
      },
    });
    mind.now = run({ state: "done", timing, summary: "converged after 1 round" });
    cancelButton(page)!.click();
    await detect();
    expect(text(page.querySelector('[data-test="mind-progress"]'))).toContain("Finished in 0:30");
    expect(page.querySelector('[data-test="mind-stopped"]')).toBeNull();
    expect(page.querySelector("mr-cancel")).toBeNull();
  });

  it("row 8, before the server has a job: Cancel during the tab's sign-in check means the run is never started", async () => {
    const check = deferred<MindSignIn>();
    const { page, mind, detect } = await drawMind(run(), { mindSignIn: (c: boolean) => (c ? check.promise : Promise.resolve(signedOut)) });
    click(buttonSaying(page, "Check the sign-in"));
    await detect();
    expect(text(page)).toContain("Checking the Mind sign-in…");
    expect(page.querySelectorAll("mr-cancel").length).toBe(1); // beside the check
    // Run in Mind pressed meanwhile: it waits for the check, with its own Cancel
    click(page.querySelector('[data-show-me="mind-run"]'));
    await detect();
    expect(mind.starts).toBe(0);
    const starting = () => page.querySelector('[data-test="mind-starting"]');
    expect(text(starting()?.querySelector("strong"))).toBe("Starting…");
    expect(page.querySelectorAll("mr-cancel").length).toBe(1); // one Cancel on screen, not two
    cancelButton(starting())!.click();
    await detect();
    expect(cancels).toEqual(["mind"]);
    expect(text(starting()?.querySelector("strong"))).toBe("Stopping…");
    expect(cancelButton(starting())!.disabled).toBe(true); // greyed until the check has answered
    // the check answers "signed in": nothing more starts
    check.resolve(signedIn);
    await detect();
    expect(mind.starts).toBe(0);
    expect(starting()).toBeNull();
    expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. Nothing was uploaded to Mind.");
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page)).toContain("Signed in to Mind"); // what the check found is still shown
    // a new click on Run in Mind is a new start
    click(page.querySelector('[data-show-me="mind-run"]'));
    await detect();
    expect(mind.starts).toBe(1);
    expect(page.querySelector('[data-test="mind-stopped"]')).toBeNull();
  });

  it("row 8, before the server has a job: Cancel on the sign-in check alone", async () => {
    const check = deferred<MindSignIn>();
    const { page, mind, detect } = await drawMind(run(), { mindSignIn: (c: boolean) => (c ? check.promise : Promise.resolve(signedOut)) });
    click(buttonSaying(page, "Check the sign-in"));
    await detect();
    const button = cancelButton(page)!;
    expect(button.closest(".row")?.querySelector("button")?.textContent).toContain("Sign in to Mind"); // in the sign-in box
    button.click();
    await detect();
    expect(cancels).toEqual(["mind"]);
    expect(text(page)).not.toContain("Checking the Mind sign-in…");
    expect(text(cancelButton(page))).toBe("Stopping…"); // the check cannot be stopped on the server: greyed until it answers
    expect(cancelButton(page)!.disabled).toBe(true);
    expect(page.querySelector('[data-test="mind-stopped"]')).toBeNull();
    check.resolve(signedIn);
    await detect();
    expect(page.querySelector("mr-cancel")).toBeNull();
    expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. Nothing was uploaded to Mind.");
    expect(mind.starts).toBe(0);
    expect((buttonSaying(page, "Check the sign-in") as HTMLButtonElement).disabled).toBe(false);
  });

  it("row 8, before the server has a job: Cancel while the options are on their way means the run is never started", async () => {
    const options = deferred<never>();
    const { page, mind, detect } = await drawMind(run(), { mindLoopOptions: () => options.promise });
    click(page.querySelector('[data-show-me="mind-run"]'));
    await detect();
    expect(mind.starts).toBe(0); // waits for what the run will send
    cancelButton(page.querySelector('[data-test="mind-starting"]'))!.click();
    await detect();
    options.reject(new Error("gone"));
    await detect();
    expect(mind.starts).toBe(0);
    expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. Nothing was uploaded to Mind.");
  });

  it("row 8, a Cancel that lands while the start itself is on its way is sent again once the run exists", async () => {
    const start = deferred<{ state: string }>();
    const { page, mind, detect } = await drawMind(run(), { startMindLoop: () => start.promise });
    click(page.querySelector('[data-show-me="mind-run"]'));
    await detect();
    cancelButton(page.querySelector('[data-test="mind-starting"]'))!.click(); // the server has no run yet: nothing to stop
    await detect();
    expect(cancels).toEqual(["mind"]);
    mind.now = run({ state: "cancelling", message: "Stopping…" });
    start.resolve({ state: "running" });
    await detect();
    expect(cancels).toEqual(["mind", "mind"]); // now there is one
    expect(text(page.querySelector('[data-test="mind-progress"] strong'))).toBe("Stopping…");
  });

  it("row 8, 'Close the sign-in window' stays as it is", async () => {
    const waiting: MindSignIn = { state: "waiting", message: "An Edge window is open on Mind: sign in there.", last_login: null, checked: null, elapsed_s: 3 };
    const { page, mind, detect } = await drawMind(run(), { mindSignIn: () => Promise.resolve(waiting) });
    const close = buttonSaying(page, "Close the sign-in window")!;
    expect(page.querySelector("mr-cancel")).toBeNull();
    close.click();
    await detect();
    expect(mind.signInStops).toBe(1);
    expect(cancels).toEqual([]);
  });

  it("row 8, after a stopped run: every project it left in Mind has its button, and the earlier rounds are named", async () => {
    const a = project("zzz_mindready_model_2");
    const b = project("zzz_mindready_model_0", { status: "run", ran: true });
    const { page } = await drawMind(run({ state: "cancelled", project: a, kept: [b, a], earlier: ["zzz_mindready_model_1"], message: "Stopped. The project zzz_mindready_model_2 is still in Mind." }));
    expect(text(page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. The project zzz_mindready_model_2 is still in Mind.");
    const rows = Array.from(page.querySelectorAll('[data-test="mind-kept-project"]'));
    expect(rows.map((r) => text(r.querySelector(".mono")))).toEqual(["zzz_mindready_model_0", "zzz_mindready_model_2"]);
    expect(rows.map((r) => text(r.querySelector("button")))).toEqual(["Delete it from Mind", "Delete it from Mind"]);
    const earlier = page.querySelector('[data-test="mind-earlier"]')!;
    expect(text(earlier)).toBe("Earlier rounds of this run are still in Mind: zzz_mindready_model_1.");
    expect(earlier.querySelector("button")).toBeNull(); // they were run: no button
    expect(page.querySelector('[data-test="mind-progress"]')).toBeNull();
  });

  it("row 8, when the tab opens later: a project of a stopped run still in Mind is listed again with its button", async () => {
    // after a restart the server knows no run (idle), only the project it wrote down
    const later = await drawMind(run({ state: "idle", kept: [project("zzz_mindready_model_1")] }));
    expect(later.page.querySelector('[data-test="mind-stopped"]')).toBeNull(); // no run was stopped on this visit
    expect(text(later.page.querySelector('[data-test="mind-kept-lead"]'))).toBe("Still in Mind from a run you stopped:");
    expect(text(later.page.querySelector('[data-test="mind-kept-project"] .mono'))).toBe("zzz_mindready_model_1");
    expect(text(later.page.querySelector('[data-test="mind-kept-project"] button'))).toBe("Delete it from Mind");
    // also under a later run that finished
    const after = await drawMind(run({ state: "done", timing, kept: [project("zzz_mindready_model_1")] }));
    expect(text(after.page.querySelector('[data-test="mind-kept-lead"]'))).toBe("Still in Mind from a run you stopped:");
    expect(text(after.page.querySelector('[data-test="mind-kept-project"] button'))).toBe("Delete it from Mind");
    // right after a Cancel the stopped run's own sentence stands above the list: no second line
    const stopped = await drawMind(run({ state: "cancelled", project: project("zzz_mindready_model_1"), kept: [project("zzz_mindready_model_1")], message: "Stopped. The project zzz_mindready_model_1 is still in Mind." }));
    expect(stopped.page.querySelector('[data-test="mind-kept-lead"]')).toBeNull();
    // a stopped run whose answer came without the sentence: the same words, built from its project
    const bare = await drawMind(run({ state: "cancelled", project: project("zzz_mindready_model_3", { status: "creating" }), kept: [project("zzz_mindready_model_3", { status: "creating" })] }));
    expect(text(bare.page.querySelector('[data-test="mind-stopped"]'))).toBe("Stopped. A project named zzz_mindready_model_3 may be in Mind.");
    // nothing left in Mind: nothing is said
    const none = await drawMind(run());
    expect(none.page.querySelector('[data-test="mind-stopped"]')).toBeNull();
    expect(none.page.querySelector('[data-test="mind-kept"]')).toBeNull();
  });

  const BY_HAND = "Mind did not delete it. Delete it by hand in Mind: open the Project manager, folder zzz_mindready.";
  const deletes: { result: MindDelete["result"]; message: string; stays: boolean; link: boolean }[] = [
    { result: "deleted", message: "Deleted from Mind.", stays: false, link: false },
    { result: "not found", message: "It was not in Mind: nothing to delete.", stays: false, link: false },
    { result: "not found", message: "It is not in Mind yet. Try again in a minute.", stays: true, link: false },
    { result: "still present", message: BY_HAND, stays: true, link: true },
    { result: "refused", message: BY_HAND, stays: true, link: true },
    { result: "unreachable", message: "Mind did not open. Sign in to Mind, then try again.", stays: true, link: false },
  ];

  for (const c of deletes) {
    it(`row 8, Delete it from Mind answers "${c.message}": the button ${c.stays ? "stays" : "goes"}${c.link ? ", with the link Open Mind" : ""}`, async () => {
      const p = project("zzz_mindready_model_1", { status: c.message.includes("not in Mind yet") ? "creating" : "model added" });
      const call = deferred<MindDelete>();
      const asked: string[] = [];
      const { page, mind, detect } = await drawMind(run({ state: "cancelled", project: p, kept: [p], message: "Stopped. The project zzz_mindready_model_1 is still in Mind." }), {
        deleteMindProject: (_id: string, name: string) => {
          asked.push(name);
          return call.promise;
        },
      });
      const row = () => page.querySelector('[data-test="mind-kept-project"]')!;
      click(row().querySelector("button"));
      await detect();
      expect(asked).toEqual(["zzz_mindready_model_1"]); // nothing is deleted without the click, and only this project
      expect(text(row().querySelector('[data-test="mind-delete-note"]'))).toBe("Deleting…");
      expect((row().querySelector("button") as HTMLButtonElement).disabled).toBe(true);
      expect(page.querySelector("mr-cancel")).toBeNull(); // about a minute, no Cancel
      // Mind answered; the server's own list says whether the project is still there
      mind.now = run({ ...mind.now, kept: c.stays ? [p] : [] });
      call.resolve({ name: p.name, result: c.result, message: c.message, url: "https://mind.example/app" });
      await detect();
      expect(text(row().querySelector(".mono"))).toBe("zzz_mindready_model_1");
      expect(text(row().querySelector('[data-test="mind-delete-note"]'))).toBe(c.message);
      const button = row().querySelector("button") as HTMLButtonElement | null;
      if (c.stays) {
        expect(text(button)).toBe("Delete it from Mind");
        expect(button!.disabled).toBe(false);
      } else {
        expect(button).toBeNull();
      }
      const link = row().querySelector("a") as HTMLAnchorElement | null;
      if (c.link) {
        expect(text(link)).toBe("Open Mind");
        expect(link!.getAttribute("href")).toBe("https://mind.example/app");
        expect(link!.getAttribute("target")).toBe("_blank"); // a new tab
      } else {
        expect(link).toBeNull();
      }
    });
  }

  it("row 8, Delete it from Mind: when the list cannot be read again, the answer itself says whether the button goes", async () => {
    const gone = project("zzz_gone");
    const there = project("zzz_there");
    const results: Record<string, MindDelete> = {
      zzz_gone: { name: "zzz_gone", result: "deleted", message: "Deleted from Mind.", url: "https://mind.example/app" },
      zzz_there: { name: "zzz_there", result: "unreachable", message: "Mind did not open. Sign in to Mind, then try again.", url: "https://mind.example/app" },
    };
    let down = false;
    const { page, detect } = await drawMind(run({ state: "cancelled", kept: [gone, there], message: "Stopped. The project zzz_there is still in Mind." }), {
      deleteMindProject: (_id: string, name: string) => {
        down = true;
        return Promise.resolve(results[name]);
      },
      mindLoopStatus: () => (down ? Promise.reject(new Error("no answer")) : Promise.resolve(run({ state: "cancelled", kept: [gone, there], message: "Stopped. The project zzz_there is still in Mind." }))),
    });
    const rows = () => Array.from(page.querySelectorAll('[data-test="mind-kept-project"]'));
    click(rows()[0].querySelector("button"));
    await detect();
    click(rows()[1].querySelector("button"));
    await detect();
    expect(rows().map((r) => text(r.querySelector('[data-test="mind-delete-note"]')))).toEqual(["Deleted from Mind.", "Mind did not open. Sign in to Mind, then try again."]);
    expect(rows().map((r) => !!r.querySelector("button"))).toEqual([false, true]);
  });

  it("row 8, Delete it from Mind refused while Mind is busy: the server's sentence, and the button stays", async () => {
    const p = project("zzz_mindready_model_1");
    const { page, detect } = await drawMind(run({ state: "cancelled", project: p, kept: [p], message: "Stopped. The project zzz_mindready_model_1 is still in Mind." }), {
      deleteMindProject: () => Promise.reject({ status: 409, error: { detail: "Mind is busy. Try again in a moment." } }),
    });
    click(page.querySelector('[data-test="mind-kept-project"] button'));
    await detect();
    expect(text(page.querySelector('[data-test="mind-delete-note"]'))).toBe("Mind is busy. Try again in a moment.");
    expect(text(page.querySelector('[data-test="mind-kept-project"] button'))).toBe("Delete it from Mind");
    expect(page.querySelector('[data-test="mind-open"]')).toBeNull();
  });
});
