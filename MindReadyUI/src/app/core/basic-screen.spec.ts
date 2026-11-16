import { CommonModule } from "@angular/common";
import { Component, Input, NO_ERRORS_SCHEMA } from "@angular/core";
import { ComponentFixture, TestBed } from "@angular/core/testing";
import { onlyRecalcLeft } from "./recalc-now";
import { ActivatedRoute, provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import type { RecentWorkbooks } from "../pages/recent-workbooks";
import { ApiService } from "./api.service";
import { finishState } from "./finish-banner";
import { ChainEstimate, Readiness, SavedWorkbook, SessionStart } from "./models";
import type { RunBar } from "./run-bar";
import { planShowMe, PAGE_LINKS, showMeFor } from "./show-me";
import { SessionStore } from "./session.store";
import { Thinking } from "./thinking";

// ---- 1.15.0: three screens (Basic - Custom - Advanced), the end-of-check banner, Remove all -------------

const saved = (id: string): SavedWorkbook => ({ id, name: `${id}.xlsm`, uploaded_at: 1_800_000_000, last_opened: 1_800_000_000, versions: 2, current_version: "ver-002", bytes: 10, has_result: true, pending: false });

const started = (id: string): SessionStart => ({
  sessionId: id,
  summary: { file_name: `${id}.xlsm`, file_type: "xlsm", sheet_count: 2, grid_count: 1, formula_count: 10 },
  report: { status: "PASS", findings: [] },
  plan: [],
  version: { id: "ver-003", major: 3, minor: 0, label: "v3", file_name: `${id}_v3.xlsm`, source: "prep" },
  versions: [],
  recalc: { status: "PASS", message: "clean", ran: true, formula_errors: [], addin_gap_errors: [], version_id: "ver-003" },
});

const readiness = (state: Readiness["state"], blocking = 0): Readiness => ({ state, headline: "h", blocking: [], blocking_count: blocking, manual_blocking_count: 0, optional_count: 0 });

const settle = () => new Promise<void>((r) => setTimeout(r, 0));

/** Something the test ends when it chooses: a step of Start held open, an estimate that has not come yet. */
function hold<T>(): { done: Promise<T>; end: (value: T) => void } {
  let end!: (value: T) => void;
  const done = new Promise<T>((r) => (end = r));
  return { done, end };
}

/** Do everything's estimate for a workbook with no outside files: four steps, 470 seconds in all. */
const estimate = (): ChainEstimate => ({
  steps: [
    { id: "fixAll", low_s: 50, mid_s: 100, high_s: 200, why: "" },
    { id: "prepare", low_s: 60, mid_s: 120, high_s: 240, why: "" },
    { id: "loop", low_s: 60, mid_s: 120, high_s: 240, why: "" },
    { id: "recalc", low_s: 65, mid_s: 130, high_s: 260, why: "" },
  ],
  total: { low_s: 235, mid_s: 470, high_s: 940 },
  caveat: "",
  basis: { file_mb: 1, scan_s: 1, scan_measured: true, stored_errors: 0, blocking_ops: 0 },
});

describe("the three screens: routes and Show me", () => {
  it("has Basic at the root, Custom at /custom and Advanced; each page is told its mode", async () => {
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    const { routes } = await import("../app.routes");
    expect(routes.map((r) => r.path)).toEqual(["", "custom", "advanced", "**"]);
    expect(routes[0].data).toEqual({ mode: "basic" });
    expect(routes[1].data).toEqual({ mode: "custom" });
    expect(routes[0].component).toBe(routes[1].component);
  });

  it("sends the places that were on the main screen to Custom, from Basic and from Advanced", () => {
    expect(planShowMe({ id: "outside-card" }, "advanced")[0]).toMatchObject({ kind: "page", page: "custom" });
    expect(planShowMe({ id: "outside-card" }, "basic")[0]).toMatchObject({ kind: "page", page: "custom" });
    expect(planShowMe({ id: "recalc-run" }, "custom")).toEqual([{ kind: "place", targets: [{ el: "main-recalc", text: "Click Recalculate" }] }]);
    expect(PAGE_LINKS.custom.map((t) => t.el)).toEqual(["nav-custom", "link-custom"]);
    expect(PAGE_LINKS.basic.map((t) => t.el)).toEqual(["nav-basic", "link-basic"]);
  });

  it("points at Go on Basic and at Do everything on Custom; Download and the recent list exist on both", () => {
    expect(planShowMe({ id: "do-everything" }, "basic")).toEqual([{ kind: "place", targets: [{ el: "do-everything", text: "Click Go" }] }]);
    expect(planShowMe({ id: "do-everything" }, "custom")[0].targets[0].text).toBe("Click Do everything");
    expect(planShowMe({ id: "download" }, "basic").map((s) => s.kind)).toEqual(["place"]);
    expect(planShowMe({ id: "recent-workbooks" }, "basic").map((s) => s.kind)).toEqual(["place"]);
    expect(planShowMe({ id: "recent-workbooks" }, "custom").map((s) => s.kind)).toEqual(["place"]);
    // the server's sentence about a closed session still leads to the list
    expect(showMeFor("Open it again from Recent workbooks on the Basic screen.")?.id).toBe("recent-workbooks");
  });
});

describe("the end-of-check banner", () => {
  it("says the workbook can be uploaded when it is acceptable", () => {
    expect(finishState(readiness("ready"))).toEqual({ acceptable: true, text: "The check is finished. The workbook can be uploaded to Mind.", reason: "" });
  });

  it("says in one line how many blocking problems are left, and greys Upload with the reason", () => {
    const s = finishState(readiness("blocked", 3));
    expect(s?.acceptable).toBe(false);
    expect(s?.text).toBe("3 blocking problems are left.");
    expect(s?.reason).toMatch(/blocking problems are left/);
    expect(finishState(readiness("blocked", 1))?.text).toBe("1 blocking problem is left.");
  });

  it("says the check is not finished when no recalculation verified the workbook", () => {
    const s = finishState(readiness("unverified"));
    expect(s?.acceptable).toBe(false);
    expect(s?.reason).toContain("Press Recalculate");
    expect(finishState(null)).toBeNull();
  });

  it("says how many new errors the recalculation found, and where they come from after a loop (16 Nov 2026)", () => {
    const r = readiness("unverified");
    const ran = { ...r, recalc: { ran: true, clean: false, new_errors: 667, formula_errors: 667, loop_note: "These errors may come from the Mind loop" } };
    expect(finishState(ran)).toMatchObject({ acceptable: false, text: "The recalculation found 667 new formula errors.", reason: "These errors may come from the Mind loop" });
    const plain = { ...r, recalc: { ran: true, clean: false, new_errors: 1 } };
    expect(finishState(plain)?.text).toBe("The recalculation found 1 new formula error.");
    expect(finishState(plain)?.reason).toContain("Advanced");
  });

  it("offers Recalculate only when the recalculation is the one thing left", () => {
    expect(onlyRecalcLeft(readiness("unverified"))).toBe(true);
    expect(onlyRecalcLeft({ ...readiness("unverified"), recalc: { ran: true, clean: false } })).toBe(false);
    expect(onlyRecalcLeft(readiness("blocked", 1))).toBe(false);
    expect(onlyRecalcLeft(readiness("ready"))).toBe(false);
    expect(onlyRecalcLeft(null)).toBe(false);
  });
});

describe("the Basic screen", () => {
  let Page: typeof HomePage;
  let Banner: typeof import("./finish-banner").FinishBanner;
  let Bar: typeof RunBar;
  let page: HomePage;
  let store: SessionStore;

  beforeAll(async () => {
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
    Banner = (await import("./finish-banner")).FinishBanner;
    Bar = (await import("./run-bar")).RunBar;
  });

  function open(mode: "basic" | "custom", api: Partial<ApiService> = {}): void {
    const none = () => Promise.reject(new Error("none"));
    TestBed.configureTestingModule({
      providers: [
        provideRouter([]),
        { provide: ActivatedRoute, useValue: { snapshot: { data: { mode } } } },
        { provide: ApiService, useValue: { estimate: none, mindLoopPlan: none, autoFixStatus: none, outsideInputs: none, nextQuarter: none, resume: none, downloadUrl: () => "/x", ...api } },
      ],
    });
    store = TestBed.inject(SessionStore);
    page = TestBed.runInInjectionContext(() => new Page());
  }

  afterEach(() => page?.ngOnDestroy());

  it("is Basic by route data, never auto-runs, and has its one tick box ticked", () => {
    localStorage.setItem("mindReady.autoRun", "1");
    open("basic");
    expect(page.basic).toBe(true);
    expect(page.autoRun()).toBe(false);
    expect(page.buildLoop()).toBe(true);
    localStorage.removeItem("mindReady.autoRun");
  });

  it("is Custom with its own route data: everything the main screen had, the auto-run tick included", () => {
    localStorage.setItem("mindReady.autoRun", "1");
    open("custom");
    expect(page.basic).toBe(false);
    expect(page.autoRun()).toBe(true);
    expect(page.loopWanted()).toBe(true);
    localStorage.removeItem("mindReady.autoRun");
  });

  function stub(results: { fixAll?: boolean; prepare?: boolean; loop?: boolean; recalc?: boolean } = {}): string[] {
    const ran: string[] = [];
    const p = page as unknown as Record<string, unknown>;
    for (const [name, key] of [["runFixAll", "fixAll"], ["runPrepare", "prepare"], ["runLoop", "loop"], ["runRecalc", "recalc"]] as const) {
      p[name] = async () => (ran.push(key), results[key] ?? true);
    }
    page.canBuildLoop = () => true; // the workbook has a loop
    return ran;
  }

  async function opened(): Promise<void> {
    open("basic");
    page.openSaved(started("m1"));
    await settle();
  }

  it("Go runs the loop step too when the box is ticked, then recalculates", async () => {
    await opened();
    const ran = stub();
    page.buildLoop.set(true);
    await page.runAll();
    expect(ran).toEqual(["fixAll", "prepare", "loop", "recalc"]);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
  });

  it("Go leaves the loop out when the box is not ticked, and says so", async () => {
    await opened();
    const ran = stub();
    page.buildLoop.set(false);
    await page.runAll();
    expect(ran).toEqual(["fixAll", "prepare", "recalc"]);
    expect(page.steps().find((s) => s.id === "loop")).toMatchObject({ state: "skipped", line: "Left out: the box was not ticked" });
  });

  it("Go stops at the first failed check, in Go's own words", async () => {
    await opened();
    const ran = stub({ prepare: false });
    await page.runAll();
    expect(ran).toEqual(["fixAll", "prepare"]);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "failed", "waiting", "waiting"]);
    expect(page.basicLine()).toMatch(/^Stopped at "Fix the formulas and the structure"/);
    expect(page.basicLine()).toContain("press Start again.");
    expect(page.basicLine()).not.toContain("Do everything");
  });

  it("Custom still runs the loop whatever the box says", () => {
    open("custom");
    page.buildLoop.set(false);
    expect(page.loopWanted()).toBe(true);
  });

  it("Upload to Mind asks for the sign-in when it is missing, and does not start", async () => {
    const calls: string[] = [];
    open("basic", {
      mindSignIn: async () => (calls.push("signIn"), { state: "done", message: null, last_login: null, checked: { signed_in: false, at: null }, elapsed_s: null }),
      startMindLoop: async () => (calls.push("start"), { state: "running" }),
    });
    store.setSession(started("m1"));
    const f = TestBed.createComponent(Banner);
    f.componentRef.setInput("readiness", readiness("ready"));
    f.detectChanges();
    await f.componentInstance.upload();
    expect(calls).toEqual(["signIn"]);
    expect(f.componentInstance.askSignIn()).toBe(true);
    expect(f.componentInstance.running()).toBe(false);
    expect(f.componentInstance.signInLine()).toContain("Not signed in to Mind");
  });

  it("Upload to Mind makes the same call as Run in Mind once signed in, and reads the answer back", async () => {
    const calls: string[] = [];
    open("basic", {
      mindSignIn: async () => ({ state: "done", message: null, last_login: null, checked: { signed_in: true, at: null }, elapsed_s: null }),
      mindLoopOptions: async () => ({ choices: [{ id: "a", default_on: true }, { id: "b", default_on: false }] }) as never,
      startMindLoop: async (_id: string, o: unknown) => (calls.push(JSON.stringify(o)), { state: "running" }),
      mindLoopStatus: async () => ({ state: "done", events: [], next: 1, report: { verdict: "converged", reason: "Mind ran it.", iterations: [] } }) as never,
    });
    store.setSession(started("m1"));
    const f = TestBed.createComponent(Banner);
    f.componentRef.setInput("readiness", readiness("ready"));
    f.detectChanges();
    await f.componentInstance.upload();
    expect(calls).toEqual([JSON.stringify({ maxIterations: 4, runModel: true, useAssistant: true, actions: ["a"] })]);
    expect(f.componentInstance.outcome()).toBe("Mind ran it.");
    expect(f.componentInstance.failed()).toBe(false);
    f.destroy();
  });

  it("draws Upload greyed with the reason, and Download usable, when blocking problems are left", () => {
    open("basic");
    const f = TestBed.createComponent(Banner);
    f.componentRef.setInput("readiness", readiness("blocked", 2));
    f.componentRef.setInput("downloadUrl", "/api/sessions/m1/files/m1_v3.xlsm");
    f.detectChanges();
    const el = f.nativeElement as HTMLElement;
    expect(el.querySelector(".finish__text")?.textContent).toBe("2 blocking problems are left.");
    expect((el.querySelector(".finish__upload") as HTMLButtonElement).disabled).toBe(true);
    expect(el.querySelector(".finish__reason")?.textContent).toMatch(/blocking problems are left/);
    expect(el.querySelector(".finish__download")?.getAttribute("href")).toBe("/api/sessions/m1/files/m1_v3.xlsm");
    f.componentRef.setInput("readiness", readiness("ready"));
    f.detectChanges();
    expect(el.querySelector(".finish__text")?.textContent).toBe("The check is finished. The workbook can be uploaded to Mind.");
    expect((el.querySelector(".finish__upload") as HTMLButtonElement).disabled).toBe(false);
    expect(el.querySelector(".finish__upload")?.textContent?.trim()).toBe("Upload to Mind");
  });

  // ---- 16 Nov 2026 (night): the loading bar while Start runs, and the words when a step runs late --------
  /** 22:03:00 on the wall clock: the page reads the time from here, and the test moves it */
  const T0 = new Date(2026, 10, 16, 22, 3, 0).getTime();
  let now = T0;

  afterEach(() => vi.restoreAllMocks());

  function clockAt(t0: number): void {
    now = t0;
    vi.spyOn(Date, "now").mockImplementation(() => now);
  }

  /** So many seconds later: the page's own tick (every 10 seconds), by hand. */
  function later(seconds: number): void {
    now += seconds * 1000;
    (page as unknown as { clock: { set(value: number): void } }).clock.set(now);
  }

  type Gates = Record<"fixAll" | "prepare" | "loop" | "recalc", ReturnType<typeof hold<boolean>>>;

  /** Every step of Start held open until the test ends it (the workbook has a loop). */
  function held(): Gates {
    const gates: Gates = { fixAll: hold(), prepare: hold(), loop: hold(), recalc: hold() };
    const p = page as unknown as Record<string, unknown>;
    for (const [name, key] of [["runFixAll", "fixAll"], ["runPrepare", "prepare"], ["runLoop", "loop"], ["runRecalc", "recalc"]] as const) {
      p[name] = () => gates[key].done;
    }
    page.canBuildLoop = () => true;
    return gates;
  }

  async function endStep(gates: Gates, step: keyof Gates): Promise<void> {
    gates[step].end(true);
    await settle();
  }

  it("says the step, the time left and that it is an estimate, while the step is on time", async () => {
    clockAt(T0);
    open("basic", { estimate: async () => estimate() });
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    await settle();
    later(50);
    expect(page.basicLines()).toEqual([
      "Step 1 of 4: fixing the formula errors, so the workbook calculates cleanly.",
      "About 7 min left (around 22:10).",
      "An estimate only: big models can take longer.",
    ]);
    expect(page.basicLine()).toBe("Step 1 of 4: fixing the formula errors, so the workbook calculates cleanly. About 7 min left (around 22:10). An estimate only: big models can take longer.");
    expect(page.basicEta()).toBe("About 7 min left (around 22:10). An estimate only: big models can take longer.");
    for (const step of ["fixAll", "prepare", "loop", "recalc"] as const) await endStep(gates, step);
    await run;
  });

  it("says one thing when the step runs late: longer than usual, and no 'an estimate only'", async () => {
    clockAt(T0);
    open("basic", { estimate: async () => estimate() });
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    await settle();
    later(250); // the step's own longest time was 200 seconds
    expect(page.basicLines()).toEqual(["Step 1 of 4: fixing the formula errors, so the workbook calculates cleanly.", "This step is taking longer than usual; about 7 min left (around 22:13), maybe more."]);
    expect(page.basicLine()).not.toContain("An estimate only");
    expect(page.basicLine()).not.toMatch(/usual\. /);
    // the next step starts on time: the disclaimer is back
    await endStep(gates, "fixAll");
    expect(page.basicLines()[0]).toMatch(/^Step 2 of 4: /);
    expect(page.basicLines()[2]).toBe("An estimate only: big models can take longer.");
    for (const step of ["prepare", "loop", "recalc"] as const) await endStep(gates, step);
    await run;
  });

  it("never says 'about under a minute'", async () => {
    clockAt(T0);
    const short = estimate();
    short.steps = [{ id: "recalc", low_s: 20, mid_s: 40, high_s: 80, why: "" }];
    open("basic", { estimate: async () => short });
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    await settle();
    for (const step of ["fixAll", "prepare", "loop"] as const) await endStep(gates, step);
    expect(page.basicLines()).toEqual(["Step 4 of 4: recalculating in Excel and checking that no value changed.", "Under a minute left (around 22:03).", "An estimate only: big models can take longer."]);
    later(100);
    expect(page.basicLines()[1]).toBe("This step is taking longer than usual; under a minute left (around 22:05), maybe more.");
    await endStep(gates, "recalc");
    await run;
  });

  it("Custom's timer says the same one thing when the step runs late", async () => {
    clockAt(T0);
    open("custom", { estimate: async () => estimate() });
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    await settle();
    later(50);
    expect(page.timerLine()).toBe("under a minute so far · about 7 min left");
    later(200);
    expect(page.timerLine()).toBe("4 min so far · this step is taking longer than usual; about 7 min left, maybe more");
    for (const step of ["fixAll", "prepare", "loop", "recalc"] as const) await endStep(gates, step);
    await run;
  });

  it("has no bar when nothing runs; while Start runs the bar fills by the time worked and the time left", async () => {
    clockAt(T0);
    open("basic", { estimate: async () => estimate() });
    page.openSaved(started("m1"));
    await settle();
    expect(page.basicBar()).toBeNull();
    const gates = held();
    const run = page.runAll();
    await settle();
    expect(page.basicBar()).toEqual({ percent: 0 });
    later(50); // 50 s worked, 50 + 370 s left
    expect(page.basicBar()).toEqual({ percent: 10 });
    later(200); // 250 s worked, 30 + 370 s left
    expect(page.basicBar()).toEqual({ percent: 38 });
    // the first step took 2.5 times its estimate, so the rest is expected to: more is left, and the bar holds
    await endStep(gates, "fixAll");
    expect(page.basicLines()[1]).toBe("About 15 min left (around 22:22).");
    expect(page.basicBar()).toEqual({ percent: 38 });
    later(400); // 650 s worked, 30 + 625 s left
    expect(page.basicBar()).toEqual({ percent: 49 });
    // far past every estimate, in the last step: never 100 before the end
    for (const step of ["prepare", "loop"] as const) await endStep(gates, step);
    later(90_000);
    expect(page.chainRunning()).toBe(true);
    expect(page.basicBar()).toEqual({ percent: 99 });
    await endStep(gates, "recalc");
    await run;
    expect(page.basicBar()).toBeNull();
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
  });

  it("only moves while the estimate has not come, then fills", async () => {
    clockAt(T0);
    const asked = hold<ChainEstimate>();
    open("basic", { estimate: () => asked.done });
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    await settle();
    expect(page.basicLines()).toEqual(["Getting ready…", "Working out how long this will take."]);
    expect(page.basicBar()).toEqual({ percent: null });
    asked.end(estimate());
    await settle();
    expect(page.basicLines()[1]).toBe("About 8 min left (around 22:10).");
    expect(page.basicBar()).toEqual({ percent: 0 });
    for (const step of ["fixAll", "prepare", "loop", "recalc"] as const) await endStep(gates, step);
    await run;
  });

  it("fills by step when the app has no estimate for the run, and promises none", async () => {
    clockAt(T0);
    open("basic"); // the estimate is refused
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    // the first moment: the answer is not in yet
    expect(page.basicBar()).toEqual({ percent: null });
    expect(page.basicLine()).toBe("Getting ready… Working out how long this will take.");
    await settle();
    expect(page.basicLines()).toEqual(["Step 1 of 4: fixing the formula errors, so the workbook calculates cleanly."]);
    const seen = [page.basicBar()?.percent];
    for (const step of ["fixAll", "prepare", "loop"] as const) {
      await endStep(gates, step);
      seen.push(page.basicBar()?.percent);
    }
    expect(seen).toEqual([12, 37, 62, 87]); // each finished step whole, the running one half
    await endStep(gates, "recalc");
    await run;
    expect(page.basicBar()).toBeNull();
  });

  it("shows no bar while Start waits for the input files (nothing runs)", async () => {
    open("basic");
    page.chainRunning.set(true);
    expect(page.basicBar()).toEqual({ percent: null });
    page.outsideWaiting.set(true);
    expect(page.basicBar()).toBeNull();
    expect(page.basicLines()).toEqual(["Waiting for the input files."]);
    page.outsideWaiting.set(false);
    page.chainRunning.set(false);
  });

  /** The Basic screen drawn: the page's own template with the dots and the bar (Carbon stays inert). */
  async function draw(api: Partial<ApiService>): Promise<{ fixture: ComponentFixture<HomePage>; el: HTMLElement; look: () => Promise<void> }> {
    const none = () => Promise.reject(new Error("none"));
    TestBed.configureTestingModule({
      providers: [
        provideRouter([]),
        { provide: ActivatedRoute, useValue: { snapshot: { data: { mode: "basic" } } } },
        { provide: ApiService, useValue: { estimate: none, mindLoopPlan: none, autoFixStatus: none, outsideInputs: none, nextQuarter: none, resume: none, downloadUrl: () => "/x", ...api } },
      ],
    });
    TestBed.overrideComponent(Page, { set: { imports: [CommonModule, Thinking, Bar], schemas: [NO_ERRORS_SCHEMA] } });
    store = TestBed.inject(SessionStore);
    const fixture = TestBed.createComponent(Page);
    page = fixture.componentInstance;
    const look = async () => {
      fixture.detectChanges();
      await settle();
      fixture.detectChanges();
    };
    page.openSaved(started("m1"));
    await look();
    return { fixture, el: fixture.nativeElement as HTMLElement, look };
  }

  const sentences = (el: HTMLElement) => Array.from(el.querySelectorAll(".basic__line .basic__sentence")).map((s) => s.textContent?.trim());

  it("draws the dots on the button and the bar under the line while Start runs, and neither when nothing runs", async () => {
    clockAt(T0);
    const asked = hold<ChainEstimate>();
    const { fixture, el, look } = await draw({ estimate: () => asked.done });
    // nothing runs: the Start button, no line, no bar
    expect(el.querySelector(".basic__go")?.textContent?.trim()).toBe("Start");
    expect(el.querySelector(".basic__line")).toBeNull();
    expect(el.querySelector("mr-run-bar")).toBeNull();

    const gates = held();
    const run = page.runAll();
    await look();
    // no estimate yet: the dots and "Working" on the button, the bar only moves and shows no number
    const button = el.querySelector(".basic__go") as HTMLButtonElement;
    expect(button.disabled).toBe(true);
    expect(button.classList.contains("basic__go--working")).toBe(true);
    expect(button.textContent?.trim()).toBe("Working");
    expect(button.querySelectorAll("mr-thinking .thinking__dot")).toHaveLength(3);
    expect(button.querySelector("mr-thinking")?.getAttribute("aria-hidden")).toBe("true");
    expect(sentences(el)).toEqual(["Getting ready…", "Working out how long this will take."]);
    expect(el.querySelector(".basic__line")?.getAttribute("role")).toBe("status");
    expect(el.querySelector("mr-run-bar .run-bar--moving")).not.toBeNull();
    expect(el.querySelector("mr-run-bar cds-progress-bar")).not.toBeNull(); // Carbon's own bar
    expect(el.querySelector("mr-run-bar .run-bar__percent")).toBeNull();
    // the bar comes after the line
    expect(el.querySelector(".basic__line")!.compareDocumentPosition(el.querySelector("mr-run-bar")!) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();

    // the estimate is in and the first step runs: a percentage
    asked.end(estimate());
    await look();
    later(50);
    await look();
    expect(sentences(el)).toEqual(["Step 1 of 4: fixing the formula errors, so the workbook calculates cleanly.", "About 7 min left (around 22:10).", "An estimate only: big models can take longer."]);
    expect(el.querySelector("mr-run-bar .run-bar--moving")).toBeNull();
    expect(el.querySelector("mr-run-bar .run-bar__percent")?.textContent).toBe("10%");

    // the step runs late: one sentence about the time, no disclaimer; the bar goes on
    later(200);
    await look();
    expect(sentences(el)).toEqual(["Step 1 of 4: fixing the formula errors, so the workbook calculates cleanly.", "This step is taking longer than usual; about 7 min left (around 22:13), maybe more."]);
    expect(el.querySelector("mr-run-bar .run-bar__percent")?.textContent).toBe("38%");

    // the run ends: the bar goes, the button is back
    for (const step of ["fixAll", "prepare", "loop", "recalc"] as const) await endStep(gates, step);
    await run;
    await look();
    expect(el.querySelector("mr-run-bar")).toBeNull();
    expect(el.querySelector(".basic__go")?.textContent?.trim()).toBe("Start again");
    expect(sentences(el)).toHaveLength(1);
    expect(sentences(el)[0]).toMatch(/^Everything ran and every check passed/);
    fixture.destroy();
  });
});

describe("Recent workbooks: collapsed on Basic, with Remove all", () => {
  let List: typeof RecentWorkbooks;
  beforeAll(async () => {
    List = (await import("../pages/recent-workbooks")).RecentWorkbooks;
  });

  function draw(list: SavedWorkbook[], removeAll?: (state: { list: SavedWorkbook[] }) => Promise<{ removed: number; kept: number }>, collapsed = true) {
    const state = { list };
    const api = {
      savedWorkbooks: () => Promise.resolve([...state.list]),
      removeAllSavedWorkbooks: () => (removeAll ? removeAll(state) : Promise.resolve({ removed: list.length, kept: 0 }).then((r) => ((state.list = []), r))),
    } as unknown as ApiService;
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: api }] });
    const fixture = TestBed.createComponent(List);
    fixture.componentRef.setInput("collapsible", collapsed);
    fixture.detectChanges();
    return { fixture, el: fixture.nativeElement as HTMLElement, comp: fixture.componentInstance, state };
  }

  it("shows only a click-to-open line on Basic, then the list with Reopen, Remove and Remove all", async () => {
    const { fixture, el } = draw([saved("a"), saved("b")]);
    await settle();
    fixture.detectChanges();
    expect(el.querySelector(".recent__toggle")?.textContent?.trim()).toBe("Show recent workbooks (2)");
    expect(el.querySelector(".recent__list")).toBeNull();
    (el.querySelector(".recent__toggle") as HTMLButtonElement).click();
    fixture.detectChanges();
    expect(el.querySelectorAll(".recent__row")).toHaveLength(2);
    expect(el.querySelector(".recent__all-remove")?.textContent?.trim()).toBe("Remove all");
    expect(el.querySelector(".recent__toggle")?.textContent?.trim()).toBe("Hide recent workbooks (2)");
  });

  it("is open as before on Custom (no toggle)", async () => {
    const { fixture, el } = draw([saved("a")], undefined, false);
    await settle();
    fixture.detectChanges();
    expect(el.querySelector(".recent__toggle")).toBeNull();
    expect(el.querySelector(".recent__title")?.textContent).toBe("Recent workbooks");
    expect(el.querySelectorAll(".recent__row")).toHaveLength(1);
  });

  it("Remove all asks the question with the count and removes nothing until Yes", async () => {
    let called = 0;
    const { fixture, el, comp } = draw([saved("a"), saved("b")], async (state) => (called++, (state.list = []), { removed: 2, kept: 0 }), false);
    await settle();
    fixture.detectChanges();
    (el.querySelector(".recent__all-remove") as HTMLButtonElement).click();
    fixture.detectChanges();
    expect(el.querySelector(".recent__all-question")?.textContent).toBe("Remove all 2 saved workbooks?");
    expect(called).toBe(0);
    (el.querySelector(".recent__all-no") as HTMLButtonElement).click();
    fixture.detectChanges();
    expect(el.querySelector(".recent__all-question")).toBeNull();
    expect(comp.rows()).toHaveLength(2);
    (el.querySelector(".recent__all-remove") as HTMLButtonElement).click();
    fixture.detectChanges();
    (el.querySelector(".recent__all-yes") as HTMLButtonElement).click();
    await settle();
    fixture.detectChanges();
    expect(called).toBe(1);
    expect(comp.rows()).toEqual([]);
    expect(el.querySelector(".recent")).toBeNull();
  });

  it("says how many were kept when another copy of the app holds them", async () => {
    const { fixture, el, comp } = draw([saved("a"), saved("b")], async (state) => ((state.list = [saved("b")]), { removed: 1, kept: 1 }), false);
    await settle();
    comp.askRemoveAll();
    await comp.confirmRemoveAll();
    fixture.detectChanges();
    expect(comp.rows().map((r) => r.id)).toEqual(["b"]);
    expect(el.querySelector(".recent__note")?.textContent).toBe("Removed 1. 1 kept: it is open in another copy of the app or still running.");
  });
});

// ---- 16 Nov 2026 (night): the bar itself ----------------------------------------------------------------
/** Carbon's progress bar as far as the run bar relies on it (carbon-components-angular 5.72, `ProgressBar`):
 *  the track is the role="progressbar", named by the label's id (the label can be hidden from the eye
 *  only); without a value the host is "indeterminate" (Carbon's moving bar) and the track carries no
 *  numbers. Carbon itself cannot load under Node here. */
@Component({
  selector: "cds-progress-bar",
  standalone: true,
  host: { "[class.cds--progress-bar--indeterminate]": "value === undefined" },
  template: `@if (label) {<div class="cds--progress-bar__label" [class.cds--visually-hidden]="hideLabel" [id]="id">{{ label }}</div>}
    <div class="cds--progress-bar__track" role="progressbar" [attr.aria-labelledby]="id" [attr.aria-valuemin]="value === undefined ? null : 0" [attr.aria-valuemax]="value === undefined ? null : max" [attr.aria-valuenow]="value === undefined ? null : value"></div>`,
})
class CarbonBar {
  @Input() value: number | undefined;
  @Input() label = "";
  @Input() hideLabel = false;
  @Input() max = 100;
  id = "progress-bar-7";
}

// Last in this file: it lets the loaded modules go (the bar is compiled against the stand-in above), so the
// describes before it must have loaded theirs already.
describe("the run bar", () => {
  let Bar: typeof RunBar;

  beforeAll(async () => {
    vi.resetModules();
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : name === "ProgressBar" ? CarbonBar : Inert), has: () => true });
    });
    Bar = (await import("./run-bar")).RunBar;
  });

  // the files that come after load their own Carbon stand-ins
  afterAll(() => {
    vi.doUnmock("carbon-components-angular");
    vi.resetModules();
  });

  function draw(percent: number | null) {
    const fixture = TestBed.createComponent(Bar);
    fixture.componentRef.setInput("percent", percent);
    fixture.detectChanges();
    const el = fixture.nativeElement as HTMLElement;
    return { fixture, el, track: el.querySelector('[role="progressbar"]') as HTMLElement };
  }

  it("is a progress bar with a plain name and the value for a screen reader, and shows the percentage beside it", () => {
    const { el, track } = draw(38);
    expect(track.getAttribute("aria-valuenow")).toBe("38");
    expect(track.getAttribute("aria-valuemin")).toBe("0");
    expect(track.getAttribute("aria-valuemax")).toBe("100");
    const name = el.querySelector(`[id="${track.getAttribute("aria-labelledby")}"]`);
    expect(name?.textContent).toBe("Work done so far");
    expect(name?.classList.contains("cds--visually-hidden")).toBe(true);
    const shown = el.querySelector(".run-bar__percent");
    expect(shown?.textContent).toBe("38%");
    expect(shown?.getAttribute("aria-hidden")).toBe("true"); // the bar already says it
    expect(el.querySelector(".run-bar--moving")).toBeNull();
  });

  it("only moves, with no number, while there is no percentage", () => {
    const { fixture, el, track } = draw(null);
    expect(el.querySelector(".cds--progress-bar--indeterminate")).not.toBeNull();
    expect(el.querySelector(".run-bar--moving")).not.toBeNull();
    expect(track.hasAttribute("aria-valuenow")).toBe(false);
    expect(el.querySelector(".run-bar__percent")).toBeNull();
    expect(el.querySelector(`[id="${track.getAttribute("aria-labelledby")}"]`)?.textContent).toBe("Work done so far");
    // the percentage arrives: the same bar fills
    fixture.componentRef.setInput("percent", 12);
    fixture.detectChanges();
    expect(el.querySelector(".cds--progress-bar--indeterminate")).toBeNull();
    expect(track.getAttribute("aria-valuenow")).toBe("12");
    expect(el.querySelector(".run-bar__percent")?.textContent).toBe("12%");
  });

  it("keeps to whole numbers from 0 to 100", () => {
    expect(draw(37.6).el.querySelector(".run-bar__percent")?.textContent).toBe("38%");
    expect(draw(-5).el.querySelector(".run-bar__percent")?.textContent).toBe("0%");
    expect(draw(250).el.querySelector(".run-bar__percent")?.textContent).toBe("100%");
    expect(draw(Number.NaN).el.querySelector(".run-bar__percent")).toBeNull();
  });

  it("stops its motion for people who ask their system for less motion", () => {
    draw(null);
    const css = Array.from(document.querySelectorAll("style")).map((s) => s.textContent ?? "").find((text) => text.includes(".run-bar"));
    const calm = css?.split("prefers-reduced-motion")[1] ?? "";
    expect(calm).toMatch(/animation:\s*none/);
    expect(calm).toMatch(/transition:\s*none/);
  });
});
