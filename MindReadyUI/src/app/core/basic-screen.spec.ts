import { CommonModule } from "@angular/common";
import { Component, Input, NO_ERRORS_SCHEMA } from "@angular/core";
import { ComponentFixture, TestBed } from "@angular/core/testing";
import { onlyRecalcLeft } from "./recalc-now";
import { ActivatedRoute, provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import type { RecentWorkbooks } from "../pages/recent-workbooks";
import { ApiService } from "./api.service";
import { CancelButton } from "./cancel-button";
import type { FeedbackDialog } from "./feedback-dialog";
import { FeedbackService } from "./feedback.service";
import { finishState } from "./finish-banner";
import { ChainEstimate, LoopPlan, OutsideInputs, PrepAction, PrepAsk, PrepRunStatus, Readiness, STOPPED_NOTHING_CHANGED, SavedWorkbook, ScanStatus, SessionStart, Version, WorkJob } from "./models";
import type { RunBar } from "./run-bar";
import { planShowMe, PAGE_LINKS, ShowMeRequest, showMeFor } from "./show-me";
import { ShowMeLink } from "./show-me-link";
import { ShowMeService } from "./show-me.service";
import { SessionStore } from "./session.store";
import { Thinking } from "./thinking";
import { WorkService } from "./work.service";

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

// ---- 17 Nov 2026 (Model A): the errors that show only for the Mind loop's first value never stop Start ----------
const LOOP_LINE =
  "667 cells show an error only for the Mind loop's first value, for example Decrements!M704: #N/A means the value is not available (often: a lookup found nothing). " +
  "The original never computed these cells, and Mind will show the same for that value. Check the results in Mind.";
const OTHER_LINE = "The other 10 cells show an error only for the Mind loop's first value, for example Decrements!M704: #N/A means the value is not available (often: a lookup found nothing).";
const naCells = (n: number) => Array.from({ length: n }, (_, i) => ({ sheet: "Decrements", cell: `M${704 + i}`, error: "#N/A", loop_only: true }));
/** The Recalculate step's answer after the loop: `real` errors that also show without the loop, `loop` that do not. */
const afterLoop = (real: number, loop: number) => ({
  status: "ERROR" as const,
  message: "recalculated",
  ran: true,
  formula_errors: [...naCells(loop), ...Array.from({ length: real }, () => ({ sheet: "Results", cell: "B5", error: "#DIV/0!" }))],
  addin_gap_errors: [],
  new_errors: real + loop,
  preexisting_errors: 0,
  loop_check: { ran: true, from_loop: loop, also_without: real, seconds: 95 },
  readiness: {
    ...readiness(real ? "unverified" : "ready"),
    recalc: { ran: true, clean: !real, new_errors: real, formula_errors: real + loop, loop_errors: loop, loop_note: real ? OTHER_LINE : LOOP_LINE },
  },
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

  // (17 Nov 2026: the bubble said "Click Go"; the button says Start)
  it("points at Start on Basic and at Do everything on Custom; Download and the recent list exist on both", () => {
    expect(planShowMe({ id: "do-everything" }, "basic")).toEqual([{ kind: "place", targets: [{ el: "do-everything", text: "Click Start" }] }]);
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

  it("finishes to an acceptable banner when every new error shows only for the loop's first value, and says so in one line", () => {
    const r = afterLoop(0, 667).readiness;
    expect(finishState(r)).toEqual({ acceptable: true, text: "The check is finished. The workbook can be uploaded to Mind.", reason: "", note: LOOP_LINE });
    // without errors of the loop: as before, no line
    expect(finishState(readiness("ready"))).toEqual({ acceptable: true, text: "The check is finished. The workbook can be uploaded to Mind.", reason: "" });
    // some errors also show without the loop: still stopped, and only those are counted
    expect(finishState(afterLoop(1, 10).readiness)).toMatchObject({ acceptable: false, text: "The recalculation found 1 new formula error.", reason: OTHER_LINE });
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
  let Dialog: typeof FeedbackDialog;
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
    Dialog = (await import("./feedback-dialog")).FeedbackDialog;
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

  it("Start stops at the first failed check, in Start's own words", async () => {
    await opened();
    const ran = stub({ prepare: false });
    await page.runAll();
    expect(ran).toEqual(["fixAll", "prepare"]);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "failed", "waiting", "waiting"]);
    expect(page.basicLine()).toMatch(/^Stopped at "Fix the formulas and the structure"/);
    expect(page.basicLine()).toContain("press Start again.");
    expect(page.basicLine()).not.toContain("Do everything");
    // 17 Nov 2026: and a sentence that begins with the button keeps its capital (it came out as "press Start again.")
    const again = stub({ fixAll: false });
    page.fixAll.set({ running: false, fraction: null, line: "Not started", detail: "", tone: "error" });
    await page.runAll();
    expect(again).toEqual(["fixAll"]);
    expect(page.chainLine()).toBe('Stopped at "Fix all issues": Not started. Press Do everything again.');
    expect(page.basicLine()).toBe('Stopped at "Fix all issues": Not started. Press Start again.');
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

  it("19 Nov 2026, a sign-in that could not start: Basic shows the plain sentence and Send feedback with the incident", async () => {
    const sentence = "Microsoft Edge was not found. Run in Mind needs Edge: install it, then press Sign in to Mind again.";
    open("basic", {
      mindSignIn: async () => ({ state: "failed", message: sentence, last_login: null, checked: null, elapsed_s: 1, detail: "RuntimeError: Chromium distribution 'msedge' is not found", feedback_ref: "R9PS-0727", kept_in: "C:\\data\\edge-profile" }),
    });
    store.setSession(started("m1"));
    const f = TestBed.createComponent(Banner);
    f.componentRef.setInput("readiness", readiness("ready"));
    f.detectChanges();
    await f.componentInstance.upload();
    f.detectChanges();
    const box = f.nativeElement as HTMLElement;
    expect(f.componentInstance.signInLine()).toBe(sentence);
    expect(box.textContent).toContain(sentence);
    expect(box.textContent).not.toContain("RuntimeError");
    expect(box.textContent).not.toContain("edge-profile"); // where it is kept is said under Advanced only
    (box.querySelector('[data-test="finish-sign-in-feedback"]') as HTMLButtonElement).click();
    expect(TestBed.inject(FeedbackService).dialog()).toEqual({ ref: "R9PS-0727", note: sentence });
    TestBed.inject(FeedbackService).close();
    // without an incident (the person closed the window): no button
    f.componentInstance.signIn.set({ state: "failed", message: "Not signed in: the Edge window was closed before the sign-in finished.", last_login: null, checked: null, elapsed_s: 4, feedback_ref: null });
    f.detectChanges();
    expect(box.querySelector('[data-test="finish-sign-in-feedback"]')).toBeNull();
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

  it("draws the loop's line under an acceptable banner, with Upload usable", () => {
    open("basic");
    const f = TestBed.createComponent(Banner);
    f.componentRef.setInput("readiness", afterLoop(0, 667).readiness);
    f.detectChanges();
    const el = f.nativeElement as HTMLElement;
    expect(el.querySelector(".finish__text")?.textContent).toBe("The check is finished. The workbook can be uploaded to Mind.");
    expect(el.querySelector(".finish__note")?.textContent).toBe(LOOP_LINE);
    expect(el.querySelector(".finish__reason")).toBeNull();
    expect((el.querySelector(".finish__upload") as HTMLButtonElement).disabled).toBe(false);
    f.componentRef.setInput("readiness", readiness("ready"));
    f.detectChanges();
    expect(el.querySelector(".finish__note")).toBeNull();
    f.destroy();
  });

  /** Every step but the recalculation stands in and succeeds; the recalculation is the page's own, on `out`. */
  function realRecalc(out: ReturnType<typeof afterLoop>): void {
    const p = page as unknown as Record<string, unknown>;
    for (const name of ["runFixAll", "runPrepare", "runLoop"]) p[name] = async () => true;
    page.canBuildLoop = () => true;
    void out;
  }

  it("Basic goes on to the end when the only new errors are the loop's first value's: the line once, on the banner", async () => {
    const out = afterLoop(0, 667);
    open("basic", { recalculate: async () => out } as unknown as Partial<ApiService>);
    page.openSaved(started("m1"));
    await settle();
    realRecalc(out);
    page.buildLoop.set(true);
    await page.runAll();
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.recalc()).toMatchObject({ line: LOOP_LINE, tone: "ok" });
    const lines = page.basicLines();
    expect(lines[0]).toMatch(/^Everything ran and every check passed \(/);
    // (not under the end sentence too: the finish banner right below says it, see the banner spec above)
    expect(lines.slice(1)).toEqual([]);
    expect(store.readiness()?.state).toBe("ready");
    expect(finishState(store.readiness())?.note).toBe(LOOP_LINE);
    // Custom says it on the Recalculate step's own line, and its end sentence stays as it was
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    open("custom", { recalculate: async () => out } as unknown as Partial<ApiService>);
    page.openSaved(started("m1"));
    await settle();
    realRecalc(out);
    await page.runAll();
    expect(page.steps().find((s) => s.id === "recalc")).toMatchObject({ state: "done", line: LOOP_LINE });
    expect(page.chainLine()).toMatch(/^Everything ran and every check passed \(/);
  });

  it("Basic still stops for the errors that also show without the loop, counts only those, and names the others", async () => {
    const out = afterLoop(1, 10);
    open("basic", { recalculate: async () => out } as unknown as Partial<ApiService>);
    page.openSaved(started("m1"));
    await settle();
    realRecalc(out);
    page.buildLoop.set(true);
    await page.runAll();
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "failed"]);
    expect(page.basicLine()).toBe('Stopped at "Recalculate": Recalculated, 1 formula errors (the other 10 only for the Mind loop\'s first value). Fix that under Advanced › Recalculate, then press Start again.');
    // the general stop keeps its Show me, to the list of errors
    expect(page.stopPlace()).toEqual(showMeFor('Fix that under Advanced › Recalculate, then press Do everything again.'));
    expect(page.recalc().detail).toContain(OTHER_LINE);
  });

  it("counts only the errors that also show without the loop on the stop line, not the ones the original already had (review LBr)", async () => {
    const base = afterLoop(1, 10);
    // two more errors the original already had: they never stop Start, and the verdict leaves them out
    const own = [{ sheet: "Inputs", cell: "C9", error: "#DIV/0!", preexisting: true }, { sheet: "Inputs", cell: "C10", error: "#DIV/0!", preexisting: true }];
    const out = { ...base, formula_errors: [...base.formula_errors, ...own], preexisting_errors: 2 };
    open("basic", { recalculate: async () => out } as unknown as Partial<ApiService>);
    page.openSaved(started("m1"));
    await settle();
    realRecalc(out);
    page.buildLoop.set(true);
    await page.runAll();
    // it said "Recalculated, 3 formula errors (the other 10 ...)", while the banner said 1 new
    expect(page.basicLine()).toBe('Stopped at "Recalculate": Recalculated, 1 formula errors (the other 10 only for the Mind loop\'s first value). Fix that under Advanced › Recalculate, then press Start again.');
    expect(finishState(store.readiness())?.text).toBe("The recalculation found 1 new formula error.");
    expect(page.recalc().detail).toContain("2 already in the original");
  });

  it("says plainly, on the banner under the stop, that the loop's errors could not be checked when the version before the loop did not recalculate", async () => {
    const COULD_NOT = "Whether these errors come only from the Mind loop's first value could not be checked: Excel could not recalculate the version before the loop.";
    const base = afterLoop(0, 0);
    const out = {
      ...base,
      formula_errors: naCells(667).map((e) => ({ sheet: e.sheet, cell: e.cell, error: e.error })),
      new_errors: 667,
      loop_check: { ran: false, seconds: 3, message: "Excel COM recalculation failed" },
      readiness: { ...readiness("unverified"), recalc: { ran: true, clean: false, new_errors: 667, formula_errors: 667, loop_errors: 0, loop_note: COULD_NOT } },
    };
    open("basic", { recalculate: async () => out } as unknown as Partial<ApiService>);
    page.openSaved(started("m1"));
    await settle();
    realRecalc(out as unknown as ReturnType<typeof afterLoop>);
    page.buildLoop.set(true);
    await page.runAll();
    // the general stop sentence, as Joseph approved it, with its Show me; the banner right under it says why
    expect(page.basicLine()).toBe('Stopped at "Recalculate": Recalculated, 667 formula errors. Fix that under Advanced › Recalculate, then press Start again.');
    expect(finishState(store.readiness())).toMatchObject({ acceptable: false, text: "The recalculation found 667 new formula errors.", reason: COULD_NOT });
    expect(page.recalc().detail).toContain(COULD_NOT);
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
    expect(el.querySelector("mr-cancel")).toBeNull(); // 17 Nov 2026: nothing runs, so there is nothing to cancel

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
    // 17 Nov 2026: Cancel stands beside the "Working" button, in the same row, after it (its own tests:
    // cancel-button.spec.ts, and cancel-places.spec.ts for the click)
    const cancel = el.querySelector(".basic__run mr-cancel");
    expect(cancel).not.toBeNull();
    expect(button.parentElement).toBe(cancel?.parentElement);
    expect(button.compareDocumentPosition(cancel!) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
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
    expect(el.querySelector("mr-cancel")).toBeNull();
    expect(el.querySelector(".basic__go")?.textContent?.trim()).toBe("Start again");
    expect(sentences(el)).toHaveLength(1);
    expect(sentences(el)[0]).toMatch(/^Everything ran and every check passed/);
    fixture.destroy();
  });

  // ---- 17 Nov 2026: never stop for a person when the app can fix it -------------------------------------
  // The Prep step is one job on the server (POST and GET .../prep-run): rounds while a round still changes
  // something, then the assistant, then the numbers check. The page starts it, follows it, and when
  // something is left it asks in two plain sentences, with "Show me" and "Send feedback".

  const ver = (n: number): Version => ({ id: `ver-00${n}`, major: n, minor: 0, label: `v${n}`, file_name: `m1_v${n}.xlsm`, source: "prep" });

  /** The question as the server words it (app/handoff.py `ask_lines`), on Basic. */
  const ASK_LINES = ["One thing is left that the app could not fix: a formula reads a lone text cell, on sheet Roll, cell W25.", "Fix it by hand, then press Start again."];
  const NO_KEY_LINES = ["Something is left that needs the assistant, and no AI key is saved.", "Paste an AI key under Advanced › Assistant, then press Start again."];
  const WHY = "The assistant's fix changed a result, so it was taken back.";
  const NOTE = "Start could not fix this by itself: a formula reads a lone text cell. The assistant tried twice: its change altered a result and was undone.";
  const question = (over: Partial<PrepAsk> = {}): PrepAsk => ({ lines: ASK_LINES, why: WHY, stops_start: true, show_me: { id: "prep-action", rule_id: "STR-008", sheet: "Roll", cell: "W25" }, feedback: { ref: "ABCD-1234", note: NOTE }, ...over });
  const noKey = (): PrepAsk => ({ lines: NO_KEY_LINES, why: "The assistant was not asked.", stops_start: false, show_me: { id: "assistant-key", rule_id: null, sheet: "", cell: "" }, feedback: null });

  const round = (n: number, changes = 966): PrepRunStatus => ({ state: "running", stage: "pass", pass_no: n, message: `Round ${n}: writing ${changes} changes through Excel…` });
  const done = (applied: number, fixed = 0, refused = 0): PrepRunStatus => ({ state: "done", stage: "done", applied, refused, assistant_fixed: fixed, left: 0, ask: null });
  const asks = (ask: PrepAsk, applied = 954): PrepRunStatus => ({ state: "ask", stage: "done", applied, refused: 0, assistant_fixed: 0, left: 1, ask });

  /** A stand-in for the server's Prep job. POST .../prep-run starts it; GET .../prep-run answers the next
   *  of `looks` (the last one repeats; an entry may be held open by the test); GET /api/sessions/{id}
   *  gives `after`. Before the start the job reads "idle", unless `pickUp`: then it runs already. */
  interface PrepFake {
    api: Partial<ApiService>;
    /** "start", "look", "resume", "cancel:prep", in order */
    calls: string[];
    /** what each POST .../prep-run was sent */
    starts: unknown[];
    /** the step's line on screen when each look was made */
    seen: string[];
    looks: (PrepRunStatus | Promise<PrepRunStatus>)[];
    after: SessionStart;
    pickUp: boolean;
  }

  function prepFake(looks: (PrepRunStatus | Promise<PrepRunStatus>)[], after: Partial<SessionStart> = {}): PrepFake {
    const f: PrepFake = { api: {}, calls: [], starts: [], seen: [], looks, after: { ...started("m1"), ...after }, pickUp: false };
    let last: PrepRunStatus | Promise<PrepRunStatus> = { state: "idle" };
    f.api = {
      startPrepRun: (_id: string, o: unknown) => (f.calls.push("start"), f.starts.push(o), Promise.resolve({ state: "running" })),
      prepRun: () => {
        if (!f.pickUp && !f.calls.includes("start")) return Promise.resolve({ state: "idle" } as PrepRunStatus);
        f.calls.push("look");
        f.seen.push(page.prepare().line);
        last = f.looks.shift() ?? last;
        return Promise.resolve(last);
      },
      resume: () => (f.calls.push("resume"), Promise.resolve(f.after)),
      scanStatus: () => Promise.resolve({ state: "ready" }),
      cancel: (_id: string, kind: string | null = null) => (f.calls.push(`cancel:${kind}`), Promise.resolve({ cancelled: kind ? [kind] : [], late: [], stopping: [], ahead: false, message: STOPPED_NOTHING_CHANGED, session: null, mind: null })),
    } as Partial<ApiService>;
    return f;
  }

  const noWaits = () => ((page as unknown as { sleep: () => Promise<void> }).sleep = () => Promise.resolve()); // no real second between two looks

  async function openedFor(mode: "basic" | "custom", f: PrepFake): Promise<void> {
    open(mode, f.api);
    noWaits();
    page.openSaved(started("m1"));
    await settle();
  }

  /** Start with the real Prep step; the other steps stand in (the workbook has a loop). */
  function others(results: { fixAll?: boolean; loop?: boolean; recalc?: boolean } = {}): string[] {
    const ran: string[] = [];
    const p = page as unknown as Record<string, unknown>;
    for (const [name, key] of [["runFixAll", "fixAll"], ["runLoop", "loop"], ["runRecalc", "recalc"]] as const) {
      p[name] = async () => (ran.push(key), results[key] ?? true);
    }
    page.canBuildLoop = () => true;
    return ran;
  }

  async function until(cond: () => boolean): Promise<void> {
    for (let i = 0; i < 300 && !cond(); i++) await settle();
    expect(cond()).toBe(true);
  }

  it("the Prep step is one job on the server: started once, followed to its end, then the workbook is read back", async () => {
    const f = prepFake(
      [round(1), { state: "running", stage: "assistant", problem_no: 1, problem_count: 2, message: "The assistant is working on what is left (1 of 2)." }, { state: "running", stage: "numbers", message: "Checking that no result changed." }, done(966)],
      { version: ver(5), versions: [ver(3), ver(4), ver(5)] },
    );
    await openedFor("basic", f);
    expect(await page.runPrepare()).toBe(true);
    expect(f.calls).toEqual(["start", "look", "look", "look", "look", "resume"]);
    // the job is told the name of the button, for the sentences it writes ("... then press Start again.")
    expect(f.starts).toEqual([{ button: "Start" }]);
    // the step's line while it ran: the job's own sentence, one at a time
    expect(f.seen).toEqual(["Looking at what blocks the file…", "Round 1: writing 966 changes through Excel…", "The assistant is working on what is left (1 of 2).", "Checking that no result changed."]);
    expect(page.prepare()).toMatchObject({ running: false, line: "966 changes applied", tone: "ok" });
    // the versions the job made are in the store, from the workbook read back
    expect(store.currentVersion()?.label).toBe("v5");
    expect(store.versions().map((x) => x.label)).toEqual(["v3", "v4", "v5"]);
    expect(page.busy()).toBe(false);
  });

  it("Custom's job is told its button is Do everything", async () => {
    const f = prepFake([done(3)]);
    await openedFor("custom", f);
    expect(await page.runPrepare()).toBe(true);
    expect(f.starts).toEqual([{ button: "Do everything" }]);
  });

  it("says what the step did in today's words, and the assistant's fixes in a sentence that leads to Versions", async () => {
    const cases: [PrepRunStatus, string, string | null][] = [
      [done(966), "966 changes applied", null],
      [done(0), "Nothing blocked the file", null],
      [done(966, 2), "966 changes applied. The assistant fixed 2 more: see Advanced › Versions.", "versions"],
      [done(0, 1), "The assistant fixed 1: see Advanced › Versions.", "versions"],
    ];
    for (const [end, line, place] of cases) {
      const f = prepFake([end]);
      await openedFor("custom", f);
      expect(await page.runPrepare()).toBe(true);
      expect(page.prepare().line).toBe(line);
      expect(showMeFor(page.prepare().line)?.id ?? null).toBe(place);
      page.ngOnDestroy();
      TestBed.resetTestingModule();
    }
    // changes Excel refused on the way are said under "View details", in plain words
    const f = prepFake([done(12, 0, 3)], { version: ver(4) });
    await openedFor("custom", f);
    await page.runPrepare();
    expect(page.prepare()).toMatchObject({ line: "12 changes applied", detail: "v4 · Excel refused 3 changes" });
  });

  it("the bar of a round being written is the Apply's own; while the assistant works it only moves", async () => {
    const second = hold<PrepRunStatus>();
    const third = hold<PrepRunStatus>();
    const f = prepFake([round(1), second.done, third.done]);
    let apply: ScanStatus["apply"] = { state: "running", fraction: 0.4, title: "Writing the changes", message: "390 of 966" };
    f.api.scanStatus = () => Promise.resolve({ state: "ready", apply });
    await openedFor("custom", f);
    const run = page.runPrepare();
    await until(() => f.calls.filter((c) => c === "look").length === 2 && page.prepare().fraction === 0.4);
    expect(page.prepare()).toMatchObject({ running: true, line: "Round 1: writing 966 changes through Excel…", fraction: 0.4, detail: "Writing the changes · 390 of 966" });
    // the rounds are over: the assistant's turn has no Apply, so no fraction and no words of an Apply
    apply = { state: "done" };
    second.end({ state: "running", stage: "assistant", message: "The assistant is working on what is left (1 of 1)." });
    await until(() => f.calls.filter((c) => c === "look").length === 3);
    expect(page.prepare()).toMatchObject({ running: true, line: "The assistant is working on what is left (1 of 1).", fraction: null, detail: "" });
    third.end(done(966, 1));
    expect(await run).toBe(true);
  });

  it("a question that does not stop Start: the step is true, Start goes on, and the question is the chain's last line", async () => {
    const ask = question({ stops_start: false, show_me: { id: "assistant-ask", rule_id: "FRM-009", sheet: "Calc", cell: "B2" } });
    const f = prepFake([round(1, 12), asks(ask, 12)], { version: ver(4) });
    await openedFor("basic", f);
    const ran = others();
    await page.runAll();
    expect(ran).toEqual(["fixAll", "loop", "recalc"]); // Start went on past the Prep step
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.steps()[1]).toMatchObject({ id: "prepare", line: "12 changes applied" });
    // in place of "Everything ran and every check passed": the two sentences, one per line on Basic
    expect(page.basicLines()).toEqual(ASK_LINES);
    expect(page.chainLine()).toBe(ASK_LINES.join(" "));
    expect(page.basicLine()).not.toContain("Everything ran");
    expect(page.chainAsk()).toBe(ask);
    expect(page.askPlace(ask)).toBe("assistant-ask");
    expect(page.chainStoppedByCancel()).toBe(false);
  });

  it("a later step that stops Start keeps the Prep question's first sentence after its own (20 Nov 2026)", async () => {
    const ask = question({ stops_start: false, show_me: { id: "assistant-ask", rule_id: "FRM-009", sheet: "Calc", cell: "B2" } });
    const f = prepFake([asks(ask, 12)]);
    await openedFor("basic", f);
    const ran = others();
    const still: Readiness = { ...readiness("blocked", 1), blocking: [{ rule_id: "FRM-009", status: "ERROR", message: "m", fix: "assistant" }] };
    (page as unknown as Record<string, unknown>)["runLoop"] = async () => (ran.push("loop"), store.readiness.set(still), false);
    await page.runAll();
    expect(ran).toEqual(["fixAll", "loop"]);
    expect(page.chainLine()).toContain("Stopped at");
    expect(page.chainLine().endsWith(` ${ASK_LINES[0]}`)).toBe(true);
    // its problem gone by then: the stop sentence alone
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const g = prepFake([asks(ask, 12)]);
    await openedFor("basic", g);
    others();
    (page as unknown as Record<string, unknown>)["runLoop"] = async () => (store.readiness.set(readiness("blocked", 1)), false);
    await page.runAll();
    expect(page.chainLine()).toContain("Stopped at");
    expect(page.chainLine()).not.toContain(ASK_LINES[0]);
  });

  it("a question whose problem the later steps made go is not asked: Everything ran", async () => {
    const ask = question({ stops_start: false, show_me: { id: "assistant-ask", rule_id: "FRM-009", sheet: "Calc", cell: "B2" } });
    const f = prepFake([asks(ask, 12)]);
    await openedFor("basic", f);
    others();
    // the recalculation's answer brings the readiness as it is at the end: nothing blocking is left
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(readiness("ready")), true);
    await page.runAll();
    expect(page.chainAsk()).toBeNull();
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
    // while it is still blocking at the end, it is asked
    const g = prepFake([asks(ask, 12)]);
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    await openedFor("basic", g);
    others();
    const still: Readiness = { ...readiness("blocked", 1), blocking: [{ rule_id: "FRM-009", status: "ERROR", message: "m", fix: "assistant" }] };
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(still), true);
    await page.runAll();
    expect(page.basicLines()).toEqual(ASK_LINES);
  });

  it("a server that says stop still stops Start: the step is false, nothing follows, and the stop line is the two sentences", async () => {
    // since 20 Nov 2026 the server never says so (Joseph: "App tries, never stops"); the screen still obeys it if it does
    const ask = question({ stops_start: true });
    const f = prepFake([round(1, 954), asks(ask)], { version: ver(4) });
    await openedFor("basic", f);
    const ran = others();
    await page.runAll();
    expect(ran).toEqual(["fixAll"]);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "failed", "waiting", "waiting"]);
    // the step says what it wrote, and why it could not finish is in its details
    expect(page.steps()[1]).toMatchObject({ line: "954 changes applied", detail: `v4 · ${WHY}` });
    expect(page.basicLines()).toEqual(ASK_LINES);
    expect(page.chainLine()).toBe(ASK_LINES.join(" "));
    expect(page.chainLine()).not.toContain("Stopped at"); // no count, no step name in quotes
    expect(page.chainAsk()).toBe(ask);
    expect(page.askPlace(ask)).toBe("prep-action");
    expect(store.currentVersion()?.label).toBe("v4"); // the rounds that ran are kept
  });

  it("with no AI key Start goes on and says so at its end while the workbook is still blocked, in the server's two sentences or in the page's own", async () => {
    // 20 Nov 2026 (Joseph: "App tries, never stops"): until then Start stopped here
    const still: Readiness = { ...readiness("blocked", 1), blocking: [{ rule_id: "FRM-009", status: "ERROR", message: "m", fix: "assistant" }] };
    const f = prepFake([{ state: "no_key", stage: "done", applied: 3, left: 1, ask: noKey() }]);
    await openedFor("basic", f);
    const ran = others();
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (ran.push("recalc"), store.readiness.set(still), true);
    await page.runAll();
    expect(ran).toEqual(["fixAll", "loop", "recalc"]);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.basicLines()).toEqual(NO_KEY_LINES);
    expect(page.askPlace(page.chainAsk()!)).toBe("assistant-key");
    expect(page.chainAsk()?.feedback).toBeNull(); // nothing went wrong in the app: nothing to send
    // a server that sends the state alone: the same sentences, with the screen's own button
    for (const [mode, button] of [["basic", "Start"], ["custom", "Do everything"]] as const) {
      page.ngOnDestroy();
      TestBed.resetTestingModule();
      const g = prepFake([{ state: "no_key", ask: null }]);
      await openedFor(mode, g);
      others();
      (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(still), true);
      await page.runAll();
      expect(page.chainAsk()?.lines).toEqual([NO_KEY_LINES[0], `Paste an AI key under Advanced › Assistant, then press ${button} again.`]);
      expect(page.askPlace(page.chainAsk()!)).toBe("assistant-key");
      expect(page.steps()[1].state).toBe("done");
    }
    // nothing blocks at the end (the loop or the recalculation made it go): no question, Everything ran
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const h = prepFake([{ state: "no_key", stage: "done", applied: 3, left: 1, ask: noKey() }]);
    await openedFor("basic", h);
    others();
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(readiness("ready")), true);
    await page.runAll();
    expect(page.chainAsk()).toBeNull();
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
  });

  it("Cancel in the second round throws the whole Prep step away: one job, so nothing of it is kept", async () => {
    const end = hold<PrepRunStatus>();
    const f = prepFake([round(1), round(2, 3), end.done]);
    await openedFor("basic", f);
    const ran = others();
    const run = page.runAll();
    await until(() => f.seen.includes("Round 2: writing 3 changes through Excel…"));
    page.onChainStopping(); // the click on Cancel beside the Working button
    expect(page.basicLines()).toEqual(["Stopping…"]);
    expect(page.chainRunning()).toBe(true); // the page ends nothing at the click: the job says "cancelled" itself
    end.end({ state: "cancelled", message: STOPPED_NOTHING_CHANGED });
    page.onChainStopped(); // the button heard from the server: the workbook is back as it was before the step
    await run;
    expect(page.prepare()).toMatchObject({ running: false, line: STOPPED_NOTHING_CHANGED, tone: "idle" });
    expect(page.steps().map((s) => [s.state, s.line])).toEqual([["done", ""], ["stopped", "Stopped."], ["waiting", ""], ["waiting", ""]]);
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was after the formula errors were fixed.");
    expect(page.chainStoppedByCancel()).toBe(true);
    expect(page.chainAsk()).toBeNull();
    // nothing of the job reached the store (the workbook comes from the Cancel's own answer), and nothing followed
    expect(f.calls).not.toContain("resume");
    expect(store.currentVersion()?.label).toBe("v3");
    expect(ran).toEqual(["fixAll"]);
  });

  it("a job that ends in an error stops Start with the sentence that names the step, in Basic's words", async () => {
    const f = prepFake([round(1), { state: "error", stage: "pass", error: "Excel did not start." }]);
    await openedFor("basic", f);
    others();
    await page.runAll();
    // the line does not guess what failed: the server's own sentence is under it
    expect(page.prepare()).toMatchObject({ line: "Stopped with an error", detail: "Excel did not start.", tone: "error" });
    expect(page.basicLine()).toBe('Stopped at "Fix the formulas and the structure": Stopped with an error. Fix that under Advanced › Prep actions, then press Start again.');
    expect(page.chainAsk()).toBeNull();
    // the rounds that had finished stay on the server: the workbook is read back after an error too
    expect(f.calls).toContain("resume");
    // a round that was written, but whose check of the workbook did not run: not "could not write"
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const written = "The changes were written, but the check of the workbook that follows did not run: Excel closed.";
    const g = prepFake([{ state: "error", stage: "pass", error: written }]);
    await openedFor("basic", g);
    expect(await page.runPrepare()).toBe(false);
    expect(page.prepare()).toMatchObject({ line: "Stopped with an error", detail: written, tone: "error" });
  });

  it("a start the server refuses is 'Not started' with its words; a job that runs already is followed instead", async () => {
    const busy = { status: 409, error: { detail: "Wait a moment: changes are still being written." } };
    const f = prepFake([]);
    f.api.startPrepRun = () => (f.calls.push("start"), Promise.reject(busy));
    f.looks.push({ state: "idle" });
    await openedFor("basic", f);
    expect(await page.runPrepare()).toBe(false);
    expect(page.prepare()).toMatchObject({ running: false, line: "Not started", detail: "Wait a moment: changes are still being written.", tone: "error" });

    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const g = prepFake([round(2), round(2), done(5)]);
    g.api.startPrepRun = () => (g.calls.push("start"), Promise.reject({ status: 409, error: { detail: "Wait a moment: the formulas and the structure are still being fixed." } }));
    await openedFor("basic", g);
    expect(await page.runPrepare()).toBe(true);
    expect(page.prepare().line).toBe("5 changes applied");
  });

  it("a look that fails is tried again; after five the job is lost", async () => {
    const f = prepFake([round(1)]);
    let failures = 4;
    const look = f.api.prepRun!;
    f.api.prepRun = (id: string) => (f.calls.includes("start") && f.seen.length && failures-- > 0 ? Promise.reject(new Error("offline")) : look(id));
    f.looks.push(done(7));
    await openedFor("basic", f);
    expect(await page.runPrepare()).toBe(true); // four failed looks in a row, then the server was back
    expect(page.prepare().line).toBe("7 changes applied");

    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const g = prepFake([]);
    g.api.prepRun = () => Promise.reject({ status: 0, message: "offline" });
    await openedFor("basic", g);
    expect(await page.runPrepare()).toBe(false);
    expect(page.prepare()).toMatchObject({ running: false, line: "Lost the job", detail: "offline", tone: "error" });
  });

  it("a Prep job that was running before a reload is picked up: followed to its end, with its question", async () => {
    const end = hold<PrepRunStatus>();
    const f = prepFake([round(2), end.done], { version: ver(5) });
    f.pickUp = true;
    open("custom", f.api);
    noWaits();
    page.openSaved(started("m1"));
    const picked = (page as unknown as { resumePrepRun: () => Promise<void> }).resumePrepRun();
    await until(() => page.prepare().running && f.calls.filter((c) => c === "look").length >= 2);
    // the step stands alone in the list, running; nothing else may start meanwhile
    expect(f.calls).not.toContain("start");
    expect(page.steps().map((s) => [s.id, s.state])).toEqual([["prepare", "running"]]);
    expect(page.prepare().line).toBe("Round 2: writing 966 changes through Excel…");
    expect(page.busy()).toBe(true);
    expect(page.busyHint()).toBe("Fix the formulas and the structure is running");
    expect(page.chainRunning()).toBe(false); // no chain in this page: the steps after it do not follow by themselves
    // the job was started by Start on Basic (its sentences say so) and is picked up here on Custom, after
    // the person went from one screen to the other: the question names the button this screen has
    const ask = question();
    end.end(asks(ask));
    await picked;
    expect(page.busy()).toBe(false);
    expect(store.currentVersion()?.label).toBe("v5");
    expect(page.steps()[0]).toMatchObject({ state: "failed", line: "954 changes applied" });
    expect(page.chainAsk()).toEqual({ ...ask, lines: [ASK_LINES[0], "Fix it by hand, then press Do everything again."] });
    expect(page.chainLine()).toBe(`${ASK_LINES[0]} Fix it by hand, then press Do everything again.`);
  });

  it("a picked-up job that ends clean, or is stopped, says so; one that does not run is left alone", async () => {
    const f = prepFake([round(1), done(4)]);
    f.pickUp = true;
    open("basic", f.api);
    noWaits();
    page.openSaved(started("m1"));
    await (page as unknown as { resumePrepRun: () => Promise<void> }).resumePrepRun();
    expect(page.steps()[0]).toMatchObject({ id: "prepare", state: "done", line: "4 changes applied" });
    expect(page.chainLine()).toBe("");

    // stopped from the strip while this page followed it: Basic's one line says so, with no Show me
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const g = prepFake([round(1), { state: "cancelled" }]);
    g.pickUp = true;
    open("basic", g.api);
    noWaits();
    page.openSaved(started("m1"));
    await (page as unknown as { resumePrepRun: () => Promise<void> }).resumePrepRun();
    expect(page.steps()[0]).toMatchObject({ state: "stopped", line: STOPPED_NOTHING_CHANGED });
    expect(page.basicLine()).toBe(STOPPED_NOTHING_CHANGED);
    expect(page.chainStoppedByCancel()).toBe(true);

    // a job that ended before the reload (its question may be out of date by now) starts nothing here
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const h = prepFake([asks(question())]);
    h.pickUp = true;
    open("basic", h.api);
    page.openSaved(started("m1"));
    await (page as unknown as { resumePrepRun: () => Promise<void> }).resumePrepRun();
    expect(page.steps()).toEqual([]);
    expect(page.chainAsk()).toBeNull();
    expect(page.prepare().running).toBe(false);
  });

  // ---- the fixes after the loop: the same job, and its question replaces the Prep step's ----------------
  const blockingLeft = (): PrepAction => ({ id: "compare_blank", title: "Compare with a blank", rule_ids: ["FRM-013"], default_on: true, level: "blocking", operations: [{ op: "set_formula", sheet: "S", cell: "B2" }], count: 1 });
  const loopPlan = (): LoopPlan => ({ found: true, loops: [{ name: "i", status: "ready", first: 1, last: 5, cells: [] }], outputs: [], clears: [], operation_count: 3 });

  /** Start with the real Prep step and the real loop step (its build answers at once); `afterLoop` is what
   *  the workbook reads back once the loop is in: the Prep plan of the new version. */
  async function withLoop(f: PrepFake, afterLoop: Promise<SessionStart>): Promise<string[]> {
    let resumes = 0;
    f.api.mindLoopPlan = () => Promise.resolve({ version_id: "ver-003", plan: loopPlan() });
    f.api.applyMindLoopPlan = (() => (f.calls.push("loop"), Promise.resolve({ result: { status: "ok", applied: 3, failed: 0 }, version: ver(5), plan: loopPlan() }))) as unknown as ApiService["applyMindLoopPlan"];
    // the first read is the Prep step's own; the loop's comes second, and may be slow
    f.api.resume = () => (f.calls.push("resume"), ++resumes === 2 ? afterLoop : Promise.resolve(f.after));
    await openedFor("basic", f);
    await (page as unknown as { loadLoopPlan: () => Promise<void> }).loadLoopPlan();
    const ran: string[] = [];
    const p = page as unknown as Record<string, unknown>;
    for (const [name, key] of [["runFixAll", "fixAll"], ["runRecalc", "recalc"]] as const) p[name] = async () => (ran.push(key), true);
    return ran;
  }

  it("the fixes after the loop start from the plan as it is after the loop, not the one from before it", async () => {
    const read = hold<SessionStart>();
    const f = prepFake([done(954), done(5)], { version: ver(4) });
    await withLoop(f, read.done);
    const run = page.runAll();
    await until(() => f.calls.includes("loop") && f.calls.filter((c) => c === "resume").length === 2);
    await settle();
    // the workbook has not been read back yet: the page waits, it does not decide from the old plan
    expect(f.calls.filter((c) => c === "start")).toHaveLength(1);
    expect(page.chainStep()?.id).toBe("loop");
    read.end({ ...started("m1"), version: ver(5), plan: [blockingLeft()] }); // the loop uncovered something Prep repairs
    await run;
    expect(f.calls.filter((c) => c === "start")).toHaveLength(2);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.steps()[2].line).toBe("3 loop formulas written; then 5 changes applied for what the loop uncovered");
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
  });

  it("the question of the fixes after the loop replaces the Prep step's; an answer of 'done' takes it away", async () => {
    const first = question({ stops_start: false, lines: ["One thing is left that the app could not fix: a formula has a broken reference, on sheet Calc, cell B2.", ASK_LINES[1]], show_me: { id: "assistant-ask", rule_id: "REF-001", sheet: "Calc", cell: "B2" } });
    const second = question({ stops_start: false, show_me: { id: "assistant-ask", rule_id: "STR-008", sheet: "Roll", cell: "W25" } });
    const f = prepFake([asks(first, 10), asks(second, 5)]);
    const ran = await withLoop(f, Promise.resolve({ ...started("m1"), version: ver(5), plan: [blockingLeft()] }));
    await page.runAll();
    expect(ran).toEqual(["fixAll", "recalc"]);
    expect(page.chainAsk()).toBe(second);
    expect(page.basicLines()).toEqual(ASK_LINES);

    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const g = prepFake([asks(first, 10), done(5, 1)]);
    await withLoop(g, Promise.resolve({ ...started("m1"), version: ver(5), plan: [blockingLeft()] }));
    await page.runAll();
    expect(page.chainAsk()).toBeNull();
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
    expect(page.steps()[2].line).toBe("3 loop formulas written; then 5 changes applied for what the loop uncovered. The assistant fixed 1 more: see Advanced › Versions.");
  });

  it("a question of the fixes after the loop that stops Start is the stop line, not 'Stopped at the loop'", async () => {
    const ask = question();
    const f = prepFake([done(954), asks(ask, 5)]);
    const ran = await withLoop(f, Promise.resolve({ ...started("m1"), version: ver(5), plan: [blockingLeft()] }));
    await page.runAll();
    expect(ran).toEqual(["fixAll"]); // no recalculation after it
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "failed", "waiting"]);
    expect(page.steps()[2].line).toBe("3 loop formulas written; the fixes after the loop stopped: 5 changes applied");
    expect(page.chainAsk()).toBe(ask);
    expect(page.basicLines()).toEqual(ASK_LINES);
  });

  it("the Cancel of a Prep step that stands alone stops that kind of work: prep", () => {
    open("custom");
    expect(page.cancelKind("prepare")).toBe("prep");
    expect(page.cancelKind("fixFormulas")).toBe("apply"); // the Fix the formulas tile still sends one Apply
  });

  // ---- drawn: Show me on every stop message of Basic, the question's buttons, Custom's "View details" ----
  /** The page drawn on one of the two screens, with the real "Show me" link and the real Cancel button. */
  async function drawOn(mode: "basic" | "custom", api: Partial<ApiService>): Promise<{ fixture: ComponentFixture<HomePage>; el: HTMLElement; look: () => Promise<void>; shown: ShowMeRequest[] }> {
    const none = () => Promise.reject(new Error("none"));
    TestBed.configureTestingModule({
      providers: [
        provideRouter([]),
        { provide: ActivatedRoute, useValue: { snapshot: { data: { mode } } } },
        { provide: ApiService, useValue: { estimate: none, mindLoopPlan: none, autoFixStatus: none, outsideInputs: none, nextQuarter: none, resume: none, downloadUrl: () => "/x", ...api } },
      ],
    });
    TestBed.overrideComponent(Page, { set: { imports: [CommonModule, Thinking, Bar, ShowMeLink, CancelButton], schemas: [NO_ERRORS_SCHEMA] } });
    store = TestBed.inject(SessionStore);
    // the pointer itself has its own tests (show-me.spec.ts): here, what each link asks it for
    const shown: ShowMeRequest[] = [];
    vi.spyOn(TestBed.inject(ShowMeService), "start").mockImplementation((r) => (shown.push(r), true));
    const fixture = TestBed.createComponent(Page);
    page = fixture.componentInstance;
    noWaits();
    const look = async () => {
      fixture.detectChanges();
      await settle();
      fixture.detectChanges();
    };
    page.openSaved(started("m1"));
    await look();
    return { fixture, el: fixture.nativeElement as HTMLElement, look, shown };
  }

  const linkIn = (root: Element, where: string) => root.querySelector(`${where} mr-show-me button`) as HTMLButtonElement | null;
  const buttonsIn = (root: Element | null) => Array.from(root?.querySelectorAll("button") ?? []).map((b) => b.textContent?.trim());

  const twoFiles = (): OutsideInputs => ({
    imports: true,
    macro: "calcCF",
    calls: [],
    rows: [{ cell: "Ctrl!H35", expected: null, keys: ["BBA_01"], target_sheet: null, file: "BBA_01.xlsm", how: "key", candidates: [] }],
    matched: 1,
    missing: [],
    extra: [],
    summary: "1 of 1 input files matched.",
    brought_in: false,
    brought_in_reason: null,
  });

  type StepName = "outside" | "fixAll" | "prepare" | "loop" | "recalc";
  interface Stop {
    name: string;
    step: StepName;
    /** how the step that stops ended */
    line: string;
    tone: "warn" | "error";
    unchecked?: boolean;
    /** Basic's sentence, whole */
    says: string;
    /** the place its Show me opens */
    to: string;
  }

  /** Every way Start stops by itself with a sentence of its own (`stoppedWhere`, and the outside step's amber stop). */
  const STOPS: Stop[] = [
    { name: "the outside files did not come in", step: "outside", line: "No input file matched yet", tone: "error", says: 'Stopped at "Bring in the outside files": No input file matched yet. See the Outside files card on the Custom screen, then press Start again.', to: "outside-card" },
    {
      name: "the outside files came in but the values were not checked",
      step: "outside",
      line: "The outside files are brought in, but the computed values were not checked",
      tone: "warn",
      unchecked: true,
      says: 'Stopped at "Bring in the outside files": the computed values were not checked. Check them under Advanced › Deliver (Check the numbers), or press Start again to go on without the check.',
      to: "deliver-numbers",
    },
    { name: "Fix all issues did not run", step: "fixAll", line: "Not started", tone: "error", says: 'Stopped at "Fix all issues": Not started. Press Start again.', to: "do-everything" },
    {
      name: "Fix all issues could not compare with the original",
      step: "fixAll",
      line: "The errors left could not be checked against the original (the comparison did not run)",
      tone: "warn",
      unchecked: true,
      says: 'Stopped at "Fix all issues": The errors left could not be checked against the original (the comparison did not run). Check the numbers under Advanced › Deliver, then press Start again.',
      to: "deliver-numbers",
    },
    { name: "Fix all issues left errors", step: "fixAll", line: "Fixed what it could (stuck)", tone: "warn", says: 'Stopped at "Fix all issues": Fixed what it could (stuck). Fix that under Advanced › Recalculate, then press Start again.', to: "recalc-errors" },
    { name: "the formulas and the structure failed", step: "prepare", line: "Stopped with an error", tone: "error", says: 'Stopped at "Fix the formulas and the structure": Stopped with an error. Fix that under Advanced › Prep actions, then press Start again.', to: "prep-apply" },
    { name: "the Mind loop was not written", step: "loop", line: "Could not write the loop", tone: "error", says: 'Stopped at "Build the Mind loop": Could not write the loop. Fix that under Advanced › Mind loop, then press Start again.', to: "loop-choices" },
    { name: "the recalculation found new errors", step: "recalc", line: "Recalculated, 4 formula errors", tone: "warn", says: 'Stopped at "Recalculate": Recalculated, 4 formula errors. Fix that under Advanced › Recalculate, then press Start again.', to: "recalc-errors" },
  ];

  /** Every step stands in and succeeds, but `stop.step`, which ends as `stop` says. */
  function stopsAt(stop: Stop): void {
    const p = page as unknown as Record<string, unknown>;
    const runner: Record<StepName, string> = { outside: "runOutside", fixAll: "runFixAll", prepare: "runPrepare", loop: "runLoop", recalc: "runRecalc" };
    for (const step of Object.keys(runner) as StepName[]) {
      p[runner[step]] = async () => {
        if (step !== stop.step) return true;
        page[step].set({ running: false, fraction: null, line: stop.line, detail: "", tone: stop.tone, unchecked: step === "fixAll" && !!stop.unchecked });
        if (step === "outside" && stop.unchecked) page.outsideUnchecked.set(true);
        return false;
      };
    }
    page.canBuildLoop = () => true;
  }

  it.each(STOPS)("Basic shows Show me when $name, and it opens the place by itself", async (stop) => {
    const { fixture, el, look, shown } = await drawOn("basic", {});
    if (stop.step === "outside") store.outsideInputs.set(twoFiles());
    stopsAt(stop);
    await page.runAll();
    await look();
    // one sentence, in Basic's words: its button is Start, and the Outside files card is on Custom
    expect(sentences(el)).toEqual([stop.says]);
    expect(el.textContent).not.toContain("Do everything");
    const link = linkIn(el, ".basic__next")!;
    expect(link.textContent).toBe("Show me");
    // under the line, outside the part a screen reader announces as the status
    expect(el.querySelector(".basic__line mr-show-me")).toBeNull();
    expect(el.querySelector(".basic__line")!.compareDocumentPosition(link) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    link.click();
    expect(shown).toEqual([{ id: stop.to, open: true }]);
    expect(el.querySelectorAll("mr-show-me button")).toHaveLength(1);
    expect(el.querySelector('[data-test="question"]')).toBeNull(); // an ordinary stop: no question, no Send feedback
    fixture.destroy();
  });

  it("no stop sentence on Basic starts with a small letter, and none names a button Basic does not have", async () => {
    for (const stop of STOPS) {
      open("basic");
      page.openSaved(started("m1"));
      if (stop.step === "outside") store.outsideInputs.set(twoFiles());
      stopsAt(stop);
      await page.runAll();
      expect(page.basicLine(), stop.name).toBe(stop.says);
      // each sentence: after a full stop and a space comes a capital, a digit or a quote, never a small letter
      expect(page.basicLine(), stop.name).not.toMatch(/[.!?]\s+[a-z]/);
      expect(page.basicLine(), stop.name).toMatch(/^[A-Z]/);
      expect(page.basicLine(), stop.name).not.toContain("Do everything");
      page.ngOnDestroy();
      TestBed.resetTestingModule();
    }
  });

  it("the outside step's sentence names the card where it is: on the Custom screen from Basic, above on Custom", async () => {
    const stop = STOPS[0];
    open("custom");
    page.openSaved(started("m1"));
    store.outsideInputs.set(twoFiles());
    stopsAt(stop);
    await page.runAll();
    expect(page.chainLine()).toBe('Stopped at "Bring in the outside files": No input file matched yet. See the Outside files card above, then press Do everything again.');
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    open("basic");
    page.openSaved(started("m1"));
    store.outsideInputs.set(twoFiles());
    stopsAt(stop);
    await page.runAll();
    expect(page.basicLine()).toBe(stop.says);
    expect(showMeFor(page.chainLine())).toEqual({ id: "outside-card" }); // the card, on Custom: the pointer goes there
  });

  it("Basic draws the question one sentence per line, with Show me to the problem's cell and Send feedback with the note in it", async () => {
    const f = prepFake([round(1, 954), asks(question())]);
    const { fixture, el, look, shown } = await drawOn("basic", f.api);
    others();
    await page.runAll();
    await look();
    expect(sentences(el)).toEqual(ASK_LINES);
    const row = el.querySelector('.basic__said [data-test="question"]')!;
    expect(buttonsIn(row)).toEqual(["Show me", "Send feedback"]); // why the assistant failed is not for Basic: no "View details"
    expect(el.textContent).not.toContain(WHY);
    expect(el.textContent).not.toContain("STR-008"); // no rule code on screen
    // Show me: the Prep action that repairs the rule, the cell's own line, and the pointer goes there itself
    linkIn(el, '[data-test="question"]')!.click();
    expect(shown).toEqual([{ id: "prep-action", ruleId: "STR-008", where: { sheet: "Roll", cell: "W25" }, open: true }]);
    // Send feedback: the dialog opens for the incident the server wrote, with the problem already in the note box
    const feedback = TestBed.inject(FeedbackService);
    expect(feedback.dialog()).toBeNull();
    (Array.from(row.querySelectorAll("button")).find((b) => b.textContent?.trim() === "Send feedback") as HTMLButtonElement).click();
    expect(feedback.dialog()).toEqual({ ref: "ABCD-1234", note: NOTE });
    feedback.close();
    expect(el.querySelectorAll("mr-show-me button")).toHaveLength(1); // one Show me: the question's own
    fixture.destroy();
  });

  it("the question's Send feedback opens the real dialog with the problem in its note box, and the email starts with it", async () => {
    // the dialog drawn beside the page, as the shell holds both: one FeedbackService between them
    const sent: unknown[] = [];
    const f = prepFake([asks(question())]);
    f.api.assistantKey = () => Promise.resolve({ saved: false, last4: null });
    f.api.feedback = (body) => (sent.push(body), Promise.resolve({ ref: "ABCD-1234", text: "Mind Ready 1.15.0 · ref ABCD-1234\n", subject: "Mind Ready feedback · ref ABCD-1234", email: "someone@example.com", file_name: "f.txt", assistant: false }));
    TestBed.overrideComponent(Dialog, { set: { imports: [], schemas: [NO_ERRORS_SCHEMA] } });
    const { fixture, el, look } = await drawOn("basic", f.api);
    const box = TestBed.createComponent(Dialog);
    const boxEl = box.nativeElement as HTMLElement;
    const lookBox = async () => {
      box.detectChanges();
      await settle();
      box.detectChanges();
    };
    await lookBox();
    expect(boxEl.querySelector("#fb-note")).toBeNull(); // closed: nothing drawn
    others();
    await page.runAll();
    await look();
    (el.querySelector('[data-test="question"] .chain__feedback') as HTMLButtonElement).click();
    await lookBox();
    // the problem is already written, in plain words: no rule code, no version, no sheet or cell name
    expect(boxEl.querySelector<HTMLTextAreaElement>("#fb-note")?.value).toBe(NOTE);
    expect(NOTE).not.toMatch(/[A-Z]{2,5}-\d{3}|ver-\d|W25|Roll/);
    await box.componentInstance.prepare();
    await lookBox();
    expect(sent).toEqual([{ ref: "ABCD-1234", note: NOTE, useAssistant: false, sessionId: "m1" }]); // the report the server wrote for the question
    Object.defineProperty(navigator, "clipboard", { value: { writeText: () => Promise.resolve() }, configurable: true });
    const mails: string[] = [];
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
      mails.push(this.href);
    });
    await box.componentInstance.send();
    expect(mails).toHaveLength(1);
    const body = decodeURIComponent(mails[0].split("&body=")[1]);
    expect(body.startsWith(NOTE)).toBe(true);
    expect(body).toBe(`${NOTE}\r\n\r\nPaste the Mind Ready message here (Ctrl+V), then send.`);
    box.destroy();
    fixture.destroy();
  });

  it("Basic draws 'no AI key' with Show me to the key box, and no Send feedback", async () => {
    const f = prepFake([{ state: "no_key", stage: "done", applied: 3, left: 1, ask: noKey() }]);
    const { fixture, el, look, shown } = await drawOn("basic", f.api);
    others();
    const still: Readiness = { ...readiness("blocked", 1), blocking: [{ rule_id: "FRM-009", status: "ERROR", message: "m", fix: "assistant" }] };
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(still), true);
    await page.runAll();
    await look();
    expect(sentences(el)).toEqual(NO_KEY_LINES);
    expect(buttonsIn(el.querySelector('[data-test="question"]'))).toEqual(["Show me"]);
    linkIn(el, '[data-test="question"]')!.click();
    expect(shown).toEqual([{ id: "assistant-key", open: true }]);
    fixture.destroy();
  });

  it("Basic draws the question kept for the end of Start the same way, no Show me when everything ran, and after a Cancel one to the step that was stopped", async () => {
    const ask = question({ stops_start: false, show_me: { id: "assistant-ask", rule_id: "FRM-009", sheet: "Calc", cell: "B2" }, feedback: { ref: null, note: NOTE } });
    const f = prepFake([asks(ask, 12)]);
    const { fixture, el, look, shown } = await drawOn("basic", f.api);
    others();
    await page.runAll();
    await look();
    expect(sentences(el)).toEqual(ASK_LINES);
    expect(buttonsIn(el.querySelector('[data-test="question"]'))).toEqual(["Show me", "Send feedback"]);
    linkIn(el, '[data-test="question"]')!.click();
    expect(shown).toEqual([{ id: "assistant-ask", ruleId: "FRM-009", where: { sheet: "Calc", cell: "B2" }, open: true }]);
    (el.querySelector('[data-test="question"] .chain__feedback') as HTMLButtonElement).click();
    expect(TestBed.inject(FeedbackService).dialog()).toEqual({ ref: null, note: NOTE }); // a question that came with no incident
    TestBed.inject(FeedbackService).close();

    // Start again, and this time nothing is left: one calm sentence, nothing to show
    f.looks.push(done(0));
    f.calls.length = 0;
    await page.runAll();
    await look();
    expect(sentences(el)[0]).toMatch(/^Everything ran and every check passed/);
    expect(el.querySelector(".basic__next")).toBeNull();
    expect(el.querySelector('[data-test="question"]')).toBeNull();
    expect(el.querySelector(".basic__said mr-show-me")).toBeNull();

    // and after a Cancel: the person stopped it, there is no place to fix
    const end = hold<PrepRunStatus>();
    f.looks.push(end.done);
    const run = page.runAll();
    await until(() => f.seen.length > 0 && page.chainStep()?.id === "prepare");
    page.onChainStopping();
    end.end({ state: "cancelled" });
    page.onChainStopped();
    await run;
    await look();
    expect(sentences(el)).toEqual(["Stopped. Your workbook is as it was after the formula errors were fixed."]);
    // 19 Nov 2026 (Joseph): one "Show me", to the step that was stopped ("Fix the formulas and the structure": the Prep actions)
    expect(buttonsIn(el.querySelector('[data-test="cancelled-show-me"]'))).toEqual(["Show me"]);
    expect(el.querySelectorAll(".basic__said mr-show-me")).toHaveLength(1);
    linkIn(el, '[data-test="cancelled-show-me"]')!.click();
    expect(shown[shown.length - 1]).toEqual({ id: "prep-apply", open: true });
    expect(el.querySelector('[data-test="question"]')).toBeNull();
    fixture.destroy();
  });

  it("Custom draws the same question under its chain, with why under View details", async () => {
    const ask = question({ lines: [ASK_LINES[0], "Fix it by hand, then press Do everything again."] });
    const f = prepFake([round(1, 954), asks(ask)]);
    const { fixture, el, look, shown } = await drawOn("custom", f.api);
    others();
    await page.runAll();
    await look();
    const text = (selector: string) => Array.from(el.querySelectorAll(selector)).map((x) => (x.textContent ?? "").replace(/\s+/g, " ").trim());
    expect(text(".chain > .chain__text strong")).toEqual([`${ASK_LINES[0]} Fix it by hand, then press Do everything again.`]);
    // the question has its own Show me: the sentence's plain one is not drawn beside it
    expect(el.querySelector(".chain > .chain__text mr-show-me")).toBeNull();
    const row = el.querySelector('.chain > [data-test="question"]')!;
    expect(buttonsIn(row)).toEqual(["Show me", "Send feedback", "View details"]);
    expect(el.querySelector(".chain__why")).toBeNull();
    (Array.from(row.querySelectorAll("button")).find((b) => b.textContent?.trim() === "View details") as HTMLButtonElement).click();
    await look();
    expect(text(".chain__why")).toEqual([WHY]);
    expect(buttonsIn(row)).toEqual(["Show me", "Send feedback", "Hide details"]);
    linkIn(el, '.chain > [data-test="question"]')!.click();
    expect(shown).toEqual([{ id: "prep-action", ruleId: "STR-008", where: { sheet: "Roll", cell: "W25" }, open: true }]);
    (row.querySelector(".chain__feedback") as HTMLButtonElement).click();
    expect(TestBed.inject(FeedbackService).dialog()).toEqual({ ref: "ABCD-1234", note: NOTE });
    TestBed.inject(FeedbackService).close();
    // the step in the list: what it wrote, in amber or red, and no question of its own
    expect(text(".chain__step--failed .chain__line")[0]).toContain("954 changes applied");
    // the next run starts clean: the question and its details go
    f.looks.push(done(2));
    await page.runAll();
    await look();
    expect(el.querySelector('[data-test="question"]')).toBeNull();
    expect(el.querySelector(".chain__why")).toBeNull();
    fixture.destroy();
  });

  it("Custom draws a Prep job picked up after a reload alone in the list, with its own Cancel, which stops the kind prep", async () => {
    const end = hold<PrepRunStatus>();
    const f = prepFake([round(2), end.done]);
    f.pickUp = true;
    const { fixture, el, look } = await drawOn("custom", f.api);
    await until(() => page.prepare().running);
    await look();
    // picked up by the page itself when the workbook opened: nobody pressed anything
    expect(f.calls).not.toContain("start");
    expect(Array.from(el.querySelectorAll(".chain__step")).map((s) => s.querySelector(".chain__label")?.textContent?.trim().split(/\s+check:/)[0])).toEqual(["Fix the formulas and the structure"]);
    const head = el.querySelector(".chain__buttons")!;
    expect(head.querySelector("mr-cancel")).toBeNull(); // Do everything is not running: its own Cancel is not there
    const cancel = el.querySelector(".chain__steps .action__progress mr-cancel button") as HTMLButtonElement;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    expect(el.querySelectorAll("mr-cancel")).toHaveLength(1);
    cancel.click();
    await look();
    expect(f.calls).toContain("cancel:prep");
    end.end({ state: "cancelled" });
    await until(() => !page.prepare().running);
    await look();
    expect(el.querySelector("mr-cancel")).toBeNull();
    expect(page.steps()[0]).toMatchObject({ state: "stopped", line: STOPPED_NOTHING_CHANGED });
    fixture.destroy();
  });

  // ---- 17 Nov 2026: the morning's review and live tests (screens) -------------------------------------
  const blockedBy = (n: number, fix: "assistant" | "prep_skipped" = "assistant"): Readiness => ({
    ...readiness("blocked", n),
    blocking: [{ rule_id: "FRM-011", rule_name: "a formula compares an empty cell with a number", status: "ERROR", message: "m", location: { sheet: "Calc", cell: "B2" }, fix }],
  });
  const running = (kind: string, title: string): WorkJob => ({ kind, title, started: 1_763_330_000, elapsed_s: 3, stopping: false });

  it("review R2-3: a blocking problem the loop made that no Prep action repairs goes to the assistant too: the Prep job runs again", async () => {
    const f = prepFake([done(954), done(0, 1)]);
    f.after = { ...f.after, readiness: readiness("ready") };
    // after the loop nothing is left for Prep to write, but the readiness lists a problem only the assistant can fix
    await withLoop(f, Promise.resolve({ ...started("m1"), version: ver(5), plan: [], readiness: blockedBy(1) }));
    await page.runAll();
    expect(f.calls.filter((c) => c === "start")).toHaveLength(2);
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.steps()[2].line).toBe("3 loop formulas written. The assistant fixed 1: see Advanced › Versions.");
    // only the outside files left: those are the outside step's, not the assistant's
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const g = prepFake([done(954)]);
    const outside: Readiness = { ...readiness("blocked", 1), blocking: [{ rule_id: "INP-001", status: "ERROR", message: "m", fix: "outside_inputs" }] };
    await withLoop(g, Promise.resolve({ ...started("m1"), version: ver(5), plan: [], readiness: outside }));
    await page.runAll();
    expect(g.calls.filter((c) => c === "start")).toHaveLength(1);
  });

  it("review R2-3: every step ran but the workbook is still blocked: it says how many problems are left, with Show me to the first, never 'every check passed'", async () => {
    const f = prepFake([done(3)]);
    const { fixture, el, look, shown } = await drawOn("basic", f.api);
    others();
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(blockedBy(2)), true);
    await page.runAll();
    await look();
    expect(sentences(el)).toEqual(["Everything ran, but 2 blocking problems are left. The first: a formula compares an empty cell with a number."]);
    expect(page.basicLine()).not.toContain("every check passed");
    linkIn(el, ".basic__next")!.click();
    expect(shown).toEqual([{ id: "assistant-ask", ruleId: "FRM-011", where: { sheet: "Calc", cell: "B2" }, open: true }]);
    fixture.destroy();
    TestBed.resetTestingModule();
    // one problem; a cell Prep skipped leads to its Prep action (its "By hand" entry)
    open("custom");
    const r = blockedBy(1, "prep_skipped");
    page.openSaved(started("m1"));
    await settle();
    stub();
    (page as unknown as Record<string, unknown>)["runRecalc"] = async () => (store.readiness.set(r), true);
    await page.runAll();
    expect(page.chainLine()).toBe("Everything ran, but 1 blocking problem is left: a formula compares an empty cell with a number.");
    expect(page.stopPlace()).toEqual({ id: "prep-action", ruleId: "FRM-011", where: { sheet: "Calc", cell: "B2" }, open: true });
  });

  it("live test L2b-9: on Basic the assistant's own fixes are said under 'Everything ran', with Show me to Versions; Custom says it on the step", async () => {
    const f = prepFake([done(0, 1)]);
    const { fixture, el, look, shown } = await drawOn("basic", f.api);
    others();
    await page.runAll();
    await look();
    expect(sentences(el)).toEqual([expect.stringMatching(/^Everything ran and every check passed \(/), "The assistant fixed 1: see Advanced › Versions."]);
    linkIn(el, ".basic__next")!.click();
    expect(shown).toEqual([{ id: "versions", open: true }]);
    // the next run: the assistant fixed nothing, so the sentence goes
    f.looks.push(done(5));
    await page.runAll();
    await look();
    expect(sentences(el)).toHaveLength(1);
    fixture.destroy();
    TestBed.resetTestingModule();
    // after changes of the Prep rounds: "more", as on Custom's step
    const g = prepFake([done(954, 2)]);
    await openedFor("basic", g);
    others();
    await page.runAll();
    expect(page.basicLines()[1]).toBe("The assistant fixed 2 more: see Advanced › Versions.");
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const h = prepFake([done(954, 2)]);
    await openedFor("custom", h);
    others();
    await page.runAll();
    expect(page.chainNote()).toBe("");
    expect(page.steps()[1].line).toBe("954 changes applied. The assistant fixed 2 more: see Advanced › Versions.");
  });

  it("live test L2b-6: the step count does not change in the middle of a run", async () => {
    open("basic");
    page.openSaved(started("m1"));
    await settle();
    // a workbook with no loop: three steps from the first one on (it said "Step 2 of 4", then "Step 3 of 3")
    let gates = held();
    page.canBuildLoop = () => false;
    let run = page.runAll();
    await settle();
    expect(page.basicLines()[0]).toMatch(/^Step 1 of 3: fixing/);
    expect(page.steps().find((s) => s.id === "loop")).toMatchObject({ state: "skipped", line: "No run loop in this workbook" });
    await endStep(gates, "fixAll");
    expect(page.basicLines()[0]).toMatch(/^Step 2 of 3: changing/);
    await endStep(gates, "prepare");
    expect(page.basicLines()[0]).toMatch(/^Step 3 of 3: recalculating/);
    await endStep(gates, "recalc");
    await run;
    // a loop that can no longer be built when its turn comes keeps its number
    gates = held();
    let buildable = true;
    page.canBuildLoop = () => buildable;
    run = page.runAll();
    await settle();
    expect(page.basicLines()[0]).toMatch(/^Step 1 of 4: /);
    buildable = false;
    await endStep(gates, "fixAll");
    expect(page.basicLines()[0]).toMatch(/^Step 2 of 4: /);
    await endStep(gates, "prepare");
    expect(page.basicLines()[0]).toMatch(/^Step 4 of 4: recalculating/);
    await endStep(gates, "recalc");
    await run;
    // the box not ticked: left out from the start
    gates = held();
    page.buildLoop.set(false);
    run = page.runAll();
    await settle();
    expect(page.basicLines()[0]).toMatch(/^Step 1 of 3: /);
    for (const step of ["fixAll", "prepare", "recalc"] as const) await endStep(gates, step);
    await run;
  });

  it("live test L2b-5: a step is late only when its own time is over, and never in its first minute", async () => {
    clockAt(T0);
    const e = estimate();
    e.steps = e.steps.map((s) => (s.id === "loop" ? { ...s, low_s: 2, mid_s: 3, high_s: 4 } : s)); // one round through Excel
    open("basic", { estimate: async () => e });
    page.openSaved(started("m1"));
    await settle();
    const gates = held();
    const run = page.runAll();
    await settle();
    for (const step of ["fixAll", "prepare"] as const) await endStep(gates, step);
    expect(page.basicLines()[0]).toMatch(/^Step 3 of 4: building the Mind loop/);
    later(5); // it said "taking longer than usual" here
    expect(page.basicLine()).not.toContain("longer than usual");
    later(50);
    expect(page.basicLine()).not.toContain("longer than usual");
    later(10);
    expect(page.basicLines()[1]).toMatch(/^This step is taking longer than usual; /);
    await endStep(gates, "loop");
    // a step estimated as nothing (as if skipped) has no time of its own: never late
    (page as unknown as { estimate: { set(v: ChainEstimate): void } }).estimate.set({ ...e, steps: e.steps.map((s) => (s.id === "recalc" ? { ...s, low_s: 0, mid_s: 0, high_s: 0 } : s)) });
    later(600);
    expect(page.basicLine()).not.toContain("longer than usual");
    await endStep(gates, "recalc");
    await run;
  });

  it("live test L2b-2: Start is greyed with the reason while other work of this workbook runs", async () => {
    const { fixture, el, look } = await drawOn("basic", {});
    const work = TestBed.inject(WorkService);
    const go = () => el.querySelector(".basic__go") as HTMLButtonElement;
    expect(go().disabled).toBe(false);
    work.jobs.set([running("recalc", "Recalculating")]); // Recalculate, pressed under Advanced
    await look();
    expect(go().disabled).toBe(true);
    expect(go().getAttribute("title")).toBe("Still working: Recalculating. Press Start once it has finished.");
    expect(el.querySelector('[data-test="start-wait"]')?.textContent?.trim()).toBe("Still working: Recalculating. Press Start once it has finished.");
    work.jobs.set([]);
    await look();
    expect(go().disabled).toBe(false);
    expect(el.querySelector('[data-test="start-wait"]')).toBeNull();
    // while Start waits for the input files after a reload, the bring-in's own wait is not "other work"
    page.outsideWaiting.set(true);
    work.jobs.set([running("outside", "Bringing in the input files")]);
    expect(page.startBlocked()).toBe(false);
    page.outsideWaiting.set(false);
    work.jobs.set([]);
    fixture.destroy();
  });

  it("live test L2b-2: a step the server would not start says the server's sentence, not 'Not started.'", async () => {
    open("basic", { startAutoFix: () => Promise.reject({ status: 409, error: { detail: "a recalculation is running for this workbook: wait for it to finish" } }) });
    page.openSaved(started("m1"));
    await settle();
    page.canBuildLoop = () => true;
    await page.runAll();
    expect(page.basicLine()).toBe("A recalculation is running for this workbook: wait for it to finish. Press Start again once that has finished.");
    expect(page.basicLine()).not.toContain("Not started");
    expect(page.basicLine()).not.toContain("Fix all issues");
    // and the server's "Wait a moment: ..." of any step, on Custom in its own words
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    open("custom");
    page.openSaved(started("m1"));
    await settle();
    stub();
    (page as unknown as Record<string, unknown>)["runPrepare"] = async () => (page.prepare.set({ running: false, fraction: null, line: "Not started", detail: "Wait a moment: changes are still being written.", tone: "error" }), false);
    await page.runAll();
    expect(page.chainLine()).toBe("Wait a moment: changes are still being written. Press Do everything again once that has finished.");
  });

  it("live test L2b-3: when Start ends, its steps leave the list of running work before its Cancel leaves the screen", async () => {
    await opened();
    stub();
    const work = TestBed.inject(WorkService);
    const run = page.runAll();
    work.jobs.set([running("recalc", "Recalculating"), running("mind", "Uploading to Mind")]); // the last look while Start ran
    await run;
    expect(work.jobs().map((j) => j.kind)).toEqual(["mind"]); // a Run in Mind started elsewhere still shows in the strip
  });

  it("live test L2b-4: after a reload in the middle of Start, one line says the rest of Start did not run", async () => {
    sessionStorage.removeItem("mindReady.startRunning");
    // Start sets the mark while it runs, and takes it away when it ends here
    await opened();
    stub();
    const run = page.runAll();
    expect(JSON.parse(sessionStorage.getItem("mindReady.startRunning") ?? "null")).toMatchObject({ id: "m1", left: 0 });
    await run;
    expect(sessionStorage.getItem("mindReady.startRunning")).toBeNull();
    page.ngOnDestroy();
    TestBed.resetTestingModule();

    // the page was reloaded while the Prep step ran: the step still runs on the server
    sessionStorage.setItem("mindReady.startRunning", "m1");
    let jobs: WorkJob[] = [running("prep", "Fixing the formulas and the structure")];
    open("basic", { jobs: async () => jobs });
    page.openSaved(started("m1"));
    await (page as unknown as { checkInterrupted: () => Promise<void> }).checkInterrupted();
    expect(page.basicLines()).toEqual(["Start was interrupted when the page was reloaded or left. Press Start again when this step ends."]);
    expect(page.stopPlace()?.id).toBe("do-everything"); // Show me: the Start button
    // the step ended (the shell's list no longer has it): press Start to finish; the tab forgets it
    const work = TestBed.inject(WorkService);
    work.jobs.set([]);
    TestBed.tick();
    expect(page.basicLines()).toEqual(["Start was interrupted when the page was reloaded or left. Press Start again to finish."]);
    expect(sessionStorage.getItem("mindReady.startRunning")).toBeNull();
    // a later reload says nothing; nor does a workbook Start never ran on
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    jobs = [];
    open("basic", { jobs: async () => jobs });
    page.openSaved(started("m1"));
    await (page as unknown as { checkInterrupted: () => Promise<void> }).checkInterrupted();
    expect(page.basicLines()).toEqual([]);
    sessionStorage.setItem("mindReady.startRunning", "other");
    await (page as unknown as { checkInterrupted: () => Promise<void> }).checkInterrupted();
    expect(page.basicLines()).toEqual([]);
    sessionStorage.removeItem("mindReady.startRunning");
  });

  // ---- 17 Nov 2026, repair round: what the check of the first fixes found ----------------------------------
  const MARK = "mindReady.startRunning";
  const checkInterrupted = () => (page as unknown as { checkInterrupted: () => Promise<void> }).checkInterrupted();
  const job = (kind: string, title: string, startedMs: number): WorkJob => ({ kind, title, started: startedMs / 1000, elapsed_s: 3, stopping: false });

  it("repair round (L2b-4): a page left or reloaded during Start's last step leaves no 'interrupted' line; one left before it does", async () => {
    sessionStorage.removeItem(MARK);
    await opened();
    let gates = held();
    let run = page.runAll();
    await settle();
    for (const step of ["fixAll", "prepare", "loop"] as const) await endStep(gates, step);
    expect(page.chainStep()?.id).toBe("recalc");
    expect(JSON.parse(sessionStorage.getItem(MARK) ?? "null")).toMatchObject({ id: "m1", kinds: ["recalc"], last: true });
    page.ngOnDestroy(); // left for Advanced while Recalculate runs: nothing of Start is left once it ends
    expect(sessionStorage.getItem(MARK)).toBeNull();
    await endStep(gates, "recalc");
    await run;
    TestBed.resetTestingModule();
    open("basic", { jobs: async () => [] });
    page.openSaved(started("m1"));
    await checkInterrupted();
    expect(page.basicLines()).toEqual([]); // it said "Start was interrupted ... Press Start again to finish."
    page.ngOnDestroy();
    TestBed.resetTestingModule();

    // a reload (no ngOnDestroy, only "pagehide") in the last step: the same
    await opened();
    gates = held();
    run = page.runAll();
    await settle();
    for (const step of ["fixAll", "prepare", "loop"] as const) await endStep(gates, step);
    window.dispatchEvent(new Event("pagehide"));
    expect(sessionStorage.getItem(MARK)).toBeNull();
    const reloaded = page;
    TestBed.resetTestingModule();
    open("basic", { jobs: async () => [job("recalc", "Recalculating", Date.now() - 5000)] });
    page.openSaved(started("m1"));
    await checkInterrupted();
    expect(page.basicLines()).toEqual([]);
    reloaded.ngOnDestroy();
    await endStep(gates, "recalc");
    await run;
    page.ngOnDestroy();
    TestBed.resetTestingModule();

    // left in an earlier step: the rest of Start did not run, and the next page says so
    await opened();
    gates = held();
    run = page.runAll();
    await settle();
    await endStep(gates, "fixAll");
    page.ngOnDestroy();
    expect(JSON.parse(sessionStorage.getItem(MARK) ?? "null")).toMatchObject({ id: "m1", kinds: ["prep"], last: false });
    await endStep(gates, "prepare");
    await run;
    TestBed.resetTestingModule();
    open("basic", { jobs: async () => [] });
    page.openSaved(started("m1"));
    await checkInterrupted();
    expect(page.basicLines()).toEqual(["Start was interrupted when the page was reloaded or left. Press Start again to finish."]);
    page.ngOnDestroy();
    TestBed.resetTestingModule();

    // a mark of the last step whose leaving was never heard (the tab was ended): nothing is said either
    sessionStorage.setItem(MARK, JSON.stringify({ id: "m1", kinds: ["recalc"], last: true, left: 0 }));
    open("basic", { jobs: async () => [] });
    page.openSaved(started("m1"));
    await checkInterrupted();
    expect(page.basicLines()).toEqual([]);
    expect(sessionStorage.getItem(MARK)).toBeNull();
  });

  it("repair round (L2b-4): after Start was left, only the step that ran then is Start's: a Recalculate or a Prep run pressed under Advanced afterwards is not", async () => {
    sessionStorage.removeItem(MARK);
    clockAt(T0);
    await opened();
    const work = TestBed.inject(WorkService);
    const gates = held();
    const run = page.runAll();
    await settle();
    await endStep(gates, "fixAll");
    expect(page.chainStep()?.id).toBe("prepare");
    work.jobs.set([job("prep", "Fixing the formulas and the structure", T0 + 1000)]); // the shell's last look
    later(30);
    page.ngOnDestroy(); // left for Advanced, 30 s into the Prep step
    expect(JSON.parse(sessionStorage.getItem(MARK) ?? "null")).toEqual({ id: "m1", kinds: ["prep"], last: false, left: (T0 + 30_000) / 1000 });
    gates.prepare.end(true); // the page stopped following the job: it may still run on the server
    await run;
    expect(work.jobs().map((j) => j.kind)).toEqual(["prep"]); // the list keeps it (it was dropped, and the strip went)
    TestBed.resetTestingModule();

    // under Advanced, once the Prep step ended: Recalculate pressed at T0 + 60 s
    let jobs: WorkJob[] = [job("recalc", "Recalculating", T0 + 60_000)];
    open("basic", { jobs: async () => jobs });
    page.openSaved(started("m1"));
    await checkInterrupted();
    expect(page.basicLines()).toEqual(["Start was interrupted when the page was reloaded or left. Press Start again to finish."]);
    expect(sessionStorage.getItem(MARK)).toBeNull(); // said once
    page.ngOnDestroy();
    TestBed.resetTestingModule();

    // a Prep run pressed under Advanced after the page was left is not Start's either; Start's own one is
    for (const [startedAt, line] of [
      [T0 + 60_000, "Start was interrupted when the page was reloaded or left. Press Start again to finish."],
      [T0 + 1000, "Start was interrupted when the page was reloaded or left. Press Start again when this step ends."],
    ] as const) {
      sessionStorage.setItem(MARK, JSON.stringify({ id: "m1", kinds: ["prep"], last: false, left: (T0 + 30_000) / 1000 }));
      jobs = [job("prep", "Fixing the formulas and the structure", startedAt)];
      open("basic", { jobs: async () => jobs });
      page.openSaved(started("m1"));
      await checkInterrupted();
      expect(page.basicLines()).toEqual([line]);
      // the shell's next look: Advanced's Recalculate is no step of that Start
      if (startedAt === T0 + 1000) {
        TestBed.inject(WorkService).jobs.set([job("recalc", "Recalculating", T0 + 90_000)]);
        TestBed.tick();
        expect(page.basicLines()).toEqual(["Start was interrupted when the page was reloaded or left. Press Start again to finish."]);
      }
      page.ngOnDestroy();
      TestBed.resetTestingModule();
    }
    sessionStorage.removeItem(MARK);
  });

  it("repair round (L2b-2): an assistant's answer or a Run in Mind does not grey Start; work Start cannot begin beside does", () => {
    open("basic");
    page.openSaved(started("m1"));
    const work = TestBed.inject(WorkService);
    for (const [kind, title] of [["chat", "The assistant is answering"], ["mind", "Uploading to Mind"], ["numbers", "Comparing the numbers"], ["finder", "Looking for instances"]]) {
      work.jobs.set([running(kind, title)]);
      expect(page.startBlocked()).toBe(false);
      expect(page.startHint()).toBe("");
    }
    for (const [kind, title] of [["recalc", "Recalculating"], ["apply", "Writing the changes"], ["names", "Suggesting names"], ["next_quarter", "Making next quarter's input files"]]) {
      work.jobs.set([running(kind, title)]);
      expect(page.startBlocked()).toBe(true);
      expect(page.startHint()).toBe(`Still working: ${title}. Press Start once it has finished.`);
    }
    // both at once: the one that holds Start back is named
    work.jobs.set([running("mind", "Uploading to Mind"), running("recalc", "Recalculating")]);
    expect(page.startHint()).toBe("Still working: Recalculating. Press Start once it has finished.");
    work.jobs.set([]);
  });

  it("repair round (R2-3): a problem the Prep step's assistant already tried does not start the Prep job again after the loop; a new one does", async () => {
    const ask = question({ stops_start: false, show_me: { id: "assistant-ask", rule_id: "FRM-011", sheet: "Calc", cell: "B2" } });
    const f = prepFake([asks(ask, 10), done(0)]);
    f.after = { ...f.after, readiness: blockedBy(1) };
    // after the loop the same problem is still there, and nothing new
    await withLoop(f, Promise.resolve({ ...started("m1"), version: ver(5), plan: [], readiness: blockedBy(1) }));
    await page.runAll();
    expect(f.calls.filter((c) => c === "start")).toHaveLength(1); // it was 2: a whole Prep job for no new answer
    expect(page.steps().map((s) => s.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.chainAsk()).toBe(ask); // Start asks about it at the end
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    // the loop made another problem as well: the job runs again for it
    const g = prepFake([asks(ask, 10), done(0, 1)]);
    g.after = { ...g.after, readiness: blockedBy(1) };
    const two: Readiness = { ...blockedBy(2), blocking: [...blockedBy(1).blocking, { rule_id: "FRM-011", rule_name: "r", status: "ERROR", message: "m", location: { sheet: "Calc", cell: "B9" }, fix: "assistant" }] };
    await withLoop(g, Promise.resolve({ ...started("m1"), version: ver(5), plan: [], readiness: two }));
    await page.runAll();
    expect(g.calls.filter((c) => c === "start")).toHaveLength(2);
  });

  it("repair round (L2b-6): Start pressed while the loop plan is still being read counts the loop step, and the count stays", async () => {
    for (const found of [true, false]) {
      const read = hold<{ version_id: string; plan: LoopPlan }>();
      open("basic", { mindLoopPlan: () => read.done });
      page.openSaved(started("m1"));
      await settle();
      const gates = held();
      page.canBuildLoop = () => !!page.loopPlan()?.found;
      const run = page.runAll();
      await settle();
      expect(page.basicLines()[0]).toMatch(/^Step 1 of 4: /); // it said "Step 1 of 3", then "Step 3 of 4" once the plan came
      await endStep(gates, "fixAll");
      read.end({ version_id: "ver-002", plan: { ...loopPlan(), found } });
      await settle();
      await endStep(gates, "prepare");
      if (found) {
        expect(page.basicLines()[0]).toMatch(/^Step 3 of 4: building the Mind loop/);
        await endStep(gates, "loop");
      }
      expect(page.basicLines()[0]).toMatch(/^Step 4 of 4: recalculating/);
      await endStep(gates, "recalc");
      await run;
      page.ngOnDestroy();
      TestBed.resetTestingModule();
    }
  });

  it("review R2-5: Basic says Start where Custom says Do everything: the waiting box and the hint on a greyed button", () => {
    open("basic");
    store.outsideInputs.set({ ...twoFiles(), rows: [{ ...twoFiles().rows[0], file: null }], missing: ["BBA_01"], matched: 0 });
    page.chainRunning.set(true);
    page.outsideWaiting.set(true);
    expect(page.waitingLine()).toBe("Start is waiting for 1 input file: BBA_01. Pick it.");
    expect(page.busyHint()).toBe("Start is waiting for the input files");
    page.outsideWaiting.set(false);
    expect(page.busyHint()).toBe("Start is running");
    page.chainRunning.set(false);
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    // Custom keeps its own words
    open("custom");
    page.chainRunning.set(true);
    expect(page.busyHint()).toBe("Do everything is running");
    page.chainRunning.set(false);
  });

  it("review R2-4: the thinking dots go beside the words of the assistant at work", () => {
    open("basic");
    expect(page.aiWords("Making the missing input files with the assistant…")).toBe(true);
    expect(page.aiWords("Make them with AI")).toBe(true);
    expect(page.aiWords("Waiting for 2 input files: BBA_01, BBA_02")).toBe(false);
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
