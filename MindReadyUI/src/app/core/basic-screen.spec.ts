import { TestBed } from "@angular/core/testing";
import { ActivatedRoute, provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import type { RecentWorkbooks } from "../pages/recent-workbooks";
import { ApiService } from "./api.service";
import { finishState } from "./finish-banner";
import { Readiness, SavedWorkbook, SessionStart } from "./models";
import { planShowMe, PAGE_LINKS, showMeFor } from "./show-me";
import { SessionStore } from "./session.store";

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
    expect(showMeFor("Open it again from Recent workbooks on the main screen.")?.id).toBe("recent-workbooks");
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
    expect(s?.reason).toContain("Press Go");
    expect(finishState(null)).toBeNull();
  });
});

describe("the Basic screen", () => {
  let Page: typeof HomePage;
  let Banner: typeof import("./finish-banner").FinishBanner;
  let page: HomePage;
  let store: SessionStore;

  beforeAll(async () => {
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
    Banner = (await import("./finish-banner")).FinishBanner;
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
    expect(page.basicLine()).toContain("press Go again.");
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
