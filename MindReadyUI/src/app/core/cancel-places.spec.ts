import { CommonModule } from "@angular/common";
import { NO_ERRORS_SCHEMA } from "@angular/core";
import { ComponentFixture, TestBed } from "@angular/core/testing";
import { ActivatedRoute, provideRouter } from "@angular/router";
import type { App } from "../app";
import type { HomePage } from "../pages/home.page";
import { ApiService } from "./api.service";
import { CancelButton } from "./cancel-button";
import { FinishBanner } from "./finish-banner";
import {
  ApplyOutcome, AutoFixStatus, CancelResult, LoopPlan, MindDelete, MindLoopStatus, MindProject, MindSignIn, NextQuarterStatus, OutsideInputs, OutsideRunStatus, PrepAction, PrepRunStatus, Readiness, RecalcResult, SessionStart, Version, WorkJob,
  appliedCount, failedCount,
} from "./models";
import { RecalcNow } from "./recalc-now";
import { SessionStore } from "./session.store";
import { Thinking } from "./thinking";
import { WorkService } from "./work.service";

// ---- 17 Nov 2026: Cancel, place by place (DESIGN.md section 4, rows 1 to 7 and 18) ----------------------
// Each place: the button is drawn while that work runs, it asks the server to stop that kind of work,
// and the place is left as the table says. Then the chain (Start / Do everything): which sentence, and
// that no next step starts. The Advanced panels (rows 8 to 17) have their own file.

const STOPPED = "Stopped. Nothing was changed.";

const ver = (n: number): Version => ({ id: `ver-00${n}`, major: n, minor: 0, label: `v${n}`, file_name: `m1_v${n}.xlsm`, source: "prep" });

/** One blocking Prep action (the chain's second step writes it) and one formula rewrite (the Fix the formulas tile). */
const blocking = (): PrepAction => ({ id: "title_grids", title: "Title the grids", rule_ids: ["STR-001"], default_on: true, level: "blocking", operations: [{ op: "set_value", sheet: "S", cell: "A1" }], count: 1 });
const formulaFix = (): PrepAction => ({ id: "fix_broken_refs", title: "Fix broken references", rule_ids: ["REF-001"], default_on: false, level: "optional", operations: [{ op: "set_formula", sheet: "S", cell: "B2" }], count: 1 });

const started = (version = 3, plan: PrepAction[] = [blocking(), formulaFix()]): SessionStart => ({
  sessionId: "m1",
  summary: { file_name: "m1.xlsm", file_type: "xlsm", sheet_count: 2, grid_count: 1, formula_count: 10 },
  report: { status: "PASS", findings: [] },
  plan,
  version: ver(version),
  versions: Array.from({ length: version }, (_x, i) => ver(i + 1)),
  recalc: null,
});

const readiness = (state: Readiness["state"], blockingCount = 0): Readiness => ({ state, headline: "h", blocking: [], blocking_count: blockingCount, manual_blocking_count: 0, optional_count: 0 });

const loopPlan = (): LoopPlan => ({ found: true, loops: [{ name: "i", status: "ready", first: 1, last: 5, cells: [] }], outputs: [], clears: [], operation_count: 3 });

const applied = (version: number, plan: PrepAction[] = [], failed = 0): ApplyOutcome => ({ result: { status: "ok", applied: 1, failed }, version: ver(version), plan });
const loopApplied = (version: number) => ({ result: { status: "ok", applied: 3, failed: 0 }, version: ver(version), plan: loopPlan() });
const recalculated = (): RecalcResult => ({ status: "PASS", message: "clean", ran: true, formula_errors: [], addin_gap_errors: [], new_errors: 0 });

const fixerRunning = (): AutoFixStatus => ({ state: "running", title: "Recalculating", fraction: 0.2 });
const fixerDone = (): AutoFixStatus => ({ state: "done", result: { status: "clean", errors_before: 2, errors_after: 0, left: [] }, version: null });
const fixerCancelled = (): AutoFixStatus => ({ state: "cancelled", message: STOPPED, error: null, version: null });

/** How the server answers a request whose work was stopped with Cancel. */
const cancelled = (kind: string) => ({ status: 409, error: { detail: STOPPED, cancelled: true, kind } });

const stopped = (extra: Partial<CancelResult> = {}): CancelResult => ({ cancelled: [], late: [], stopping: [], ahead: false, message: STOPPED, session: null, mind: null, ...extra });

const tick = () => new Promise<void>((r) => setTimeout(r, 0));
async function until(cond: () => boolean): Promise<void> {
  for (let i = 0; i < 500 && !cond(); i++) await tick();
  expect(cond()).toBe(true);
}

interface Hold<T> {
  done: Promise<T>;
  end: (value: T) => void;
  fail: (error: unknown) => void;
}
function hold<T>(): Hold<T> {
  let end!: (value: T) => void;
  let fail!: (error: unknown) => void;
  const done = new Promise<T>((res, rej) => {
    end = res;
    fail = rej;
  });
  return { done, end, fail };
}

const texts = (root: Element, selector: string) => Array.from(root.querySelectorAll(selector)).map((x) => (x.textContent ?? "").replace(/\s+/g, " ").trim());
const buttons = (root: Element) => texts(root, "button");

/** The outside files as the server describes them: two rows, both matched (or the second one missing). */
const outsideFiles = (extra: Partial<OutsideInputs> = {}, missing = false): OutsideInputs => ({
  imports: true,
  macro: "calcCF",
  calls: [],
  rows: [
    { cell: "Ctrl!H35", expected: null, keys: ["BBA_01"], target_sheet: null, file: "BBA_01.xlsm", how: "key", candidates: [] },
    { cell: "Ctrl!H36", expected: null, keys: ["BBA_02"], target_sheet: "BBA_02", file: missing ? null : "BBA_02.xlsm", how: missing ? null : "key", candidates: [] },
  ],
  matched: missing ? 1 : 2,
  missing: missing ? ["Ctrl!H36"] : [],
  extra: [],
  summary: missing ? "1 of 2 input files matched." : "2 of 2 input files matched.",
  brought_in: false,
  brought_in_reason: null,
  ...extra,
});
const outsideStatus = (o: OutsideInputs, extra: Partial<OutsideRunStatus> = {}): OutsideRunStatus => ({ outsideInputs: o, running: false, state: "idle", stage: null, result: null, error: null, ...extra });
const quarterStatus = (extra: Partial<NextQuarterStatus> = {}): NextQuarterStatus => ({ grids: [], available: true, why_not: null, running: false, state: "idle", stage: null, result: null, error: null, ...extra });

/** A stand-in for the server: the calls that start or stop work are written down, the long ones are held
 *  until the test answers them, and the background jobs say what the test tells them to. */
interface Server {
  api: ApiService;
  /** every call that starts or stops work, in order: "startAutoFix", "apply", "loop", "recalc", "cancel:apply" ... */
  calls: string[];
  /** the workbook each Cancel was sent for */
  cancelFor: string[];
  /** the long request of each name that is waiting for its answer */
  pending: Record<string, Hold<unknown>>;
  /** GET .../auto-fix */
  fixer: AutoFixStatus;
  /** GET .../prep-run: where the Prep step's job is (17 Nov 2026, the no-stopping flow) */
  prep: PrepRunStatus;
  /** GET .../outside-inputs, in order (the last one repeats) */
  outside: OutsideRunStatus[];
  outsideLooks: number;
  /** GET .../next-quarter-inputs, in order (the last one repeats) */
  quarter: NextQuarterStatus[];
  /** GET /api/sessions/{id} (null: offline) */
  resumed: SessionStart | null;
  /** what POST .../cancel does and answers (default: "Stopped. Nothing was changed.", no session) */
  onCancel: ((kind: string | null) => Promise<CancelResult>) | null;
}

function server(): Server {
  const s: Server = { api: null as unknown as ApiService, calls: [], cancelFor: [], pending: {}, fixer: { state: "idle" }, prep: { state: "idle" }, outside: [], outsideLooks: 0, quarter: [], resumed: null, onCancel: null };
  const held = <T>(name: string): Promise<T> => {
    s.calls.push(name);
    const h = hold<T>();
    s.pending[name] = h as Hold<unknown>;
    return h.done;
  };
  let lastOutside: OutsideRunStatus | undefined;
  let lastQuarter: NextQuarterStatus | undefined;
  s.api = {
    estimate: () => Promise.reject(new Error("no estimate")),
    downloadUrl: (id: string, name: string) => `/api/sessions/${id}/files/${encodeURIComponent(name)}`,
    mindLoopPlan: () => Promise.resolve({ version_id: "ver-003", plan: loopPlan() }),
    resume: () => (s.resumed ? Promise.resolve(s.resumed) : Promise.reject(new Error("offline"))),
    scanStatus: () => Promise.resolve({ state: "ready" }),
    startAutoFix: () => {
      s.calls.push("startAutoFix");
      s.fixer = fixerRunning();
      return Promise.resolve(s.fixer);
    },
    autoFixStatus: () => Promise.resolve(s.fixer),
    stopAutoFix: () => (s.calls.push("stopAutoFix"), Promise.resolve(s.fixer)),
    apply: () => held<ApplyOutcome>("apply"),
    // 17 Nov 2026 (the no-stopping flow, builder D): the chain's Prep step is one job on the server now
    // (POST and GET .../prep-run), no longer a request of its own for each round. The tests below still
    // hold it and answer it as "apply", as they did the round's request: an ApplyOutcome ends the job
    // "done" (its version and plan are what the workbook then reads back), a 409 "cancelled" ends it
    // "cancelled", any other failure ends it "error". The Fix the formulas tile still sends a real apply.
    startPrepRun: () => {
      const job = held<ApplyOutcome>("apply");
      s.prep = { state: "running", stage: "pass", pass_no: 1, message: "Round 1: writing 1 changes through Excel…" };
      job.then(
        (out) => {
          const n = out.version.major;
          s.resumed = { ...(s.resumed ?? started()), version: out.version, versions: Array.from({ length: n }, (_x, i) => ver(i + 1)), plan: out.plan ?? [] };
          s.prep = { state: "done", stage: "done", applied: appliedCount(out.result), refused: failedCount(out.result), assistant_fixed: 0, left: 0, ask: null };
        },
        (e: { error?: { cancelled?: boolean; detail?: string } }) => {
          s.prep = e?.error?.cancelled === true ? { state: "cancelled", message: STOPPED } : { state: "error", stage: "pass", error: e?.error?.detail ?? "failed" };
        },
      );
      return Promise.resolve({ state: "running" });
    },
    prepRun: () => Promise.resolve(s.prep),
    applyMindLoopPlan: () => held("loop"),
    recalculate: () => held<RecalcResult>("recalc"),
    analyze: () => held<SessionStart>("analyze"),
    cancel: (id: string, kind: string | null = null) => {
      s.calls.push(`cancel:${kind}`);
      s.cancelFor.push(id);
      return s.onCancel ? s.onCancel(kind) : Promise.resolve(stopped());
    },
    outsideInputs: () => {
      s.outsideLooks++;
      lastOutside = s.outside.shift() ?? lastOutside;
      return lastOutside ? Promise.resolve(lastOutside) : Promise.reject(new Error("no status"));
    },
    startOutsideInputs: (_id: string, _standIns = false, wait = false) => (s.calls.push(wait ? "startOutside:wait" : "startOutside"), Promise.resolve({ running: !wait })),
    stopOutsideInputs: () => {
      s.calls.push("stopWaiting");
      const error = "Stopped while waiting for the input files. Nothing was changed: pick the files, then press Do everything again.";
      s.outside.push(outsideStatus(lastOutside?.outsideInputs ?? outsideFiles(), { state: "stopped", error }));
      return Promise.resolve({ running: false, state: "stopped" });
    },
    startNextQuarter: () => (s.calls.push("startNextQuarter"), Promise.resolve({ running: true, raw: [], grid_files: [], notes: [] })),
    nextQuarter: () => {
      lastQuarter = s.quarter.shift() ?? lastQuarter;
      return lastQuarter ? Promise.resolve(lastQuarter) : Promise.reject(new Error("no status"));
    },
  } as unknown as ApiService;
  return s;
}

describe("Cancel on Basic and Custom", () => {
  let Page: typeof HomePage;
  let page: HomePage;
  let store: SessionStore;
  let fixture: ComponentFixture<HomePage> | null = null;

  beforeAll(async () => {
    // Carbon (a package Node cannot load as it is) stands in as inert classes; the page is loaded after that
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
  });

  afterEach(() => {
    if (fixture) fixture.destroy();
    else page?.ngOnDestroy();
    fixture = null;
  });

  const providers = (mode: "basic" | "custom", s: Server) => [provideRouter([]), { provide: ActivatedRoute, useValue: { snapshot: { data: { mode } } } }, { provide: ApiService, useValue: s.api }];

  /** The page's logic, without drawing it: the workbook m1 is open at v3 and has a loop. */
  async function opened(mode: "basic" | "custom", s: Server): Promise<void> {
    TestBed.configureTestingModule({ providers: providers(mode, s) });
    store = TestBed.inject(SessionStore);
    page = TestBed.runInInjectionContext(() => new Page());
    (page as unknown as { sleep: () => Promise<void> }).sleep = tick; // no real waits between looks
    page.canBuildLoop = () => true;
    page.openSaved(started());
    await tick();
  }

  /** The page drawn with its own template; the Cancel button, the dots and the plain parts are real, Carbon is inert. */
  async function drawn(mode: "basic" | "custom", s: Server, open = true): Promise<{ el: HTMLElement; look: () => Promise<void> }> {
    TestBed.configureTestingModule({ providers: providers(mode, s) });
    TestBed.overrideComponent(Page, { set: { imports: [CommonModule, Thinking, CancelButton], schemas: [NO_ERRORS_SCHEMA] } });
    store = TestBed.inject(SessionStore);
    fixture = TestBed.createComponent(Page);
    page = fixture.componentInstance;
    (page as unknown as { sleep: () => Promise<void> }).sleep = tick;
    const f = fixture;
    const look = async () => {
      f.detectChanges();
      await tick();
      f.detectChanges();
    };
    if (open) page.openSaved(started());
    await look();
    await look();
    return { el: fixture.nativeElement as HTMLElement, look };
  }

  const cancelIn = (root: Element, where: string) => root.querySelector(`${where} mr-cancel button`) as HTMLButtonElement | null;
  const sentences = (el: HTMLElement) => texts(el, ".basic__line .basic__sentence");
  const count = (s: Server, name: string) => s.calls.filter((c) => c === name).length;

  // ---- row 1: Basic, Start ----------------------------------------------------------------------
  it("row 1, Basic: Cancel stands beside the Working button; one real click stops everything on the server, and the screen says the workbook is as it was", async () => {
    const s = server();
    const { el, look } = await drawn("basic", s);
    expect(el.querySelector("mr-cancel")).toBeNull(); // nothing runs: no Cancel
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    await look();
    // beside the "Working" button, in the same row
    const row = el.querySelector(".basic__run")!;
    expect(row.querySelector(".basic__go--working")?.textContent?.trim()).toBe("Working");
    const cancel = cancelIn(el, ".basic__run")!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    expect(cancel.disabled).toBe(false);
    expect(el.querySelectorAll("mr-cancel")).toHaveLength(1);

    // the click: the server is asked to stop Start's own steps (17 Nov 2026, review R1-5: it asked for
    // everything, so a Run in Mind or an assistant's answer started on another screen was stopped too)
    const answer = hold<CancelResult>();
    s.onCancel = () => answer.done;
    cancel.click();
    await look();
    expect(s.calls).toContain("cancel:outside,autofix,prep,apply,recalc,scan");
    expect(s.calls).not.toContain("cancel:null");
    expect(s.cancelFor).toEqual(["m1"]);
    // while it stops: the button is greyed and says so, the line says so, the bar is gone
    expect(cancel.disabled).toBe(true);
    expect(cancel.textContent?.trim()).toBe("Stopping…");
    expect(sentences(el)).toEqual(["Stopping…"]);
    expect(el.querySelector("mr-run-bar")).toBeNull();
    expect(page.chainRunning()).toBe(true); // the chain did not end at the click: the step ends by itself

    // the server stopped the fixer and answers with the workbook as it is now
    s.fixer = fixerCancelled();
    answer.end(stopped({ cancelled: ["autofix"], session: started() }));
    await run;
    await look();
    expect(sentences(el)).toEqual(["Stopped. Your workbook is as it was before you pressed Start."]);
    expect(page.steps().map((x) => x.state)).toEqual(["stopped", "waiting", "waiting", "waiting"]);
    expect(page.steps()[0].line).toBe("Stopped.");
    expect(el.querySelector("mr-cancel")).toBeNull();
    expect(el.querySelector(".basic__go")?.textContent?.trim()).toBe("Start again");
    // no next step started
    expect(count(s, "apply")).toBe(0);
    expect(count(s, "recalc")).toBe(0);
    expect(store.currentVersion()?.label).toBe("v3");
  });

  // ---- row 2: Custom, Do everything -----------------------------------------------------------------
  it("row 2, Custom: Cancel stands beside 'Stop, keep what is fixed' while the fixer's step runs, and beside 'Running…' in the other steps", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    await look();
    const head = el.querySelector(".chain__buttons")!;
    // the soft stop is still there, as before, with Cancel beside it
    expect(buttons(head)).toEqual(["Stop, keep what is fixed", "Cancel"]);
    // (inside Do everything the head's Cancel serves: the tile and the list have none of their own)
    expect(el.querySelectorAll("mr-cancel")).toHaveLength(1);
    expect(buttons(el.querySelector(".actions")!)).toContain("Stop, keep what is fixed");

    // the fixer ends; the next step is one long request: "Running…" with Cancel beside it
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    await look();
    expect(buttons(head)).toEqual(["Running…", "Cancel"]);

    // the click: Do everything's own steps (17 Nov 2026, review R1-5: not "everything for this workbook");
    // the request answers "cancelled" by itself
    s.onCancel = () => {
      s.pending["apply"].fail(cancelled("apply"));
      return Promise.resolve(stopped({ cancelled: ["apply"], session: started() }));
    };
    cancelIn(el, ".chain__buttons")!.click();
    await run;
    await look();
    expect(s.calls).toContain("cancel:outside,autofix,prep,apply,recalc,scan");
    expect(texts(el, ".chain > .chain__text strong")).toEqual(["Stopped. Your workbook is as it was after the formula errors were fixed."]);
    expect(page.chainStoppedByCancel()).toBe(true); // no "Show me" on this sentence: there is no place to fix
    expect(el.querySelector(".chain > .chain__text mr-show-me")).toBeNull();
    expect(page.steps().map((x) => x.state)).toEqual(["done", "stopped", "waiting", "waiting"]);
    expect(buttons(head)).toEqual(["Do everything again"]);
    expect(count(s, "loop")).toBe(0);
  });

  it("'Stop, keep what is fixed' still does what it did: it asks the fixer to stop softly, not the Cancel", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    await look();
    const stop = Array.from(el.querySelectorAll(".chain__buttons button")).find((b) => b.textContent?.trim() === "Stop, keep what is fixed") as HTMLButtonElement;
    stop.click();
    await look();
    expect(s.calls).toContain("stopAutoFix");
    expect(s.calls.some((c) => c.startsWith("cancel"))).toBe(false);
    expect(page.fixAll().stopping).toBe(true);
    // the fixer ends its round and keeps what it fixed: an ordinary end of the step, not a Cancel
    s.fixer = { state: "done", result: { status: "partial", errors_before: 5, errors_after: 2, left: [], stopped: "on request" }, version: null };
    await run;
    expect(page.chainStoppedByCancel()).toBe(false);
    expect(page.chainLine()).toMatch(/^Stopped at "Fix all issues"/);
  });

  // ---- row 3: the waiting box -------------------------------------------------------------------------
  it("row 3, the waiting box: its button says Cancel, still ends the wait, and the sentence says nothing had finished", async () => {
    const s = server();
    const some = outsideFiles({}, true);
    const { el, look } = await drawn("custom", s);
    store.outsideInputs.set(some);
    s.outside = [outsideStatus(some), outsideStatus(some, { state: "waiting_for_inputs", waiting: true, stage: "Waiting for the input files", message: "Waiting for 1 input file: BBA_02" })];
    const run = page.runAll();
    await until(() => page.outsideWaiting());
    await look();
    const box = el.querySelector(".chain__waiting")!;
    expect(buttons(box)).toContain("Cancel");
    expect(buttons(box)).not.toContain("Stop");
    // one Cancel on screen: the box's own (the head's is not drawn while Do everything waits)
    expect(el.querySelectorAll("mr-cancel")).toHaveLength(0);
    const cancel = Array.from(box.querySelectorAll("button")).find((b) => b.textContent?.trim() === "Cancel") as HTMLButtonElement;
    cancel.click();
    await run;
    await look();
    // it ended the wait as before (the server's own route for that), and nothing ran
    expect(count(s, "stopWaiting")).toBe(1);
    expect(s.calls.some((c) => c.startsWith("cancel"))).toBe(false);
    expect(count(s, "startAutoFix")).toBe(0);
    expect(page.outsideWaiting()).toBe(false);
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was before you pressed Do everything.");
    expect(page.steps()[0]).toMatchObject({ id: "outside", state: "stopped", line: "Stopped." });
    // the card says so calmly: not an error, no "Try again"
    expect(page.outsideCard()).toMatchObject({ line: STOPPED, tone: "idle" });
    expect(page.outsideFailed()).toBe(false);
  });

  it("row 3 on Basic: the same box, and the sentence in Basic's words", async () => {
    const s = server();
    const some = outsideFiles({}, true);
    const { el, look } = await drawn("basic", s);
    store.outsideInputs.set(some);
    s.outside = [outsideStatus(some), outsideStatus(some, { state: "waiting_for_inputs", waiting: true, stage: "Waiting for the input files", message: "Waiting for 1 input file: BBA_02" })];
    const run = page.runAll();
    await until(() => page.outsideWaiting());
    await look();
    expect(el.querySelectorAll("mr-cancel")).toHaveLength(0);
    const cancel = Array.from(el.querySelectorAll(".chain__waiting button")).find((b) => b.textContent?.trim() === "Cancel") as HTMLButtonElement;
    cancel.click();
    await run;
    await look();
    expect(sentences(el)).toEqual(["Stopped. Your workbook is as it was before you pressed Start."]);
  });

  // ---- row 4: the Custom tiles ------------------------------------------------------------------------
  it("row 4, Fix all issues: Cancel under the bar stops the fixer (kind autofix); 'Stop, keep what is fixed' stays beside it", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const run = page.runFixAll();
    await until(() => s.calls.includes("startAutoFix"));
    await look();
    const tiles = el.querySelector(".actions")!;
    expect(buttons(tiles)).toContain("Stop, keep what is fixed");
    const cancel = cancelIn(el, ".actions .action__progress")!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    expect(el.querySelectorAll("mr-cancel")).toHaveLength(1);
    s.onCancel = () => ((s.fixer = fixerCancelled()), Promise.resolve(stopped({ cancelled: ["autofix"], session: started() })));
    cancel.click();
    expect(await run).toBe(false);
    await look();
    expect(s.calls).toContain("cancel:autofix");
    expect(page.fixAll()).toMatchObject({ running: false, line: STOPPED, tone: "idle", detail: "" });
    // the sentence is on the tile at once, in no colour
    expect(texts(tiles, ".action__result strong")).toEqual([STOPPED]);
    expect(tiles.querySelector(".action__result--error, .action__result--warn, .action__result--ok")).toBeNull();
    expect(el.querySelector("mr-cancel")).toBeNull();
    expect(buttons(tiles)).not.toContain("Stop, keep what is fixed");
  });

  it("row 4, Fix the formulas: Cancel under the bar stops the writing (kind apply); nothing was changed", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const run = page.runFixFormulas();
    await until(() => s.calls.includes("apply"));
    await look();
    const cancel = cancelIn(el, ".actions .action__progress")!;
    s.onCancel = () => (s.pending["apply"].fail(cancelled("apply")), Promise.resolve(stopped({ cancelled: ["apply"], session: started() })));
    cancel.click();
    expect(await run).toBe(false);
    await look();
    expect(s.calls).toContain("cancel:apply");
    expect(page.fixFormulas()).toMatchObject({ running: false, line: STOPPED, tone: "idle" });
    expect(texts(el, ".actions .action__result strong")).toEqual([STOPPED]);
    expect(store.currentVersion()?.label).toBe("v3");
  });

  it("row 4, Build the Mind loop: Cancel under the bar stops the writing (kind apply); the loop was not written", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    await until(() => page.canBuildLoop());
    const run = page.runLoop();
    await until(() => s.calls.includes("loop"));
    await look();
    const cancel = cancelIn(el, ".actions .action__progress")!;
    s.onCancel = () => (s.pending["loop"].fail(cancelled("apply")), Promise.resolve(stopped({ cancelled: ["apply"], session: started() })));
    cancel.click();
    expect(await run).toBe(false);
    await look();
    expect(s.calls).toContain("cancel:apply");
    expect(page.loop()).toMatchObject({ running: false, line: STOPPED, tone: "idle" });
    expect(texts(el, ".actions .action__result strong")).toEqual([STOPPED]);
  });

  it("row 4, Recalculate: Cancel under the bar stops the recalculation (kind recalc); the last good one stays", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const before = { status: "PASS", message: "clean", ran: true, formula_errors: [], addin_gap_errors: [], version_id: "ver-003" } as RecalcResult;
    store.recalc.set(before);
    const run = page.runRecalc();
    await until(() => s.calls.includes("recalc"));
    await look();
    const cancel = cancelIn(el, ".actions .action__progress")!;
    s.onCancel = () => (s.pending["recalc"].fail(cancelled("recalc")), Promise.resolve(stopped({ cancelled: ["recalc"] })));
    cancel.click();
    expect(await run).toBe(false);
    await look();
    expect(s.calls).toContain("cancel:recalc");
    expect(page.recalc()).toMatchObject({ running: false, line: STOPPED, tone: "idle" });
    expect(store.recalc()).toBe(before);
  });

  it("row 4: each tile's Cancel names its own kind of work", async () => {
    const s = server();
    await opened("custom", s);
    // (17 Nov 2026, builder D: the Prep step is a job of its own kind now, "prep"; it was "apply")
    expect((["fixAll", "fixFormulas", "prepare", "loop", "recalc", "outside", "nextQuarter"] as const).map((w) => page.cancelKind(w))).toEqual(["autofix", "apply", "prep", "apply", "recalc", "outside", "next_quarter"]);
  });

  // ---- row 4b: the Outside files card, and next quarter ------------------------------------------------
  it("row 4b, Bring in the outside files: Cancel under the bar stops the bring-in (kind outside); the card says so and asks the files again", async () => {
    const s = server();
    const all = outsideFiles();
    const { el, look } = await drawn("custom", s);
    store.outsideInputs.set(all);
    s.outside = [outsideStatus(all, { running: true, state: "running", stage: "Running the workbook's import macro" })];
    const run = page.runOutside();
    await until(() => s.calls.includes("startOutside") && page.outside().running && s.outsideLooks > 0);
    await look();
    const cancel = cancelIn(el, ".outside .action__progress")!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    let looksAtCancel = 0;
    s.onCancel = () => {
      looksAtCancel = s.outsideLooks;
      s.outside.push(outsideStatus(all, { state: "cancelled", running: false, error: null, versions: [], message: STOPPED }));
      return Promise.resolve(stopped({ cancelled: ["outside"], session: started() }));
    };
    cancel.click();
    expect(await run).toBe(false);
    await look();
    expect(s.calls).toContain("cancel:outside");
    expect(page.outsideCard()).toMatchObject({ running: false, line: STOPPED, tone: "idle" });
    expect(texts(el, ".outside__result")).toEqual([STOPPED]);
    // not a failure: no "Try again", and the files are as they were (the match was read again)
    expect(page.outsideFailed()).toBe(false);
    expect(page.outsideButton(store.outsideInputs()!)).toBe("Bring in the outside files");
    expect(s.outsideLooks).toBeGreaterThanOrEqual(looksAtCancel + 2); // the job's last look, then the files again
    expect(el.querySelector("mr-cancel")).toBeNull();
  });

  it("row 4b, Make them with AI: the same Cancel while the assistant makes the missing files", async () => {
    const s = server();
    const some = outsideFiles({ assistant: true }, true);
    const { el, look } = await drawn("custom", s);
    store.outsideInputs.set(some);
    s.outside = [outsideStatus(some, { running: true, state: "running", stage: "Making the missing input files with the assistant" })];
    const run = page.runOutside(false, true);
    await until(() => s.calls.includes("startOutside") && s.outsideLooks > 0);
    await look();
    const cancel = cancelIn(el, ".outside .action__progress")!;
    s.onCancel = () => (s.outside.push(outsideStatus(some, { state: "cancelled", running: false, error: null, versions: [] })), Promise.resolve(stopped({ cancelled: ["outside"] })));
    cancel.click();
    expect(await run).toBe(false);
    expect(s.calls).toContain("cancel:outside");
    expect(page.outsideCard()).toMatchObject({ line: STOPPED, tone: "idle" });
  });

  it("row 4b, next quarter's files: Cancel under the bar stops the job (kind next_quarter); no zip, and the card is emptied", async () => {
    const s = server();
    const brought = outsideFiles({ brought_in: true, mind_inputs: ["IN_BBA_01"] });
    const { el, look } = await drawn("custom", s);
    store.outsideInputs.set(brought);
    s.quarter = [quarterStatus({ running: true, state: "running", stage: "Running the model's import macro with the new files" })];
    const run = page.makeNextQuarter([new File(["x"], "2026_Q2 - BBA_01.csv")]);
    await until(() => s.calls.includes("startNextQuarter") && page.nextQuarter().running);
    await look();
    const cancel = cancelIn(el, ".outside__quarter .action__progress")!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    s.onCancel = () => (s.quarter.push(quarterStatus({ state: "cancelled", running: false, error: null, result: null })), Promise.resolve(stopped({ cancelled: ["next_quarter"] })));
    cancel.click();
    expect(await run).toBe(false);
    await look();
    expect(s.calls).toContain("cancel:next_quarter");
    expect(page.nextQuarter()).toMatchObject({ running: false, line: STOPPED, tone: "idle" });
    expect(texts(el, ".outside__quarter .outside__result")).toEqual([STOPPED]);
    expect(page.nextQuarterZipUrl()).toBeNull();
    expect(page.nextQuarterFiles()).toEqual([]);
    expect(page.busy()).toBe(false);
  });

  // ---- row 5: the first check after an upload ------------------------------------------------------------
  it("row 5, the first check: Cancel under the stages stops the check (kind scan) for the workbook being read, and the upload screen says where the workbook is", async () => {
    const s = server();
    const { el, look } = await drawn("basic", s, false);
    const scan = page.scanSession("p9", []);
    await until(() => s.calls.includes("analyze"));
    await look();
    expect(texts(el, ".page-title")).toEqual(["Reading the workbook"]);
    const cancel = cancelIn(el, ".scan")!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    s.onCancel = () => (s.pending["analyze"].fail(cancelled("scan")), Promise.resolve(stopped({ cancelled: ["scan"], session: null })));
    cancel.click();
    await scan;
    await look();
    // the store holds no workbook yet: the Cancel was sent for the one being checked
    expect(s.calls).toContain("cancel:scan");
    expect(s.cancelFor).toEqual(["p9"]);
    // back on the upload screen, with one calm sentence and no error box
    expect(page.phase()).toBe("idle");
    expect(texts(el, ".scan__stopped")).toEqual(["Stopped. The workbook was not checked. It is in Recent workbooks."]);
    expect(page.uploadError()).toBe("");
    expect(el.querySelector("cds-inline-notification")).toBeNull();
    expect(el.querySelector("mr-recent-workbooks")).not.toBeNull(); // the list the sentence names
    expect(store.sessionId()).toBeNull();
    expect(el.querySelector("mr-cancel")).toBeNull();
  });

  it("row 5: a check that fails by itself is still an error, and the next upload clears the sentence", async () => {
    const s = server();
    await drawn("basic", s, false);
    const scan = page.scanSession("p9", []);
    await until(() => s.calls.includes("analyze"));
    s.pending["analyze"].fail({ status: 500, error: { detail: "The workbook could not be read." } });
    await scan;
    expect(page.uploadError()).toBe("The workbook could not be read.");
    expect(page.scanNote()).toBe("");
    // after a Cancel, a new check starts clean
    const again = page.scanSession("p9", []);
    await until(() => count(s, "analyze") === 2);
    s.pending["analyze"].fail(cancelled("scan"));
    await again;
    expect(page.scanNote()).toBe("Stopped. The workbook was not checked. It is in Recent workbooks.");
    const third = page.scanSession("p9", []);
    expect(page.scanNote()).toBe("");
    await until(() => count(s, "analyze") === 3);
    s.pending["analyze"].fail(cancelled("scan"));
    await third;
  });

  // ---- a late answer never reaches the store ----------------------------------------------------------
  it("the answer of a cancelled call comes late and never reaches the store: the workbook is the one the Cancel gave back", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const run = page.runFixFormulas();
    await until(() => s.calls.includes("apply"));
    await look();
    // the Cancel answers first, with the workbook as it is now (the server went back to v2, say)
    s.onCancel = () => Promise.resolve(stopped({ cancelled: ["apply"], session: started(2, [blocking()]) }));
    cancelIn(el, ".actions .action__progress")!.click();
    await look();
    expect(store.currentVersion()?.label).toBe("v2");
    expect(page.fixFormulas().running).toBe(true); // the page abandoned nothing: its own request is still out
    // now the stopped request answers (409 "cancelled", with whatever it carries)
    s.pending["apply"].fail({ status: 409, error: { detail: STOPPED, cancelled: true, kind: "apply", version: ver(4), plan: [], versions: [ver(1), ver(2), ver(3), ver(4)] } });
    expect(await run).toBe(false);
    await look();
    expect(store.currentVersion()?.label).toBe("v2");
    expect(store.versions().map((v) => v.label)).toEqual(["v1", "v2"]);
    expect(store.plan().map((a) => a.id)).toEqual(["title_grids"]);
    expect(page.fixFormulas()).toMatchObject({ running: false, line: STOPPED, tone: "idle" });
  });

  it("a call that answers normally after the click is handled as if no Cancel was pressed: its result stands and is shown", async () => {
    const s = server();
    const { el, look } = await drawn("custom", s);
    const run = page.runFixFormulas();
    await until(() => s.calls.includes("apply"));
    await look();
    s.onCancel = () => Promise.resolve(stopped({ late: ["apply"], message: "It had already finished.", session: started(4, []) }));
    cancelIn(el, ".actions .action__progress")!.click();
    s.pending["apply"].end(applied(4));
    expect(await run).toBe(true);
    await look();
    expect(page.fixFormulas()).toMatchObject({ running: false, line: "1 formulas changed", tone: "ok" });
    expect(store.currentVersion()?.label).toBe("v4");
  });

  // ---- the chain: which sentence, and that nothing new starts --------------------------------------------
  /** The click on the Cancel beside Start, and the button hearing from the server, as the template wires them. */
  const click = () => page.onChainStopping();
  const heard = () => page.onChainStopped();

  it("the chain: Cancel before the first step ends says the workbook is as it was before Start, on both screens", async () => {
    for (const [mode, word] of [["basic", "Start"], ["custom", "Do everything"]] as const) {
      const s = server();
      await opened(mode, s);
      const run = page.runAll();
      await until(() => s.calls.includes("startAutoFix"));
      click();
      expect(page.chainStopping()).toBe(true);
      s.fixer = fixerCancelled();
      heard();
      await run;
      // the chain's own sentence; Basic shows it in its own words ("Start" for "Do everything")
      expect(page.chainLine()).toBe("Stopped. Your workbook is as it was before you pressed Do everything.");
      if (mode === "basic") expect(page.basicLine()).toBe(`Stopped. Your workbook is as it was before you pressed ${word}.`);
      expect(page.steps().map((x) => x.state)).toEqual(["stopped", "waiting", "waiting", "waiting"]);
      expect(page.fixAll()).toMatchObject({ line: STOPPED, tone: "idle" });
      expect(page.chainRunning()).toBe(false);
      expect(page.chainStopping()).toBe(false);
      expect(s.calls).toEqual(["startAutoFix"]);
      page.ngOnDestroy();
      TestBed.resetTestingModule();
    }
  });

  it("the chain: Cancel after one step names that step, and only the running step is thrown away", async () => {
    const s = server();
    await opened("basic", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    click();
    s.pending["apply"].fail(cancelled("apply"));
    heard();
    await run;
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was after the formula errors were fixed.");
    // the finished step keeps its state, the running one says "Stopped."
    expect(page.steps().map((x) => [x.state, x.line])).toEqual([["done", "Every formula error is fixed"], ["stopped", "Stopped."], ["waiting", ""], ["waiting", ""]]);
    expect(s.calls).toEqual(["startAutoFix", "apply"]);
  });

  it("the chain: Cancel after three steps names the third (the Mind loop)", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    s.pending["apply"].end(applied(4));
    await until(() => s.calls.includes("loop"));
    s.pending["loop"].end(loopApplied(5));
    await until(() => s.calls.includes("recalc"));
    click();
    s.pending["recalc"].fail(cancelled("recalc"));
    heard();
    await run;
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was after the Mind loop was built.");
    expect(page.steps().map((x) => x.state)).toEqual(["done", "done", "done", "stopped"]);
    expect(page.recalc()).toMatchObject({ line: STOPPED, tone: "idle" });
    expect(store.currentVersion()?.label).toBe("v5");
  });

  it("the chain: a skipped step does not count as the last one that finished", async () => {
    const s = server();
    await opened("basic", s);
    page.buildLoop.set(false); // the loop is left out
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    s.pending["apply"].end(applied(4));
    await until(() => s.calls.includes("recalc"));
    click();
    s.pending["recalc"].fail(cancelled("recalc"));
    heard();
    await run;
    expect(page.steps().map((x) => x.state)).toEqual(["done", "done", "skipped", "stopped"]);
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was after the formulas and the structure were fixed.");
  });

  it("the chain: Cancel between two steps, with nothing running on the server, ends it the same way and starts no next step", async () => {
    // before the first step: Start is still getting ready (its time estimate is on its way)
    const s = server();
    const estimate = hold<never>();
    (s.api as unknown as { estimate: () => Promise<never> }).estimate = () => estimate.done;
    await opened("basic", s);
    const run = page.runAll();
    await tick();
    expect(page.chainRunning()).toBe(true);
    click();
    heard(); // the server had nothing to stop: "Nothing was running."
    estimate.fail(new Error("no estimate"));
    await run;
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was before you pressed Start.");
    expect(page.steps().map((x) => x.state)).toEqual(["waiting", "waiting", "waiting", "waiting"]);
    expect(s.calls).toEqual([]);

    // after a step: the click lands when the first step has just answered
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    const s2 = server();
    await opened("basic", s2);
    const p = page as unknown as Record<string, unknown>;
    const fixAll = page.runFixAll.bind(page);
    p["runFixAll"] = async () => {
      const ok = await fixAll();
      click();
      heard();
      return ok;
    };
    const run2 = page.runAll();
    await until(() => s2.calls.includes("startAutoFix"));
    s2.fixer = fixerDone();
    await run2;
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was after the formula errors were fixed.");
    expect(page.steps().map((x) => x.state)).toEqual(["done", "waiting", "waiting", "waiting"]);
    expect(s2.calls).toEqual(["startAutoFix"]);
  });

  it("the chain: Cancel during the fixes after the loop throws those away; the loop counts as finished", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    s.pending["apply"].end(applied(4));
    await until(() => s.calls.includes("loop"));
    // the loop uncovers something blocking: one more round of fixes follows it
    s.resumed = started(5, [blocking()]);
    s.pending["loop"].end(loopApplied(5));
    await until(() => count(s, "apply") === 2);
    expect(page.steps()[2]).toMatchObject({ id: "loop", state: "running" });
    click();
    s.pending["apply"].fail(cancelled("apply"));
    heard();
    await run;
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was after the Mind loop was built.");
    expect(page.steps().map((x) => x.state)).toEqual(["done", "done", "stopped", "waiting"]);
    expect(page.steps()[2].line).toBe("3 loop formulas written; the fixes after the loop were stopped");
    expect(count(s, "recalc")).toBe(0);
    expect(store.currentVersion()?.label).toBe("v5");
  });

  it("the chain: after a Cancel the fixes after the loop do not even start", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    s.pending["apply"].end(applied(4));
    await until(() => s.calls.includes("loop"));
    s.resumed = started(5, [blocking()]);
    click(); // while the loop itself is being written; its answer comes normally
    s.pending["loop"].end(loopApplied(5));
    heard();
    await run;
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was after the Mind loop was built.");
    expect(page.steps().map((x) => x.state)).toEqual(["done", "done", "done", "waiting"]);
    expect(count(s, "apply")).toBe(1); // no round of fixes was begun
    expect(count(s, "recalc")).toBe(0);
  });

  it("the chain: a step whose call answers normally after the click counts as finished and is named", async () => {
    const s = server();
    await opened("basic", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    click();
    s.pending["apply"].end(applied(4)); // the click reached the server too late: the round was written
    heard();
    await run;
    expect(page.steps().map((x) => [x.state, x.line])).toEqual([["done", "Every formula error is fixed"], ["done", "1 changes applied"], ["waiting", ""], ["waiting", ""]]);
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was after the formulas and the structure were fixed.");
    expect(store.currentVersion()?.label).toBe("v4"); // its result stands
    expect(count(s, "loop")).toBe(0); // and the chain stops after it
  });

  it("the chain: the fixer ending 'done' after the click counts as finished and is named", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    click();
    s.fixer = fixerDone();
    heard();
    await run;
    expect(page.steps().map((x) => x.state)).toEqual(["done", "waiting", "waiting", "waiting"]);
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was after the formula errors were fixed.");
    expect(count(s, "apply")).toBe(0);
  });

  // (17 Nov 2026, builder D: the rounds are inside one job on the server now, so the page starts none itself;
  // what is left of this case: a Prep job that ends by itself after the click stands whole)
  it("the chain: no new Prep round starts after the click, and the round that was written stands", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    click();
    s.pending["apply"].end(applied(4, [blocking()])); // this round was written; another one would be needed
    heard();
    await run;
    expect(count(s, "apply")).toBe(1);
    expect(page.steps()[1]).toMatchObject({ id: "prepare", line: "1 changes applied" });
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was after the formulas and the structure were fixed.");
  });

  it("the chain: a step that fails by itself after the click keeps its own stop sentence", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    click();
    s.pending["apply"].fail({ status: 500, error: { detail: "Excel did not start." } });
    heard();
    await run;
    expect(page.steps()[1].state).toBe("failed");
    expect(page.chainStoppedByCancel()).toBe(false);
    // (17 Nov 2026, builder D: the Prep job's own error says "Stopped with an error", with the server's sentence
    // under it; it was "Could not write the changes", which is untrue for a round that was written but not checked)
    expect(page.chainLine()).toBe('Stopped at "Fix the formulas and the structure": Stopped with an error. Fix that under Advanced › Prep actions, then press Do everything again.');
    expect(count(s, "loop")).toBe(0);
  });

  it("the chain: a step stopped by another Cancel (a card's, the strip's) ends the chain the same way, without waiting for a button", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    s.fixer = fixerCancelled(); // stopped from somewhere else: this page's own Cancel was never clicked
    await run;
    expect(page.chainLine()).toBe("Stopped. Your workbook is as it was before you pressed Do everything.");
    expect(page.steps()[0]).toMatchObject({ state: "stopped", line: "Stopped." });
    expect(count(s, "apply")).toBe(0);
  });

  it("the chain: when the server cannot be asked, nothing new starts and the chain does not wait for ever", async () => {
    const s = server();
    await opened("basic", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    click();
    page.onChainStopFailed(); // the Cancel call itself failed: the button is back for another try
    expect(page.chainStopping()).toBe(false);
    s.fixer = fixerDone(); // the step went to its own end
    await run;
    expect(page.basicLine()).toBe("Stopped. Your workbook is as it was after the formula errors were fixed.");
    expect(count(s, "apply")).toBe(0);
  });

  it("the chain: Start again after a Cancel is a new run: every step runs, and the sentence is the ordinary one", async () => {
    const s = server();
    await opened("basic", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    click();
    s.fixer = fixerCancelled();
    heard();
    await run;
    expect(page.chainStoppedByCancel()).toBe(true);
    const again = page.runAll();
    await until(() => count(s, "startAutoFix") === 2);
    expect(page.chainStoppedByCancel()).toBe(false);
    expect(page.chainStopping()).toBe(false);
    s.fixer = fixerDone();
    await until(() => s.calls.includes("apply"));
    s.pending["apply"].end(applied(4));
    await until(() => s.calls.includes("loop"));
    s.pending["loop"].end(loopApplied(5));
    await until(() => s.calls.includes("recalc"));
    s.pending["recalc"].end(recalculated());
    await again;
    expect(page.steps().map((x) => x.state)).toEqual(["done", "done", "done", "done"]);
    expect(page.basicLine()).toMatch(/^Everything ran and every check passed/);
  });

  it("Custom's timer and Basic's bar give way while a Cancel is stopping the run", async () => {
    const s = server();
    await opened("custom", s);
    const run = page.runAll();
    await until(() => s.calls.includes("startAutoFix"));
    expect(page.basicBar()).not.toBeNull();
    click();
    expect(page.timerLine()).toMatch(/so far · stopping…$/);
    expect(page.basicBar()).toBeNull();
    expect(page.basicLines()).toEqual(["Stopping…"]);
    s.fixer = fixerCancelled();
    heard();
    await run;
    expect(page.timerLine()).toBe("");
  });
});

// ---- row 6: Recalculate in the banners -------------------------------------------------------------------
describe("Cancel beside Recalculate in the banners", () => {
  let store: SessionStore;
  let fixture: ComponentFixture<RecalcNow>;

  afterEach(() => fixture?.destroy());

  function draw(api: Partial<ApiService>): { el: HTMLElement; look: () => Promise<void> } {
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: api }] });
    store = TestBed.inject(SessionStore);
    store.setSession(started());
    store.readiness.set(readiness("unverified")); // the recalculation is the one thing left
    fixture = TestBed.createComponent(RecalcNow);
    fixture.detectChanges();
    const look = async () => {
      fixture.detectChanges();
      await tick();
      fixture.detectChanges();
    };
    return { el: fixture.nativeElement as HTMLElement, look };
  }

  it("row 6: Cancel stands beside 'Recalculating…', stops the recalculation (kind recalc), and the button is back to Recalculate", async () => {
    const answer = hold<RecalcResult>();
    const calls: string[] = [];
    const { el, look } = draw({
      recalculate: () => (calls.push("recalc"), answer.done),
      cancel: (_id: string, kind: string | null = null) => (calls.push(`cancel:${kind}`), answer.fail(cancelled("recalc")), Promise.resolve(stopped({ cancelled: ["recalc"] }))),
    });
    const button = () => el.querySelector(".recalc-now__button") as HTMLButtonElement;
    expect(button().textContent?.trim()).toBe("Recalculate");
    expect(el.querySelector("mr-cancel")).toBeNull();
    button().click();
    await look();
    expect(button().textContent?.trim()).toBe("Recalculating…");
    const cancel = el.querySelector(".recalc-now mr-cancel button") as HTMLButtonElement;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    cancel.click();
    await look();
    await look();
    expect(calls).toEqual(["recalc", "cancel:recalc"]);
    expect(button().textContent?.trim()).toBe("Recalculate");
    expect(button().disabled).toBe(false);
    expect(el.querySelector("mr-cancel")).toBeNull();
    // one calm sentence, not an error; nothing reached the store
    expect(texts(el, ".recalc-now__note")).toEqual([STOPPED]);
    expect(el.querySelector(".recalc-now__note--error")).toBeNull();
    expect(fixture.componentInstance.error()).toBe("");
    expect(store.recalc()).toBeNull();
  });

  it("row 6: a recalculation that fails by itself is still an error, and one that answers after the click is kept", async () => {
    const failing = hold<RecalcResult>();
    const first = draw({ recalculate: () => failing.done });
    (first.el.querySelector(".recalc-now__button") as HTMLButtonElement).click();
    failing.fail({ status: 500, error: { detail: "Excel did not start." } });
    await first.look();
    expect(texts(first.el, ".recalc-now__note--error")).toEqual(["Excel did not start."]);
    fixture.destroy();
    TestBed.resetTestingModule();

    const late = hold<RecalcResult>();
    const second = draw({ recalculate: () => late.done, cancel: () => Promise.resolve(stopped({ late: ["recalc"], message: "It had already finished." })) });
    (second.el.querySelector(".recalc-now__button") as HTMLButtonElement).click();
    await second.look();
    (second.el.querySelector("mr-cancel button") as HTMLButtonElement).click();
    late.end({ ...recalculated(), readiness: readiness("ready") });
    await second.look();
    await second.look();
    expect(store.recalc()?.status).toBe("PASS"); // its result stands and is shown
    expect(store.readiness()?.state).toBe("ready");
    expect(fixture.componentInstance.note()).toBe("");
  });
});

// ---- row 7: Upload to Mind in the finish banner ---------------------------------------------------------
describe("Cancel in the finish banner (Upload to Mind)", () => {
  let store: SessionStore;
  let fixture: ComponentFixture<FinishBanner>;
  const everyMs = FinishBanner.everyMs;

  beforeEach(() => (FinishBanner.everyMs = 1));
  afterEach(() => {
    fixture?.destroy();
    FinishBanner.everyMs = everyMs;
  });

  const signedIn = (yes: boolean, state: MindSignIn["state"] = "done"): MindSignIn => ({ state, message: null, last_login: null, checked: { signed_in: yes, at: null }, elapsed_s: null });
  const project = (name = "zzz_mindready_ab12", status = "model added"): MindProject => ({ name, status, ran: false, deleted: null, url: "https://mind.example/" });
  const run = (state: MindLoopStatus["state"], extra: Partial<MindLoopStatus> = {}): MindLoopStatus => ({ state, events: [], next: 1, report: null, ...extra });
  const STILL = "Stopped. The project zzz_mindready_ab12 is still in Mind.";

  interface Mind {
    api: Partial<ApiService>;
    calls: string[];
    /** GET .../mind-loop, in order (the last one repeats); the first is the banner's look when it opens */
    statuses: MindLoopStatus[];
    check: Hold<MindSignIn> | null;
    options: Hold<unknown> | null;
    start: Hold<{ state: string }> | null;
    /** what the sign-in window says at each look */
    window: MindSignIn;
    onCancel: (() => Promise<CancelResult>) | null;
    onDelete: ((name: string) => Promise<MindDelete>) | null;
  }

  function mind(statuses: MindLoopStatus[] = [run("idle")]): Mind {
    const m: Mind = { api: {}, calls: [], statuses, check: null, options: null, start: null, window: signedIn(false, "waiting"), onCancel: null, onDelete: null };
    let last: MindLoopStatus | undefined;
    m.api = {
      mindSignIn: (check = false) => {
        m.calls.push(check ? "check" : "look");
        return check ? (m.check ? m.check.done : Promise.resolve(signedIn(true))) : Promise.resolve(m.window);
      },
      startMindSignIn: () => (m.calls.push("window"), Promise.resolve(signedIn(false, "waiting"))),
      mindLoopOptions: () => (m.calls.push("options"), m.options ? m.options.done : Promise.resolve({ choices: [{ id: "a", default_on: true }] })),
      startMindLoop: () => (m.calls.push("start"), m.start ? m.start.done : Promise.resolve({ state: "running" })),
      mindLoopStatus: () => {
        last = m.statuses.shift() ?? last;
        return last ? Promise.resolve(last) : Promise.reject(new Error("no status"));
      },
      cancel: (_id: string, kind: string | null = null) => (m.calls.push(`cancel:${kind}`), m.onCancel ? m.onCancel() : Promise.resolve(stopped({ message: "Nothing was running." }))),
      deleteMindProject: (_id: string, name: string) => (m.calls.push(`delete:${name}`), m.onDelete ? m.onDelete(name) : Promise.reject(new Error("no delete"))),
    } as unknown as Partial<ApiService>;
    return m;
  }

  async function draw(m: Mind, state: Readiness["state"] = "ready"): Promise<{ el: HTMLElement; look: () => Promise<void>; banner: FinishBanner }> {
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: m.api }] });
    store = TestBed.inject(SessionStore);
    store.setSession(started());
    fixture = TestBed.createComponent(FinishBanner);
    fixture.componentRef.setInput("readiness", readiness(state, state === "blocked" ? 1 : 0));
    const look = async () => {
      fixture.detectChanges();
      await tick();
      fixture.detectChanges();
    };
    await look();
    await look();
    return { el: fixture.nativeElement as HTMLElement, look, banner: fixture.componentInstance };
  }

  const upload = (el: HTMLElement) => el.querySelector(".finish__upload") as HTMLButtonElement;
  const cancelButton = (el: HTMLElement, where = ".finish__working") => el.querySelector(`${where} mr-cancel button`) as HTMLButtonElement | null;
  const outcome = (el: HTMLElement) => texts(el, ".finish__outcome");

  it("row 7: Cancel stands beside the progress line, stops the run (kind mind), says Stopping…, then the server's sentence with the project and its Delete button", async () => {
    const m = mind();
    const { el, look, banner } = await draw(m);
    expect(el.querySelector("mr-cancel")).toBeNull();
    upload(el).click();
    m.statuses.push(run("running", { timing: { stage_title: "Uploading the workbook" } as never }));
    await until(() => m.calls.includes("start") && banner.progress() === "Uploading the workbook…");
    await look();
    expect(texts(el, ".finish__progress")).toEqual(["Uploading the workbook…"]);
    const cancel = cancelButton(el)!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    // beside the progress line: the same row
    expect(cancel.closest(".finish__working")?.querySelector(".finish__progress")).not.toBeNull();

    const answer = hold<CancelResult>();
    m.onCancel = () => answer.done;
    cancel.click();
    await look();
    expect(m.calls).toContain("cancel:mind");
    expect(texts(el, ".finish__progress")).toEqual(["Stopping…"]);
    expect(cancel.disabled).toBe(true);
    expect(cancel.textContent?.trim()).toBe("Stopping…");
    expect(banner.running()).toBe(true); // the look does not end at the click

    // the server: stopping, then stopped, with the project it keeps in Mind
    m.statuses.push(run("cancelling", { message: "Stopping…" }));
    await tick();
    await look();
    expect(texts(el, ".finish__progress")).toEqual(["Stopping…"]);
    m.statuses.push(run("cancelled", { message: STILL, project: project(), kept: [project()], earlier: [] }));
    answer.end(stopped({ cancelled: ["mind"], message: STILL, mind: project() }));
    await until(() => !banner.running());
    await look();
    expect(outcome(el)).toEqual([STILL]);
    expect(el.querySelector(".finish__outcome--error")).toBeNull();
    expect(el.querySelector(".finish__progress")).toBeNull();
    expect(el.querySelector("mr-cancel")).toBeNull();
    // under it: the project by name, with its button; nothing was deleted
    expect(texts(el, ".finish__project-name")).toEqual(["zzz_mindready_ab12"]);
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
    expect(el.querySelector(".finish__kept-lead")).toBeNull(); // the run's own sentence stands above the list
    expect(m.calls.some((c) => c.startsWith("delete"))).toBe(false);
    expect(upload(el).disabled).toBe(false);
    expect(upload(el).textContent?.trim()).toBe("Upload to Mind");
  });

  it("Cancel during the sign-in check: startMindLoop is never called, and the line says nothing was uploaded", async () => {
    const m = mind();
    m.check = hold<MindSignIn>();
    const { el, look, banner } = await draw(m);
    upload(el).click();
    await look();
    expect(texts(el, ".finish__progress")).toEqual(["Checking the Mind sign-in…"]);
    const cancel = cancelButton(el)!;
    cancel.click();
    await look();
    expect(m.calls).toContain("cancel:mind");
    // the check cannot be stopped half-way: the line says "Stopping…" until it is back
    expect(texts(el, ".finish__progress")).toEqual(["Stopping…"]);
    expect(cancelButton(el)!.textContent?.trim()).toBe("Stopping…");
    expect(banner.running()).toBe(true);
    m.check.end(signedIn(true)); // signed in: the upload would have started
    await look();
    await look();
    expect(m.calls).not.toContain("start");
    expect(m.calls).not.toContain("options");
    expect(outcome(el)).toEqual(["Stopped. Nothing was uploaded to Mind."]);
    expect(el.querySelector(".finish__outcome--error")).toBeNull();
    expect(banner.running()).toBe(false);
    expect(el.querySelector("mr-cancel")).toBeNull();
    expect(banner.askSignIn()).toBe(false);
  });

  it("Cancel during the sign-in check that then finds no sign-in: no question about signing in either", async () => {
    const m = mind();
    m.check = hold<MindSignIn>();
    const { el, look, banner } = await draw(m);
    upload(el).click();
    await look();
    cancelButton(el)!.click();
    m.check.end(signedIn(false));
    await look();
    await look();
    expect(banner.askSignIn()).toBe(false);
    expect(outcome(el)).toEqual(["Stopped. Nothing was uploaded to Mind."]);
    expect(m.calls).not.toContain("start");
  });

  it("Cancel while the run's options are read: startMindLoop is never called", async () => {
    const m = mind();
    m.options = hold<unknown>();
    const { el, look } = await draw(m);
    upload(el).click();
    await until(() => m.calls.includes("options"));
    await look();
    expect(texts(el, ".finish__progress")).toEqual(["Starting the upload…"]);
    cancelButton(el)!.click();
    await look();
    m.options.end({ choices: [] });
    await look();
    await look();
    expect(m.calls).not.toContain("start");
    expect(outcome(el)).toEqual(["Stopped. Nothing was uploaded to Mind."]);
  });

  it("Cancel while the sign-in window waits: the upload never starts when the sign-in is done, and the line says nothing was uploaded", async () => {
    const m = mind();
    m.check = hold<MindSignIn>();
    const { el, look, banner } = await draw(m);
    upload(el).click();
    m.check.end(signedIn(false));
    await look();
    await look();
    // not signed in: the banner asks; the person opens the window; the upload waits for it
    expect(banner.askSignIn()).toBe(true);
    expect(el.querySelector(".finish__signin mr-cancel")).toBeNull(); // nothing waits yet
    (el.querySelector(".finish__signin-open") as HTMLButtonElement).click();
    await look();
    await look();
    expect(m.calls).toContain("window");
    const cancel = cancelButton(el, ".finish__signin")!;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    cancel.click();
    await look();
    expect(m.calls).toContain("cancel:mind"); // the server closes the window
    expect(outcome(el)).toEqual(["Stopped. Nothing was uploaded to Mind."]);
    expect(el.querySelector(".finish__signin")).toBeNull();
    // the person signs in anyway, a moment later: nothing starts by itself any more
    m.window = signedIn(true);
    const checks = m.calls.filter((c) => c === "check").length;
    for (let i = 0; i < 20; i++) await tick();
    await look();
    expect(m.calls).not.toContain("start");
    expect(m.calls.filter((c) => c === "check")).toHaveLength(checks);
    expect(banner.running()).toBe(false);
    expect(outcome(el)).toEqual(["Stopped. Nothing was uploaded to Mind."]);
  });

  it("without a Cancel the sign-in window still leads to the upload by itself", async () => {
    const m = mind();
    m.check = hold<MindSignIn>();
    const { el, look, banner } = await draw(m);
    upload(el).click();
    m.check.end(signedIn(false));
    await look();
    await look();
    (el.querySelector(".finish__signin-open") as HTMLButtonElement).click();
    await look();
    m.check = null; // the next check finds the sign-in
    m.window = signedIn(true);
    m.statuses.push(run("done", { report: { verdict: "converged", reason: "Mind ran it.", iterations: [] } as never }));
    await until(() => m.calls.includes("start") && !banner.running());
    await look();
    expect(outcome(el)).toEqual(["Mind ran it."]);
  });

  it("a click that reaches the server while the start is on its way is sent again once the run exists", async () => {
    const m = mind();
    m.start = hold<{ state: string }>();
    const { el, look, banner } = await draw(m);
    upload(el).click();
    await until(() => m.calls.includes("start"));
    await look();
    cancelButton(el)!.click(); // the server has no run to stop yet: "Nothing was running."
    await look();
    expect(m.calls.filter((c) => c === "cancel:mind")).toHaveLength(1);
    m.statuses.push(run("cancelled", { message: "Stopped. Nothing was uploaded to Mind.", project: null, kept: [] }));
    m.start.end({ state: "running" });
    await until(() => !banner.running());
    await look();
    expect(m.calls.filter((c) => c === "cancel:mind")).toHaveLength(2);
    expect(outcome(el)).toEqual(["Stopped. Nothing was uploaded to Mind."]);
    expect(el.querySelector(".finish__project")).toBeNull();
  });

  it("a Cancel that came too late: the run's result stands and is shown", async () => {
    const m = mind();
    const { el, look, banner } = await draw(m);
    upload(el).click();
    m.statuses.push(run("running", { timing: { stage_title: "Running the model" } as never }));
    await until(() => banner.progress() === "Running the model…");
    await look();
    m.onCancel = () => Promise.resolve(stopped({ late: ["mind"], message: "It had already finished." }));
    cancelButton(el)!.click();
    m.statuses.push(run("done", { report: { verdict: "converged", reason: "Mind ran it.", iterations: [] } as never }));
    await until(() => !banner.running());
    await look();
    expect(outcome(el)).toEqual(["Mind ran it."]);
    expect(banner.failed()).toBe(false);
    expect(banner.stoppingNow()).toBe(false);
  });

  /** A banner under a run that was just stopped: one project kept in Mind. */
  async function stoppedRun(m: Mind, extra: Partial<MindLoopStatus> = {}) {
    const d = await draw(m);
    upload(d.el).click();
    m.statuses.push(run("cancelled", { message: STILL, project: project(), kept: [project()], earlier: [], ...extra }));
    await until(() => m.calls.includes("start") && !d.banner.running());
    await d.look();
    return d;
  }

  const deleteButton = (el: HTMLElement) => el.querySelector(".finish__delete") as HTMLButtonElement | null;
  const answerOf = (result: MindDelete["result"], message: string): MindDelete => ({ name: "zzz_mindready_ab12", result, message, url: "https://mind.example/" });

  it("Delete it from Mind: 'deleted' says so and the button goes", async () => {
    const m = mind();
    const { el, look } = await stoppedRun(m);
    const answer = hold<MindDelete>();
    m.onDelete = () => answer.done;
    deleteButton(el)!.click();
    await look();
    // while it works: "Deleting…", greyed, and no second upload meanwhile
    expect(m.calls).toContain("delete:zzz_mindready_ab12");
    expect(deleteButton(el)!.textContent?.trim()).toBe("Deleting…");
    expect(deleteButton(el)!.disabled).toBe(true);
    expect(upload(el).disabled).toBe(true);
    // the server's list after the delete, and its sentence about the stopped run
    const DELETED = "Stopped. The project zzz_mindready_ab12 was deleted from Mind.";
    m.statuses.push(run("cancelled", { message: DELETED, project: { ...project(), deleted: "deleted" }, kept: [], earlier: [] }));
    answer.end(answerOf("deleted", "Deleted from Mind."));
    await look();
    await look();
    expect(deleteButton(el)).toBeNull();
    expect(outcome(el)).toEqual([DELETED]); // never "is still in Mind" beside "Deleted from Mind."
    // the project's line stays, with what Mind did, and no button
    expect(texts(el, ".finish__project-name")).toEqual(["zzz_mindready_ab12"]);
    expect(texts(el, ".finish__delete-note")).toEqual(["Deleted from Mind."]);
    expect(el.querySelector(".finish__open-mind")).toBeNull();
    expect(upload(el).disabled).toBe(false);
  });

  it("Delete it from Mind: 'not found' says there was nothing to delete and the button goes", async () => {
    const m = mind();
    const { el, look } = await stoppedRun(m);
    const NOT_THERE = "Stopped. The project zzz_mindready_ab12 is not in Mind.";
    m.onDelete = () => (m.statuses.push(run("cancelled", { message: NOT_THERE, kept: [] })), Promise.resolve(answerOf("not found", "It was not in Mind: nothing to delete.")));
    deleteButton(el)!.click();
    await look();
    await look();
    expect(deleteButton(el)).toBeNull();
    expect(texts(el, ".finish__delete-note")).toEqual(["It was not in Mind: nothing to delete."]);
    expect(outcome(el)).toEqual([NOT_THERE]);
  });

  it("Delete it from Mind: a project that was still being made is 'not in Mind yet': the button stays for another try", async () => {
    const m = mind();
    const making = project("zzz_mindready_ab12", "creating");
    const { el, look } = await stoppedRun(m, { message: "Stopped. A project named zzz_mindready_ab12 may be in Mind.", project: making, kept: [making] });
    expect(outcome(el)).toEqual(["Stopped. A project named zzz_mindready_ab12 may be in Mind."]);
    m.onDelete = () => Promise.resolve(answerOf("not found", "It is not in Mind yet. Try again in a minute.")); // the server still lists it
    deleteButton(el)!.click();
    await look();
    await look();
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
    expect(deleteButton(el)!.disabled).toBe(false);
    expect(texts(el, ".finish__delete-note")).toEqual(["It is not in Mind yet. Try again in a minute."]);
    expect(el.querySelector(".finish__open-mind")).toBeNull();
  });

  it("Delete it from Mind: 'still present' and 'refused' say to delete it by hand, keep the button, and offer Open Mind in a new tab", async () => {
    for (const result of ["still present", "refused"] as const) {
      const m = mind();
      const { el, look } = await stoppedRun(m);
      const byHand = "Mind did not delete it. Delete it by hand in Mind: open the Project manager, folder zzz_mindready.";
      m.onDelete = () => Promise.resolve(answerOf(result, byHand));
      deleteButton(el)!.click();
      await look();
      await look();
      expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
      expect(texts(el, ".finish__delete-note")).toEqual([byHand]);
      const link = el.querySelector(".finish__open-mind") as HTMLAnchorElement;
      expect(link.textContent?.trim()).toBe("Open Mind");
      expect(link.getAttribute("href")).toBe("https://mind.example/");
      expect(link.getAttribute("target")).toBe("_blank");
      fixture.destroy();
      TestBed.resetTestingModule();
    }
  });

  it("Delete it from Mind: 'unreachable' says to sign in and keeps the button, with no link", async () => {
    const m = mind();
    const { el, look } = await stoppedRun(m);
    m.onDelete = () => Promise.resolve(answerOf("unreachable", "Mind did not open. Sign in to Mind, then try again."));
    deleteButton(el)!.click();
    await look();
    await look();
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
    expect(texts(el, ".finish__delete-note")).toEqual(["Mind did not open. Sign in to Mind, then try again."]);
    expect(el.querySelector(".finish__open-mind")).toBeNull();
  });

  it("Delete it from Mind: when Mind is busy the server's words are shown and the button stays", async () => {
    const m = mind();
    const { el, look } = await stoppedRun(m);
    m.onDelete = () => Promise.reject({ status: 409, error: { detail: "Mind is busy. Try again in a moment." } });
    deleteButton(el)!.click();
    await look();
    await look();
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
    expect(texts(el, ".finish__delete-note")).toEqual(["Mind is busy. Try again in a moment."]);
  });

  it("Delete it from Mind: when the list cannot be read again, the answer itself decides whether the button goes", async () => {
    const m = mind();
    const { el, look } = await stoppedRun(m);
    (m.api as unknown as { mindLoopStatus: () => Promise<never> }).mindLoopStatus = () => Promise.reject(new Error("offline"));
    m.onDelete = () => Promise.resolve(answerOf("deleted", "Deleted from Mind."));
    deleteButton(el)!.click();
    await look();
    await look();
    expect(deleteButton(el)).toBeNull();
    expect(texts(el, ".finish__delete-note")).toEqual(["Deleted from Mind."]);
    expect(outcome(el)).toEqual([]); // the sentence that said "is still in Mind" is not left standing
  });

  it("names the projects of earlier rounds that are still in Mind, with no button", async () => {
    const m = mind();
    const { el } = await stoppedRun(m, { earlier: ["zzz_mindready_r1", "zzz_mindready_r2"] });
    expect(texts(el, ".finish__earlier")).toEqual(["Earlier rounds of this run are still in Mind: zzz_mindready_r1, zzz_mindready_r2."]);
    expect(el.querySelectorAll(".finish__delete")).toHaveLength(1); // only the stopped run's own project
  });

  it("every kept project has its own line and its own button", async () => {
    const m = mind();
    const { el } = await stoppedRun(m, { kept: [project("zzz_mindready_old1"), project()] });
    expect(texts(el, ".finish__project-name")).toEqual(["zzz_mindready_old1", "zzz_mindready_ab12"]);
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind", "Delete it from Mind"]);
  });

  it("when the banner opens it lists again what stopped runs left in Mind, and nothing was deleted without a click", async () => {
    const m = mind([run("idle", { kept: [project()], earlier: [] })]);
    const { el } = await draw(m);
    expect(texts(el, ".finish__kept-lead")).toEqual(["Still in Mind from a run you stopped:"]);
    expect(texts(el, ".finish__project-name")).toEqual(["zzz_mindready_ab12"]);
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
    expect(outcome(el)).toEqual([]);
    expect(m.calls).toEqual([]);
  });

  it("when the banner opens after a stopped run, that run's sentence stands above its project again", async () => {
    const m = mind([run("cancelled", { message: STILL, project: project(), kept: [project()], earlier: ["zzz_mindready_r1"] })]);
    const { el } = await draw(m);
    expect(outcome(el)).toEqual([STILL]);
    expect(el.querySelector(".finish__outcome--error")).toBeNull();
    expect(el.querySelector(".finish__kept-lead")).toBeNull();
    expect(texts(el, ".finish__delete")).toEqual(["Delete it from Mind"]);
    expect(texts(el, ".finish__earlier")).toEqual(["Earlier rounds of this run are still in Mind: zzz_mindready_r1."]);
    fixture.destroy();
    TestBed.resetTestingModule();
    // a stopped run that left nothing in Mind says nothing when the banner opens later
    const m2 = mind([run("cancelled", { message: "Stopped. Nothing was uploaded to Mind.", project: null, kept: [] })]);
    const second = await draw(m2);
    expect(outcome(second.el)).toEqual([]);
    expect(second.el.querySelector(".finish__project")).toBeNull();
  });

  it("when the banner opens it follows a run that still runs, with Cancel beside it, and one that is stopping", async () => {
    const m = mind([run("running", { timing: { stage_title: "Running the model" } as never })]);
    const { el, look, banner } = await draw(m);
    expect(banner.running()).toBe(true);
    expect(texts(el, ".finish__progress")).toEqual(["Running the model…"]);
    expect(cancelButton(el)!.textContent?.trim()).toBe("Cancel");
    expect(upload(el).disabled).toBe(true);
    m.statuses.push(run("done", { report: { verdict: "converged", reason: "Mind ran it.", iterations: [] } as never }));
    await until(() => !banner.running());
    await look();
    expect(outcome(el)).toEqual(["Mind ran it."]);
    expect(m.calls).toEqual([]); // nothing was started by the banner itself
    fixture.destroy();
    TestBed.resetTestingModule();

    const m2 = mind([run("cancelling", { message: "Stopping…" })]);
    const second = await draw(m2);
    expect(texts(second.el, ".finish__progress")).toEqual(["Stopping…"]);
    expect(cancelButton(second.el)!.disabled).toBe(true);
    expect(cancelButton(second.el)!.textContent?.trim()).toBe("Stopping…");
    m2.statuses.push(run("cancelled", { message: STILL, project: project(), kept: [project()] }));
    await until(() => !second.banner.running());
    await second.look();
    expect(outcome(second.el)).toEqual([STILL]);
    expect(texts(second.el, ".finish__delete")).toEqual(["Delete it from Mind"]);
  });

  it("offers no Cancel while nothing runs, whatever the verdict", async () => {
    const m = mind();
    const { el } = await draw(m, "blocked");
    expect(el.querySelector("mr-cancel")).toBeNull();
    expect(upload(el).disabled).toBe(true);
  });
});

// ---- row 18: the strip in the shell --------------------------------------------------------------------
describe("the strip under the header: work left running by a page that was left", () => {
  let Shell: typeof App;
  let fixture: ComponentFixture<App>;
  let store: SessionStore;
  let work: WorkService;
  const everyMs = WorkService.everyMs;

  beforeAll(async () => {
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Shell = (await import("../app")).App;
  });

  beforeEach(() => (WorkService.everyMs = 1));
  afterEach(() => {
    work?.stop();
    fixture?.destroy();
    WorkService.everyMs = everyMs;
  });

  const job = (kind: string, title: string, stopping = false): WorkJob => ({ kind, title, started: 1_763_330_000, elapsed_s: 12, stopping });

  interface Fake {
    api: Partial<ApiService>;
    jobs: WorkJob[];
    asked: string[];
    cancels: [string, string | null][];
    onCancel: (() => Promise<CancelResult>) | null;
  }

  function fake(jobs: WorkJob[] = []): Fake {
    const f: Fake = { api: {}, jobs, asked: [], cancels: [], onCancel: null };
    f.api = {
      health: () => Promise.resolve({ version: "1.15.0" }),
      plainCodes: () => Promise.resolve({ rules: {}, errors: {} }),
      jobs: (id: string) => (f.asked.push(id), Promise.resolve([...f.jobs])),
      cancel: (id: string, kind: string | null = null) => (f.cancels.push([id, kind]), f.onCancel ? f.onCancel() : Promise.resolve(stopped())),
    } as unknown as Partial<ApiService>;
    return f;
  }

  async function draw(f: Fake, open = true): Promise<{ el: HTMLElement; look: () => Promise<void> }> {
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: f.api }] });
    TestBed.overrideComponent(Shell, { set: { imports: [CancelButton], schemas: [NO_ERRORS_SCHEMA] } });
    store = TestBed.inject(SessionStore);
    work = TestBed.inject(WorkService);
    if (open) store.setSession(started());
    fixture = TestBed.createComponent(Shell);
    const look = async () => {
      fixture.detectChanges();
      await tick();
      fixture.detectChanges();
    };
    await look();
    await look();
    return { el: fixture.nativeElement as HTMLElement, look };
  }

  const strip = (el: HTMLElement) => el.querySelector(".shell__work");

  it("row 18: shows 'Still working: …' with Cancel when a job runs and no Cancel is drawn; the click stops that work, and the strip goes", async () => {
    const f = fake([job("autofix", "Fixing the formula errors")]);
    const { el, look } = await draw(f);
    await until(() => f.asked.length > 0);
    await look();
    expect(f.asked[0]).toBe("m1");
    expect(texts(el, ".shell__work-text")).toEqual(["Still working: Fixing the formula errors."]);
    expect(strip(el)?.getAttribute("role")).toBe("status");
    const cancel = el.querySelector(".shell__work mr-cancel button") as HTMLButtonElement;
    expect(cancel.textContent?.trim()).toBe("Cancel");
    expect(cancel.classList.contains("cds--btn--sm")).toBe(true);
    // the strip's own button is not "a Cancel on screen for the work": it would hide the strip it stands in
    expect(work.shown()).toBe(0);
    // above the page itself
    expect(strip(el)!.compareDocumentPosition(el.querySelector("router-outlet")!) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();

    f.onCancel = () => ((f.jobs = []), Promise.resolve(stopped({ cancelled: ["autofix"], session: started(2) })));
    cancel.click();
    await look();
    await look();
    // (17 Nov 2026: the work it names, and only that: it said null, everything for this workbook, which
    // would also stop a step of Start while the strip shows a Run in Mind left on another screen)
    expect(f.cancels).toEqual([["m1", "autofix"]]);
    expect(strip(el)).toBeNull();
    expect(store.currentVersion()?.label).toBe("v2"); // the workbook as the server holds it now
  });

  it("shows no strip while nothing runs, and none without a workbook", async () => {
    const f = fake([]);
    const { el, look } = await draw(f);
    await until(() => f.asked.length > 1);
    await look();
    expect(strip(el)).toBeNull();
    fixture.destroy();
    work.stop();
    TestBed.resetTestingModule();

    const g = fake([job("apply", "Writing the changes")]);
    const closed = await draw(g, false);
    for (let i = 0; i < 10; i++) await tick();
    await closed.look();
    expect(g.asked).toEqual([]); // no workbook is open: nothing is asked
    expect(strip(closed.el)).toBeNull();
  });

  it("shows no strip while a Cancel for that work is drawn on the page, and shows it once that page is left", async () => {
    const f = fake([job("apply", "Writing the changes")]);
    const { el, look } = await draw(f);
    // the page that started the work draws its own Cancel (for everything, or for that kind of work)
    const everything = work.show(() => null);
    await until(() => f.asked.length > 0);
    await look();
    expect(strip(el)).toBeNull();
    everything();
    const own = work.show(() => "apply");
    await look();
    expect(strip(el)).toBeNull();
    // the page is left: its Cancel goes, the work goes on, the strip comes
    own();
    await look();
    expect(texts(el, ".shell__work-text")).toEqual(["Still working: Writing the changes."]);
  });

  it("a Cancel for another kind of work does not hide the strip: the job on the server would have no Cancel at all", async () => {
    const f = fake([job("autofix", "Fixing the formula errors")]);
    const { el, look } = await draw(f);
    const other = work.show(() => "chat"); // the assistant is answering on screen, with its own Cancel
    await until(() => f.asked.length > 0);
    await look();
    expect(texts(el, ".shell__work-text")).toEqual(["Still working: Fixing the formula errors."]);
    other();
  });

  it("follows the workbook that is open: another workbook's jobs are asked for, and a closed one empties the strip", async () => {
    const f = fake([job("apply", "Writing the changes")]);
    const { el, look } = await draw(f);
    await until(() => f.asked.includes("m1"));
    await look();
    expect(strip(el)).not.toBeNull();
    store.reset();
    await look();
    await look();
    expect(strip(el)).toBeNull();
    const before = f.asked.length;
    for (let i = 0; i < 10; i++) await tick();
    expect(f.asked.length).toBe(before); // nothing is asked while no workbook is open
  });

  it("while the strip's Cancel is stopping a job that takes long, the strip stays and says so", async () => {
    const f = fake([job("mind", "Uploading to Mind")]);
    const { el, look } = await draw(f);
    await until(() => f.asked.length > 0);
    await look();
    const answer = hold<CancelResult>();
    f.onCancel = () => ((f.jobs = [job("mind", "Uploading to Mind", true)]), answer.done);
    const cancel = el.querySelector(".shell__work mr-cancel button") as HTMLButtonElement;
    cancel.click();
    await look();
    expect(cancel.disabled).toBe(true);
    expect(cancel.textContent?.trim()).toBe("Stopping…");
    expect(texts(el, ".shell__work-text")).toEqual(["Still working: Uploading to Mind."]);
    f.jobs = [];
    answer.end(stopped({ cancelled: ["mind"] }));
    await until(() => f.jobs.length === 0 && work.jobs().length === 0);
    await look();
    expect(strip(el)).toBeNull();
  });
});
