import { provideHttpClient } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { ComponentFixture, TestBed } from "@angular/core/testing";
import { ApiService, wasCancelled } from "./api.service";
import { CancelButton } from "./cancel-button";
import { CancelResult, ChatStreamEvent, START_KINDS, SessionStart, WorkJob, chainStoppedLine, mindDeleteByHand, mindDeleteGone, mindEarlierLine, mindStoppedLine } from "./models";
import { SessionStore } from "./session.store";
import { WorkService } from "./work.service";

// ---- 17 Nov 2026: the Cancel button, the same one wherever the app works -------------------------------
// One click stops the work on the server; the workbook goes back to what it was before that work began.

const session = (id: string, version: number, recalc = true): SessionStart => ({
  sessionId: id,
  summary: { file_name: `${id}.xlsm`, file_type: "xlsm", sheet_count: 2, grid_count: 1, formula_count: 10 },
  report: { status: "PASS", findings: [] },
  plan: [],
  version: { id: `ver-00${version}`, major: version, minor: 0, label: `v${version}`, file_name: `${id}_v${version}.xlsm`, source: "prep" },
  versions: Array.from({ length: version }, (_x, i) => ({ id: `ver-00${i + 1}`, major: i + 1, minor: 0, label: `v${i + 1}`, file_name: `${id}_v${i + 1}.xlsm`, source: "prep" })),
  recalc: recalc ? { status: "PASS", message: "clean", ran: true, formula_errors: [], addin_gap_errors: [], version_id: `ver-00${version}` } : null,
});

const result = (extra: Partial<CancelResult> = {}): CancelResult => ({ cancelled: [], late: [], stopping: [], ahead: false, message: "Stopped. Nothing was changed.", session: null, mind: null, ...extra });

const job = (kind: string, stopping = true): WorkJob => ({ kind, title: "Writing the changes", started: 1_763_330_000, elapsed_s: 41.2, stopping });

const settle = () => new Promise<void>((r) => setTimeout(r, 0));

/** Something the test ends when it chooses: the server's answer to the click. */
function hold<T>(): { done: Promise<T>; end: (value: T) => void; fail: (error: unknown) => void } {
  let end!: (value: T) => void;
  let fail!: (error: unknown) => void;
  const done = new Promise<T>((res, rej) => {
    end = res;
    fail = rej;
  });
  return { done, end, fail };
}

describe("the Cancel button", () => {
  let store: SessionStore;
  let work: WorkService;
  let fixture: ComponentFixture<CancelButton>;
  /** what the page around the button hears, in order, with what the store held at that moment */
  let heard: string[];
  const everyMs = CancelButton.everyMs;

  afterEach(() => {
    CancelButton.everyMs = everyMs;
    fixture?.destroy();
  });

  function draw(api: Partial<ApiService>, inputs: Record<string, unknown> = {}, open: string | null = "m1"): { el: HTMLElement; button: () => HTMLButtonElement; look: () => Promise<void> } {
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: api }] });
    store = TestBed.inject(SessionStore);
    work = TestBed.inject(WorkService);
    if (open) store.setSession(session(open, 5));
    fixture = TestBed.createComponent(CancelButton);
    for (const [name, value] of Object.entries(inputs)) fixture.componentRef.setInput(name, value);
    heard = [];
    const at = () => store.currentVersion()?.label ?? "none";
    fixture.componentInstance.stopping.subscribe(() => heard.push(`stopping at ${at()}`));
    fixture.componentInstance.stopped.subscribe((r) => heard.push(`stopped at ${at()}: ${r.message}`));
    fixture.componentInstance.failed.subscribe((why) => heard.push(`failed: ${why}`));
    fixture.detectChanges();
    const el = fixture.nativeElement as HTMLElement;
    const look = async () => {
      fixture.detectChanges();
      await settle();
      fixture.detectChanges();
    };
    return { el, button: () => el.querySelector("button") as HTMLButtonElement, look };
  }

  it("is one red-outlined Carbon button that says Cancel, in the size asked for, and Show me can point at it", () => {
    const { button } = draw({}, { size: "sm" });
    expect(button().textContent?.trim()).toBe("Cancel");
    expect(button().disabled).toBe(false);
    expect(button().getAttribute("data-show-me")).toBe("cancel");
    expect(button().getAttribute("type")).toBe("button");
    expect([...button().classList]).toEqual(expect.arrayContaining(["cds--btn", "cds--btn--danger--tertiary", "cds--btn--sm"]));
    fixture.componentRef.setInput("size", "lg");
    fixture.detectChanges();
    expect(button().classList.contains("cds--btn--lg")).toBe(true);
    expect(button().classList.contains("cds--btn--sm")).toBe(false);
  });

  it("one click is one call with its kind; while it waits it is greyed and says Stopping…, and a second click does nothing", async () => {
    const answer = hold<CancelResult>();
    const calls: unknown[][] = [];
    const { button, look } = draw({ cancel: (id: string, kind: string | null = null) => (calls.push([id, kind]), answer.done) }, { kind: "apply" });
    button().click();
    await look();
    expect(calls).toEqual([["m1", "apply"]]);
    expect(button().disabled).toBe(true);
    expect(button().textContent?.trim()).toBe("Stopping…");
    expect(fixture.componentInstance.working()).toBe(true);
    // a second click (a greyed button takes none; the method itself refuses too)
    button().click();
    await fixture.componentInstance.cancel();
    expect(calls).toHaveLength(1);
    expect(heard).toEqual(["stopping at v5"]);
    answer.end(result({ cancelled: ["apply"] }));
    await look();
    expect(button().disabled).toBe(false);
    expect(button().textContent?.trim()).toBe("Cancel");
    expect(heard).toEqual(["stopping at v5", "stopped at v5: Stopped. Nothing was changed."]);
  });

  it("with no kind it stops everything that runs for the workbook", async () => {
    const calls: unknown[][] = [];
    const { button, look } = draw({ cancel: (id: string, kind: string | null = null) => (calls.push([id, kind]), Promise.resolve(result())) });
    expect(fixture.componentInstance.kind()).toBeNull();
    button().click();
    await look();
    expect(calls).toEqual([["m1", null]]);
  });

  it("says stopping at the click, before the server is asked", async () => {
    const order: string[] = [];
    const { button, look } = draw({ cancel: () => (order.push("the server is asked"), Promise.resolve(result())) }, { kind: "recalc" });
    fixture.componentInstance.stopping.subscribe(() => order.push("stopping"));
    button().click();
    await look();
    expect(order).toEqual(["stopping", "the server is asked"]);
  });

  it("puts the workbook as the server holds it now into the store, and only then says stopped", async () => {
    const restored = session("m1", 3);
    const { button, look } = draw({ cancel: () => Promise.resolve(result({ cancelled: ["apply"], session: restored })) }, { kind: "apply" });
    store.recalc.set(null); // the page had no recalculation on screen
    expect(store.currentVersion()?.label).toBe("v5");
    button().click();
    await look();
    // the version the cancelled work had made is gone, and the last good recalculation is back
    expect(store.currentVersion()?.label).toBe("v3");
    expect(store.versions().map((v) => v.label)).toEqual(["v1", "v2", "v3"]);
    expect(store.recalc()?.version_id).toBe("ver-003");
    expect(heard).toEqual(["stopping at v5", "stopped at v3: Stopped. Nothing was changed."]);
  });

  it("empties the recalculation on screen when the workbook as it is now has none", async () => {
    const { button, look } = draw({ cancel: () => Promise.resolve(result({ session: session("m1", 3, false) })) });
    expect(store.recalc()).not.toBeNull();
    button().click();
    await look();
    expect(store.recalc()).toBeNull();
  });

  it("while kinds are still stopping it stays on Stopping…, looks at the jobs until they are gone, then reads the workbook, and only then says stopped", async () => {
    CancelButton.everyMs = 1;
    const order: string[] = [];
    const jobs = [[job("apply"), job("scan", false)], [job("apply")], [job("scan", false)]]; // the scan is another job's: not waited for
    const { button, look } = draw(
      {
        cancel: () => (order.push("cancel"), Promise.resolve(result({ stopping: ["apply"], message: "Stopping…", session: null }))),
        jobs: () => (order.push("jobs"), Promise.resolve(jobs.shift() ?? [])),
        resume: () => (order.push("resume"), Promise.resolve(session("m1", 3))),
      },
      { kind: "apply" },
    );
    button().click();
    await look();
    expect(button().textContent?.trim()).toBe("Stopping…");
    expect(button().disabled).toBe(true);
    for (let i = 0; i < 200 && fixture.componentInstance.working(); i++) await settle();
    await look();
    expect(order).toEqual(["cancel", "jobs", "jobs", "jobs", "resume"]);
    expect(store.currentVersion()?.label).toBe("v3");
    // the store held the restored workbook before the page was told
    expect(heard).toEqual(["stopping at v5", "stopped at v3: Stopping…"]);
    expect(button().textContent?.trim()).toBe("Cancel");
  });

  it("a first check that is still stopping is waited for too; there is no workbook to read back, and it still says stopped", async () => {
    CancelButton.everyMs = 1;
    const order: string[] = [];
    const jobs = [[job("scan")], []];
    const { button, look } = draw(
      {
        cancel: () => (order.push("cancel"), Promise.resolve(result({ stopping: ["scan"], message: "Stopping…", session: null }))),
        jobs: (id: string) => (order.push(`jobs ${id}`), Promise.resolve(jobs.shift() ?? [])),
        resume: () => (order.push("resume"), Promise.reject({ error: { detail: "The workbook has not been analysed yet." } })),
      },
      { kind: "scan", session: "p9" },
      null,
    );
    button().click();
    for (let i = 0; i < 200 && order.length < 3; i++) await settle();
    await look();
    expect(order).toEqual(["cancel", "jobs p9", "jobs p9"]); // nothing is read back for a workbook that was never checked
    expect(heard).toEqual(["stopping at none", "stopped at none: Stopping…"]);
    expect(button().textContent?.trim()).toBe("Cancel");
  });

  it("when the server cannot be asked it says why, stops nothing, and is back to Cancel", async () => {
    const { button, look } = draw({ cancel: () => Promise.reject({ error: { detail: "This workbook is not open any more." } }) }, { kind: "apply" });
    button().click();
    await look();
    expect(heard).toEqual(["stopping at v5", "failed: This workbook is not open any more."]);
    expect(button().disabled).toBe(false);
    expect(button().textContent?.trim()).toBe("Cancel");
    expect(store.currentVersion()?.label).toBe("v5");
  });

  it("goes on after the page took it off the screen: the store still gets the workbook, and nothing is said to a page that is gone", async () => {
    const answer = hold<CancelResult>();
    const { button } = draw({ cancel: () => answer.done }, { kind: "apply" });
    button().click();
    fixture.destroy(); // the page's own request answered first and the page removed the button
    answer.end(result({ cancelled: ["apply"], session: session("m1", 3) }));
    await settle();
    expect(store.currentVersion()?.label).toBe("v3");
    expect(heard).toEqual(["stopping at v5"]);
  });

  it("takes the workbook it is given while the store holds none (the first check after an upload)", async () => {
    const calls: unknown[][] = [];
    const { button, look } = draw({ cancel: (id: string, kind: string | null = null) => (calls.push([id, kind]), Promise.resolve(result({ cancelled: ["scan"], session: null }))) }, { kind: "scan", session: "p9" }, null);
    button().click();
    await look();
    expect(calls).toEqual([["p9", "scan"]]);
    expect(store.sessionId()).toBeNull();
    expect(store.report()).toBeNull();
    expect(heard).toEqual(["stopping at none", "stopped at none: Stopped. Nothing was changed."]);
  });

  it("does nothing without a workbook", async () => {
    const calls: unknown[] = [];
    const { button, look } = draw({ cancel: () => (calls.push(1), Promise.resolve(result())) }, {}, null);
    button().click();
    await look();
    expect(calls).toEqual([]);
    expect(heard).toEqual([]);
  });

  it("never puts one workbook's state into the store of another", async () => {
    const answer = hold<CancelResult>();
    const { button } = draw({ cancel: () => answer.done }, { kind: "apply" });
    button().click();
    store.setSession(session("other", 2)); // another workbook was opened meanwhile
    answer.end(result({ session: session("m1", 3) }));
    await settle();
    expect(store.sessionId()).toBe("other");
    expect(store.currentVersion()?.label).toBe("v2");
  });

  it("stays greyed on Stopping… while the page says the work is still on its way out", async () => {
    const calls: unknown[] = [];
    const { button, look } = draw({ cancel: () => (calls.push(1), Promise.resolve(result())) }, { busy: true });
    expect(button().disabled).toBe(true);
    expect(button().textContent?.trim()).toBe("Stopping…");
    await fixture.componentInstance.cancel();
    expect(calls).toEqual([]);
    fixture.componentRef.setInput("busy", false);
    await look();
    expect(button().textContent?.trim()).toBe("Cancel");
  });

  // 17 Nov 2026 (review R2-6): a Cancel whose page was left kept asking GET .../jobs every second for as
  // long as the job ran (minutes, for a page of Mind that hangs), then read the workbook back
  it("stops looking at the jobs as soon as its page took it away, and reads nothing back", async () => {
    CancelButton.everyMs = 1;
    const order: string[] = [];
    const { button } = draw(
      {
        cancel: () => (order.push("cancel"), Promise.resolve(result({ stopping: ["mind"], message: "Stopping…", session: null }))),
        jobs: () => (order.push("jobs"), Promise.resolve([job("mind")])), // the job does not stop
        resume: () => (order.push("resume"), Promise.resolve(session("m1", 3))),
      },
      { kind: "mind" },
    );
    button().click();
    for (let i = 0; i < 50 && !order.includes("jobs"); i++) await settle();
    expect(order.slice(0, 2)).toEqual(["cancel", "jobs"]);
    fixture.destroy(); // the page was left
    const seen = order.length;
    for (let i = 0; i < 50; i++) await new Promise<void>((r) => setTimeout(r, 2));
    expect(order.length).toBeLessThanOrEqual(seen + 1); // at most the look that was on its way
    expect(order).not.toContain("resume");
    expect(store.currentVersion()?.label).toBe("v5");
    expect(heard).toEqual(["stopping at v5"]);
  });

  // 17 Nov 2026 (review R1-5): the Cancel beside Start stops Start's own steps only
  it("with a list of kinds it stops those kinds only, and counts on screen for each of them", async () => {
    const calls: unknown[][] = [];
    const { button, look } = draw({ cancel: (id: string, kind: unknown = null) => (calls.push([id, kind]), Promise.resolve(result())) }, { kind: START_KINDS });
    button().click();
    await look();
    expect(calls).toEqual([["m1", START_KINDS]]);
    // the strip: a Run in Mind or an assistant's answer left on another screen has no Cancel here
    work.jobs.set([job("apply"), job("mind")]);
    expect(work.left()?.kind).toBe("mind");
    work.jobs.set([job("recalc"), job("prep"), job("chat")]);
    expect(work.left()?.kind).toBe("chat");
    work.jobs.set([job("outside"), job("autofix"), job("scan")]);
    expect(work.left()).toBeNull();
  });

  it("counts as a Cancel on screen while it is drawn; the strip's own button does not", () => {
    draw({});
    expect(work.shown()).toBe(1);
    fixture.destroy();
    expect(work.shown()).toBe(0);
    TestBed.resetTestingModule();
    draw({}, { counted: false });
    expect(work.shown()).toBe(0);
  });
});

describe("a call that was stopped with Cancel", () => {
  it("is known by the server's own mark, and by nothing else", () => {
    expect(wasCancelled({ status: 409, error: { detail: "Stopped. Nothing was changed.", cancelled: true, kind: "apply" } })).toBe(true);
    expect(wasCancelled({ status: 409, error: { detail: "Wait a moment: changes are still being written." } })).toBe(false);
    expect(wasCancelled({ status: 409, error: { cancelled: "true" } })).toBe(false);
    expect(wasCancelled({ status: 500, error: null })).toBe(false);
    expect(wasCancelled(new Error("Stopped. Nothing was changed."))).toBe(false);
    expect(wasCancelled(null)).toBe(false);
    expect(wasCancelled(undefined)).toBe(false);
    expect(wasCancelled("cancelled")).toBe(false);
  });
});

describe("the words after a Cancel", () => {
  it("names the last step that finished, or says nothing had", () => {
    expect(chainStoppedLine(null)).toBe("Stopped. Your workbook is as it was before you pressed Do everything.");
    expect(chainStoppedLine("outside")).toBe("Stopped. Your workbook is as it was after the input files were brought in.");
    expect(chainStoppedLine("fixAll")).toBe("Stopped. Your workbook is as it was after the formula errors were fixed.");
    expect(chainStoppedLine("prepare")).toBe("Stopped. Your workbook is as it was after the formulas and the structure were fixed.");
    expect(chainStoppedLine("loop")).toBe("Stopped. Your workbook is as it was after the Mind loop was built.");
    expect(chainStoppedLine("recalc")).toBe("Stopped. Your workbook is as it was after the recalculation.");
  });

  it("says what a stopped run left in Mind: the server's sentence, or the same three built from the project", () => {
    const project = (status: string) => ({ name: "zzz_mindready_ab12", status, ran: false, deleted: null, url: "https://mind.example" });
    expect(mindStoppedLine({ message: "Stopped. The project zzz_mindready_ab12 is still in Mind." })).toBe("Stopped. The project zzz_mindready_ab12 is still in Mind.");
    expect(mindStoppedLine({ project: null })).toBe("Stopped. Nothing was uploaded to Mind.");
    expect(mindStoppedLine(null)).toBe("Stopped. Nothing was uploaded to Mind.");
    expect(mindStoppedLine({ project: project("creating") })).toBe("Stopped. A project named zzz_mindready_ab12 may be in Mind.");
    expect(mindStoppedLine({ project: project("model added") })).toBe("Stopped. The project zzz_mindready_ab12 is still in Mind.");
    expect(mindEarlierLine(["zzz_a", "zzz_b"])).toBe("Earlier rounds of this run are still in Mind: zzz_a, zzz_b.");
    expect(mindEarlierLine([])).toBe("");
    expect(mindEarlierLine(undefined)).toBe("");
  });

  it("knows after each of the five answers of Delete whether the button goes and whether Open Mind is offered", () => {
    const d = (r: string, message: string) => ({ name: "p", result: r, message, url: "https://mind.example" }) as never;
    expect(mindDeleteGone(d("deleted", "Deleted from Mind."))).toBe(true);
    expect(mindDeleteGone(d("not found", "It was not in Mind: nothing to delete."))).toBe(true);
    expect(mindDeleteGone(d("not found", "It is not in Mind yet. Try again in a minute."))).toBe(false);
    expect(mindDeleteGone(d("still present", "Mind did not delete it."))).toBe(false);
    expect(mindDeleteGone(d("refused", "Mind did not delete it."))).toBe(false);
    expect(mindDeleteGone(d("unreachable", "Mind did not open. Sign in to Mind, then try again."))).toBe(false);
    expect(mindDeleteGone(null)).toBe(false);
    expect(["deleted", "not found", "still present", "refused", "unreachable"].map((r) => mindDeleteByHand(d(r, "")))).toEqual([false, false, true, true, false]);
  });
});

describe("ApiService and Cancel", () => {
  let api: ApiService;
  let http: HttpTestingController;

  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideHttpClient(), provideHttpClientTesting()] });
    api = TestBed.inject(ApiService);
    http = TestBed.inject(HttpTestingController);
  });

  afterEach(() => {
    http.verify();
    vi.unstubAllGlobals();
  });

  it("asks the server to stop a list of kinds (Start's own steps) as {kinds: [...]}", async () => {
    const some = api.cancel("s1", START_KINDS);
    const req = http.expectOne("/api/sessions/s1/cancel");
    expect(req.request.body).toEqual({ kinds: ["outside", "autofix", "prep", "apply", "recalc", "scan"] });
    req.flush(result({ cancelled: ["prep"] }));
    expect((await some).cancelled).toEqual(["prep"]);
    expect(START_KINDS).not.toContain("mind");
    expect(START_KINDS).not.toContain("chat");
  });

  it("asks the server to stop one kind of work, or everything", async () => {
    const one = api.cancel("s 1", "apply");
    const first = http.expectOne("/api/sessions/s%201/cancel");
    expect(first.request.method).toBe("POST");
    expect(first.request.body).toEqual({ kind: "apply" });
    first.flush(result({ cancelled: ["apply"] }));
    expect((await one).cancelled).toEqual(["apply"]);
    const all = api.cancel("s1");
    const second = http.expectOne("/api/sessions/s1/cancel");
    expect(second.request.body).toEqual({});
    second.flush(result({ message: "Nothing was running." }));
    expect((await all).message).toBe("Nothing was running.");
  });

  it("reads what runs for the workbook as a plain list", async () => {
    const got = api.jobs("s1");
    const req = http.expectOne("/api/sessions/s1/jobs");
    expect(req.request.method).toBe("GET");
    req.flush({ jobs: [job("apply", false)] });
    expect(await got).toEqual([job("apply", false)]);
    const none = api.jobs("s1");
    http.expectOne("/api/sessions/s1/jobs").flush({});
    expect(await none).toEqual([]);
  });

  it("asks Mind to delete a project by its name", async () => {
    const got = api.deleteMindProject("s1", "zzz_mindready_ab12");
    const req = http.expectOne("/api/sessions/s1/mind-loop/delete-project");
    expect(req.request.method).toBe("POST");
    expect(req.request.body).toEqual({ name: "zzz_mindready_ab12" });
    req.flush({ name: "zzz_mindready_ab12", result: "deleted", message: "Deleted from Mind.", url: "https://mind.example" });
    expect((await got).message).toBe("Deleted from Mind.");
  });

  /** the server's stream, as the lines it sends */
  function streamOf(lines: object[], then?: () => Promise<never>): { ok: boolean; body: { getReader: () => { read: () => Promise<{ done: boolean; value?: Uint8Array }> } } } {
    const chunks = lines.map((l) => new TextEncoder().encode(JSON.stringify(l) + "\n"));
    return { ok: true, body: { getReader: () => ({ read: () => (chunks.length ? Promise.resolve({ done: false, value: chunks.shift() }) : then ? then() : Promise.resolve({ done: true })) }) } };
  }

  it("gives the assistant's whole answer, and null when the server stopped it with its cancelled line", async () => {
    const seen: ChatStreamEvent[] = [];
    vi.stubGlobal("fetch", () => Promise.resolve(streamOf([{ type: "phase", phase: "thinking" }, { type: "delta", text: "The " }, { type: "done", text: "The answer", proposal: null }])));
    expect(await api.chatStream("s1", [], "why?", undefined, (ev) => seen.push(ev))).toEqual({ text: "The answer", proposal: null, provenance: undefined });
    expect(seen.map((e) => e.type)).toEqual(["phase", "delta"]);
    vi.stubGlobal("fetch", () => Promise.resolve(streamOf([{ type: "delta", text: "The " }, { type: "cancelled" }])));
    expect(await api.chatStream("s1", [], "why?", undefined, () => undefined)).toBeNull();
  });

  it("gives null when the page aborted the answer, before the stream began or in the middle of it", async () => {
    const abort = () => Promise.reject(Object.assign(new Error("The operation was aborted."), { name: "AbortError" }));
    let given: AbortSignal | undefined;
    vi.stubGlobal("fetch", (_url: string, init: { signal?: AbortSignal }) => ((given = init.signal), abort()));
    const stop = new AbortController();
    expect(await api.chatStream("s1", [], "why?", undefined, () => undefined, stop.signal)).toBeNull();
    expect(given).toBe(stop.signal); // the request itself carries the signal
    vi.stubGlobal("fetch", () => Promise.resolve(streamOf([{ type: "delta", text: "The " }], abort)));
    expect(await api.chatStream("s1", [], "why?", undefined, () => undefined, stop.signal)).toBeNull();
    // any other failure is still a failure
    vi.stubGlobal("fetch", () => Promise.reject(new Error("offline")));
    await expect(api.chatStream("s1", [], "why?", undefined, () => undefined)).rejects.toThrow("offline");
    vi.stubGlobal("fetch", () => Promise.resolve(streamOf([{ type: "error", detail: "the gateway refused" }])));
    await expect(api.chatStream("s1", [], "why?", undefined, () => undefined)).rejects.toThrow("the gateway refused");
  });

  it("gives null for an answer the server refused as cancelled, and puts no problem box up", async () => {
    vi.stubGlobal("fetch", () => Promise.resolve({ ok: false, status: 409, statusText: "Conflict", body: null, json: () => Promise.resolve({ detail: "Stopped. Nothing was changed.", cancelled: true, kind: "chat" }) }));
    expect(await api.chatStream("s1", [], "why?", undefined, () => undefined)).toBeNull();
  });
});

// 17 Nov 2026 (live test L2b-3): when Start ends, its steps leave the list of running work at once
describe("the list of running work when the page's own work has ended", () => {
  it("drops those kinds now, keeps the others, and drops a look that was asked before", async () => {
    const answers: ReturnType<typeof hold<WorkJob[]>>[] = [];
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: { jobs: () => { const a = hold<WorkJob[]>(); answers.push(a); return a.done; } } }] });
    const store = TestBed.inject(SessionStore);
    const work = TestBed.inject(WorkService);
    store.setSession(session("m1", 5));
    work.jobs.set([job("recalc"), job("mind")]);
    work.ended(START_KINDS);
    expect(work.jobs().map((j) => j.kind)).toEqual(["mind"]);
    // the shell follows the workbook: a look on its way while the step still ran answers late
    work.watch();
    TestBed.tick();
    await settle();
    expect(answers.length).toBeGreaterThan(0);
    const early = answers[answers.length - 1]; // the first look, still on its way
    work.jobs.set([job("recalc"), job("mind")]); // the list as a look before it had left it
    work.ended(START_KINDS);
    expect(work.jobs().map((j) => j.kind)).toEqual(["mind"]);
    early.end([job("recalc"), job("mind")]);
    await settle();
    expect(work.jobs().map((j) => j.kind)).toEqual(["mind"]); // the old look is not taken
    answers[answers.length - 1].end([job("mind")]);
    await settle();
    expect(work.jobs().map((j) => j.kind)).toEqual(["mind"]);
    work.stop();
  });
});
