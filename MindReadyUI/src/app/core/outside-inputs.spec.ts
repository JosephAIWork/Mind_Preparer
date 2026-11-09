import { provideHttpClient } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { TestBed } from "@angular/core/testing";
import { provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import { ApiService } from "./api.service";
import {
  ChainEstimate, OutsideInputRow, OutsideInputs, OutsideRunResult, OutsideRunStatus, firstClause, inputFilesProblem, outsideButtonLabel, outsideOutcome, outsideReasonLine, outsideRowLabel,
  outsideSkipReason, pickedFilesProblem, withOutsideStep,
} from "./models";
import { SessionStore } from "./session.store";

const row = (cell: string, extra: Partial<OutsideInputRow> = {}): OutsideInputRow => ({ cell, expected: null, keys: [], target_sheet: null, file: null, how: null, candidates: [], ...extra });

/** The outside files as the server describes them: 2 of 2 matched, not brought in. */
const outside = (extra: Partial<OutsideInputs> = {}): OutsideInputs => ({
  imports: true,
  macro: "calcCF",
  calls: [],
  rows: [row("Ctrl_01!H35", { keys: ["BBA_01"], file: "2026Q1_2026_BBA_01.xlsm", how: "key" }), row("Ctrl_01!H36", { keys: ["BBA_02"], file: "2026Q1_2026_BBA_02.xlsm", how: "key" })],
  matched: 2,
  missing: [],
  extra: [],
  summary: "2 of 2 input files matched.",
  brought_in: false,
  brought_in_reason: null,
  ...extra,
});

const REFUSED = "Making the Mind inputs would change 3 computed values; the outside data stays inside the model (v2); nothing else was changed. For example Calc!A3 would show 0 instead of 120.";
const NOT_CHECKED = { ran: false, clean: null, verdict: "Excel is not available here: the numbers could not be compared." };

const result = (extra: Partial<OutsideRunResult> = {}): OutsideRunResult => ({
  import: {
    macro: "calcCF",
    seconds: 10.8,
    files: [
      { cell: "Ctrl_01!H35", file: "2026Q1_2026_BBA_01.xlsm", expected: "2026Q1_2026_BBA_01.xlsm", exact: true, opened: true, areas: [{ sheet: "BBA_01", ref: "A5:Z900" }] },
      { cell: "Ctrl_01!H36", file: "2026Q1_2026_BBA_02.xlsm", expected: "2026Q1_2026_BBA_02.xlsm", exact: true, opened: true, areas: [{ sheet: "BBA_02", ref: "A5:Z900" }] },
    ],
    kept: [],
    dialogs: [],
    stopped_early: false,
    error: null,
  },
  mindInputs: {
    inputs: [
      { name: "IN_BBA_01", sheet: "BBA_01", ref: "A5:Z900", rows: 896, cols: 26, csv: "IN_BBA_01.csv" },
      { name: "IN_BBA_02", sheet: "BBA_02", ref: "A5:Z900", rows: 896, cols: 26, csv: "IN_BBA_02.csv" },
    ],
    kept: [],
    formulas_changed: 46800,
    names_changed: 0,
    sums_split: 10800,
  },
  versions: ["ver-002", "ver-003"],
  download: { zip: "Mind inputs.zip", files: ["IN_BBA_01.csv", "IN_BBA_02.csv"] },
  summary: "2 outside files brought in; 2 Mind inputs made (IN_BBA_01, IN_BBA_02); no computed value changed.",
  ...extra,
});

describe("picking the model and its input files", () => {
  it("takes one workbook as before, and refuses one file that is not a workbook", () => {
    expect(pickedFilesProblem(["Model.xlsm"])).toBe("");
    expect(pickedFilesProblem(["data.csv"])).toBe("Choose an .xlsx, .xlsm or .xlsb workbook.");
    expect(pickedFilesProblem([])).toBe("");
  });

  it("takes the model with its input files in any language", () => {
    expect(pickedFilesProblem(["SBI生命_IFRS17計算ツール.xlsm", "2026Q1_2026_BBA_01.xlsm", "תחזית.csv", "data.txt", "old.xls"])).toBe("");
  });

  it("wants a workbook among several files, and only Excel or text files", () => {
    expect(pickedFilesProblem(["a.csv", "b.txt"])).toMatch(/model workbook/);
    expect(pickedFilesProblem(["Model.xlsm", "notes.pdf"])).toBe("notes.pdf is not an Excel or text file: pick .xlsx, .xlsm, .xlsb, .xls, .csv or .txt files.");
    expect(pickedFilesProblem(["Model.xlsm", "a.pdf", "b.docx"])).toMatch(/^a\.pdf, b\.docx are not Excel or text files/);
  });

  it("refuses the same name twice, as the server compares names (full-width, case)", () => {
    expect(inputFilesProblem(["Ｆｉｌｅ（１）.xlsm", "file(1).XLSM"])).toMatch(/^Two files are named/);
    expect(inputFilesProblem(["BBA_01.xlsm", "BBA_02.xlsm"])).toBe("");
  });
});

describe("the import table in words", () => {
  it("names a row by the file it expects, else its key, else its cell", () => {
    expect(outsideRowLabel(row("Ctrl_01!H35", { expected: "2026Q1_2026_BBA_01.xlsm", keys: ["BBA_01"] }))).toBe("2026Q1_2026_BBA_01.xlsm");
    expect(outsideRowLabel(row("Ctrl_01!H35", { keys: ["2026Q1", "BBA_01"] }))).toBe("2026Q1 / BBA_01");
    expect(outsideRowLabel(row("Ctrl_01!H35"))).toBe("Ctrl_01!H35");
  });
});

describe("the outcome of Bring in the outside files", () => {
  it("goes on with the server's summary when every file was read", () => {
    const o = outsideOutcome(result());
    expect(o.ok).toBe(true);
    expect(o.tone).toBe("ok");
    expect(o.line).toBe("2 outside files brought in; 2 Mind inputs made (IN_BBA_01, IN_BBA_02); no computed value changed.");
    expect(o.detail).toContain("2 of 2 files read by the macro calcCF (11 s)");
    expect(o.detail).toContain("46,800 formulas now read the Mind inputs (10,800 sums split in two)");
  });

  it("stops when the macro did not read a picked file", () => {
    const r = result();
    r.import!.files[1].opened = false;
    r.import!.stopped_early = true;
    const o = outsideOutcome(r);
    expect(o.ok).toBe(false);
    expect(o.tone).toBe("error");
    expect(o.line).toBe("The import macro stopped before it read every file (not read: 2026Q1_2026_BBA_02.xlsm)");
    expect(o.detail).toContain("1 of 2 files read");
  });

  it("stops when the app had to answer No to the macro's question", () => {
    const r = result();
    r.import!.stopped_early = true;
    r.import!.dialogs = [{ title: "Microsoft Excel", message: "上書きしますか？", buttons: ["はい(&Y)", "いいえ(&N)"], pressed: "いいえ(&N)" }];
    const o = outsideOutcome(r);
    expect(o.ok).toBe(false);
    expect(o.line).toBe("The import macro asked a question and the app answered いいえ, so it may have stopped early");
    expect(o.detail).toContain('1 message box answered: "上書きしますか？"');
  });

  it("goes on, with a warning, when an area stays inside the model", () => {
    const r = result();
    r.mindInputs!.kept = [{ sheet: "VFA_01", ref: "A5:Z900", name: "IN_VFA_01", why: "the name Prem covers part of it" }];
    const o = outsideOutcome(r);
    expect(o.ok).toBe(true);
    expect(o.tone).toBe("warn");
    expect(o.detail).toContain("stays inside the model: VFA_01!A5:Z900 (the name Prem covers part of it)");
  });

  it("says the server's notes in its words, instead of its own reading of the reports", () => {
    const r = result({ notes: ["1 area(s) stay inside the model (Mind cannot swap them): the name Prem covers part of it"] });
    r.mindInputs!.kept = [{ sheet: "VFA_01", ref: "A5:Z900", why: "the name Prem covers part of it" }];
    const o = outsideOutcome(r);
    expect(o.detail).toContain("1 area(s) stay inside the model (Mind cannot swap them)");
    expect(o.detail).not.toContain("stays inside the model: VFA_01");
  });

  it("stops when a computed value changed: values must never change", () => {
    const o = outsideOutcome(result({ numbers: { ran: true, clean: false }, summary: "2 outside files brought in; 2 Mind inputs made (IN_BBA_01, IN_BBA_02); 3 computed values changed: see Compare the numbers." }));
    expect(o.ok).toBe(false);
    expect(o.tone).toBe("error");
    expect(o.line).toMatch(/3 computed values changed/);
    expect(outsideOutcome(result({ numbers: { ran: true, clean: true } })).ok).toBe(true);
    expect(outsideOutcome(result()).ok).toBe(true); // no comparison asked for: nothing to stop on
  });

  it("stops in amber when the comparison of the computed values could not run, and says so plainly", () => {
    const summary = "2 outside files brought in; 2 Mind inputs made (IN_BBA_01, IN_BBA_02); the computed values were NOT checked (Excel is not available here: the numbers could not be compared).";
    const o = outsideOutcome(result({ numbers: NOT_CHECKED, summary }));
    expect(o.ok).toBe(false);
    expect(o.unchecked).toBe(true);
    expect(o.tone).toBe("warn");
    expect(o.line).toBe("The outside files are brought in, but the computed values were not checked");
    expect(o.detail.startsWith("Excel is not available here: the numbers could not be compared. Check them under Advanced › Deliver (Check the numbers).")).toBe(true);
    expect(o.detail).toContain(summary);
    expect(outsideOutcome(result({ numbers: { ran: true, clean: true } })).unchecked).toBe(false);
    // a macro that stopped early is red whatever the comparison
    const stopped = result({ numbers: NOT_CHECKED });
    stopped.import!.stopped_early = true;
    expect(outsideOutcome(stopped)).toMatchObject({ ok: false, tone: "error", unchecked: false });
  });
});

describe("the Outside files card in words", () => {
  it("takes the first clause of the server's message for the line, quotes kept whole", () => {
    expect(firstClause(REFUSED)).toBe("Making the Mind inputs would change 3 computed values");
    expect(firstClause('The workbook\'s import macro stopped early: "Cannot open. Check the path.". Not opened: a.xlsm. Nothing was changed.')).toBe(
      'The workbook\'s import macro stopped early: "Cannot open. Check the path."',
    );
    expect(firstClause("v3 holds the outside data, but the rest did not finish: a scan is already running for this session")).toBe(
      "v3 holds the outside data, but the rest did not finish: a scan is already running for this session",
    );
    expect(firstClause("x".repeat(300))).toHaveLength(160);
    expect(firstClause("x".repeat(300)).endsWith("…")).toBe(true);
    expect(firstClause("")).toBe("");
  });

  it("labels the button: Try again after a failed run, the matched files, Bring in again", () => {
    expect(outsideButtonLabel(outside(), false)).toBe("Bring in the outside files");
    expect(outsideButtonLabel(outside(), true)).toBe("Try again");
    expect(outsideButtonLabel(outside({ brought_in_reason: "the input files changed since they were brought in" }), false)).toBe("Bring in again");
    expect(outsideButtonLabel(outside({ brought_in: true }), false)).toBe("Bring in again");
    expect(outsideButtonLabel(outside({ matched: 1, missing: ["Ctrl_01!H36"] }), false)).toBe("Bring in the 1 matched file");
    expect(outsideButtonLabel(outside({ matched: 1, missing: ["Ctrl_01!H36"] }), true)).toBe("Try again");
  });

  it("runs the step in Do everything whenever the server says not brought in, after a failure too", () => {
    expect(outsideSkipReason(outside())).toBe("");
    expect(outsideSkipReason(outside({ brought_in_reason: "the last bring-in did not finish" }))).toBe("");
    expect(outsideSkipReason(outside({ brought_in_reason: "the input files changed since they were brought in" }))).toBe("");
    expect(outsideSkipReason(outside({ brought_in: true }))).toBe("Already brought in");
    expect(outsideSkipReason(outside({ brought_in: true }), true)).toBe("Already brought in (the computed values were not checked)");
    expect(outsideSkipReason(null)).toBe("The model's macros open no outside files");
    expect(outsideSkipReason(outside({ rows: [] }))).toMatch(/^No cell names the files/);
    expect(outsideSkipReason(outside({ macro: null }))).toMatch(/^Which macro brings the outside files in is not clear/);
  });

  it("says why the files are not brought in any more", () => {
    expect(outsideReasonLine(outside({ brought_in_reason: "the input files changed since they were brought in" }))).toBe("The input files changed since they were brought in: bring them in again.");
    expect(outsideReasonLine(outside({ brought_in_reason: "the last bring-in did not finish." }))).toBe("The last bring-in did not finish: bring them in again.");
    expect(outsideReasonLine(outside())).toBe("");
    expect(outsideReasonLine(outside({ brought_in: true, brought_in_reason: "ignored" }))).toBe("");
    expect(outsideReasonLine(null)).toBe("");
  });
});

describe("the estimate of Do everything", () => {
  const estimate: ChainEstimate = {
    steps: [{ id: "fixAll", low_s: 60, mid_s: 120, high_s: 600, why: "" }],
    total: { low_s: 60, mid_s: 120, high_s: 600 },
    caveat: "",
    basis: { file_mb: 10, scan_s: 40, scan_measured: true, stored_errors: 0, blocking_ops: 0 },
  };

  it("adds the outside step when it runs and the server left it out", () => {
    const e = withOutsideStep(estimate, true)!;
    expect(e.steps.map((s) => s.id)).toEqual(["outside", "fixAll"]);
    const step = e.steps[0];
    expect(step.low_s).toBeLessThan(step.mid_s);
    expect(step.mid_s).toBeLessThan(step.high_s);
    expect(e.total.mid_s).toBe(120 + step.mid_s);
  });

  it("keeps the server's estimate when the step does not run, or when the server counted it", () => {
    expect(withOutsideStep(estimate, false)).toBe(estimate);
    const counted: ChainEstimate = { ...estimate, steps: [{ id: "outside", low_s: 1, mid_s: 2, high_s: 3, why: "" }, ...estimate.steps] };
    expect(withOutsideStep(counted, true)).toBe(counted);
    expect(withOutsideStep(null, true)).toBeNull();
  });
});

describe("ApiService and the input files", () => {
  let api: ApiService;
  let http: HttpTestingController;

  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideHttpClient(), provideHttpClientTesting()] });
    api = TestBed.inject(ApiService);
    http = TestBed.inject(HttpTestingController);
  });

  afterEach(() => http.verify());

  it("sends one workbook as the field file, as before", async () => {
    const done = api.upload(new File(["x"], "Model.xlsm"), "plan", true);
    const req = http.expectOne("/api/sessions");
    const form = req.request.body as FormData;
    expect((form.get("file") as File).name).toBe("Model.xlsm");
    expect(form.getAll("files")).toEqual([]);
    expect(form.get("defer")).toBe("1");
    req.flush({ pending: true, sessionId: "s1" });
    await done;
  });

  it("sends several files as repeated files, names in any language", async () => {
    const names = ["SBI生命_IFRS17.xlsm", "2026Q1_2026_BBA_01.xlsm", "קלט.csv"];
    const done = api.upload(names.map((n) => new File(["x"], n)), "plan", true);
    const req = http.expectOne("/api/sessions");
    const form = req.request.body as FormData;
    expect(form.get("file")).toBeNull();
    expect((form.getAll("files") as File[]).map((f) => f.name)).toEqual(names);
    req.flush({ pending: true, sessionId: "s1", main: names[0], inputs: [{ name: names[1], bytes: 1 }, { name: names[2], bytes: 1 }] });
    const out = await done;
    expect(out.inputs?.length).toBe(2);
  });

  it("adds input files, starts the job and asks where it is", async () => {
    const added = api.addInputs("s 1", [new File(["x"], "2026Q1_2026_VFA_01.xlsm")]);
    const req = http.expectOne("/api/sessions/s%201/inputs");
    expect(req.request.method).toBe("POST");
    expect(((req.request.body as FormData).getAll("files") as File[]).map((f) => f.name)).toEqual(["2026Q1_2026_VFA_01.xlsm"]);
    req.flush({ inputs: [], outsideInputs: null });
    await added;

    const started = api.startOutsideInputs("s1");
    const run = http.expectOne("/api/sessions/s1/outside-inputs/run");
    expect(run.request.method).toBe("POST");
    run.flush({ running: true }, { status: 202, statusText: "Accepted" });
    expect((await started).running).toBe(true);

    const status = api.outsideInputs("s1");
    const get = http.expectOne("/api/sessions/s1/outside-inputs");
    expect(get.request.method).toBe("GET");
    get.flush({ outsideInputs: null, running: false, stage: null, result: null, error: null });
    expect((await status).running).toBe(false);
  });
});

describe("SessionStore and the outside files", () => {
  it("keeps the outside files of the upload when a later answer has none", () => {
    TestBed.configureTestingModule({});
    const store = TestBed.inject(SessionStore);
    const oi = { imports: true, macro: "calcCF", calls: [], rows: [row("Ctrl_01!H35")], matched: 0, missing: ["Ctrl_01!H35"], extra: [], summary: "0 of 1 input files matched; missing: Ctrl_01!H35.", brought_in: false };
    store.outsideInputs.set(oi);
    store.applyAnalysis({ outsideInputs: null });
    expect(store.outsideInputs()).toBe(oi);
    store.applyAnalysis({ outsideInputs: { ...oi, brought_in: true } });
    expect(store.outsideInputs()?.brought_in).toBe(true);
    // a corrected file added since: the server's answer says not brought in any more, and why
    store.applyAnalysis({ outsideInputs: { ...oi, brought_in: false, brought_in_reason: "the input files changed since they were brought in" } });
    expect(store.outsideInputs()?.brought_in).toBe(false);
    expect(store.outsideInputs()?.brought_in_reason).toBe("the input files changed since they were brought in");
    store.reset();
    expect(store.outsideInputs()).toBeNull();
  });
});

// ---- the main screen: the Outside files card and Do everything, against a fake server ----------------
interface Fake {
  api: ApiService;
  /** the answers of GET .../outside-inputs, in order (the last one repeats) */
  statuses: OutsideRunStatus[];
  starts: number;
  autoFixStarts: number;
}

function fakeServer(statuses: OutsideRunStatus[]): Fake {
  const f: Fake = { api: null as unknown as ApiService, statuses, starts: 0, autoFixStarts: 0 };
  let last: OutsideRunStatus | undefined;
  f.api = {
    outsideInputs: () => {
      last = f.statuses.shift() ?? last;
      return last ? Promise.resolve(last) : Promise.reject(new Error("no status"));
    },
    startOutsideInputs: () => {
      f.starts++;
      return Promise.resolve({ running: true });
    },
    // Do everything went on past the outside step: the fake stops it at the next one
    startAutoFix: () => {
      f.autoFixStarts++;
      return Promise.reject(new Error("the test stops here"));
    },
    estimate: () => Promise.reject(new Error("no estimate")),
    resume: () => Promise.reject(new Error("offline")),
    mindLoopPlan: () => Promise.reject(new Error("no loop plan")),
    downloadUrl: (id: string, name: string) => `/api/sessions/${id}/files/${encodeURIComponent(name)}`,
  } as unknown as ApiService;
  return f;
}

const status = (o: OutsideInputs, extra: Partial<OutsideRunStatus> = {}): OutsideRunStatus => ({ outsideInputs: o, running: false, state: "idle", stage: null, result: null, error: null, ...extra });
const running = (o: OutsideInputs) => status(o, { running: true, state: "running", stage: "Running the workbook's import macro" });

describe("the Outside files card and Do everything", () => {
  let Page: typeof HomePage;
  let page: HomePage;
  let store: SessionStore;

  beforeAll(async () => {
    // The main screen's logic, without drawing it: Carbon (a package Node cannot load as it is) stands in
    // as inert classes, whatever name the page asks for; the page is loaded after that.
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
  });

  function open(f: Fake, o: OutsideInputs): void {
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: f.api }] });
    store = TestBed.inject(SessionStore);
    page = TestBed.runInInjectionContext(() => new Page());
    (page as unknown as { sleep: () => Promise<void> }).sleep = () => Promise.resolve(); // no real waits between looks
    store.sessionId.set("s1");
    store.outsideInputs.set(o);
  }

  afterEach(() => page?.ngOnDestroy());

  it("runs the step again after a bring-in whose Mind inputs were refused, and offers Try again", async () => {
    const failed = outside({ brought_in_reason: "the last bring-in did not finish" });
    const f = fakeServer([status(failed, { state: "error", error: REFUSED }), running(failed), status(failed, { state: "error", error: REFUSED })]);
    open(f, failed);
    await page.runAll();
    expect(f.starts).toBe(1); // not skipped: the server says not brought in
    expect(page.steps()[0]).toMatchObject({ id: "outside", state: "failed", line: "Making the Mind inputs would change 3 computed values" });
    expect(page.chainLine()).toBe('Stopped at "Bring in the outside files": Making the Mind inputs would change 3 computed values. See the Outside files card above, then press Do everything again.');
    expect(f.autoFixStarts).toBe(0);
    // the card: the server's words, open, and a button to try again
    expect(page.outsideCard()).toMatchObject({ tone: "error", line: "Making the Mind inputs would change 3 computed values", detail: REFUSED, open: true });
    expect(page.outsideFailed()).toBe(true);
    expect(page.outsideButton(store.outsideInputs()!)).toBe("Try again");
    expect(page.zipUrl()).toBeNull();
  });

  it("follows the server after a reload: files changed since the bring-in come in again", async () => {
    const stale = outside({ brought_in_reason: "the input files changed since they were brought in" });
    const done = outside({ brought_in: true });
    // the last job's result (a success, before a corrected file was added), then Do everything
    const f = fakeServer([status(stale, { state: "done", result: result() }), status(stale, { state: "done", result: result() }), running(stale), status(done, { state: "done", result: result() })]);
    open(f, stale);
    await (page as unknown as { resumeOutside: () => Promise<void> }).resumeOutside();
    // not the old success: why, and the button to bring them in again; no stale CSVs offered
    expect(page.outsideCard()).toMatchObject({ tone: "warn", line: "The input files changed since they were brought in: bring them in again." });
    expect(page.outsideButton(stale)).toBe("Bring in again");
    expect(page.zipUrl()).toBeNull();
    expect(page.outsideWillRun()).toBe(true);
    await page.runAll();
    expect(f.starts).toBe(1);
    expect(page.steps()[0].state).toBe("done");
    expect(f.autoFixStarts).toBe(1); // and the chain went on
    expect(store.outsideInputs()?.brought_in).toBe(true);
    expect(page.outsideCard().tone).toBe("ok");
    expect(page.zipUrl()).toBe("/api/sessions/s1/files/Mind%20inputs.zip");
  });

  it("stops Do everything in amber when the computed values were not checked; going on is a choice", async () => {
    const before = outside();
    const after = outside({ brought_in: true });
    const unchecked = result({ numbers: NOT_CHECKED, summary: "2 outside files brought in; 2 Mind inputs made (IN_BBA_01, IN_BBA_02); the computed values were NOT checked (Excel is not available here: the numbers could not be compared)." });
    const f = fakeServer([status(before), running(before), status(after, { state: "done", result: unchecked })]);
    open(f, before);
    await page.runAll();
    expect(f.starts).toBe(1);
    expect(page.steps()[0]).toMatchObject({ state: "stopped", line: "The outside files are brought in, but the computed values were not checked" });
    expect(page.steps()[1]).toMatchObject({ id: "fixAll", state: "waiting" });
    expect(f.autoFixStarts).toBe(0);
    expect(page.chainLine()).toBe(
      'Stopped at "Bring in the outside files": the computed values were not checked. Check them under Advanced › Deliver (Check the numbers), or press Do everything again to go on without the check.',
    );
    expect(page.outsideCard().tone).toBe("warn");
    expect(page.outsideUnchecked()).toBe(true);
    expect(page.outsideFailed()).toBe(false);
    // on purpose: Do everything again (or "Go on without the check") skips the step and goes on
    f.statuses.push(status(after, { state: "done", result: unchecked }));
    await page.runAll();
    expect(f.starts).toBe(1);
    expect(page.steps()[0]).toMatchObject({ state: "skipped", line: "Already brought in (the computed values were not checked)" });
    expect(f.autoFixStarts).toBe(1);
  });
});
