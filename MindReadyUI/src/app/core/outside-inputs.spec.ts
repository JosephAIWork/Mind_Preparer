import { provideHttpClient } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { TestBed } from "@angular/core/testing";
import { provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import { ApiService } from "./api.service";
import {
  ChainEstimate, OutsideInputRow, OutsideInputs, OutsideRunResult, OutsideRunStatus, firstClause, inputFilesProblem, outsideButtonLabel, outsideOutcome, outsideReasonLine, outsideRowLabel,
  outsideAddLine, outsideSkipReason, outsideStandInLine, outsideStepEstimate, pickedFilesProblem, withOutsideStep,
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

describe("stand-in files made with the assistant, in words", () => {
  it("warns while stand-ins are in use: best results come with the actual files", () => {
    expect(outsideStandInLine(outside())).toBe("");
    expect(outsideStandInLine(null)).toBe("");
    expect(outsideStandInLine(outside({ stand_ins: ["2026Q1_2026_BBA_01.xlsm", "2026Q1_2026_BBA_02.xlsm"] }))).toBe(
      "All input files are stand-ins the AI assistant made from the numbers the model already holds. Best results come with the actual input files: add them, then bring them in again.",
    );
    expect(outsideStandInLine(outside({ stand_ins: ["2026Q1_2026_BBA_02.xlsm"] }))).toMatch(/^1 of 2 input files are stand-ins/);
  });

  it("says what to do when files are missing: add them, or let the assistant make them", () => {
    expect(outsideAddLine(outside({ assistant: true }))).toMatch(/^Add the model's input files .* or press "Make them with AI" there/);
    expect(outsideAddLine(outside({ assistant: true }))).toContain("best results come with the actual files");
    expect(outsideAddLine(outside({ assistant: false }))).toBe("Add the model's input files on the Outside files card (Add files).");
  });

  it("counts the assistant's work in the outside step's time", () => {
    const basis = { file_mb: 10, scan_s: 40, scan_measured: true, stored_errors: 0, blocking_ops: 0 };
    const plain = outsideStepEstimate(basis);
    const made = outsideStepEstimate(basis, true);
    expect(made.low_s).toBeGreaterThan(plain.low_s);
    expect(made.mid_s).toBeGreaterThan(plain.mid_s + 200);
    expect(made.why).toMatch(/^the AI assistant making the missing input files/);
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
    expect(run.request.body).toEqual({});
    run.flush({ running: true }, { status: 202, statusText: "Accepted" });
    expect((await started).running).toBe(true);

    const status = api.outsideInputs("s1");
    const get = http.expectOne("/api/sessions/s1/outside-inputs");
    expect(get.request.method).toBe("GET");
    get.flush({ outsideInputs: null, running: false, stage: null, result: null, error: null });
    expect((await status).running).toBe(false);

    const made = api.startOutsideInputs("s1", true);
    const withAi = http.expectOne("/api/sessions/s1/outside-inputs/run");
    expect(withAi.request.body).toEqual({ standIns: true });
    withAi.flush({ running: true }, { status: 202, statusText: "Accepted" });
    expect((await made).running).toBe(true);
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
  /** the standIns choice of each start */
  standIns: boolean[];
  /** the waitForInputs choice of each start (Do everything asks the server to wait for missing files) */
  waits: boolean[];
  /** POST .../outside-inputs/continue: the makeStandIns choice of each call */
  continues: boolean[];
  stops: number;
  /** the files added through POST .../inputs, by name */
  added: string[];
  autoFixStarts: number;
}

function fakeServer(statuses: OutsideRunStatus[], onContinue?: (makeStandIns: boolean) => OutsideRunStatus[] | Error, onStop?: () => OutsideRunStatus[]): Fake {
  const f: Fake = { api: null as unknown as ApiService, statuses, starts: 0, standIns: [], waits: [], continues: [], stops: 0, added: [], autoFixStarts: 0 };
  let last: OutsideRunStatus | undefined;
  f.api = {
    outsideInputs: () => {
      last = f.statuses.shift() ?? last;
      return last ? Promise.resolve(last) : Promise.reject(new Error("no status"));
    },
    startOutsideInputs: (_id: string, standIns = false, waitForInputs = false) => {
      f.starts++;
      f.standIns.push(standIns);
      f.waits.push(waitForInputs);
      return Promise.resolve({ running: !waitForInputs });
    },
    continueOutsideInputs: (_id: string, makeStandIns = false) => {
      f.continues.push(makeStandIns);
      const next = onContinue?.(makeStandIns);
      if (next instanceof Error) return Promise.reject(next);
      f.statuses.push(...(next ?? []));
      return Promise.resolve({ running: true });
    },
    stopOutsideInputs: () => {
      f.stops++;
      f.statuses.push(...(onStop?.() ?? []));
      return Promise.resolve({ running: false, state: "stopped" });
    },
    addInputs: (_id: string, files: File[]) => {
      f.added.push(...files.map((x) => x.name));
      // the server matches the files again: every row now has one
      const o = last?.outsideInputs ?? outside();
      const rows = o.rows.map((r) => ({ ...r, file: r.file ?? files.find((x) => x.name.includes(r.keys[0] ?? ""))?.name ?? r.file }));
      const matched = rows.filter((r) => r.file).length;
      const again: OutsideInputs = { ...o, rows, matched, missing: rows.filter((r) => !r.file).map((r) => r.cell) };
      if (last) last = { ...last, outsideInputs: again };
      return Promise.resolve({ inputs: [], outsideInputs: again });
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
/** 1.14.0: Do everything waits for the input files (the server's job holds, nothing runs) */
const waiting = (o: OutsideInputs, message: string) =>
  status(o, {
    state: "waiting_for_inputs",
    waiting: true,
    stage: "Waiting for the input files",
    message,
    stages: ["waiting"],
    needed: o.rows.map((r) => ({ name: outsideRowLabel(r), cell: r.cell, present: !!r.file, file: r.file, fills: r.target_sheet })),
  });
const STOPPED = "Stopped while waiting for the input files. Nothing was changed: pick the files, then press Do everything again.";
/** a real turn of the event loop: the page's look-every-second must not starve the test while the server waits */
const tick = () => new Promise<void>((r) => setTimeout(r, 0));
async function until(cond: () => boolean): Promise<void> {
  for (let i = 0; i < 500 && !cond(); i++) await tick();
  expect(cond()).toBe(true);
}

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
    TestBed.resetTestingModule(); // a test may open a second page against another fake
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

  // ---- 1.14.0: with files missing, Do everything stops at its first step and asks for them ----------
  /** one file of two picked; the server's job waits for the other */
  const someMissing = (assistant: boolean) =>
    outside({ rows: [row("Ctrl_01!H35", { keys: ["BBA_01"], file: "2026Q1_2026_BBA_01.xlsm", how: "key" }), row("Ctrl_01!H36", { keys: ["BBA_02"], target_sheet: "BBA_02" })], matched: 1, missing: ["Ctrl_01!H36"], assistant });

  function openWaiting(f: Fake, o: OutsideInputs): void {
    open(f, o);
    (page as unknown as { sleep: () => Promise<void> }).sleep = tick; // the wait is a look every turn, not a hot loop
  }

  it("with a file missing, Do everything waits: one box with the names, nothing runs, Go on greyed out, Stop ends it", async () => {
    const some = someMissing(false);
    const f = fakeServer([status(some), waiting(some, "Waiting for 1 input file: BBA_02")], undefined, () => [status(some, { state: "stopped", error: STOPPED })]);
    openWaiting(f, some);
    const run = page.runAll();
    await until(() => page.outsideWaiting());
    // the server was asked to wait, not to make stand-ins; the chain is held at its first step
    expect(f.starts).toBe(1);
    expect(f.waits).toEqual([true]);
    expect(f.standIns).toEqual([false]);
    expect(page.chainRunning()).toBe(true);
    expect(page.steps()[0]).toMatchObject({ id: "outside", state: "running" });
    expect(page.busyHint()).toBe("Do everything is waiting for the input files");
    // the box: the names, which are there, and the ways on (no AI key: no stand-ins offered)
    expect(page.waitingLine()).toBe("Do everything is waiting for 1 input file: BBA_02. Pick it.");
    expect(page.waitingFiles()).toEqual([{ name: "BBA_01", present: true, fills: null }, { name: "BBA_02", present: false, fills: "BBA_02" }]);
    expect(page.canGoOn()).toBe(false);
    expect(page.canMakeStandIns()).toBe(false);
    // the countdown pauses: nothing runs, so no "taking longer than usual" and no time counted against the step
    expect(page.timerLine()).toMatch(/so far · paused: waiting for the input files$/);
    // the card points at the box
    expect(page.outsideCard()).toMatchObject({ line: "Waiting for 1 input file: BBA_02", detail: "See the box under Do everything." });
    expect(f.autoFixStarts).toBe(0);
    // Go on does nothing while a file is missing
    await page.goOn();
    expect(f.continues).toEqual([]);
    // Stop: the server's job ends, the chain stops at this step, nothing ran
    await page.stopWaiting();
    await run;
    expect(f.stops).toBe(1);
    expect(page.outsideWaiting()).toBe(false);
    expect(page.chainRunning()).toBe(false);
    expect(page.steps()[0]).toMatchObject({ state: "failed", line: "Stopped while waiting for the input files" });
    expect(page.chainLine()).toBe('Stopped at "Bring in the outside files": Stopped while waiting for the input files. See the Outside files card above, then press Do everything again.');
    expect(f.autoFixStarts).toBe(0);
    expect(page.outsideFailed()).toBe(true);
  });

  it("the files picked in the box, then Go on: the bring-in runs from where it stopped and the other steps follow", async () => {
    const some = someMissing(true);
    const done = outside({ brought_in: true, assistant: true });
    const f = fakeServer([status(some), waiting(some, "Waiting for 1 input file: BBA_02")], () => [running(some), status(done, { state: "done", result: result() })]);
    openWaiting(f, some);
    const run = page.runAll();
    await until(() => page.outsideWaiting());
    expect(page.canMakeStandIns()).toBe(true); // an AI key is saved: the other way on is offered
    expect(page.waitingLine()).toBe("Do everything is waiting for 1 input file: BBA_02. Pick it, or let the assistant make stand-ins.");
    // picking is allowed while it waits (the other buttons are not)
    await page.addInputs([new File(["x"], "2026Q1_2026_BBA_02.xlsm")]);
    expect(f.added).toEqual(["2026Q1_2026_BBA_02.xlsm"]);
    expect(page.inputsError()).toBe("");
    expect(page.waitingMissing()).toEqual([]);
    expect(page.waitingLine()).toBe("Every input file is there: press Go on.");
    expect(page.canGoOn()).toBe(true);
    expect(page.canMakeStandIns()).toBe(false); // nothing is missing: nothing for the assistant to make
    // a double click sends one continue: the second click finds the first on its way
    const first = page.goOn();
    expect(page.goingOn()).toBe(true);
    await page.goOn();
    await first;
    expect(f.continues).toEqual([false]);
    expect(page.goingOn()).toBe(false);
    expect(page.outsideWaiting()).toBe(false); // the box went with the server's yes, not a second later
    expect(page.inputsError()).toBe("");
    await run;
    expect(f.starts).toBe(1); // the step was not started over: the server's job went on
    expect(f.standIns).toEqual([false]);
    expect(page.outsideWaiting()).toBe(false);
    expect(page.steps()[0].state).toBe("done");
    expect(f.autoFixStarts).toBe(1); // and the chain went on
    expect(store.outsideInputs()?.brought_in).toBe(true);
  });

  it("Make them with AI in the box takes the stand-in path on purpose, with its warning", async () => {
    const alone = outside({ rows: [row("Ctrl_01!H35", { keys: ["BBA_01"] }), row("Ctrl_01!H36", { keys: ["BBA_02"] })], matched: 0, missing: ["Ctrl_01!H35", "Ctrl_01!H36"], assistant: true });
    const made = outside({
      rows: [row("Ctrl_01!H35", { keys: ["BBA_01"], file: "2026Q1_2026_BBA_01.xlsm", stand_in: true }), row("Ctrl_01!H36", { keys: ["BBA_02"], file: "2026Q1_2026_BBA_02.xlsm", stand_in: true })],
      matched: 2, brought_in: true, assistant: true, stand_ins: ["2026Q1_2026_BBA_01.xlsm", "2026Q1_2026_BBA_02.xlsm"],
    });
    const f = fakeServer([status(alone), waiting(alone, "Waiting for 2 input files: BBA_01, BBA_02")], (ai) =>
      ai ? [running(alone), status(made, { state: "done", result: result() })] : new Error("2 input files still missing: BBA_01, BBA_02. Pick them, or let the assistant make stand-ins."),
    );
    openWaiting(f, alone);
    const run = page.runAll();
    await until(() => page.outsideWaiting());
    expect(f.standIns).toEqual([false]); // nothing made quietly
    expect(page.waitingLine()).toBe("Do everything is waiting for 2 input files: BBA_01, BBA_02. Pick them, or let the assistant make stand-ins.");
    expect(page.canGoOn()).toBe(false);
    await page.goOn(true);
    expect(f.continues).toEqual([true]);
    await run;
    expect(page.steps()[0].state).toBe("done");
    expect(f.autoFixStarts).toBe(1);
    expect(page.standInLine(store.outsideInputs()!)).toMatch(/^All input files are stand-ins the AI assistant made/);
  });

  it("a Go on the server refuses keeps the box, with its words", async () => {
    const some = someMissing(true);
    const f = fakeServer([status(some), waiting(some, "Waiting for 1 input file: BBA_02")], () => new Error("1 input file still missing: BBA_02. Pick it, or let the assistant make stand-ins."), () => [status(some, { state: "stopped", error: STOPPED })]);
    openWaiting(f, some);
    const run = page.runAll();
    await until(() => page.outsideWaiting());
    await page.goOn(true); // the assistant's try is refused by the server (say, the key expired)
    expect(page.outsideWaiting()).toBe(true);
    expect(page.inputsError()).toBe("1 input file still missing: BBA_02. Pick it, or let the assistant make stand-ins.");
    await page.stopWaiting();
    await run;
    expect(page.outsideWaiting()).toBe(false);
  });

  it("after a reload the box is back; Go on starts Do everything again, which goes straight through the step", async () => {
    const all = outside({ assistant: false }); // the files were picked before the reload
    const f = fakeServer([waiting(all, "Every input file is there: press Go on."), waiting(all, "Every input file is there: press Go on."), running(all), status(outside({ brought_in: true }), { state: "done", result: result() })]);
    openWaiting(f, all);
    await (page as unknown as { resumeOutside: () => Promise<void> }).resumeOutside();
    expect(page.outsideWaiting()).toBe(true);
    expect(page.chainRunning()).toBe(false);
    expect(page.busy()).toBe(false); // nothing runs: the box is the one thing to do
    expect(page.canGoOn()).toBe(true);
    expect(page.waitingLine()).toBe("Every input file is there: press Go on.");
    await page.goOn();
    expect(f.continues).toEqual([]); // no chain was held here: Do everything again
    expect(f.starts).toBe(1);
    expect(f.waits).toEqual([false]); // every file there: the plain bring-in, no wait
    expect(page.outsideWaiting()).toBe(false);
    expect(page.steps()[0].state).toBe("done");
    expect(f.autoFixStarts).toBe(1);
  });

  it("after a reload, Make them with AI starts Do everything again on the stand-in path; Stop ends the wait on the card", async () => {
    const alone = outside({ rows: [row("Ctrl_01!H35", { keys: ["BBA_01"] })], matched: 0, missing: ["Ctrl_01!H35"], assistant: true });
    const made = outside({ rows: [row("Ctrl_01!H35", { keys: ["BBA_01"], file: "2026Q1_2026_BBA_01.xlsm", stand_in: true })], matched: 1, brought_in: true, assistant: true, stand_ins: ["2026Q1_2026_BBA_01.xlsm"] });
    const f = fakeServer([waiting(alone, "Waiting for 1 input file: BBA_01"), waiting(alone, "Waiting for 1 input file: BBA_01"), running(alone), status(made, { state: "done", result: result() })]);
    openWaiting(f, alone);
    await (page as unknown as { resumeOutside: () => Promise<void> }).resumeOutside();
    expect(page.outsideWaiting()).toBe(true);
    await page.goOn(true);
    expect(f.continues).toEqual([]);
    expect(f.starts).toBe(1);
    expect(f.standIns).toEqual([true]);
    expect(f.waits).toEqual([false]);
    expect(page.steps()[0].state).toBe("done");
    // and Stop with no chain held: the card says it stopped, the box goes
    const g = fakeServer([waiting(alone, "Waiting for 1 input file: BBA_01")]);
    openWaiting(g, alone);
    await (page as unknown as { resumeOutside: () => Promise<void> }).resumeOutside();
    expect(page.outsideWaiting()).toBe(true);
    await page.stopWaiting();
    expect(g.stops).toBe(1);
    expect(page.outsideWaiting()).toBe(false);
    expect(page.outsideCard()).toMatchObject({ tone: "error", line: "Stopped while waiting for the input files" });
  });

  it("the card's button makes the missing files with the assistant, the others stay as picked", async () => {
    const some = outside({ rows: [row("Ctrl_01!H35", { keys: ["BBA_01"], file: "2026Q1_2026_BBA_01.xlsm", how: "key" }), row("Ctrl_01!H36", { keys: ["BBA_02"] })], matched: 1, missing: ["Ctrl_01!H36"], assistant: true });
    const f = fakeServer([running(some), status(outside({ brought_in: true, assistant: true }), { state: "done", result: result() })]);
    open(f, some);
    expect(await page.runOutside(false, true)).toBe(true);
    expect(f.standIns).toEqual([true]);
  });

  it("with every file there, nothing is made: the plain bring-in", async () => {
    const all = outside({ assistant: true });
    const g = fakeServer([running(all), status(outside({ brought_in: true }), { state: "done", result: result() })]);
    open(g, all);
    await page.runOutside(false, true);
    expect(g.standIns).toEqual([false]);
  });
});
