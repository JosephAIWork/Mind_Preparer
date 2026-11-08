import { provideHttpClient } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { TestBed } from "@angular/core/testing";
import { ApiService } from "./api.service";
import { ChainEstimate, OutsideInputRow, OutsideRunResult, inputFilesProblem, outsideOutcome, outsideRowLabel, pickedFilesProblem, withOutsideStep } from "./models";
import { SessionStore } from "./session.store";

const row = (cell: string, extra: Partial<OutsideInputRow> = {}): OutsideInputRow => ({ cell, expected: null, keys: [], target_sheet: null, file: null, how: null, candidates: [], ...extra });

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
    expect(outsideOutcome(result({ numbers: { ran: false, clean: null } })).ok).toBe(true);
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
    store.reset();
    expect(store.outsideInputs()).toBeNull();
  });
});
