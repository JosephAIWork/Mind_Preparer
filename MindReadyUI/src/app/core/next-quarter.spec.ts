import { provideHttpClient } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { TestBed } from "@angular/core/testing";
import { provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import { ApiService } from "./api.service";
import { NEXT_QUARTER_LEAD, NextQuarterResult, NextQuarterStatus, OutsideInputs, nextQuarterLines, nextQuarterOutcome } from "./models";
import { SessionStore } from "./session.store";

/** Next quarter's Mind input files: the new quarter's RAW files go through the model's own import
 *  macro on the server; the screen sends them, follows the job and offers the zip. */

const result = (extra: Partial<NextQuarterResult> = {}): NextQuarterResult => ({
  ok: true,
  summary: "2 Mind input files made (IN_BBA_01.csv, IN_VFA_01.txt): put them in Mind's Inputs window, then Run.",
  files: [
    { file: "IN_BBA_01.csv", grid: "IN_BBA_01", from: "2026_Q2 - BBA_01.csv", source: "macro", rows: 1333, cols: 93, separator: "comma", changed: 1234 },
    { file: "IN_VFA_01.txt", grid: "IN_VFA_01", from: "IN_VFA_01.txt", source: "as_given", rows: 16382, cols: 717, separator: "TAB", changed: 0 },
  ],
  not_in_zip: [{ grid: "IN_BBA_02", why: "no new file for it: Mind keeps the data it has" }],
  notes: [],
  unused: [],
  zip: "Next quarter Mind inputs.zip",
  readme: "# Next quarter's Mind input files",
  error: null,
  ...extra,
});

const status = (extra: Partial<NextQuarterStatus> = {}): NextQuarterStatus => ({
  grids: [{ name: "IN_BBA_01", sheet: "BBA_01", ref: "AC3:DQ1335", rows: 1333, cols: 93, file: "IN_BBA_01.csv" }],
  available: true,
  why_not: null,
  running: false,
  state: "idle",
  stage: null,
  result: null,
  error: null,
  ...extra,
});

describe("next quarter's Mind input files in words", () => {
  it("says why raw files cannot go into Mind, and what the button does", () => {
    expect(NEXT_QUARTER_LEAD).toMatch(/^Raw input files cannot go into Mind directly\. Mind takes one file per input grid, shaped like the grid\./);
    expect(NEXT_QUARTER_LEAD).toMatch(/model's own import macro/);
    expect(NEXT_QUARTER_LEAD).toMatch(/No new model upload\.$/);
  });

  it("lists which file goes into which grid, and what changed", () => {
    expect(nextQuarterLines(result())).toEqual([
      "IN_BBA_01.csv → grid IN_BBA_01, from 2026_Q2 - BBA_01.csv; 1,234 values changed",
      "IN_VFA_01.txt → grid IN_VFA_01, from IN_VFA_01.txt (as it is); same values as before",
    ]);
    const one = result({ files: [{ ...result().files[0], changed: 1 }] });
    expect(nextQuarterLines(one)).toEqual(["IN_BBA_01.csv → grid IN_BBA_01, from 2026_Q2 - BBA_01.csv; 1 value changed"]);
    expect(nextQuarterLines(result({ files: [{ ...result().files[0], changed: null }] }))[0]).toBe("IN_BBA_01.csv → grid IN_BBA_01, from 2026_Q2 - BBA_01.csv");
  });

  it("is green with every file made, amber when a file did not fit, red when nothing was made", () => {
    expect(nextQuarterOutcome(result())).toEqual({
      tone: "ok",
      line: result().summary,
      detail: "IN_BBA_02: no new file for it: Mind keeps the data it has",
    });
    const raw = "IN_BBA_01.csv has 1,340 columns; the Mind input grid IN_BBA_01 has 93. It looks like a raw input file, so the app reads it with the model's import macro.";
    expect(nextQuarterOutcome(result({ notes: [raw], not_in_zip: [] }))).toEqual({ tone: "ok", line: result().summary, detail: raw });
    const big = result({ not_in_zip: [{ grid: "IN_VFA_01", why: "v.xlsx filled VFA_01!B3:C9, more than the Mind input grid IN_VFA_01 holds" }] });
    expect(nextQuarterOutcome(big).tone).toBe("warn");
    expect(nextQuarterOutcome(result({ unused: [{ file: "notes.csv", why: "no row of the model's import table names it" }] })).tone).toBe("warn");
    const failed = result({ ok: false, files: [], zip: null, error: 'The model\'s import macro did not take the new files: "93項目まで". No file was made: check the input files, then try again.' });
    expect(nextQuarterOutcome(failed)).toEqual({ tone: "error", line: "The model's import macro did not take the new files: \"93項目まで\"", detail: failed.error });
  });
});

describe("ApiService and next quarter's files", () => {
  let api: ApiService;
  let http: HttpTestingController;

  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideHttpClient(), provideHttpClientTesting()] });
    api = TestBed.inject(ApiService);
    http = TestBed.inject(HttpTestingController);
  });

  afterEach(() => http.verify());

  it("sends the raw files as repeated files and asks where the job is", async () => {
    const names = ["2026_Q2 - BBA_01.csv", "2026_Q2 - VFA_01.xlsx"];
    const started = api.startNextQuarter("s 1", names.map((n) => new File(["x"], n)));
    const req = http.expectOne("/api/sessions/s%201/next-quarter-inputs");
    expect(req.request.method).toBe("POST");
    expect(((req.request.body as FormData).getAll("files") as File[]).map((f) => f.name)).toEqual(names);
    req.flush({ running: true, raw: names, grid_files: [], notes: [] }, { status: 202, statusText: "Accepted" });
    expect((await started).raw).toEqual(names);
    const got = api.nextQuarter("s1");
    const get = http.expectOne("/api/sessions/s1/next-quarter-inputs");
    expect(get.request.method).toBe("GET");
    get.flush(status());
    expect((await got).grids[0].file).toBe("IN_BBA_01.csv");
  });
});

describe("the Next quarter part of the Outside files card", () => {
  let Page: typeof HomePage;
  let page: HomePage;
  let store: SessionStore;

  beforeAll(async () => {
    // Carbon (a package Node cannot load as it is) stands in as inert classes, as in outside-inputs.spec
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
  });

  interface Fake {
    api: ApiService;
    statuses: NextQuarterStatus[];
    sent: string[][];
    refuse?: string;
  }

  function fake(statuses: NextQuarterStatus[], refuse?: string): Fake {
    const f: Fake = { api: null as unknown as ApiService, statuses, sent: [], refuse };
    let last: NextQuarterStatus | undefined;
    f.api = {
      startNextQuarter: (_id: string, files: File[]) => {
        f.sent.push(files.map((x) => x.name));
        return f.refuse ? Promise.reject({ error: { detail: f.refuse } }) : Promise.resolve({ running: true, raw: [], grid_files: [], notes: [] });
      },
      nextQuarter: () => {
        last = f.statuses.shift() ?? last;
        return last ? Promise.resolve(last) : Promise.reject(new Error("no status"));
      },
      estimate: () => Promise.reject(new Error("no estimate")),
      resume: () => Promise.reject(new Error("offline")),
      mindLoopPlan: () => Promise.reject(new Error("no loop plan")),
      outsideInputs: () => Promise.reject(new Error("not asked here")),
      downloadUrl: (id: string, name: string) => `/api/sessions/${id}/files/${encodeURIComponent(name)}`,
    } as unknown as ApiService;
    return f;
  }

  const broughtIn: OutsideInputs = {
    imports: true, macro: "calcCF", calls: [], rows: [], matched: 2, missing: [], extra: [], summary: "2 of 2 input files matched.", brought_in: true,
    mind_inputs: ["IN_BBA_01", "IN_VFA_01"],
  };

  function open(f: Fake): void {
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: f.api }] });
    store = TestBed.inject(SessionStore);
    page = TestBed.runInInjectionContext(() => new Page());
    (page as unknown as { sleep: () => Promise<void> }).sleep = () => Promise.resolve();
    store.sessionId.set("s1");
    store.outsideInputs.set(broughtIn);
  }

  afterEach(() => page?.ngOnDestroy());

  it("sends the raw files, follows the job and offers the zip with each file's grid", async () => {
    const f = fake([status({ running: true, state: "running", stage: "Running the model's import macro with the new files", message: "Running the workbook's import macro calcCF", elapsed_s: 12 }),
      status({ state: "done", result: result() })]);
    open(f);
    const files = [new File(["x"], "2026_Q2 - BBA_01.csv"), new File(["x"], "IN_VFA_01.txt")];
    expect(await page.makeNextQuarter(files)).toBe(true);
    expect(f.sent).toEqual([["2026_Q2 - BBA_01.csv", "IN_VFA_01.txt"]]);
    expect(page.nextQuarter()).toMatchObject({ running: false, tone: "ok", line: result().summary, open: false });
    expect(page.nextQuarterZipUrl()).toBe("/api/sessions/s1/files/Next%20quarter%20Mind%20inputs.zip");
    expect(page.nextQuarterFiles()).toEqual(nextQuarterLines(result()));
    expect(page.busy()).toBe(false);
    // no new version and no change to the Outside files card: the model in Mind stays as it is
    expect(store.outsideInputs()).toBe(broughtIn);
    expect(page.zipUrl()).toBeNull();
  });

  it("says the server's refusal in its words, and offers no zip", async () => {
    const f = fake([], "None of the picked files fills a Mind input grid (notes.csv: no row of the model's import table names it, and it is not shaped like a Mind input grid).");
    open(f);
    expect(await page.makeNextQuarter([new File(["x"], "notes.csv")])).toBe(false);
    expect(page.nextQuarter()).toMatchObject({ tone: "error", line: "None of the picked files fills a Mind input grid (notes.csv: no row of the model's import table names it, and it is not shaped like a Mind input grid)", open: true });
    expect(page.nextQuarterZipUrl()).toBeNull();
  });

  it("says the macro's own message when it did not take the new files", async () => {
    const err = "The model's import macro did not take the new files: \"BBA_01のinput領域には93項目まで入力可能です\". No file was made: check the input files, then try again.";
    const g = fake([status({ running: true, state: "running", stage: "Starting" }), status({ state: "error", error: err, result: result({ ok: false, files: [], zip: null, error: err }) })]);
    open(g);
    expect(await page.makeNextQuarter([new File(["x"], "2026_Q2 - BBA_01.xlsx")])).toBe(false);
    expect(page.nextQuarter()).toMatchObject({ tone: "error", line: "The model's import macro did not take the new files: \"BBA_01のinput領域には93項目まで入力可能です\"", detail: err });
    expect(page.nextQuarterFiles()).toEqual([]);
    expect(page.nextQuarterZipUrl()).toBeNull();
  });

  it("forgets the last job's files and zip when another workbook comes", async () => {
    const f = fake([status({ state: "done", result: result() })]);
    open(f);
    expect(await page.makeNextQuarter([new File(["x"], "2026_Q2 - BBA_01.csv")])).toBe(true);
    expect(page.nextQuarterZipUrl()).not.toBeNull();
    page.loadAnother();
    store.sessionId.set("s2");
    store.outsideInputs.set(broughtIn);
    expect(page.nextQuarter()).toMatchObject({ running: false, line: "" });
    expect(page.nextQuarterFiles()).toEqual([]);
    expect(page.nextQuarterZipUrl()).toBeNull();
    // a zip is only ever offered in the session that made it
    page.nextQuarterResult.set(result());
    expect(page.nextQuarterZipUrl()).toBeNull();
  });

  it("stops following an old job when the session changes, so its result never lands on the new card", async () => {
    const f = fake([status({ running: true, state: "running", stage: "Starting" }), status({ running: true, state: "running", stage: "Starting" }),
      status({ state: "done", result: result() })]);
    open(f);
    let rounds = 0;
    (page as unknown as { sleep: () => Promise<void> }).sleep = () => {
      if (++rounds === 1) store.sessionId.set("s2"); // another workbook opened while the job runs
      return Promise.resolve();
    };
    expect(await page.makeNextQuarter([new File(["x"], "2026_Q2 - BBA_01.csv")])).toBe(false);
    expect(page.nextQuarter().running).toBe(false);
    expect(page.nextQuarterResult()).toBeNull();
    expect(page.nextQuarterZipUrl()).toBeNull();
    expect(page.busy()).toBe(false);
  });

  it("picks a running job back up after a reload and offers its zip when it ends", async () => {
    const f = fake([status({ running: true, state: "running", stage: "Running the model's import macro with the new files", elapsed_s: 95 }),
      status({ running: true, state: "running", stage: "Writing the Mind input files", elapsed_s: 96 }), status({ state: "done", result: result() })]);
    open(f);
    const seen: string[] = [];
    (page as unknown as { sleep: () => Promise<void> }).sleep = () => {
      seen.push(`${page.nextQuarter().line} | ${page.nextQuarter().detail} | busy=${page.busy()}`);
      return Promise.resolve();
    };
    await (page as unknown as { resumeNextQuarter: () => Promise<void> }).resumeNextQuarter();
    // the first answer shows the job at once; the follow goes on from the next one
    expect(seen).toEqual(["Writing the Mind input files | 1 min 36 s so far | busy=true"]);
    expect(f.sent).toEqual([]); // nothing sent again: the same job, followed
    expect(page.nextQuarter()).toMatchObject({ running: false, tone: "ok", line: result().summary });
    expect(page.nextQuarterZipUrl()).toBe("/api/sessions/s1/files/Next%20quarter%20Mind%20inputs.zip");
  });

  it("shows the last job's result after a reload, and nothing when there was none", async () => {
    open(fake([status({ state: "done", result: result() })]));
    await (page as unknown as { resumeNextQuarter: () => Promise<void> }).resumeNextQuarter();
    expect(page.nextQuarter()).toMatchObject({ running: false, tone: "ok" });
    expect(page.nextQuarterFiles()).toEqual(nextQuarterLines(result()));
    expect(page.nextQuarterZipUrl()).not.toBeNull();
    page.ngOnDestroy();
    TestBed.resetTestingModule();
    open(fake([status()]));
    await (page as unknown as { resumeNextQuarter: () => Promise<void> }).resumeNextQuarter();
    expect(page.nextQuarter()).toMatchObject({ running: false, line: "" });
    expect(page.nextQuarterZipUrl()).toBeNull();
  });

  it("checks the picked files first and waits while another step runs", async () => {
    const f = fake([status({ state: "done", result: result() })]);
    open(f);
    expect(await page.makeNextQuarter([new File(["x"], "notes.pdf")])).toBe(false);
    expect(page.nextQuarter().line).toBe("notes.pdf is not an Excel or text file: pick .xlsx, .xlsm, .xlsb, .xls, .csv or .txt files.");
    page.outside.set({ running: true, fraction: null, line: "Running", detail: "", tone: "idle" });
    expect(await page.makeNextQuarter([new File(["x"], "2026_Q2 - BBA_01.csv")])).toBe(false);
    expect(page.nextQuarter().line).toBe("Wait: Bring in the outside files is running.");
    expect(f.sent).toEqual([]);
  });
});
