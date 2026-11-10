import { TestBed } from "@angular/core/testing";
import { provideRouter } from "@angular/router";
import type { HomePage } from "../pages/home.page";
import type { RecentWorkbooks } from "../pages/recent-workbooks";
import { ApiService } from "./api.service";
import { PendingUpload, SavedWorkbook, SessionStart } from "./models";
import { SessionStore } from "./session.store";

// ---- 1.14.0: the main screen's "Recent workbooks" (the workbooks kept in the data folder) ----------------

const saved = (id: string, extra: Partial<SavedWorkbook> = {}): SavedWorkbook => ({
  id,
  name: `${id}.xlsm`,
  uploaded_at: 1_800_000_000,
  last_opened: 1_800_000_000,
  versions: 3,
  current_version: "ver-003",
  bytes: 1234,
  has_result: true,
  pending: false,
  ...extra,
});

/** A checked workbook as the server sends it back (GET /api/sessions/{id} / POST .../open). */
const started = (id: string): SessionStart => ({
  sessionId: id,
  summary: { file_name: `${id}.xlsm`, file_type: "xlsm", sheet_count: 2, grid_count: 1, formula_count: 10 },
  report: { status: "PASS", findings: [] },
  plan: [],
  version: { id: "ver-003", major: 3, minor: 0, label: "v3 — Prep", file_name: `${id}_v3.xlsm`, source: "prep" },
  versions: [],
  recalc: { status: "PASS", message: "clean", ran: true, formula_errors: [], addin_gap_errors: [], version_id: "ver-003" },
});

interface Fake {
  api: ApiService;
  list: SavedWorkbook[];
  opened: string[];
  removed: string[];
  openFails: string | null;
}

function fakeServer(list: SavedWorkbook[]): Fake {
  const f: Fake = { api: null as unknown as ApiService, list, opened: [], removed: [], openFails: null };
  f.api = {
    savedWorkbooks: () => Promise.resolve([...f.list]),
    openSavedWorkbook: (id: string) => {
      f.opened.push(id);
      if (f.openFails) return Promise.reject({ error: { detail: f.openFails }, status: 404 });
      const row = f.list.find((r) => r.id === id);
      if (row?.pending) return Promise.resolve({ pending: true, sessionId: id, mode: "plan", size: { above_threshold: false, sheets: [{ name: "Data" }] }, version: started(id).version, versions: [] } as unknown as PendingUpload);
      return Promise.resolve(started(id));
    },
    removeSavedWorkbook: (id: string) => {
      f.removed.push(id);
      f.list = f.list.filter((r) => r.id !== id);
      return Promise.resolve({ removed: true, id });
    },
    // what the main screen asks for once a session is loaded: nothing to say here
    estimate: () => Promise.reject(new Error("no estimate")),
    mindLoopPlan: () => Promise.reject(new Error("no loop plan")),
    autoFixStatus: () => Promise.reject(new Error("no fixer")),
    outsideInputs: () => Promise.reject(new Error("no outside files")),
    nextQuarter: () => Promise.reject(new Error("no next quarter")),
    analyze: () => Promise.reject(new Error("no scan here")),
    scanStatus: () => Promise.reject(new Error("no scan here")),
    resume: () => Promise.reject(new Error("offline")),
    downloadUrl: (id: string, name: string) => `/api/sessions/${id}/files/${encodeURIComponent(name)}`,
  } as unknown as ApiService;
  return f;
}

const settle = () => new Promise<void>((r) => setTimeout(r, 0));

describe("Recent workbooks: the list under the upload box", () => {
  let List: typeof RecentWorkbooks;

  beforeAll(async () => {
    List = (await import("../pages/recent-workbooks")).RecentWorkbooks;
  });

  function draw(f: Fake) {
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: f.api }] });
    const fixture = TestBed.createComponent(List);
    fixture.detectChanges();
    return { fixture, el: fixture.nativeElement as HTMLElement, list: fixture.componentInstance };
  }

  async function drawn(f: Fake) {
    const d = draw(f);
    await settle(); // the list arrives
    d.fixture.detectChanges();
    return d;
  }

  it("is hidden when no workbook is saved", async () => {
    const { el, list } = await drawn(fakeServer([]));
    expect(list.rows()).toEqual([]);
    expect(el.querySelector(".recent")).toBeNull();
    expect(el.textContent?.trim()).toBe("");
  });

  it("is hidden too when the server has no such route (an older app)", async () => {
    const f = fakeServer([]);
    f.api.savedWorkbooks = () => Promise.reject(new Error("404"));
    const { el } = await drawn(f);
    expect(el.querySelector(".recent")).toBeNull();
  });

  it("lists each saved workbook: name, date, versions, Reopen and Remove, the last opened first", async () => {
    const f = fakeServer([saved("model_a"), saved("model_b", { versions: 1, has_result: false, pending: true })]);
    const { el, list } = await drawn(f);
    expect(el.querySelector(".recent")).not.toBeNull();
    expect(el.querySelector("[data-show-me='recent-workbooks']")).not.toBeNull();
    expect(el.querySelector(".recent__title")?.textContent).toBe("Recent workbooks");
    const rows = Array.from(el.querySelectorAll(".recent__row"));
    expect(rows.map((r) => r.querySelector(".recent__name")?.textContent)).toEqual(["model_a.xlsm", "model_b.xlsm"]);
    expect(rows[0].querySelector(".recent__meta")?.textContent).toMatch(/^\d{1,2} \w{3} 2027, \d{2}:\d{2} · 3 versions$/);
    expect(list.metaText(f.list[1])).toMatch(/· 1 version · not checked yet$/);
    expect(rows[0].querySelector(".recent__open")?.textContent?.trim()).toBe("Reopen");
    expect(rows[0].querySelector(".recent__remove")?.textContent?.trim()).toBe("Remove");
    expect(el.querySelector(".recent__ask")).toBeNull(); // nothing asked yet
  });

  it("Reopen loads the workbook and hands it to the page", async () => {
    const f = fakeServer([saved("model_a")]);
    const { el, fixture, list } = await drawn(f);
    const got: (SessionStart | PendingUpload)[] = [];
    list.opened.subscribe((s) => got.push(s));
    (el.querySelector(".recent__open") as HTMLButtonElement).click();
    fixture.detectChanges();
    expect(list.busyId()).toBe("model_a");
    expect(el.querySelector(".recent__open")?.textContent?.trim()).toBe("Opening…");
    await settle();
    fixture.detectChanges();
    expect(f.opened).toEqual(["model_a"]);
    expect(got).toHaveLength(1);
    expect((got[0] as SessionStart).sessionId).toBe("model_a");
    expect(list.busyId()).toBeNull();
    expect(list.error()).toBe("");
  });

  it("says why when a workbook cannot be opened, and reads the list again", async () => {
    const f = fakeServer([saved("model_a")]);
    const { el, fixture, list } = await drawn(f);
    f.openFails = "This workbook is not saved here any more.";
    f.list = []; // removed from another window meanwhile
    await list.reopen(saved("model_a"));
    await settle();
    fixture.detectChanges();
    expect(list.error()).toBe("Could not open model_a.xlsm: This workbook is not saved here any more.");
    expect(list.rows()).toEqual([]);
    expect(el.querySelector(".recent")).toBeNull(); // the list went with it
  });

  it("Remove asks once, in place, then removes; No keeps it", async () => {
    const f = fakeServer([saved("model_a"), saved("model_b")]);
    const { el, fixture, list } = await drawn(f);
    (el.querySelector(".recent__row[data-id='model_a'] .recent__remove") as HTMLButtonElement).click();
    fixture.detectChanges();
    const ask = el.querySelector(".recent__row[data-id='model_a'] .recent__ask");
    expect(ask?.textContent).toContain("Remove model_a.xlsm from this computer?");
    expect(ask?.querySelector(".recent__yes")?.textContent?.trim()).toBe("Yes");
    expect(ask?.querySelector(".recent__no")?.textContent?.trim()).toBe("No");
    expect(el.querySelector(".recent__row[data-id='model_b'] .recent__ask")).toBeNull(); // only that row asks
    expect(f.removed).toEqual([]); // nothing removed yet
    // No: the row stays, the question goes
    (ask?.querySelector(".recent__no") as HTMLButtonElement).click();
    fixture.detectChanges();
    expect(el.querySelector(".recent__ask")).toBeNull();
    expect(list.rows()).toHaveLength(2);
    // Yes: removed on the server, gone from the list
    (el.querySelector(".recent__row[data-id='model_a'] .recent__remove") as HTMLButtonElement).click();
    fixture.detectChanges();
    (el.querySelector(".recent__yes") as HTMLButtonElement).click();
    await settle();
    fixture.detectChanges();
    expect(f.removed).toEqual(["model_a"]);
    expect(list.rows().map((r) => r.id)).toEqual(["model_b"]);
    expect(Array.from(el.querySelectorAll(".recent__name")).map((n) => n.textContent)).toEqual(["model_b.xlsm"]);
    expect(list.askingId()).toBeNull();
  });
});

describe("Recent workbooks: the main screen opens a saved workbook where it was", () => {
  let Page: typeof HomePage;
  let page: HomePage;
  let store: SessionStore;

  beforeAll(async () => {
    // The main screen's logic without drawing it: Carbon stands in as inert classes (see outside-inputs.spec.ts).
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    Page = (await import("../pages/home.page")).HomePage;
  });

  function open(f: Fake): void {
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: f.api }] });
    store = TestBed.inject(SessionStore);
    page = TestBed.runInInjectionContext(() => new Page());
  }

  afterEach(() => page?.ngOnDestroy());

  it("shows the checked workbook with its last recalculation; nothing runs by itself", () => {
    const f = fakeServer([saved("model_a")]);
    open(f);
    page.setAutoRun(true); // "Do everything right after the upload" is for uploads, not for a reopen
    page.openSaved(started("model_a"));
    expect(store.sessionId()).toBe("model_a");
    expect(store.currentVersion()?.id).toBe("ver-003");
    expect(store.recalc()?.ran).toBe(true); // Advanced > Recalculate shows the list without recalculating
    expect(page.pending()).toBeNull();
    expect(page.uploadError()).toBe("");
    expect(page.steps()).toEqual([]); // no Do everything started
    expect(page.chainRunning()).toBe(false);
    page.setAutoRun(false);
  });

  it("takes a workbook that was uploaded but never checked to the scan", async () => {
    const f = fakeServer([saved("model_b", { pending: true, has_result: false })]);
    open(f);
    const pending = (await f.api.openSavedWorkbook("model_b")) as PendingUpload;
    page.openSaved(pending);
    expect(store.sessionId()).toBeNull();
    expect(page.phase()).toBe("scanning"); // a small one scans at once, as after an upload
    await settle();
    // a big one asks which sheets to skip first
    const big = { ...pending, size: { above_threshold: true, sheets: [{ name: "A" }, { name: "B" }], message: "big" } } as unknown as PendingUpload;
    page.openSaved(big);
    expect(page.phase()).toBe("choose");
    expect(page.pending()?.sessionId).toBe("model_b");
  });
});
