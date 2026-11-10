import { provideHttpClient, withInterceptors } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { TestBed } from "@angular/core/testing";
import { ApiService } from "./api.service";
import { SessionStart } from "./models";
import { RESTART_MARK, restartInterceptor } from "./restart.interceptor";
import { SessionStore } from "./session.store";

// ---- 1.14.0: a session the restart closed: the page lets go of it, the main screen says so above the list ----

const SENTENCE = "This workbook is not open any more (the app was restarted). Open it again from Recent workbooks on the main screen.";

const started: SessionStart = {
  sessionId: "old1",
  summary: { file_name: "old1.xlsm", file_type: "xlsm", sheet_count: 2, grid_count: 1, formula_count: 10 },
  report: { status: "PASS", findings: [] },
  plan: [],
  version: { id: "ver-001", major: 1, minor: 0, label: "v1", file_name: "old1.xlsm", source: "upload" },
  versions: [],
};

describe("a session the restart closed", () => {
  let api: ApiService;
  let http: HttpTestingController;
  let store: SessionStore;

  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideHttpClient(withInterceptors([restartInterceptor])), provideHttpClientTesting()] });
    api = TestBed.inject(ApiService);
    http = TestBed.inject(HttpTestingController);
    store = TestBed.inject(SessionStore);
  });

  afterEach(() => http.verify());

  it("lets go of the session and keeps the server's sentence for the main screen", async () => {
    store.setSession(started);
    expect(store.loaded()).toBe(true);
    const call = api.autoFixStatus("old1", false).catch((e: { status: number }) => e);
    http.expectOne((r) => r.url.endsWith("/api/sessions/old1/auto-fix")).flush({ detail: SENTENCE }, { status: 404, statusText: "Not Found" });
    const err = (await call) as { status: number };
    expect(err.status).toBe(404); // the caller still sees its failure
    expect(store.sessionId()).toBeNull();
    expect(store.loaded()).toBe(false);
    expect(store.restartNotice()).toBe(SENTENCE);
    expect(SENTENCE).toContain(RESTART_MARK);
    // the next workbook opened clears the sentence
    store.setSession({ ...started, sessionId: "new1" });
    expect(store.restartNotice()).toBe("");
  });

  it("leaves any other failure alone", async () => {
    store.setSession(started);
    const call = api.autoFixStatus("old1", false).catch((e: { status: number }) => e);
    http.expectOne((r) => r.url.endsWith("/api/sessions/old1/auto-fix")).flush({ detail: "unknown session old1" }, { status: 404, statusText: "Not Found" });
    await call;
    expect(store.sessionId()).toBe("old1");
    expect(store.restartNotice()).toBe("");
    const call2 = api.autoFixStatus("old1", false).catch((e: { status: number }) => e);
    http.expectOne((r) => r.url.endsWith("/api/sessions/old1/auto-fix")).flush({ detail: SENTENCE }, { status: 409, statusText: "Conflict" });
    await call2;
    expect(store.sessionId()).toBe("old1");
  });
});
