import { provideHttpClient, withInterceptors } from "@angular/common/http";
import { HttpTestingController, provideHttpClientTesting } from "@angular/common/http/testing";
import { NO_ERRORS_SCHEMA } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import { RouterLink, RouterLinkActive, RouterOutlet, provideRouter } from "@angular/router";
import type { App } from "../app";
import { ApiService } from "./api.service";
import { ExpiredNotice } from "./expired-notice";
import { ExpiryService, expiryInterceptor, noticeParts, shortDate, soonLine, untilLine } from "./expiry";
import { AppHealth } from "./models";

const STOPPED = "This copy of Mind Ready stopped working on 15 December 2026. Ask Joseph Gestetner for a new copy.";
const health = (extra: Partial<AppHealth> = {}): AppHealth => ({ version: "1.12.4", rules: 126, ...extra });

describe("the end date in words", () => {
  it("writes the date as it is written, whatever the time zone", () => {
    expect(shortDate("2026-12-15")).toBe("15 Dec 2026");
    expect(shortDate("2027-01-05")).toBe("5 Jan 2027");
    for (const bad of [null, undefined, "", "15/12/2026", "2026-13-01"]) expect(shortDate(bad)).toBeNull();
  });

  it("says until when the copy works, and nothing without an end date", () => {
    expect(untilLine("2026-12-15")).toBe("Works until 15 Dec 2026");
    expect(untilLine(null)).toBeNull();
  });

  it("says it gently in the last 7 days only", () => {
    expect(soonLine("2026-12-15", 7)).toBeNull();
    expect(soonLine("2026-12-15", 6)).toBe("This copy stops working on 15 Dec 2026.");
    expect(soonLine("2026-12-15", 0)).toBe("This copy stops working on 15 Dec 2026."); // the last day
    expect(soonLine("2026-12-15", -1)).toBeNull();
    expect(soonLine("2026-12-15", null)).toBeNull();
    expect(soonLine(null, 3)).toBeNull();
  });

  it("splits the message into what happened and what to do", () => {
    expect(noticeParts(STOPPED)).toEqual({ title: "This copy of Mind Ready stopped working on 15 December 2026.", text: "Ask Joseph Gestetner for a new copy." });
    expect(noticeParts("Stopped")).toEqual({ title: "Stopped", text: "" });
  });
});

describe("ExpiryService", () => {
  let expiry: ExpiryService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    expiry = TestBed.inject(ExpiryService);
  });

  it("follows /api/health: the date, then the gentle line in the last days", () => {
    expiry.fromHealth(health({ expires: "2026-12-15", days_left: 20, expired: false }));
    expect([expiry.until(), expiry.soon(), expiry.stopped()]).toEqual(["Works until 15 Dec 2026", null, null]);
    expiry.fromHealth(health({ expires: "2026-12-15", days_left: 3, expired: false }));
    expect(expiry.soon()).toBe("This copy stops working on 15 Dec 2026.");
  });

  it("shows nothing for the developer version", () => {
    expiry.fromHealth(health());
    expect([expiry.until(), expiry.soon(), expiry.stopped()]).toEqual([null, null, null]);
    expiry.fromHealth(health({ expires: null, days_left: null, expired: false, expired_message: null }));
    expect([expiry.until(), expiry.soon(), expiry.stopped()]).toEqual([null, null, null]);
  });

  it("stops on the server's message once the date passed; nothing else is shown then", () => {
    expiry.fromHealth(health({ expires: "2026-12-15", days_left: -1, expired: true, expired_message: STOPPED }));
    expect([expiry.stopped(), expiry.until(), expiry.soon()]).toEqual([STOPPED, null, null]);
  });

  it("keeps plain words when a 410 comes without any", () => {
    expiry.noteAnswer(410, { expired: true });
    expect(expiry.stopped()).toBe("This copy of Mind Ready stopped working. Ask the person who shared it for a new copy.");
  });
});

describe("a 410 from any API call", () => {
  let api: ApiService;
  let http: HttpTestingController;
  let expiry: ExpiryService;

  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideHttpClient(withInterceptors([expiryInterceptor])), provideHttpClientTesting()] });
    api = TestBed.inject(ApiService);
    http = TestBed.inject(HttpTestingController);
    expiry = TestBed.inject(ExpiryService);
  });

  afterEach(() => {
    http.verify();
    vi.unstubAllGlobals();
  });

  it("turns the screens into the notice", async () => {
    const call = api.versions("s1");
    http.expectOne("/api/sessions/s1/versions").flush({ detail: STOPPED, expired: true, expires: "2026-12-15" }, { status: 410, statusText: "Gone" });
    await expect(call).rejects.toBeTruthy();
    expect(expiry.stopped()).toBe(STOPPED);
  });

  it("leaves other refusals alone", async () => {
    const missing = api.versions("s1");
    http.expectOne("/api/sessions/s1/versions").flush({ detail: "no session s1" }, { status: 404, statusText: "Not Found" });
    await expect(missing).rejects.toBeTruthy();
    const gone = api.versions("s2");
    http.expectOne("/api/sessions/s2/versions").flush({ detail: "gone" }, { status: 410, statusText: "Gone" }); // not the end date
    await expect(gone).rejects.toBeTruthy();
    expect(expiry.stopped()).toBeNull();
  });

  it("also from the assistant's streamed answer (fetch, past the interceptor)", async () => {
    const body = JSON.stringify({ detail: STOPPED, expired: true, expires: "2026-12-15" });
    vi.stubGlobal("fetch", () => Promise.resolve(new Response(body, { status: 410, headers: { "Content-Type": "application/json" } })));
    await expect(api.chatStream("s1", [], "Why?", undefined, () => {})).rejects.toThrow(STOPPED);
    expect(expiry.stopped()).toBe(STOPPED);
  });
});

describe("the notice, drawn", () => {
  it("says what happened and what to do, with nothing else on the page", () => {
    TestBed.configureTestingModule({});
    const fixture = TestBed.createComponent(ExpiredNotice);
    fixture.componentRef.setInput("message", STOPPED);
    fixture.detectChanges();
    const page = fixture.nativeElement as HTMLElement;
    expect(page.querySelector("h1")?.textContent?.trim()).toBe("This copy of Mind Ready stopped working on 15 December 2026.");
    expect(page.querySelector(".expired__text")?.textContent?.trim()).toBe("Ask Joseph Gestetner for a new copy.");
    expect(page.querySelectorAll("button, a, input").length).toBe(0);
  });
});

// ---- the shell, drawn: its Carbon header stands in as plain elements -----------------------------------
describe("the shell and the end date", () => {
  let AppClass: typeof App;

  beforeAll(async () => {
    // Carbon (a package Node cannot load as it is) stands in as inert classes; the shell drops them below
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    AppClass = (await import("../app")).App;
  });

  async function draw(h: AppHealth): Promise<HTMLElement> {
    const api = { health: () => Promise.resolve(h), resume: () => Promise.reject(new Error("no session")) } as unknown as ApiService;
    TestBed.configureTestingModule({ providers: [provideRouter([]), { provide: ApiService, useValue: api }] });
    TestBed.overrideComponent(AppClass, { set: { imports: [RouterOutlet, RouterLink, RouterLinkActive, ExpiredNotice], schemas: [NO_ERRORS_SCHEMA] } });
    const fixture = TestBed.createComponent(AppClass);
    fixture.detectChanges();
    await fixture.whenStable();
    fixture.detectChanges();
    return fixture.nativeElement as HTMLElement;
  }

  it("says 'Works until' quietly in the header", async () => {
    const page = await draw(health({ expires: "2026-12-15", days_left: 20, expired: false }));
    expect(page.querySelector(".shell__until")?.textContent?.trim()).toBe("Works until 15 Dec 2026");
    expect(page.querySelector(".shell__soon")).toBeNull();
    expect(page.querySelector("mr-expired-notice")).toBeNull();
  });

  it("adds the gentle line in the last 7 days", async () => {
    const page = await draw(health({ expires: "2026-12-15", days_left: 3, expired: false }));
    expect(page.querySelector(".shell__soon")?.textContent?.trim()).toBe("This copy stops working on 15 Dec 2026.");
    expect(page.querySelector(".shell__until")?.textContent?.trim()).toBe("Works until 15 Dec 2026");
  });

  it("shows nothing about it in the developer version", async () => {
    const page = await draw(health());
    expect(page.querySelector(".shell__until")).toBeNull();
    expect(page.querySelector(".shell__soon")).toBeNull();
    expect(page.querySelector("cds-header")).not.toBeNull();
  });

  it("gives the whole page to the notice once the copy stopped working", async () => {
    const page = await draw(health({ expires: "2026-12-15", days_left: -1, expired: true, expired_message: STOPPED }));
    expect(page.querySelector("mr-expired-notice h1")?.textContent?.trim()).toBe("This copy of Mind Ready stopped working on 15 December 2026.");
    expect(page.querySelector("cds-header")).toBeNull();
    expect(page.querySelector("router-outlet")).toBeNull();
  });
});
