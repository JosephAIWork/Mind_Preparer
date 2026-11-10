import { NO_ERRORS_SCHEMA, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { ApiService } from "./api.service";
import type { FeedbackDialog } from "./feedback-dialog";
import type { FeedbackService } from "./feedback.service";
import { FeedbackMessage } from "./models";
import type { SessionStore } from "./session.store";

// "Send feedback" (10 November 2026): the dialog names no team. The email goes to the address the
// backend gives (feedback.email in the settings), and the sent step says so.

const message: FeedbackMessage = {
  ref: "ABCD-1234",
  text: "Mind Ready 1.14.0 (source) · ref ABCD-1234\n\nFeedback (no step failed)\n",
  subject: "Mind Ready 1.14.0 feedback · ref ABCD-1234",
  email: "someone@example.com",
  file_name: "MindReady-feedback-ABCD-1234.txt",
  assistant: false,
} as FeedbackMessage;

describe("the Send feedback dialog", () => {
  let DialogClass: typeof FeedbackDialog;
  // the injection tokens from the same module registry as the dialog (vi.resetModules makes a fresh one)
  let Api: typeof ApiService;
  let Store: typeof SessionStore;
  let Feedback: typeof FeedbackService;
  let sent: unknown[];

  beforeAll(async () => {
    vi.resetModules();
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    DialogClass = (await import("./feedback-dialog")).FeedbackDialog;
    Api = (await import("./api.service")).ApiService;
    Store = (await import("./session.store")).SessionStore;
    Feedback = (await import("./feedback.service")).FeedbackService;
  });

  async function draw(ref: string | null) {
    sent = [];
    const api = {
      assistantKey: () => Promise.resolve({ saved: false }),
      feedback: (body: unknown) => {
        sent.push(body);
        return Promise.resolve(message);
      },
    };
    TestBed.resetTestingModule();
    TestBed.configureTestingModule({ providers: [{ provide: Api, useValue: api }, { provide: Store, useValue: { sessionId: signal(null) } }] });
    TestBed.overrideComponent(DialogClass, { set: { imports: [], schemas: [NO_ERRORS_SCHEMA] } });
    const fixture = TestBed.createComponent(DialogClass);
    const feedback = TestBed.inject(Feedback);
    feedback.open(ref);
    const detect = async () => {
      fixture.detectChanges();
      await fixture.whenStable();
      fixture.detectChanges();
    };
    await detect();
    return { page: fixture.nativeElement as HTMLElement, dialog: fixture.componentInstance, detect, feedback };
  }

  it("is titled Send feedback and names no team, from the header and after a problem", async () => {
    for (const ref of [null, "ABCD-1234"]) {
      const { page } = await draw(ref);
      expect(page.querySelector("#fb-title")?.textContent).toBe("Send feedback");
      expect(page.textContent).not.toContain("Tel Aviv");
      expect(page.textContent).not.toContain("team");
      expect(page.textContent).toContain("Nothing from your workbook is in it");
    }
  });

  it("after the problem it says the message holds codes, and asks what the user was doing", async () => {
    const { page } = await draw("ABCD-1234");
    expect(page.textContent).toContain("Something went wrong. The message says where in the app it stopped and why, in codes we can read.");
    expect(page.querySelector("label[for='fb-note']")?.textContent).toBe("What were you doing? (optional)");
  });

  it("the sent step names the address the backend gave", async () => {
    const { page, dialog, detect } = await draw(null);
    dialog.note.set("Works well");
    await dialog.prepare();
    await detect();
    expect(sent).toEqual([{ ref: null, note: "Works well", useAssistant: false, sessionId: null }]);
    expect(page.querySelector(".fb__message")?.textContent).toContain("Feedback (no step failed)");
    // the send: the clipboard and the mailto link are the browser's; here they are stand-ins
    Object.defineProperty(navigator, "clipboard", { value: { writeText: () => Promise.resolve() }, configurable: true });
    const clicks: string[] = [];
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
      clicks.push(this.href);
    });
    try {
      await dialog.send();
      await detect();
    } finally {
      click.mockRestore();
    }
    expect(clicks).toHaveLength(1);
    expect(clicks[0]).toMatch(/^mailto:someone%40example\.com\?subject=/);
    const step = page.querySelector(".fb__steps li")?.textContent ?? "";
    expect(step).toContain("The message is copied.");
    expect(step).toContain("An email to someone@example.com should be open.");
    expect(page.textContent).not.toContain("Tel Aviv");
  });
});
