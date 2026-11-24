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
  text: "Mind Ready 1.14.1 (source) · ref ABCD-1234\n\nFeedback (no step failed)\n",
  subject: "Mind Ready 1.14.1 feedback · ref ABCD-1234",
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

  async function draw(ref: string | null, note?: string) {
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
    if (note === undefined) feedback.open(ref);
    else feedback.open(ref, note);
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
    // opened from the header, with no note: the body is the one sentence, as before 17 Nov 2026
    expect(decodeURIComponent(clicks[0].split("&body=")[1])).toBe("Paste the Mind Ready message here (Ctrl+V), then send.");
    const step = page.querySelector(".fb__steps li")?.textContent ?? "";
    expect(step).toContain("The message is copied.");
    expect(step).toContain("An email to someone@example.com should be open.");
    expect(page.textContent).not.toContain("Tel Aviv");
  });

  it("shows what the assistant used on the open workbook once, inside the message (1.14.1)", async () => {
    // the backend puts "AI used: ..." into the message's log part (app/feedback.py compose); the dialog
    // repeats nothing under the box (12 November 2026 review: the line showed twice in a row)
    const original = message.text;
    message.aiUsage = { calls: 7, input_tokens: 12000, output_tokens: 340, cache_read_tokens: 0, cache_write_tokens: 0 };
    message.text = original + "  Workbook: 3 sheets\n  AI used: 12,340 tokens in 7 calls\n";
    try {
      const { page, dialog, detect } = await draw(null);
      dialog.note.set("Works well");
      await dialog.prepare();
      await detect();
      const text = page.textContent ?? "";
      expect(text.split("AI used: 12,340 tokens in 7 calls").length - 1).toBe(1);
      expect(page.querySelector(".fb__message")?.textContent).toContain("AI used: 12,340 tokens in 7 calls");
      expect(page.querySelector('[data-test="fb-ai-usage"]')).toBeNull();
      expect(page.querySelector(".fb__ai")).toBeNull();
    } finally {
      message.aiUsage = null;
      message.text = original;
    }
  });

  // ---- 17 Nov 2026: "Send feedback" with the problem already written (Start could not fix something) ----
  const NOTE = "Start could not fix this by itself: a formula reads a lone text cell. The assistant tried twice: its change altered a result and was undone.";
  const PASTE = "Paste the Mind Ready message here (Ctrl+V), then send.";

  /** "Copy and open email", with stand-ins for the clipboard and the mailto link: the email's body, decoded. */
  async function emailBodyOf(dialog: FeedbackDialog, detect: () => Promise<void>): Promise<string> {
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
    return decodeURIComponent(clicks[0].split("&body=")[1]);
  }

  const noteBox = (page: HTMLElement) => page.querySelector<HTMLTextAreaElement>("#fb-note");

  it("the service opens the dialog with a ref, a note, both or neither", async () => {
    const { feedback } = await draw(null);
    expect(feedback.dialog()).toEqual({ ref: null, note: "" }); // the header's button: feedback.open()
    feedback.open("ABCD-1234");
    expect(feedback.dialog()).toEqual({ ref: "ABCD-1234", note: "" }); // the problem box under the header
    feedback.open("ABCD-1234", NOTE);
    expect(feedback.dialog()).toEqual({ ref: "ABCD-1234", note: NOTE }); // Start's question
    feedback.open(null, NOTE);
    expect(feedback.dialog()).toEqual({ ref: null, note: NOTE }); // a question that came with no incident
    feedback.close();
    expect(feedback.dialog()).toBeNull();
  });

  it("opened with a note, the note box holds it already and the message can be prepared at once", async () => {
    const { page, dialog, detect } = await draw("ABCD-1234", NOTE);
    expect(dialog.note()).toBe(NOTE);
    expect(noteBox(page)?.value).toBe(NOTE);
    expect(page.querySelector("label[for='fb-note']")?.textContent).toBe("What happened (change it or add to it if you like)");
    await dialog.prepare();
    await detect();
    expect(sent).toEqual([{ ref: "ABCD-1234", note: NOTE, useAssistant: false, sessionId: null }]);
    expect(page.querySelector(".fb__message")?.textContent).toContain("Feedback (no step failed)");

    // with no incident behind it the note is what makes the message: the button is not held back
    const plain = await draw(null, NOTE);
    const prepare = Array.from(plain.page.querySelectorAll("button")).find((b) => b.textContent?.includes("Prepare the message"));
    expect(prepare?.disabled).toBe(false);
    expect((await draw(null)).page.querySelector("label[for='fb-note']")?.textContent).toBe("Your feedback");
  });

  it("the email's body starts with the note, then an empty line, then the sentence that says to paste", async () => {
    const { dialog, detect } = await draw("ABCD-1234", NOTE);
    await dialog.prepare();
    await detect();
    const body = await emailBodyOf(dialog, detect);
    expect(body.startsWith(NOTE)).toBe(true);
    expect(body).toBe(`${NOTE}\r\n\r\n${PASTE}`);
    expect(body.split("\r\n")).toEqual([NOTE, "", PASTE]);
    // the full message is not in the email: it is on the clipboard, to paste
    expect(body).not.toContain("Feedback (no step failed)");
  });

  it("puts at most 500 characters of the note into the email", async () => {
    const mod = await import("./feedback-dialog");
    expect(mod.EMAIL_NOTE_MAX).toBe(500);
    const long = "Start could not fix this by itself. ".repeat(30); // 1080 characters
    const { dialog, detect } = await draw("ABCD-1234", long);
    await dialog.prepare();
    await detect();
    expect((sent[0] as { note: string }).note).toBe(long); // the message itself gets the whole note
    const body = await emailBodyOf(dialog, detect);
    const [note, empty, paste] = body.split("\r\n");
    expect(note.length).toBeLessThanOrEqual(500);
    expect(long.startsWith(note)).toBe(true);
    expect(note.length).toBeGreaterThan(490);
    expect([empty, paste]).toEqual(["", PASTE]);
    expect(mod.emailBody("x".repeat(600))).toBe("x".repeat(500) + "\r\n\r\n" + PASTE);
    expect(mod.emailBody("")).toBe(PASTE);
    expect(mod.emailBody("   ")).toBe(PASTE);
  });

  it("the email starts with the note as the person left it", async () => {
    const changed = await draw("ABCD-1234", NOTE);
    changed.dialog.note.set(NOTE + " It was the Model D model.");
    await changed.dialog.prepare();
    await changed.detect();
    expect(await emailBodyOf(changed.dialog, changed.detect)).toBe(`${NOTE} It was the Model D model.\r\n\r\n${PASTE}`);

    const emptied = await draw("ABCD-1234", NOTE);
    emptied.dialog.note.set("");
    await emptied.dialog.prepare();
    await emptied.detect();
    expect(await emailBodyOf(emptied.dialog, emptied.detect)).toBe(PASTE);

    // words typed into a dialog opened with none stay out of the email, as before: they are in the message
    const typed = await draw(null);
    typed.dialog.note.set("Works well");
    await typed.dialog.prepare();
    await typed.detect();
    expect(await emailBodyOf(typed.dialog, typed.detect)).toBe(PASTE);
  });

  it("stays in the shell for the whole visit: each opening brings its own note, the header's an empty box", async () => {
    const { page, dialog, detect, feedback } = await draw("ABCD-1234", NOTE);
    expect(noteBox(page)?.value).toBe(NOTE);
    dialog.close();
    await detect();
    expect(noteBox(page)).toBeNull(); // closed: nothing drawn
    feedback.open("WXYZ-5678", "Start could not fix this by itself: a formula has a broken reference. The assistant was not asked.");
    await detect();
    expect(noteBox(page)?.value).toBe("Start could not fix this by itself: a formula has a broken reference. The assistant was not asked.");
    dialog.close();
    feedback.open(); // the header's button, after a question
    await detect();
    expect(noteBox(page)?.value).toBe("");
    expect(dialog.note()).toBe("");
    // typing does not come back by itself, and is gone after Cancel
    dialog.note.set("half a sentence");
    await detect();
    expect(noteBox(page)?.value).toBe("half a sentence");
    dialog.close();
    feedback.open("ABCD-1234");
    await detect();
    expect(noteBox(page)?.value).toBe("");
  });

  // 17 Nov 2026 (live test L2b-7): opened from Start's question nothing went wrong: the app could not fix one
  // problem by itself. The rest of the dialog stays as it is.
  it("opened from Start's question it does not lead with 'Something went wrong.'", async () => {
    const lead = (page: HTMLElement) => (page.querySelector(".fb__text")?.textContent ?? "").replace(/\s+/g, " ").trim();
    const asked = await draw("ABCD-1234", NOTE);
    expect(lead(asked.page)).toBe("The app could not fix one problem by itself. The message says where in the app it stopped and why, in codes we can read. Nothing from your workbook is in it: no file name, no sheet names, no cell addresses, no values.");
    expect(asked.page.textContent).not.toContain("Something went wrong");
    // a question that came with no incident: no codes to speak of
    const plain = await draw(null, NOTE);
    expect(lead(plain.page)).toMatch(/^The app could not fix one problem by itself\. Nothing from your workbook is in it/);
    // the problem box under the header and the header's own button are as before
    expect(lead((await draw("ABCD-1234")).page)).toMatch(/^Something went wrong\. The message says where in the app it stopped and why/);
    expect(lead((await draw(null)).page)).toMatch(/^Tell us what works, what does not, and what is missing\./);
  });
});
