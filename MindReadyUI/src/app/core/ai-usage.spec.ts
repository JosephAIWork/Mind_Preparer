import { NO_ERRORS_SCHEMA, signal } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import type { AssistantPanel } from "../pages/advanced/assistant-panel";
import { ApiService } from "./api.service";
import { AiUsage, aiUsageLine } from "./models";
import { SessionStore } from "./session.store";

// 1.14.1: the app shows how many tokens its assistant used on this workbook -- counted per workbook on
// the server, kept with the saved workbook, shown under Advanced > Assistant (by the AI key box) and in
// the feedback dialog (core/feedback-dialog.spec.ts). "AI used: 12,340 tokens in 7 calls": input and
// output tokens together; before the first call, "AI not used yet on this workbook".

const usage = (calls: number, input: number, output: number): AiUsage => ({ calls, input_tokens: input, output_tokens: output, cache_read_tokens: 0, cache_write_tokens: 0 });

describe("the token line", () => {
  it("says tokens and calls, or that the assistant was not used yet", () => {
    expect(aiUsageLine(null)).toBe("AI not used yet on this workbook");
    expect(aiUsageLine(undefined)).toBe("AI not used yet on this workbook");
    expect(aiUsageLine(usage(0, 0, 0))).toBe("AI not used yet on this workbook");
    expect(aiUsageLine(usage(1, 90, 10))).toBe("AI used: 100 tokens in 1 call");
    expect(aiUsageLine(usage(7, 12000, 340))).toBe("AI used: 12,340 tokens in 7 calls");
    // the cache counts are part of the input count: never added again
    expect(aiUsageLine({ ...usage(2, 1000, 50), cache_read_tokens: 800, cache_write_tokens: 100 })).toBe("AI used: 1,050 tokens in 2 calls");
  });
});

describe("the Assistant tab's token line", () => {
  let PanelClass: typeof AssistantPanel;
  let served: AiUsage;
  let fetches: number;

  beforeAll(async () => {
    // no vi.resetModules() here: the panel must resolve the same ApiService / SessionStore tokens this file provides
    vi.doMock("carbon-components-angular", () => {
      class Inert {}
      return new Proxy({}, { get: (_t, name) => (name === "then" ? undefined : Inert), has: () => true });
    });
    PanelClass = (await import("../pages/advanced/assistant-panel")).AssistantPanel;
  });

  async function draw(sessionId: string | null = "s1") {
    fetches = 0;
    const store = { sessionId: signal(sessionId), aiUsage: signal<AiUsage | null>(null), fixTarget: signal(null), chatHistory: signal([]), currentVersion: signal(null) };
    const api = {
      aiUsage: () => {
        fetches += 1;
        return Promise.resolve({ aiUsage: served, line: aiUsageLine(served) });
      },
      assistantKey: () => Promise.resolve({ saved: false }),
      plainCodes: () => Promise.resolve({ rules: {}, errors: {} }),
      chatStream: () => Promise.resolve({ text: "Fine.", proposal: null }),
    };
    TestBed.resetTestingModule();
    TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: api }, { provide: SessionStore, useValue: store }] });
    TestBed.overrideComponent(PanelClass, { set: { imports: [], schemas: [NO_ERRORS_SCHEMA] } });
    const fixture = TestBed.createComponent(PanelClass);
    const detect = async () => {
      fixture.detectChanges();
      await fixture.whenStable();
      fixture.detectChanges();
    };
    await detect();
    return { page: fixture.nativeElement as HTMLElement, panel: fixture.componentInstance, store, detect };
  }

  const line = (page: HTMLElement) => page.querySelector('[data-test="ai-usage"]')?.textContent?.replace(/\s+/g, " ").trim();

  it("sits by the AI key box, says what it counts, and reads the count when the tab opens", async () => {
    served = usage(7, 12000, 340);
    const { page } = await draw();
    expect(fetches).toBe(1);
    expect(line(page)).toBe("AI used: 12,340 tokens in 7 calls Tokens the assistant used on this workbook since it was uploaded.");
    expect(page.querySelector('[data-test="ai-usage"] strong')?.textContent).toBe("AI used: 12,340 tokens in 7 calls");
    const key = page.querySelector('[data-show-me="assistant-key"]');
    expect(key?.nextElementSibling?.getAttribute("data-test")).toBe("ai-usage");
  });

  it("says the assistant was not used yet on a fresh workbook, and asks nothing without a workbook", async () => {
    served = usage(0, 0, 0);
    const { page } = await draw();
    expect(line(page)).toBe("AI not used yet on this workbook Tokens the assistant used on this workbook since it was uploaded.");
    const none = await draw(null);
    expect(fetches).toBe(0);
    expect(line(none.page)).toContain("AI not used yet on this workbook");
  });

  it("reads the count again after a chat turn", async () => {
    served = usage(1, 90, 10);
    const { page, panel, detect } = await draw();
    expect(line(page)).toContain("AI used: 100 tokens in 1 call");
    served = usage(2, 190, 20);
    await panel.send("What blocks this workbook?");
    await detect();
    expect(fetches).toBe(2);
    expect(line(page)).toContain("AI used: 210 tokens in 2 calls");
  });
});
