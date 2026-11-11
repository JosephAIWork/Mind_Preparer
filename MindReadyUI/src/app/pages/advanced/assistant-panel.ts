import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, inject, signal } from "@angular/core";
import { ButtonModule, ProgressBarModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { ChangeStack } from "../../core/change-stack";
import { CodeName } from "../../core/code-name";
import { ChatMessage, FixTarget, Proposal, Version, aiUsageLine, appliedCount, failedCount, failedText } from "../../core/models";
import { PlainCodes } from "../../core/plain-codes";
import { SessionStore } from "../../core/session.store";
import { chatIsAppWords } from "../../core/show-me";
import { ShowMeLink } from "../../core/show-me-link";
import { KeyBox } from "./key-box";

/** The  assistant: questions  about the workbook, or about one finding / one error it was pointed at; proposals applied through Excel. */
@Component({
  selector: "mr-assistant-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, ProgressBarModule, KeyBox, CodeName, ShowMeLink],
  templateUrl: "./assistant-panel.html",
  styleUrl: "./panels.scss",
})
export class AssistantPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly stack = inject(ChangeStack);
  private codes = inject(PlainCodes);

  readonly draft = signal("");
  readonly streaming = signal(false);
  readonly phase = signal("");
  readonly partial = signal("");
  readonly proposal = signal<Proposal | null>(null);
  /** The version the proposal was made on: an Apply from the list above may  make a newer one meanwhile. */
  private proposalVersion: Version | null = null;
  readonly applying = signal(false);
  readonly fraction = signal<number | null>(null);
  readonly applyLine = signal("");
  readonly error = signal("");
  private stop: (() => void) | null = null;

  readonly target = computed(() => this.store.fixTarget());
  /** 1.14.1: "AI used: 12,340 tokens in 7  calls" on this workbook, or "AI  not used yet on this workbook". */
  readonly usageLine = computed(() => aiUsageLine(this.store.aiUsage()));
  /** Only the  app's own words in the chat get a "Show me", never  the assistant's free text. */
  readonly appWords = chatIsAppWords;
  readonly prompts = computed(() => {
    const t = this.target();
    if (!t) return ["What blocks this workbook in Mind?", "Which changes would you make first?"];
    if (t.kind === "recalc" || t.kind === "recalc-group") return ["Propose a fix", "Why does this cell fail?", "Is this a workbook defect or the add-in gap?"];
    return ["Propose a fix", "Why does Mind reject this?", "What is the smallest change?"];
  });

  ngOnInit(): void {
    void this.refreshUsage();
  }

  ngOnDestroy(): void {
    this.stop?.();
  }

  /** The count from the  server (a step in a job -- Fix all issues, the grid namer -- adds to it too). */
  async refreshUsage(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      this.store.aiUsage.set((await this.api.aiUsage(id)).aiUsage);
    } catch {
      // the line keeps what it had
    }
  }

  /** A recalculation error in words: "#REF! at Calc!B7 and 3 cells alike: the formula points at a cell
   *  that no longer exists" (its "RECALC" is no rule: not shown). */
  targetText(t: FixTarget): string {
    const where = [t.sheet, t.cell].filter(Boolean).join("!");
    const n = t.cells?.length ? ` and ${t.cells.length} cells alike` : "";
    const meaning = this.codes.meaning(t.error);
    return `${t.error || "the error"}${where ? " at " + where : ""}${n}${meaning ? ": " + meaning : ""}`;
  }

  /** What follows a finding's code and name: where, the error, the cells  alike. */
  whereText(t: FixTarget): string {
    const where = [t.sheet, t.cell].filter(Boolean).join("!");
    const n = t.cells?.length ? ` and ${t.cells.length} cells alike` : "";
    return `${where ? ", at " + where : ""}${t.error ? " " + t.error : ""}${n}`;
  }

  clearTarget(): void {
    this.store.fixTarget.set(null);
  }

  value(event: Event): string {
    return (event.target as HTMLTextAreaElement).value;
  }

  onKey(event: KeyboardEvent): void {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void this.send(this.draft());
    }
  }

  async send(question: string): Promise<void> {
    const id = this.store.sessionId();
    const q = question.trim();
    if (!id || !q || this.streaming()) return;
    this.draft.set("");
    this.error.set("");
    this.proposal.set(null);
    const history = this.store.chatHistory();
    this.store.chatHistory.set([...history, { role: "user", content: q }]);
    this.streaming.set(true);
    this.phase.set("Thinking…");
    this.partial.set("");
    this.proposalVersion = this.store.currentVersion();
    try {
      const reply = await this.api.chatStream(id, history, q, this.target() ?? undefined, (ev) => {
        if (ev.type === "phase") this.phase.set(ev.phase === "lookup" ? "Looking into the workbook…" : ev.phase === "repair" ? "Checking the proposal…" : "Answering…");
        else this.partial.update((p) => p + ev.text);
      });
      this.store.chatHistory.update((h) => [...h, { role: "assistant", content: reply.text }]);
      this.proposal.set(reply.proposal && reply.proposal.operations?.length ? reply.proposal : null);
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.streaming.set(false);
      this.phase.set("");
      this.partial.set("");
      void this.refreshUsage();
    }
  }

  /** 1.11.6: keep the proposal for the list at the top of the page instead of writing it now. */
  addToList(): void {
    const p = this.proposal();
    if (!p) return;
    const label = `Assistant: ${p.summary || p.operations.length + " changes"}`.slice(0, 120);
    this.stack.add("assistant", label, p.operations, this.proposalVersion);
    this.store.chatHistory.update((h) => [...h, { role: "assistant", content: `${p.operations.length} changes added to the list (applied with the others from the top of the page)`, note: true, tone: "success" }]);
    this.proposal.set(null);
  }

  discard(): void {
    this.proposal.set(null);
  }

  async apply(): Promise<void> {
    const id = this.store.sessionId();
    const p = this.proposal();
    if (!id || !p || this.applying()) return;
    const now = this.store.currentVersion();
    if (this.proposalVersion && now && this.proposalVersion.id !== now.id) {
      // its cells were chosen on the older version: after an inserted row they would land one row off
      this.error.set(`This proposal was made on ${this.proposalVersion.label.split(" — ")[0]}; the workbook is now ${now.label.split(" — ")[0]}. Ask again so it is made on the current version.`);
      return;
    }
    this.applying.set(true);
    this.applyLine.set(`Writing ${p.operations.length} changes through Excel…`);
    this.stop = followApply(this.api, id, (t) => {
      this.fraction.set(t.fraction);
      if (t.text) this.applyLine.set(t.text);
    });
    try {
      const minor = this.target()?.kind === "recalc" || this.target()?.kind === "recalc-group";
      const out = await this.api.apply(id, p.operations, minor ? "minor" : "major");
      this.stop?.();
      this.store.applyAnalysis({ ...out, version: out.version });
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      const line = refused ? `${done} changed, ${refused} refused (${failedText(r)})` : `${done} changes applied`;
      this.store.chatHistory.update((h) => [...h, { role: "assistant", content: `${line} · ${out.version?.label ?? ""}`, note: true, tone: refused ? "warning" : "success" }]);
      this.proposal.set(null);
      this.applyLine.set("");
    } catch (e) {
      this.stop?.();
      this.applyLine.set("");
      this.error.set(errorText(e));
    } finally {
      this.applying.set(false);
      this.fraction.set(null);
    }
  }

  clear(): void {
    this.store.chatHistory.set([]);
    this.proposal.set(null);
    this.error.set("");
  }

  opText(o: { op: string; sheet: string; cell?: string; after?: unknown; formula?: unknown; value?: unknown }): string {
    const after = o.after ?? o.formula ?? o.value ?? "";
    return `${o.op} ${o.sheet}!${o.cell ?? ""}${after !== "" ? " → " + String(after).slice(0, 90) : ""}`;
  }

  trackMsg(i: number, m: ChatMessage): string {
    return `${i}:${m.role}`;
  }
}
