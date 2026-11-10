import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule, ProgressBarModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { LoopChoice, LoopChoices, LoopPlan, LoopPlanLoop, appliedCount, failedCount, failedText } from "../../core/models";
import { SessionStore } from "../../core/session.store";
import { ShowMeLink } from "../../core/show-me-link";

/** The box a loop still needs an answer in: the cell for the number of passes, the first value, or the last. */
export type LoopAnswerKey = "count_cell" | "first" | "last";

/** The Mind loop tab: one sentence that says what the macro runs, one button that builds it. The choices
 *  (each loop's own name, runs from / to, a cell for the number of passes; the macro; the two options;
 *  Preview) sit behind the fold "Change the choices"; what gets written behind "What gets written". A
 *  loop that still needs one answer shows only that box, above the folds. */
@Component({
  selector: "mr-loop-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule, ProgressBarModule, ShowMeLink],
  templateUrl: "./loop-panel.html",
  styleUrl: "./panels.scss",
  // the macro dropdown, styled like the panel's text inputs (as the instances panel's key dropdown)
  styles: [
    `.field select { font: inherit; font-size: 0.875rem; color: var(--cds-text-primary, #161616); background: var(--cds-field-01, #f4f4f4); border: 0; border-bottom: 1px solid var(--cds-border-strong-01, #8d8d8d); padding: 0.5rem 0.75rem; min-width: 14rem; }`,
  ],
})
export class LoopPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  /** each loop's own choices, by driver cell */
  readonly picks = signal<Record<string, LoopChoice>>({});
  /** the macro whose loops are built; empty = the plan's own (the first macro with a loop) */
  readonly macro = signal("");
  readonly clearOld = signal(false);
  readonly fixedRanges = signal(false);
  readonly plan = signal<LoopPlan | null>(null);
  readonly loading = signal(false);
  readonly writing = signal(false);
  readonly fraction = signal<number | null>(null);
  readonly line = signal("");
  readonly detail = signal("");
  readonly tone = signal<"idle" | "ok" | "warn" | "error">("idle");
  readonly showOps = signal(false);
  /** the folds: closed until the user opens them */
  readonly choicesOpen = signal(false);
  readonly writtenOpen = signal(false);
  readonly whatOpen = signal(false);
  /** the last build wrote the loop: "Done. Next: Run in Mind" */
  readonly built = signal(false);
  private stop: (() => void) | null = null;

  readonly canWrite = computed(() => {
    const p = this.plan();
    return !!p && p.found && p.operation_count > 0 && p.loops.some((l) => l.status === "ready");
  });

  /** The loops of the macro this plan builds (another macro's loops are listed in the fold only). */
  readonly ownLoops = computed(() => (this.plan()?.loops ?? []).filter((l) => !l.other_macro));

  /** The loops that still need one answer, each with the box that is missing. */
  readonly needsAnswer = computed(() => this.ownLoops().filter((l) => l.status === "needs_decision" && !!l.driver && !!this.missingKey(l)));

  /** One sentence in plain words: what the macro runs, and that Mind can do the same. */
  readonly sentence = computed(() => {
    const p = this.plan();
    if (!p || !p.found) return "";
    const parts = this.ownLoops()
      .filter((l) => l.status !== "not_built")
      .map((l) => `${l.name}${this.span(l)}`);
    const runs = parts.length ? `The macro runs ${this.joined(parts)}.` : "The macro's loop cannot be written as it is.";
    const open = this.needsAnswer().length;
    const open2 = this.ownLoops().filter((l) => l.status === "needs_decision").length;
    const asked = Math.max(open, open2);
    if (!parts.length) return runs;
    if (asked) return `${runs} Mind can do the same, after ${asked === 1 ? "one answer" : asked + " answers"} below.`;
    return `${runs} Mind can do the same.`;
  });

  /** How many things the build writes, for the fold's title: the formulas, plus the cleared areas. */
  readonly writtenCount = computed(() => {
    const p = this.plan();
    if (!p) return 0;
    return p.operation_count + (this.clearOld() ? p.clears.filter((c) => c.included).length : 0);
  });

  ngOnInit(): void {
    void this.preview();
  }

  ngOnDestroy(): void {
    this.stop?.();
  }

  /** "Scenario 0 to 3", "Cohort from 2021" (the last value still to come), "" when nothing is known */
  private span(l: LoopPlanLoop): string {
    const has = (v: number | null | undefined) => v !== null && v !== undefined;
    if (has(l.first) && has(l.last)) return ` ${l.first} to ${l.last}`;
    if (has(l.first)) return ` from ${l.first}`;
    if (has(l.last)) return ` to ${l.last}`;
    return "";
  }

  /** "A", "A and B", "A, B and C" */
  private joined(parts: string[]): string {
    if (parts.length <= 1) return parts.join("");
    return `${parts.slice(0, -1).join(", ")} and ${parts[parts.length - 1]}`;
  }

  private choices(): LoopChoices {
    const loops: Record<string, LoopChoice> = {};
    for (const [driver, c] of Object.entries(this.picks())) {
      const kept = Object.fromEntries(Object.entries(c).map(([k, v]) => [k, (v ?? "").trim()]).filter(([, v]) => v)) as LoopChoice;
      if (Object.keys(kept).length) loops[driver] = kept;
    }
    return { clear_old: this.clearOld() || undefined, fixed_ranges: this.fixedRanges() || undefined, loops: Object.keys(loops).length ? loops : undefined, macro: this.macro() || undefined };
  }

  /** Another macro's loops: Excel runs each macro on its own, so the plan builds one macro at a time. */
  pickMacro(event: Event): void {
    this.macro.set((event.target as HTMLSelectElement).value);
    void this.preview();
  }

  /** what was typed for one loop's choice */
  pick(l: LoopPlanLoop, key: keyof LoopChoice): string {
    return (l.driver && this.picks()[l.driver]?.[key]) || "";
  }

  setPick(l: LoopPlanLoop, key: keyof LoopChoice, v: string): void {
    if (!l.driver) return;
    const driver = l.driver;
    this.picks.update((all) => ({ ...all, [driver]: { ...all[driver], [key]: v } }));
  }

  /** the box for the number of passes shows when the loop asks for one, or one was typed */
  wantsCount(l: LoopPlanLoop): boolean {
    return /count cell/i.test(l.why ?? "") || !!this.pick(l, "count_cell") || !!l.suggested_count_cell;
  }

  /** The one box a loop that needs an answer is missing: the cell for the number of passes when its
   *  why asks for one, else the first or the last value the plan does not know. */
  missingKey(l: LoopPlanLoop): LoopAnswerKey | null {
    if (l.status !== "needs_decision") return null;
    if (/count cell/i.test(l.why ?? "") || l.suggested_count_cell) return "count_cell";
    if (l.first === null || l.first === undefined) return "first";
    if (l.last === null || l.last === undefined) return "last";
    return null;
  }

  /** The missing box's label: the loop's own why, in plain words, else what the box is for. */
  answerLabel(l: LoopPlanLoop): string {
    const key = this.missingKey(l);
    const why = this.plainWhy(l.why ?? "");
    if (why) return why;
    return key === "count_cell" ? `${l.name}: a cell for the number of passes` : key === "first" ? `${l.name} runs from` : `${l.name} runs to`;
  }

  /** The missing box's greyed-in words. */
  answerPlaceholder(l: LoopPlanLoop): string {
    const key = this.missingKey(l);
    if (key === "count_cell") return l.suggested_count_cell || "Sheet!A1 (an empty cell)";
    return "a number, or a cell (Sheet!A1)";
  }

  /** The plan's why with its jargon said plainly: "count cell" -> "a cell for the number of passes",
   *  "the driver" -> "the cell the macro changes", "MM_LOOP needs" -> "Mind needs". */
  plainWhy(why: string): string {
    return why
      .replace(/\s*\(count cell\)/gi, "")
      .replace(/\bMM_LOOP needs a cell that holds the count\b/g, "Mind needs a cell that holds the number of passes")
      .replace(/\bthe count cell\b/gi, "the cell for the number of passes")
      .replace(/\bcount cell\b/gi, "a cell for the number of passes")
      .replace(/\bthe driver cell\b/gi, "the cell the macro changes")
      .replace(/\bthe driver\b/gi, "the cell the macro changes")
      .replace(/\s{2,}/g, " ")
      .trim();
  }

  useSuggested(l: LoopPlanLoop): void {
    if (l.suggested_count_cell) {
      this.setPick(l, "count_cell", l.suggested_count_cell);
      void this.preview();
    }
  }

  /** "1 formula", "120 formulas" */
  formulas(n: number): string {
    return `${n} formula${n === 1 ? "" : "s"}`;
  }

  /** the value the loop has now, shown greyed in an empty box */
  shown(v: number | null | undefined): string {
    return v === null || v === undefined ? "type it" : String(v);
  }

  /** The Instances tab: one copy of the model per value, instead of one model that loops. */
  openInstances(): void {
    this.store.advancedTab?.set("instances");
  }

  async preview(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.loading.set(true);
    this.tone.set("idle");
    this.line.set("");
    this.built.set(false);
    try {
      const choices = this.choices();
      const { plan } = await this.api.mindLoopPlan(id, choices);
      this.plan.set(plan);
      // the instances panel builds its copies from loops with the same choices
      this.store.loopChoices?.set(choices.loops || choices.macro ? { loops: choices.loops, macro: choices.macro } : null);
    } catch (e) {
      this.plan.set(null);
      this.tone.set("error");
      this.line.set("Could not read the plan");
      this.detail.set(errorText(e));
    } finally {
      this.loading.set(false);
    }
  }

  async write(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.writing() || !this.canWrite()) return;
    this.writing.set(true);
    this.tone.set("idle");
    this.built.set(false);
    this.line.set("Writing the loop through Excel…");
    this.detail.set("");
    this.stop = followApply(this.api, id, (t) => {
      this.fraction.set(t.fraction);
      this.detail.set(t.text);
    });
    try {
      const out = await this.api.applyMindLoopPlan(id, this.choices());
      this.stop?.();
      const { plan: loopPlan, ...analysis } = out; // `plan` here is the LOOP plan, not the Prep actions
      this.store.applyAnalysis({ ...analysis, version: out.version });
      this.plan.set(loopPlan);
      void this.api.resume(id).then((s) => this.store.applyAnalysis({ summary: s.summary, report: s.report, plan: s.plan, readiness: s.readiness, prep_progress: s.prep_progress, versions: s.versions })).catch(() => {});
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      this.tone.set(refused ? "warn" : "ok");
      this.line.set(refused ? `${done} written, ${refused} refused` : `${done} loop formulas written`);
      this.detail.set([out.version?.label ?? "", out.plan.how_it_runs ?? "", failedText(r)].filter(Boolean).join(" · "));
      this.built.set(done > 0);
    } catch (e) {
      this.stop?.();
      this.tone.set("error");
      this.line.set("Could not write the loop");
      this.detail.set(errorText(e));
    } finally {
      this.writing.set(false);
      this.fraction.set(null);
    }
  }

  chipClass(status: string): string {
    return status === "ready" ? "chip chip--ok" : status === "needs_decision" ? "chip chip--warn" : status === "nothing_to_do" ? "chip" : "chip chip--error";
  }

  /** The status in plain words: "left as is" is not an error, the model keeps what it has there. */
  statusText(status: string): string {
    return status === "ready" ? "ready" : status === "needs_decision" ? "needs one answer" : status === "nothing_to_do" ? "nothing to do" : "left as is";
  }

  value(event: Event): string {
    return (event.target as HTMLInputElement).value;
  }
}
