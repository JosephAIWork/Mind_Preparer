import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule, ProgressBarModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { LoopChoice, LoopChoices, LoopPlan, LoopPlanLoop, appliedCount, failedCount, failedText } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** The Mind loops, each with its own choices (name, first, last, count cell), and the clearing; preview, then write. */
@Component({
  selector: "mr-loop-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule, ProgressBarModule],
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
  private stop: (() => void) | null = null;

  readonly canWrite = computed(() => {
    const p = this.plan();
    return !!p && p.found && p.operation_count > 0 && p.loops.some((l) => l.status === "ready");
  });

  ngOnInit(): void {
    void this.preview();
  }

  ngOnDestroy(): void {
    this.stop?.();
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

  /** the count cell box shows when the loop asks for one, or one was typed */
  wantsCount(l: LoopPlanLoop): boolean {
    return /count cell/i.test(l.why ?? "") || !!this.pick(l, "count_cell") || !!l.suggested_count_cell;
  }

  useSuggested(l: LoopPlanLoop): void {
    if (l.suggested_count_cell) {
      this.setPick(l, "count_cell", l.suggested_count_cell);
      void this.preview();
    }
  }

  /** the value the loop has now, shown greyed in an empty box */
  shown(v: number | null | undefined): string {
    return v === null || v === undefined ? "type it" : String(v);
  }

  async preview(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.loading.set(true);
    this.tone.set("idle");
    this.line.set("");
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

  statusText(status: string): string {
    return status === "ready" ? "ready" : status === "needs_decision" ? "needs a decision" : status === "nothing_to_do" ? "nothing to do" : "not built";
  }

  value(event: Event): string {
    return (event.target as HTMLInputElement).value;
  }
}
