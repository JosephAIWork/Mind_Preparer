import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule, ProgressBarModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { LoopChoices, LoopPlan, appliedCount, failedCount, failedText } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** The Mind loop with its choices: name, first, last, count cell, clearing; preview, then write. */
@Component({
  selector: "mr-loop-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule, ProgressBarModule],
  templateUrl: "./loop-panel.html",
  styleUrl: "./panels.scss",
})
export class LoopPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  readonly name = signal("");
  readonly first = signal("");
  readonly last = signal("");
  readonly countCell = signal("");
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

  readonly needsCount = computed(() => (this.plan()?.loops ?? []).some((l) => l.status === "needs_decision" && (l.notes ?? []).some((n) => /count/i.test(n))));
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
    return { name: this.name() || undefined, first: this.first() || undefined, last: this.last() || undefined, count_cell: this.countCell() || undefined, clear_old: this.clearOld() || undefined, fixed_ranges: this.fixedRanges() || undefined };
  }

  async preview(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.loading.set(true);
    this.tone.set("idle");
    this.line.set("");
    try {
      const { plan } = await this.api.mindLoopPlan(id, this.choices());
      this.plan.set(plan);
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
