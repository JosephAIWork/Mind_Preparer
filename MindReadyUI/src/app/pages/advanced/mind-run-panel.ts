import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, inject, signal } from "@angular/core";
import { ButtonModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { MindLoopEvent, MindLoopIteration, MindLoopReport } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** The app running itself against real Mind: prepare, gates, upload, convert, run, decide; one run at a time. */
@Component({
  selector: "mr-mind-run-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule],
  templateUrl: "./mind-run-panel.html",
  styleUrl: "./panels.scss",
})
export class MindRunPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  readonly maxIterations = signal(4);
  readonly runModel = signal(true);
  readonly useAssistant = signal(true);
  readonly state = signal<"idle" | "running" | "done" | "error">("idle");
  readonly events = signal<MindLoopEvent[]>([]);
  readonly report = signal<MindLoopReport | null>(null);
  readonly summary = signal("");
  readonly error = signal("");
  private next = 0;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private live = true;

  readonly iterations = computed<MindLoopIteration[]>(() => this.report()?.iterations ?? []);
  readonly verdictClass = computed(() => {
    const v = this.report()?.verdict;
    return v === "converged" ? "chip chip--ok" : v === "stuck" || v === "exhausted" ? "chip chip--warn" : "chip";
  });

  ngOnInit(): void {
    void this.poll();
  }

  ngOnDestroy(): void {
    this.live = false;
    if (this.timer) clearTimeout(this.timer);
  }

  async start(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.state() === "running") return;
    this.error.set("");
    try {
      await this.api.startMindLoop(id, { maxIterations: this.maxIterations(), runModel: this.runModel(), useAssistant: this.useAssistant() });
      this.events.set([]);
      this.report.set(null);
      this.next = 0;
      this.state.set("running");
      void this.poll();
    } catch (e) {
      this.error.set(errorText(e));
    }
  }

  private async poll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.live) return;
    try {
      const s = await this.api.mindLoopStatus(id, this.next);
      if (s.events?.length) this.events.update((e) => [...e, ...s.events].slice(-300));
      this.next = s.next ?? this.next;
      this.state.set(s.state);
      if (s.report) this.report.set(s.report);
      if (s.summary) this.summary.set(s.summary);
      if (s.error) this.error.set(s.error);
    } catch {
      /* the next tick tries again */
    }
    if (this.live && this.state() === "running") this.timer = setTimeout(() => void this.poll(), 3000);
  }

  gate(it: MindLoopIteration, which: "apply" | "numbers" | "names" | "convert" | "run"): { label: string; cls: string } {
    const ok = (b: boolean | null | undefined) => (b === true ? "chip chip--ok" : b === false ? "chip chip--error" : "chip");
    switch (which) {
      case "apply":
        return { label: it.apply ? `${it.apply.applied} applied` : "not prepared", cls: ok(it.apply ? it.apply.failed === 0 : undefined) };
      case "numbers":
        return { label: it.numbers?.ran ? (it.numbers.match ? "numbers match" : `${it.numbers.differences ?? "?"} differ`) : "numbers skipped", cls: ok(it.numbers?.ran ? it.numbers.match : undefined) };
      case "names":
        return { label: it.names ? `${it.names.unnamed} untitled` : "names skipped", cls: ok(it.names?.ok) };
      case "convert":
        return { label: it.mind?.skipped ? "Mind skipped" : it.mind?.convert ? (it.mind.convert.success ? "converted" : "convert failed") : "no convert", cls: ok(it.mind?.convert?.success) };
      case "run":
        return { label: it.mind?.run ? (it.mind.run.completed ? (it.mind.run.audit_consistent ? "run ok" : "run inconsistent") : "run failed") : "not run", cls: ok(it.mind?.run ? it.mind.run.completed && it.mind.run.audit_consistent : undefined) };
    }
  }

  num(event: Event): number {
    return Number((event.target as HTMLInputElement).value) || 1;
  }

  when(t: number): string {
    const d = new Date(t * 1000);
    return isNaN(d.getTime()) ? "" : d.toLocaleTimeString();
  }
}
