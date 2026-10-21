import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, inject, signal } from "@angular/core";
import { ButtonModule, ProgressBarModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { ChangeStack } from "../../core/change-stack";
import { appliedCount, failedCount, failedText } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** 1.11.6: the list of approved changes waiting to be written, with one Apply for all of them. */
@Component({
  selector: "mr-stack-bar",
  standalone: true,
  imports: [CommonModule, ButtonModule, ProgressBarModule],
  templateUrl: "./stack-bar.html",
  styleUrl: "./panels.scss",
})
export class StackBar implements OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly stack = inject(ChangeStack);

  readonly running = signal(false);
  readonly fraction = signal<number | null>(null);
  readonly line = signal("");
  readonly detail = signal("");
  readonly tone = signal<"idle" | "ok" | "warn" | "error">("idle");
  readonly open = signal(false);
  /** An entry made on an older version: its cells may have moved (a row inserted since), and the
   *  server does not compare "before", so it would be written blind. It must be removed first. */
  readonly stale = computed(() => this.stack.warnings().some((w) => w.kind === "stale"));
  private stop: (() => void) | null = null;

  ngOnDestroy(): void {
    this.stop?.();
  }

  async applyAll(): Promise<void> {
    const id = this.store.sessionId();
    const ops = this.stack.operations();
    if (!id || !ops.length || this.running() || this.stale()) return;
    // the entries written now: one added while the Apply runs stays in the list for the next one
    const keys = new Set(this.stack.entries().map((e) => e.key));
    this.running.set(true);
    this.tone.set("idle");
    this.line.set(`Writing ${ops.length} changes through Excel, in one go…`);
    this.detail.set("");
    this.fraction.set(null);
    this.stop = followApply(this.api, id, (t) => {
      this.fraction.set(t.fraction);
      this.detail.set(t.text);
    });
    try {
      const out = await this.api.apply(id, ops, "major");
      this.stop?.();
      this.stack.removeKeys(keys);
      if (this.store.sessionId() !== id) {
        // another workbook was opened meanwhile: its analysis must not be replaced with this one
        this.tone.set("idle");
        this.line.set("");
        this.detail.set("");
        return;
      }
      this.store.applyAnalysis({ ...out, version: out.version });
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      this.tone.set(refused ? "warn" : "ok");
      this.line.set(refused ? `${done} applied, ${refused} refused` : `${done} changes applied`);
      this.detail.set([out.version?.label ?? "", r.verified_opens_in_excel ? "opens in Excel" : "", failedText(r)].filter(Boolean).join(" · "));
    } catch (e) {
      this.stop?.();
      this.tone.set("error");
      this.line.set("Could not apply the list");
      this.detail.set(errorText(e));
    } finally {
      this.running.set(false);
      this.fraction.set(null);
    }
  }
}
