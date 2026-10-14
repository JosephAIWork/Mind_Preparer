import { CommonModule } from "@angular/common";
import { Component, computed, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { SessionStore } from "../../core/session.store";

/** A full recalculation in Excel, and what it found: genuine errors, the add-in gap, new vs already there. */
@Component({
  selector: "mr-recalc-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule],
  templateUrl: "./recalc-panel.html",
  styleUrl: "./panels.scss",
})
export class RecalcPanel {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly running = signal(false);
  readonly error = signal("");
  readonly showAll = signal(false);

  readonly result = computed(() => this.store.recalc());
  readonly stale = computed(() => {
    const r = this.result();
    const v = this.store.currentVersion();
    return !!r && !!v && !!r.version_id && r.version_id !== v.id;
  });
  readonly errors = computed(() => {
    const list = this.result()?.formula_errors ?? [];
    return this.showAll() ? list : list.slice(0, 200);
  });
  readonly chip = computed(() => {
    const s = this.result()?.status;
    return s === "PASS" ? "chip chip--ok" : s === "ERROR" ? "chip chip--error" : s ? "chip chip--warn" : "chip";
  });

  ask(e: { sheet: string; cell: string; error: string; formula?: string }, addinGap = false): void {
    this.store.ask({ kind: "recalc", rule_id: "RECALC", sheet: e.sheet, cell: e.cell, error: e.error, formula: e.formula, addin_gap: addinGap });
  }

  async run(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.running()) return;
    this.running.set(true);
    this.error.set("");
    try {
      const out = await this.api.recalculate(id);
      this.store.recalc.set(out);
      this.store.applyAnalysis({ report: out.report, plan: out.plan, readiness: out.readiness, version: out.version ?? undefined });
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.running.set(false);
    }
  }
}
