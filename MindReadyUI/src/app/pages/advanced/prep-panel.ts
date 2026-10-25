import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, inject, signal } from "@angular/core";
import { ButtonModule, ProgressBarModule, TagModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { ChangeStack } from "../../core/change-stack";
import { Operation, PivotMode, PrepAction, appliedCount, failedCount, failedText } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** The Prep actions one by one: tick what to apply, see the changes, apply through Excel. */
@Component({
  selector: "mr-prep-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, ProgressBarModule, TagModule],
  templateUrl: "./prep-panel.html",
  styleUrl: "./panels.scss",
})
export class PrepPanel implements OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly stack = inject(ChangeStack);

  readonly selected = signal<Set<string>>(new Set());
  readonly open = signal<Set<string>>(new Set());
  readonly running = signal(false);
  readonly fraction = signal<number | null>(null);
  readonly line = signal("");
  readonly detail = signal("");
  readonly tone = signal<"idle" | "ok" | "warn" | "error">("idle");
  private stop: (() => void) | null = null;
  private planSeen: PrepAction[] | null = null;

  readonly showEmpty = signal(false);
  readonly blocking = computed(() => this.store.plan().filter((a) => a.level === "blocking" && (a.count > 0 || this.showEmpty())));
  readonly optional = computed(() => this.store.plan().filter((a) => a.level !== "blocking" && (a.count > 0 || this.showEmpty())));
  readonly emptyCount = computed(() => this.store.plan().filter((a) => !a.count).length);
  readonly byHand = computed(() => (this.store.readiness()?.blocking ?? []).filter((b) => b.fix !== "prep"));
  readonly selection = computed(() => {
    // Defaults follow the plan (default_on) until the user changes them; a new plan resets.
    const plan = this.store.plan();
    if (plan !== this.planSeen) {
      this.planSeen = plan;
      const def = new Set(plan.filter((a) => a.default_on && a.count > 0).map((a) => a.id));
      queueMicrotask(() => this.selected.set(def));
      return def;
    }
    return this.selected();
  });
  readonly selectedOps = computed<Operation[]>(() => {
    const sel = this.selection();
    return this.store.plan().filter((a) => sel.has(a.id)).flatMap((a) => a.operations);
  });

  readonly pivotSaving = signal(false);
  readonly pivotModeList = computed(() => Object.entries(this.store.pivotModes()) as [PivotMode, string][]);

  /** 1.11.7: what Prep (and Do everything) does with the pivot tables: the plan is made again. */
  async setPivotMode(mode: PivotMode): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.pivotSaving() || mode === this.store.pivotMode()) return;
    this.pivotSaving.set(true);
    try {
      const out = await this.api.setPivotMode(id, mode);
      this.store.pivotMode.set(out.pivot_mode);
      this.store.plan.set(out.plan);
      if (out.readiness) this.store.readiness.set(out.readiness);
    } catch (e) {
      this.tone.set("error");
      this.line.set("Could not change the pivot setting");
      this.detail.set(errorText(e));
    } finally {
      this.pivotSaving.set(false);
    }
  }

  ngOnDestroy(): void {
    this.stop?.();
  }

  isOn(id: string): boolean {
    return this.selection().has(id);
  }

  toggle(id: string): void {
    const next = new Set(this.selection());
    if (next.has(id)) next.delete(id);
    else next.add(id);
    this.selected.set(next);
  }

  isOpen(id: string): boolean {
    return this.open().has(id);
  }

  toggleOpen(id: string): void {
    const next = new Set(this.open());
    if (next.has(id)) next.delete(id);
    else next.add(id);
    this.open.set(next);
  }

  opText(o: Operation): string {
    const target = [o.sheet, o["cell"] ?? o["ref"] ?? o["range"]].filter(Boolean).join("!");
    const after = o["after"] ?? o["formula"] ?? o["value"] ?? o["title"] ?? "";
    const note = o["note"] ? ` (${String(o["note"]).slice(0, 60)})` : "";
    return `${o.op} ${target}${after !== "" && after !== null ? " → " + String(after).slice(0, 80) : ""}${note}`;
  }

  /** 1.11.6: put the ticked actions on the list (Apply all writes the list in one Excel session). */
  addToList(): void {
    const sel = this.selection();
    const actions = this.store.plan().filter((a) => sel.has(a.id) && a.operations.length);
    for (const a of actions) this.stack.add("prep", `Prep: ${a.title}`, a.operations);
    const n = actions.reduce((k, a) => k + a.operations.length, 0);
    this.tone.set("ok");
    this.line.set(`${n} changes added to the list`);
    this.detail.set("Apply them with the others from the list at the top of this page.");
  }

  async apply(): Promise<void> {
    const id = this.store.sessionId();
    const ops = this.selectedOps();
    if (!id || !ops.length || this.running()) return;
    this.running.set(true);
    this.tone.set("idle");
    this.line.set(`Writing ${ops.length} changes through Excel…`);
    this.detail.set("");
    this.fraction.set(null);
    this.stop = followApply(this.api, id, (t) => {
      this.fraction.set(t.fraction);
      this.detail.set(t.text);
    });
    try {
      const out = await this.api.apply(id, ops, "major");
      this.stop?.();
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
      this.line.set("Could not apply");
      this.detail.set(errorText(e));
    } finally {
      this.running.set(false);
      this.fraction.set(null);
    }
  }
}
