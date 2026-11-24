import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, effect, inject, signal, untracked } from "@angular/core";
import { ButtonModule, ProgressBarModule, TagModule } from "carbon-components-angular";
import { ApiService, errorText, wasCancelled } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { CancelButton } from "../../core/cancel-button";
import { ChangeStack } from "../../core/change-stack";
import { CodeName } from "../../core/code-name";
import { Operation, PivotMode, PrepAction, STOPPED_NOTHING_CHANGED, appliedCount, failedCount, failedText } from "../../core/models";
import { PlainCodes } from "../../core/plain-codes";
import { SessionStore } from "../../core/session.store";
import { ShowMeWhere } from "../../core/show-me";
import { ShowMeLink } from "../../core/show-me-link";
import { ShowMeService } from "../../core/show-me.service";

/** How  many changes  an action's list shows before "… N more". */
const SHOWN_CHANGES = 80;

/** 17  Nov 2026:  which change of an action is for a cell (-1: none). The change of that very cell first;
 *  else the first one on that sheet whose own words name the cell: the  title or the label for a lone
 *  text cell is written BESIDE  the cell, and its note says  which cell it is for ("W25 =  'x' is read
 *  by ..."). A cell of another sheet named in  a note ("Calc!W25") is not this cell. */
export function changeFor(ops: Operation[], where: ShowMeWhere): number {
  const sheet = (where.sheet ?? "").trim().toLowerCase();
  const cell = (where.cell ?? "").replace(/\$/g, "").trim().toUpperCase();
  if (!sheet || !cell) return -1;
  const onSheet = (o: Operation) => String(o.sheet ?? "").trim().toLowerCase() === sheet;
  const own = ops.findIndex((o) => onSheet(o) && String(o["cell"] ?? "").replace(/\$/g, "").trim().toUpperCase() === cell);
  if (own >= 0) return own;
  const literal = cell.replace(/[^A-Z0-9]/g, "\\$&"); // a  range's ":"  and anything else that is no letter or digit
  const named = new RegExp("(^|[^A-Za-z0-9$!])" + literal + "(?![A-Za-z0-9])");
  const words = (o: Operation) => [o["note"], o["after"], o["before"]].filter((x) => typeof x === "string").join(" ");
  return ops.findIndex((o) => onSheet(o) && named.test(words(o)));
}

/** The  Prep actions  one by one: tick what to apply, see the changes, apply through Excel. */
@Component({
  selector: "mr-prep-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, ProgressBarModule, TagModule, CodeName, ShowMeLink, CancelButton],
  templateUrl: "./prep-panel.html",
  styleUrl: "./panels.scss",
})
export class PrepPanel implements OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly stack = inject(ChangeStack);
  private codes = inject(PlainCodes);

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
  readonly byHand = computed(() => (this.store.readiness()?.blocking ?? []).filter((b) => b.fix !== "prep" && b.fix !== "outside_inputs"));
  /** each  one with  what the error values its message quotes mean, once each */
  readonly byHandRows = computed(() => this.byHand().map((b) => ({ b, meanings: this.codes.meaningsIn(b.message) })));
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

  // ---- 17 Nov 2026: "Show me" from a stop  message ends on the line of the problem's own cell ----------
  private showMe = inject(ShowMeService);
  /** the  change a "Show me" pointer asked  for: its action's list is opened and  the change is put first,
   *  marked `prep-cell` (the arrow ends there: "This  change is for cell W25 on sheet  Roll. ..."). Kept
   *  after the pointer has ended (the service forgets the cell then); a plan made again holds other
   *  changes, so nothing is marked in it. */
  readonly pointed = signal<{ action: string; op: Operation } | null>(null);

  constructor() {
    effect(() => {
      const request = this.showMe.request();
      const where = this.showMe.where();
      const plan = this.store.plan();
      if (request?.id !== "prep-action" || !where) return;
      untracked(() => this.pointAt(plan, request.ruleId, where));
    });
  }

  /** Finds the change for the cell: in the action that repairs the pointer's rule first, then in any
   *  other. An action with nothing left to change is not drawn, so it is not looked at. */
  private pointAt(plan: PrepAction[], rule: string | undefined, where: ShowMeWhere): void {
    const repairs = (a: PrepAction) => !rule || (a.rule_ids ?? []).includes(rule);
    for (const a of [...plan.filter(repairs), ...plan.filter((x) => !repairs(x))]) {
      const at = a.count ? changeFor(a.operations, where) : -1;
      if (at < 0) continue;
      this.pointed.set({ action: a.id, op: a.operations[at] });
      if (!this.open().has(a.id)) this.open.set(new Set([...this.open(), a.id]));
      return;
    }
    this.pointed.set(null); // no change is for  that cell: the arrow stays on the action's card, or on the list
  }

  /** The changes an action's list shows: the first ones, with the pointed change first (also when it is
   *  further down the list than the list shows). */
  shownOps(a: PrepAction): Operation[] {
    const head = a.operations.slice(0, SHOWN_CHANGES);
    const first = this.pointedOp(a);
    return first ? [first, ...head.filter((o) => o !== first)].slice(0, SHOWN_CHANGES) : head;
  }

  /** How many changes the list leaves out ("… N more"). */
  moreOps(a: PrepAction): number {
    return Math.max(0, a.operations.length - SHOWN_CHANGES);
  }

  isPointed(a: PrepAction, o: Operation): boolean {
    return this.pointedOp(a) === o;
  }

  private pointedOp(a: PrepAction): Operation | null {
    const p = this.pointed();
    return p && p.action === a.id && a.operations.includes(p.op) ? p.op : null;
  }

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

  /** The card's words for "Show me": the rules this action repairs ("The Prep action '…'" in their
   *  findings leads here). None on  an action with nothing to  change: there is nothing to tick. */
  showMeWords(a: PrepAction): string | null {
    return a.count ? (a.rule_ids ?? []).map((r) => `prep-for-${r}`).join(" ") : null;
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

  /** 1.11.6: put  the ticked actions  on the  list (Apply all writes the list in one Excel session). */
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
      // 17 Nov 2026: stopped with Cancel is no error. No version was made; the workbook comes from the
      // Cancel's own answer, never from  this call.
      if (wasCancelled(e)) {
        this.tone.set("idle");
        this.line.set(STOPPED_NOTHING_CHANGED);
        this.detail.set("");
        return;
      }
      this.tone.set("error");
      this.line.set("Could not apply");
      this.detail.set(errorText(e));
    } finally {
      this.running.set(false);
      this.fraction.set(null);
    }
  }
}
