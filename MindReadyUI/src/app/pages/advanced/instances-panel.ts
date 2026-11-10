import { CommonModule } from "@angular/common";
import { Component, OnDestroy, OnInit, computed, effect, inject, signal, untracked } from "@angular/core";
import { ButtonModule, InlineLoadingModule, ProgressBarModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { followApply } from "../../core/apply-progress";
import { FinderDimension, FinderFits, FinderResult, InstanceChoices, InstanceLoopOption, InstancePlan, InstancePlanOperation, InstancePlanSplitGrid, appliedCount, failedCount, failedText } from "../../core/models";
import { plainWhy, splitNotes } from "../../core/plain-words";
import { SessionStore } from "../../core/session.store";
import { ShowMeLink } from "../../core/show-me-link";

/** One word by count: plural(1, "cell") = "cell", plural(3, "cell") = "cells", plural(2, "axis", "axes") = "axes". */
function plural(n: number, one: string, many = one + "s"): string {
  return n === 1 ? one : many;
}

/** The Instances tab, one thing at a time: (1) one question, "Make one copy per…" (the macro's loops to
 *  tick, or a column of a table); (2) one line of counts and one button, Build it, the rest behind
 *  "Details"; (3) after the build, the download and "Next: Run in Mind". "Not sure? Let the app look"
 *  opens the finder (where instances could help) under the question, top suggestion first. */
@Component({
  selector: "mr-instances-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule, ProgressBarModule, ShowMeLink],
  templateUrl: "./instances-panel.html",
  styleUrl: "./panels.scss",
  // the key dropdown, styled like the panel's text inputs (panels.scss has no rule for a select)
  styles: [
    `.field select { font: inherit; font-size: 0.875rem; color: var(--cds-text-primary, #161616); background: var(--cds-field-01, #f4f4f4); border: 0; border-bottom: 1px solid var(--cds-border-strong-01, #8d8d8d); padding: 0.5rem 0.75rem; min-width: 14rem; }`,
  ],
})
export class InstancesPanel implements OnInit, OnDestroy {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  /** the key column the user picked (a column letter); empty = the plan's own choice */
  readonly key = signal("");
  /** keys=all: every value of the list the macro reads, even the ones without rows */
  readonly allKeys = signal(false);
  /** the loops the copies are made from (names, in the plan's order); empty = by the key column */
  readonly loops = signal<string[]>([]);
  readonly plan = signal<InstancePlan | null>(null);
  readonly loading = signal(false);
  readonly writing = signal(false);
  /** true after a write, until the next preview: Build stays off, the download link shows */
  readonly written = signal(false);
  /** true after a write that wrote at least one change, until the next preview: the "Done. Next" line */
  readonly built = signal(false);
  readonly downloadUrl = signal<string | null>(null);
  readonly files = signal<string[]>([]);
  readonly fraction = signal<number | null>(null);
  readonly line = signal("");
  readonly detail = signal("");
  readonly tone = signal<"idle" | "ok" | "warn" | "error">("idle");
  readonly showOps = signal(false);
  /** the folds: Details (the cards, the cleared areas, how it runs), the finder, What is this? */
  readonly detailsOpen = signal(false);
  readonly finderOpen = signal(false);
  readonly notesOpen = signal(false);
  readonly whatOpen = signal(false);
  // ---- Where instances could help (GET .../instance-finder): runs only when its fold is opened, never on init ----
  readonly finder = signal<FinderResult | null>(null);
  readonly finding = signal(false);
  /** the error text of the last find */
  readonly findLine = signal("");
  /** a key is saved (api.assistantKey().saved), as the feedback dialog checks it */
  readonly assistantOn = signal(false);
  /** the tick "Explain and rank with the assistant" */
  readonly useAssistant = signal(false);
  /** every suggestion, not only the top one */
  readonly showAll = signal(false);
  /** the one card whose "Show why" is open */
  readonly openId = signal<string | null>(null);
  /** beside the key dropdown: the finder's column is not one of its text candidates; the plan uses it anyway */
  readonly keyHint = signal("");
  private stop: (() => void) | null = null;
  /** the Mind loop tab's choices (each loop's own, and the macro), sent with every plan */
  readonly loopChoices = computed(() => this.store.loopChoices?.() ?? null);
  readonly usesLoopChoices = computed(() => Object.keys(this.loopChoices()?.loops ?? {}).length > 0);

  constructor() {
    // the loop tab previewed other choices: the plan here follows them
    let first = true;
    effect(() => {
      this.loopChoices();
      if (first) {
        first = false;
        return;
      }
      untracked(() => void this.preview());
    });
  }

  readonly candidates = computed(() => this.plan()?.key?.candidates ?? []);
  readonly chosenColumn = computed(() => this.key() || this.plan()?.key?.column || "");
  /** the finder's column when it is not one of the dropdown's candidates: shown as one more option, so the dropdown shows the real choice */
  readonly finderColumn = computed(() => {
    const col = this.chosenColumn();
    return this.keyHint() && col && !this.candidates().some((c) => c.column === col) ? col : "";
  });
  readonly keyWord = computed(() => this.plan()?.key?.header || "key");
  readonly keyCounts = computed<[string, number][]>(() => Object.entries(this.plan()?.key?.counts ?? {}));
  /** the loops copies can be made from (the plan lists them whatever the mode); kept while a preview runs */
  readonly loopOptions = signal<InstanceLoopOption[]>([]);
  /** the number of copies the ticked loops make (the product of their value counts) */
  readonly loopCopies = computed(() => {
    const chosen = this.loops();
    if (!chosen.length) return 0;
    return this.loopOptions().filter((o) => chosen.includes(o.name)).reduce((n, o) => n * (o.count ?? 1), 1);
  });
  /** the question has an answer: loops ticked, a column picked, or the plan's own pick of a column (the key the macro looks up) */
  readonly chosen = computed(() => this.loops().length > 0 || !!this.key() || !!this.plan()?.key?.column);
  /** the question can be answered: the model has a loop to tick or a table column to pick */
  readonly hasChoices = computed(() => this.loopOptions().length > 0 || this.candidates().length > 0 || !!this.finderColumn());
  readonly canWrite = computed(() => {
    const p = this.plan();
    return !!p && p.found && p.status === "ready" && p.operation_count > 0 && !this.written();
  });
  /** the finder's top suggestion, the one shown first */
  readonly topSuggestion = computed<FinderDimension | null>(() => this.finder()?.dimensions[0] ?? null);
  /** The plan's notes in two parts: Prep's blockers (above Build it) and the rest (behind "Notes" once the plan is ready) */
  readonly notes = computed(() => splitNotes(this.plan()?.notes));

  ngOnInit(): void {
    void this.preview();
    this.api.assistantKey().then((k) => this.assistantOn.set(k.saved)).catch(() => {});
  }

  ngOnDestroy(): void {
    this.stop?.();
  }

  private choices(): InstanceChoices {
    const loops = this.loops();
    const lc = this.loopChoices();
    const loopChoices = lc?.loops && Object.keys(lc.loops).length ? { loop_choices: lc.loops } : {};
    if (loops.length) return { loops: loops.join(","), ...loopChoices };
    return { key: this.key() || undefined, keys: this.allKeys() ? "all" : undefined, ...loopChoices, ...(lc?.macro ? { macro: lc.macro } : {}) };
  }

  /** Tick or untick a loop the copies are made from, then preview; the ticked loops keep the plan's (the macro's) order. */
  toggleLoop(name: string): void {
    const on = new Set(this.loops());
    if (on.has(name)) on.delete(name);
    else on.add(name);
    this.loops.set(this.loopOptions().map((o) => o.name).filter((n) => on.has(n)));
    void this.preview();
  }

  /** "CohortRng (6) × Scenario (4)": the loops the copies are made from, with their value counts. */
  copiesTitle(p: InstancePlan): string {
    return (p.copies_from ?? []).map((c) => `${c.name} (${c.count ?? c.values?.length ?? 0})`).join(" × ");
  }

  /** The plan's texts with their jargon said plainly (core/plain-words): "the driver" -> "the cell the macro changes". */
  plainWhy(text: string): string {
    return plainWhy(text);
  }

  /** "6 values" / "4 values" beside a loop's tick box, or why copies cannot be made from it, in plain words. */
  loopText(o: InstanceLoopOption): string {
    if (!o.usable) return plainWhy(o.why ?? "not usable");
    const span = o.first != null && o.last != null && o.kind === "counter" ? `, ${o.first} to ${o.last}` : "";
    return `${o.count ?? "?"} ${plural(o.count ?? 0, "value")}${span}`;
  }

  /** "One copy of the model per combination of CohortRng and Scenario." / "One copy of the model per Product, a loop inside each copy over its rows." */
  perSentence(): string {
    const loops = this.loops();
    if (loops.length) return `One copy of the model per combination of ${loops.join(" and ")}.`;
    return `One copy of the model per ${this.keyWord()}, a loop inside each copy over its rows.`;
  }

  /** "24 copies · 522 result cells": the one line of step 2. */
  summary(p: InstancePlan): string {
    const copies = p.counts?.key_values ?? 0;
    const cells = p.counts?.result_cells ?? 0;
    return `${copies} ${plural(copies, "copy", "copies")} · ${cells} result ${plural(cells, "cell")}`;
  }

  /** "12 changes" / "1 change" on the Build it button. */
  changes(n: number): string {
    return `${n} ${plural(n, "change")}`;
  }

  pickKey(event: Event): void {
    this.setKey((event.target as HTMLSelectElement).value);
  }

  /** The key column into the builder, then the preview; the hint (and the extra dropdown option) stay while the column is not a candidate. */
  private setKey(column: string): void {
    const inList = !column || this.candidates().some((c) => c.column === column);
    this.keyHint.set(inList ? "" : `Column ${column} is not one of the text columns in this list. The plan uses it anyway.`);
    this.key.set(column);
    this.loops.set([]); // a key column means copies by the table, not by loops
    void this.preview();
  }

  toggleAllKeys(): void {
    this.allKeys.set(!this.allKeys());
    void this.preview();
  }

  async preview(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.loading.set(true);
    this.tone.set("idle");
    this.line.set("");
    this.written.set(false);
    this.built.set(false);
    // a new plan: the zip of an earlier write belongs to other choices
    this.downloadUrl.set(null);
    this.files.set([]);
    try {
      const { plan } = await this.api.mindInstancesPlan(id, this.choices());
      this.plan.set(plan);
      if (plan.loop_options) this.loopOptions.set(plan.loop_options);
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
    this.line.set("Writing the instances through Excel…");
    this.detail.set("");
    this.stop = followApply(this.api, id, (t) => {
      this.fraction.set(t.fraction);
      this.detail.set(t.text);
    });
    try {
      const out = await this.api.applyMindInstancesPlan(id, this.choices());
      this.stop?.();
      const { plan: instancePlan, download: _download, files, ...analysis } = out; // `plan` here is the instances plan, not the Prep actions
      this.store.applyAnalysis({ ...analysis, version: out.version });
      this.plan.set(instancePlan);
      this.files.set(files ?? []);
      // the zip exists only when the server says so (a refused change or a failed main workbook leaves none)
      this.downloadUrl.set(out.download ? this.api.mindInstancesDownloadUrl(id) : null);
      this.written.set(true);
      void this.api.resume(id).then((s) => this.store.applyAnalysis({ summary: s.summary, report: s.report, plan: s.plan, readiness: s.readiness, prep_progress: s.prep_progress, versions: s.versions })).catch(() => {});
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      const warnings = out.warnings ?? [];
      this.tone.set(refused || warnings.length ? "warn" : "ok");
      this.line.set(refused ? `${done} written, ${refused} refused` : `${done} changes written`);
      this.detail.set([out.version?.label ?? "", instancePlan?.how_it_runs ?? "", failedText(r), ...warnings].filter(Boolean).join(" · "));
      this.built.set(done > 0);
    } catch (e) {
      this.stop?.();
      this.tone.set("error");
      this.line.set("Could not write the instances");
      this.detail.set(errorText(e));
    } finally {
      this.writing.set(false);
      this.fraction.set(null);
    }
  }

  chipClass(status: string): string {
    return status === "ready" ? "chip chip--ok" : status === "needs_decision" ? "chip chip--warn" : status === "nothing_to_do" ? "chip" : "chip chip--error";
  }

  /** The status in plain words: "needs one answer" (not "a decision"), "left as is" (not "not built"). */
  statusText(status: string): string {
    return status === "ready" ? "ready" : status === "needs_decision" ? "needs one answer" : status === "nothing_to_do" ? "nothing to do" : "left as is";
  }

  chosenByText(by: string | undefined): string {
    return by === "macro" ? "the key the macro looks up" : by === "user" ? "your choice" : by === "heuristic" ? "best guess from the headers" : "";
  }

  /** What happens around the table: rows inserted above the title, the key column put first. */
  splitText(g: InstancePlanSplitGrid): string {
    const parts: string[] = [];
    if (g.inserted_rows) parts.push(`${g.inserted_rows} row${g.inserted_rows === 1 ? "" : "s"} inserted above`);
    if (g.inserted_column) parts.push("the key column inserted first");
    return parts.length ? ` (${parts.join(", ")})` : "";
  }

  /** One line per operation in the "Show the changes" list. */
  opText(op: InstancePlanOperation): string {
    if (op.op === "set_range_values") return `${op.sheet}!${op.ref ?? ""} ${op.values_count ?? op.values?.length ?? 0} values`;
    if (op.op === "insert_row") return `${op.sheet}: row ${op.row ?? ""} inserted`;
    if (op.op === "insert_column") return `${op.sheet}: column ${op.column ?? ""} inserted`;
    if (op.op === "clear_cell") return `${op.sheet}!${op.cell ?? ""} (cleared)`;
    return `${op.sheet}!${op.cell ?? ""} ${op.after ?? "(cleared)"}`;
  }

  /** "Not sure? Let the app look": opens the finder's fold and, the first time, reads the model. */
  toggleFinder(): void {
    const open = !this.finderOpen();
    this.finderOpen.set(open);
    if (open && !this.finder() && !this.finding()) void this.find();
  }

  /** The finder: where instances or a loop would help. Reads every formula; up to two minutes on a big model. */
  async find(refresh = false): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.finding()) return;
    this.finding.set(true);
    this.findLine.set("");
    this.openId.set(null);
    try {
      const { finder } = await this.api.instanceFinder(id, this.assistantOn() && this.useAssistant(), refresh);
      this.finder.set(finder);
    } catch (e) {
      this.finder.set(null);
      this.findLine.set(errorText(e));
    } finally {
      this.finding.set(false);
    }
  }

  /** "Use it": the suggestion's column into the question's dropdown, then the preview (the same path as the dropdown); the finder folds away. */
  useAsKey(d: FinderDimension): void {
    const column = d.key?.column;
    if (!column) return;
    this.finderOpen.set(false);
    this.setKey(column);
  }

  toggleWhy(id: string): void {
    this.openId.set(this.openId() === id ? null : id);
  }

  fitClass(f: FinderFits): string {
    return f === "yes" ? "chip chip--ok" : f === "partly" ? "chip chip--warn" : "chip";
  }

  /** "copies: fits" / "a loop: fits partly" / "leave as is: does not fit". */
  fitWord(o: "instances" | "loop" | "leave", f: FinderFits): string {
    const w = o === "instances" ? "copies" : o === "loop" ? "a loop" : "leave as is";
    return `${w}: ${f === "yes" ? "fits" : f === "partly" ? "fits partly" : "does not fit"}`;
  }

  /** The kinds of evidence as plain phrases: no raw kind words on screen. */
  kindWords(kinds: string[]): string {
    const w: Record<string, string> = { macro: "the macro groups by it", lookups: "lookups keyed by it", branches: "formulas that branch on it", stepping: "rows stepping through it", keyed: "a block reading its own slot", list: "a list of its values", bands: "repeated blocks", sheets: "repeated sheets", counts: "its values" };
    return kinds.map((k) => w[k] ?? k).join(", ");
  }

  /** The "To unify" line from the non-zero counts only; "" when the copies differ in nothing. */
  unifyLine(d: FinderDimension): string {
    const e = d.edits;
    if (!e) return "";
    const parts: string[] = [];
    if (e.numbers_count) parts.push(`${e.numbers_count} typed-in ${plural(e.numbers_count, "number")} to move into a table row per copy`);
    if (e.formula_numbers_count) parts.push(`${e.formula_numbers_count} ${plural(e.formula_numbers_count, "number")} inside formulas to replace by a table cell per copy`);
    if (e.logic_count) parts.push(`${e.logic_count} ${plural(e.logic_count, "cell")} of different logic (an IF on the copy's key, or an override kept per copy)`);
    if (e.size_count) parts.push(`${e.size_count} ${plural(e.size_count, "cell")} present in some copies only (instances with MM_LOOPINSTANCE inside)`);
    return parts.length ? `To unify: ${parts.join("; ")}.` : "";
  }

  /** "Also seen and left alone: …" from the non-zero counts only; "" when nothing was set aside. */
  ignoredLine(f: FinderResult): string {
    const ig = f.ignored ?? {};
    const parts: string[] = [];
    if (ig["time_axes"]) parts.push(`${ig["time_axes"]} time ${plural(ig["time_axes"], "axis", "axes")}`);
    if (ig["filled_areas"]) parts.push(`${ig["filled_areas"]} filled ${plural(ig["filled_areas"], "area")} that ${ig["filled_areas"] === 1 ? "joins" : "join"} nothing`);
    if (ig["small_fills"]) parts.push(`${ig["small_fills"]} small ${plural(ig["small_fills"], "fill")}`);
    return parts.length ? `Also seen and left alone: ${parts.join(", ")}.` : "";
  }

  /** "1,200 formulas · 3 places found · 0 set aside · 2.5 s": what the finder read. */
  finderLine(f: FinderResult): string {
    const c = f.counts;
    return `${c.formulas.toLocaleString("en-US")} formulas · ${c.dimensions} ${plural(c.dimensions, "place")} found · ${c.chains} set aside · ${c.seconds.toFixed(1)} s`;
  }

  /** The builder's key next to the finder's top one: "the builder picked the same key" / "… another key: column X"; "" when there is nothing to compare. */
  builderLine(f: FinderResult): string {
    const b = f.builder;
    if (!b || b.matches_top == null || !b.column) return "";
    const who = b.source === "plan" ? "the builder picked" : "the builder's rule picks";
    return b.matches_top ? `${who} the same key` : `${who} another key: column ${b.column}`;
  }

  /** The top suggestion only, or every one after "Show all". */
  shown(): FinderDimension[] {
    const d = this.finder()?.dimensions ?? [];
    return this.showAll() ? d : d.slice(0, 1);
  }
}
