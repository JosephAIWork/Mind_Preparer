import { Injectable, computed, inject, signal } from "@angular/core";
import { Operation } from "./models";
import { SessionStore } from "./session.store";

/** 1.11.6: changes approved in Advanced (Prep actions, assistant proposals) wait here and are written
 *  together, in the order they were added: one Excel session and one re-analysis instead of one each. */
export interface StackEntry {
  key: string;
  source: "prep" | "assistant";
  label: string;
  operations: Operation[];
  sessionId: string;
  versionId: string | null;
  versionLabel: string;
}

export interface StackWarning {
  kind: "overlap" | "stale";
  text: string;
}

interface Target {
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function colNumber(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** The cells an operation writes: its range (an array formula) or its cell; null when it names no cells. */
export function opTarget(o: Operation): Target | null {
  const ref = String(o["range"] ?? o["ref"] ?? o["cell"] ?? "").replace(/\$/g, "").toUpperCase();
  const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(ref);
  if (!m || !o.sheet) return null;
  const c1 = colNumber(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colNumber(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  return { sheet: String(o.sheet), r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

function sameChange(a: Operation, b: Operation): boolean {
  return JSON.stringify({ ...a, action_id: undefined, note: undefined }) === JSON.stringify({ ...b, action_id: undefined, note: undefined });
}

function where(t: Target, o: Operation): string {
  return `${t.sheet}!${String(o["range"] ?? o["ref"] ?? o["cell"])}`;
}

const BIG = 10_000; // a target this many cells or more is compared one by one, not cell by cell

@Injectable({ providedIn: "root" })
export class ChangeStack {
  private store = inject(SessionStore);
  private all = signal<StackEntry[]>([]);
  private n = 0;

  /** The entries of the loaded session (a list made for another session is not shown). */
  readonly entries = computed(() => this.all().filter((e) => e.sessionId === this.store.sessionId()));
  readonly count = computed(() => this.entries().reduce((n, e) => n + e.operations.length, 0));

  /** Every operation in order; an exact repeat of an earlier one is dropped. */
  readonly operations = computed<Operation[]>(() => {
    const out: Operation[] = [];
    const seen = new Set<string>();
    for (const e of this.entries())
      for (const o of e.operations) {
        const k = JSON.stringify({ ...o, action_id: undefined, note: undefined });
        if (!seen.has(k)) {
          seen.add(k);
          out.push(o);
        }
      }
    return out;
  });

  /** Two entries writing the same cell differently (the later one wins, which may break what the
   *  first one fixed), and entries made on an earlier version (their "before" may no longer match). */
  readonly warnings = computed<StackWarning[]>(() => {
    const entries = this.entries();
    const out: StackWarning[] = [];
    const current = this.store.currentVersion();
    for (const e of entries)
      if (e.versionId && current && e.versionId !== current.id)
        out.push({ kind: "stale", text: `"${e.label}" was made on ${e.versionLabel}; the workbook is now ${current.label.split(" — ")[0]}. Its cells may have changed since: check it or remove it.` });

    const cells = new Map<string, { entry: number; op: Operation; t: Target }>();
    const big: { entry: number; op: Operation; t: Target }[] = [];
    const pairs = new Map<string, { a: number; b: number; first: string; n: number }>();
    const clash = (a: { entry: number; op: Operation; t: Target }, b: { entry: number; op: Operation; t: Target }) => {
      if (a.entry === b.entry || sameChange(a.op, b.op)) return;
      const [x, y] = a.entry < b.entry ? [a, b] : [b, a];
      const k = `${x.entry}|${y.entry}`;
      const p = pairs.get(k);
      if (p) p.n += 1;
      else pairs.set(k, { a: x.entry, b: y.entry, first: where(y.t, y.op), n: 1 });
    };
    const overlaps = (s: Target, t: Target) => s.sheet === t.sheet && s.r1 <= t.r2 && t.r1 <= s.r2 && s.c1 <= t.c2 && t.c1 <= s.c2;
    entries.forEach((e, i) => {
      for (const op of e.operations) {
        const t = opTarget(op);
        if (!t) continue;
        const item = { entry: i, op, t };
        for (const b of big) if (overlaps(b.t, t)) clash(b, item);
        if ((t.r2 - t.r1 + 1) * (t.c2 - t.c1 + 1) >= BIG) {
          for (const v of cells.values()) if (overlaps(t, v.t)) clash(v, item);
          big.push(item);
          continue;
        }
        for (let r = t.r1; r <= t.r2; r++)
          for (let c = t.c1; c <= t.c2; c++) {
            const k = `${t.sheet}|${r}|${c}`;
            const prev = cells.get(k);
            if (prev) clash(prev, item);
            cells.set(k, item);
          }
      }
    });
    for (const p of pairs.values())
      out.push({
        kind: "overlap",
        text: `"${entries[p.a].label}" and "${entries[p.b].label}" both change ${p.n === 1 ? p.first : `${p.n} cells (first ${p.first})`}: the later one wins, which may undo or break the first.`,
      });
    return out;
  });

  add(source: StackEntry["source"], label: string, operations: Operation[]): void {
    const id = this.store.sessionId();
    const v = this.store.currentVersion();
    if (!id || !operations.length) return;
    this.all.update((list) => [
      ...list,
      { key: `s${++this.n}`, source, label, operations, sessionId: id, versionId: v?.id ?? null, versionLabel: v?.label.split(" — ")[0] ?? "this version" },
    ]);
  }

  has(label: string): boolean {
    return this.entries().some((e) => e.label === label);
  }

  remove(key: string): void {
    this.all.update((list) => list.filter((e) => e.key !== key));
  }

  clear(): void {
    const id = this.store.sessionId();
    this.all.update((list) => list.filter((e) => e.sessionId !== id));
  }
}
