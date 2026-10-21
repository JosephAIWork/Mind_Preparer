import { Injectable, computed, inject, signal } from "@angular/core";
import { Operation, Version } from "./models";
import { SessionStore } from "./session.store";

/** 1.11.6: changes approved in Advanced (Prep actions, assistant proposals) wait here and are written
 *  together: one Excel session and one re-analysis instead of one each. The server writes the pivot
 *  tables first, then the cell changes in the order added, then the inserted rows and columns, the
 *  cells planned for after the inserts, and the sheet-level changes. */
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

const PIVOT_OPS = new Set(["pivot_to_formulas", "pivot_to_values"]);
const INSERT_OPS = new Set(["insert_row", "insert_column"]);

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

  /** Every operation in order; an exact repeat is dropped, keeping the last copy (B2 = 3, then 3.5,
   *  then 3 again must end at 3, so the earlier copy goes, not the later one). */
  readonly operations = computed<Operation[]>(() => {
    const out: Operation[] = [];
    const seen = new Set<string>();
    const flat = this.entries().flatMap((e) => e.operations);
    for (let i = flat.length - 1; i >= 0; i--) {
      const k = JSON.stringify({ ...flat[i], action_id: undefined, note: undefined });
      if (!seen.has(k)) {
        seen.add(k);
        out.push(flat[i]);
      }
    }
    return out.reverse();
  });

  /** Two entries writing the same cell differently (the later one wins, which may break what the
   *  first one fixed; a pivot table is always converted first, so the cell write wins over it), entries
   *  from Prep and from the assistant that both insert rows or columns on one sheet (the cells planned
   *  around one insert do not follow the other), and entries made on an earlier version (their
   *  "before" may no longer match). */
  readonly warnings = computed<StackWarning[]>(() => {
    const entries = this.entries();
    const out: StackWarning[] = [];
    const current = this.store.currentVersion();
    for (const e of entries)
      if (e.versionId && current && e.versionId !== current.id)
        out.push({ kind: "stale", text: `"${e.label}" was made on ${e.versionLabel}; the workbook is now ${current.label.split(" — ")[0]}. Its cells may have changed since: remove it and add it again from the current version.` });

    // An insert from Prep and one from the assistant on the same sheet: Prep's title cells are placed
    // for its own inserts only, so they may land one row or column off.
    const inserts = new Map<string, Set<number>>();
    entries.forEach((e, i) => {
      for (const op of e.operations)
        if (INSERT_OPS.has(String(op.op)) && op.sheet) {
          const list = inserts.get(String(op.sheet)) ?? new Set<number>();
          list.add(i);
          inserts.set(String(op.sheet), list);
        }
    });
    for (const [sheet, list] of inserts) {
      const prep = [...list].find((i) => entries[i].source === "prep");
      const assistant = [...list].find((i) => entries[i].source === "assistant");
      if (prep !== undefined && assistant !== undefined) {
        const [a, b] = prep < assistant ? [prep, assistant] : [assistant, prep];
        out.push({ kind: "overlap", text: `"${entries[a].label}" and "${entries[b].label}" both insert rows or columns on ${sheet}: the cells planned around one insert may land one row or column off. Apply one of them, then add the other again.` });
      }
    }

    type Item = { entry: number; op: Operation; t: Target; post: boolean };
    const cells = new Map<string, Item>();
    const big: Item[] = [];
    const pairs = new Map<string, { a: number; b: number; first: string; n: number; pivot: number | null }>();
    const clash = (a: Item, b: Item) => {
      if (a.entry === b.entry || sameChange(a.op, b.op)) return;
      const [x, y] = a.entry < b.entry ? [a, b] : [b, a];
      const k = `${x.entry}|${y.entry}`;
      const p = pairs.get(k);
      if (p) p.n += 1;
      else {
        // a pivot table is converted before any cell is written, whatever the order added
        const xp = PIVOT_OPS.has(String(x.op.op)), yp = PIVOT_OPS.has(String(y.op.op));
        pairs.set(k, { a: x.entry, b: y.entry, first: where(y.t, y.op), n: 1, pivot: xp !== yp ? (xp ? x.entry : y.entry) : null });
      }
    };
    // A cell planned for after the inserts (a title in an inserted row) is written in other
    // coordinates than a cell planned on the sheet as it is: each is compared only with its own kind.
    const overlaps = (s: Item, t: Item) => s.post === t.post && s.t.sheet === t.t.sheet && s.t.r1 <= t.t.r2 && t.t.r1 <= s.t.r2 && s.t.c1 <= t.t.c2 && t.t.c1 <= s.t.c2;
    entries.forEach((e, i) => {
      for (const op of e.operations) {
        const t = opTarget(op);
        if (!t) continue;
        const item: Item = { entry: i, op, t, post: !!op["after_inserts"] };
        for (const b of big) if (overlaps(b, item)) clash(b, item);
        if ((t.r2 - t.r1 + 1) * (t.c2 - t.c1 + 1) >= BIG) {
          for (const v of cells.values()) if (overlaps(item, v)) clash(v, item);
          big.push(item);
          continue;
        }
        for (let r = t.r1; r <= t.r2; r++)
          for (let c = t.c1; c <= t.c2; c++) {
            const k = `${t.sheet}|${item.post ? "post" : "pre"}|${r}|${c}`;
            const prev = cells.get(k);
            if (prev) clash(prev, item);
            cells.set(k, item);
          }
      }
    });
    for (const p of pairs.values()) {
      const what = p.n === 1 ? p.first : `${p.n} cells (first ${p.first})`;
      const cellEntry = p.pivot === p.a ? p.b : p.a;
      out.push({
        kind: "overlap",
        text:
          p.pivot !== null
            ? `"${entries[p.a].label}" and "${entries[p.b].label}" both change ${what}: the pivot table is converted first, so "${entries[cellEntry].label}" is written over it.`
            : `"${entries[p.a].label}" and "${entries[p.b].label}" both change ${what}: the later one wins, which may undo or break the first.`,
      });
    }
    return out;
  });

  /** `version`: the version the operations were planned on, when it may be older than the current one
   *  (an assistant proposal received before an Apply); the current version otherwise. */
  add(source: StackEntry["source"], label: string, operations: Operation[], version?: Version | null): void {
    const id = this.store.sessionId();
    const v = version === undefined ? this.store.currentVersion() : version;
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

  /** Removes the entries that were applied, by key: one added while the Apply ran stays. */
  removeKeys(keys: Set<string>): void {
    this.all.update((list) => list.filter((e) => !keys.has(e.key)));
  }

  clear(): void {
    const id = this.store.sessionId();
    this.all.update((list) => list.filter((e) => e.sessionId !== id));
  }
}
