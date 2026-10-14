import { CommonModule } from "@angular/common";
import { Component, HostListener, OnInit, computed, effect, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { Grid, GridNamerOutcome, GridNamesResult, MindFlag, SheetCells, failedCount } from "../../core/models";
import { SessionStore } from "../../core/session.store";

interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

interface PendingArea {
  sheet: string;
  ref: string;
  gridRef: string;
  name: string;
  flags: string[];
  wasNamed: string | null;
}

type Resolved = { kind: "grid"; grid: Grid; hits: Grid[] } | { kind: "caption"; grid: Grid; table: Grid; hits: Grid[] } | { kind: "none" | "many"; hits: Grid[] };

/** 1 -> A, 27 -> AA */
function colName(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseCell(ref: string | null | undefined): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec((ref ?? "").trim());
  if (!m) return null;
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: parseInt(m[2], 10), c };
}

function rectRef(s: Rect): string {
  const a = `${colName(s.c1)}${s.r1}`;
  return s.r1 === s.r2 && s.c1 === s.c2 ? a : `${a}:${colName(s.c2)}${s.r2}`;
}

/** Mirror of the backend's resolution: the grids the selection touches (their rectangle or their
 * title cell), and whether two of them are the caption-above-a-table pair that is named as one. */
function resolveSelection(grids: Grid[], sel: Rect): Resolved {
  const hits = grids.filter((g) => {
    const t = parseCell(g.title_cell);
    const inRect = !(g.last_row < sel.r1 || g.first_row > sel.r2 || g.last_col < sel.c1 || g.first_col > sel.c2);
    const onTitle = !!t && sel.r1 <= t.r && t.r <= sel.r2 && sel.c1 <= t.c && t.c <= sel.c2;
    return inRect || onTitle;
  });
  const isCap = (g: Grid) => {
    const v = g.header_values[0];
    return !g.name && g.n_cols === 1 && typeof v === "string" && v.trim() !== "" && !v.startsWith("=");
  };
  if (hits.length === 1) {
    const g = hits[0];
    if (!g.name) {
      if (isCap(g)) {
        const t = grids.find((h) => h !== g && !h.name && h.first_row === g.first_row + 1 && h.first_col === g.first_col + 1);
        if (t) return { kind: "caption", grid: g, table: t, hits };
      }
      const cap = grids.find((h) => h !== g && isCap(h) && g.first_row === h.first_row + 1 && g.first_col === h.first_col + 1);
      if (cap) return { kind: "caption", grid: cap, table: g, hits };
    }
    return { kind: "grid", grid: g, hits };
  }
  if (hits.length === 2) {
    for (const [cap, table] of [
      [hits[0], hits[1]],
      [hits[1], hits[0]],
    ] as const) {
      const capText = cap.header_values[0];
      if (!cap.name && cap.n_cols === 1 && typeof capText === "string" && capText.trim() !== "" && !capText.startsWith("=") && table.first_row === cap.first_row + 1 && table.first_col === cap.first_col + 1) {
        return { kind: "caption", grid: cap, table, hits };
      }
    }
  }
  return { kind: hits.length === 0 ? "none" : "many", hits };
}

/** "/Resize.A /Group.(My group).0.1" or "Resize.A" -> ["Resize.A", "Group.(My group).0.1"] */
function parseExtraFlags(text: string): string[] {
  const t = text.trim();
  if (!t) return [];
  if (t.includes("/")) return t.split("/").map((x) => x.trim()).filter(Boolean);
  return [t];
}

/** Drag over an area of a sheet and give it a #Name and flags; Submit writes the titles as a new version. */
@Component({
  selector: "mr-grid-namer-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule],
  templateUrl: "./grid-namer-panel.html",
  styleUrl: "./grid-namer-panel.scss",
})
export class GridNamerPanel implements OnInit {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);

  readonly sheets = computed(() => this.store.summary()?.sheets ?? []);
  readonly active = signal("");
  readonly cells = signal<SheetCells | null>(null);
  readonly loading = signal(false);
  readonly big = signal(false);
  readonly error = signal("");
  readonly sel = signal<Rect | null>(null);
  readonly dragging = signal(false);
  private dragFrom: { r: number; c: number } | null = null;

  readonly flagList = signal<MindFlag[]>([]);
  readonly allFlags = signal(false);
  /** The flags people use most, first; the chosen ones always stay visible. */
  readonly shownFlags = computed<MindFlag[]>(() => {
    const list = this.flagList();
    if (this.allFlags() || list.length <= 8) return list;
    const common = new Set(["input", "result", "parameters", "group", "resize", "step", "noheader", "readonly"]);
    const chosen = this.chosenFlags();
    return list.filter((f) => common.has(f.name.toLowerCase()) || chosen.has(f.name));
  });
  readonly name = signal("");
  readonly chosenFlags = signal<Set<string>>(new Set());
  readonly extraFlags = signal("");
  readonly pending = signal<PendingArea[]>([]);
  readonly nameRest = signal(true);
  readonly submitting = signal(false);
  readonly outcome = signal<GridNamerOutcome | null>(null);
  readonly showSkipped = signal(false);
  readonly suggesting = signal(false);
  readonly suggestions = signal<GridNamesResult | null>(null);

  readonly grids = computed<Grid[]>(() => this.sheets().find((s) => s.name === this.active())?.grids ?? []);
  readonly target = computed<Resolved | null>(() => {
    const s = this.sel();
    return s ? resolveSelection(this.grids(), s) : null;
  });
  readonly targetGrid = computed<Grid | null>(() => {
    const t = this.target();
    return t && (t.kind === "grid" || t.kind === "caption") ? t.grid : null;
  });
  readonly pendingByGrid = computed(() => new Map(this.pending().map((p) => [`${p.sheet}!${p.gridRef}`, p])));
  readonly paint = computed(() => {
    const cellOf = new Map<string, number>();
    const titleOf = new Map<string, number>();
    this.grids().forEach((g, i) => {
      for (let r = g.first_row; r <= g.last_row; r++) for (let c = g.first_col; c <= g.last_col; c++) cellOf.set(`${r}:${c}`, i);
      const t = parseCell(g.title_cell);
      if (t) titleOf.set(`${t.r}:${t.c}`, i);
    });
    return { cellOf, titleOf };
  });
  readonly columns = computed(() => {
    const n = this.cells()?.n_cols ?? 0;
    return Array.from({ length: n }, (_, i) => colName(i + 1));
  });
  readonly selRef = computed(() => {
    const s = this.sel();
    return s ? rectRef(s) : "";
  });

  constructor() {
    // A new selection prefills the name and flags from the grid it lands on.
    effect(() => {
      const g = this.targetGrid();
      if (!g) return;
      this.name.set(g.name ?? "");
      const docFlags = new Set<string>();
      const extras: string[] = [];
      for (const f of g.flags ?? []) {
        const raw = f.raw.replace(/^\//, "");
        const known = this.flagList().find((d) => d.name.toLowerCase() === f.name.toLowerCase());
        if (known && f.args.length === 0) docFlags.add(known.name);
        else extras.push(raw);
      }
      this.chosenFlags.set(docFlags);
      this.extraFlags.set(extras.map((e) => "/" + e).join(" "));
    });
    // The sheet's cells follow the active sheet, the size and the summary.
    effect(() => {
      const id = this.store.sessionId();
      const sheet = this.active();
      const big = this.big();
      this.store.summary();
      if (!id || !sheet) return;
      void this.loadCells(id, sheet, big);
    });
  }

  ngOnInit(): void {
    this.api
      .mindFlags()
      .then((f) => this.flagList.set(f))
      .catch(() => {});
    const first = this.sheets()[0]?.name ?? "";
    if (!this.active() || !this.sheets().some((s) => s.name === this.active())) this.active.set(first);
  }

  private async loadCells(id: string, sheet: string, big: boolean): Promise<void> {
    this.loading.set(true);
    this.error.set("");
    this.sel.set(null);
    try {
      const c = await this.api.sheetCells(id, sheet, big ? 2000 : 400, big ? 200 : 60);
      if (this.active() === sheet) this.cells.set(c);
    } catch (e) {
      this.cells.set(null);
      this.error.set(errorText(e));
    } finally {
      this.loading.set(false);
    }
  }

  @HostListener("window:mouseup")
  onWindowUp(): void {
    this.dragging.set(false);
    this.dragFrom = null;
  }

  down(r: number, c: number, event: MouseEvent): void {
    if (event.button !== 0) return;
    event.preventDefault();
    this.dragFrom = { r, c };
    this.dragging.set(true);
    this.sel.set({ r1: r, c1: c, r2: r, c2: c });
  }

  enter(r: number, c: number): void {
    if (!this.dragging() || !this.dragFrom) return;
    const a = this.dragFrom;
    this.sel.set({ r1: Math.min(a.r, r), r2: Math.max(a.r, r), c1: Math.min(a.c, c), c2: Math.max(a.c, c) });
  }

  cellClass(r: number, c: number): string {
    const s = this.sel();
    const inSel = !!s && s.r1 <= r && r <= s.r2 && s.c1 <= c && c <= s.c2;
    const { cellOf, titleOf } = this.paint();
    const key = `${r}:${c}`;
    const t = titleOf.get(key);
    const i = cellOf.get(key);
    const g = t !== undefined ? this.grids()[t] : i !== undefined ? this.grids()[i] : null;
    const pending = g ? this.pendingByGrid().has(`${this.active()}!${g.ref}`) : false;
    return ["cell", t !== undefined ? "cell--title" : "", pending ? "cell--pending" : i !== undefined ? (g?.name ? "cell--named" : "cell--untitled") : "", inSel ? "cell--sel" : ""].filter(Boolean).join(" ");
  }

  text(v: unknown): string {
    if (v === null || v === undefined) return "";
    const s = typeof v === "object" ? JSON.stringify(v) : String(v);
    return s.length > 24 ? s.slice(0, 23) + "…" : s;
  }

  targetText(): string {
    const t = this.target();
    if (!t) return "";
    if (t.kind === "none") return "No grid under the selection: Mind only reads grids. Pick cells inside one.";
    if (t.kind === "many") return `${t.hits.length} grids under the selection: pick one at a time.`;
    if (t.kind === "caption") return `Caption "${this.text(t.grid.header_values[0])}" above the table ${t.table.ref}: named as one grid.`;
    if (t.kind === "grid") {
      const g = t.grid;
      return `${g.name ? "Grid #" + g.name : "Untitled grid"} ${g.ref}${g.title_cell ? " (title in " + g.title_cell + ")" : ""}`;
    }
    return "";
  }

  flagOn(name: string): boolean {
    return this.chosenFlags().has(name);
  }

  toggleFlag(name: string): void {
    const next = new Set(this.chosenFlags());
    if (next.has(name)) next.delete(name);
    else next.add(name);
    this.chosenFlags.set(next);
  }

  addArea(): void {
    const g = this.targetGrid();
    const t = this.target();
    const s = this.sel();
    const name = this.name().trim();
    if (!g || !t || !s || !name) return;
    const entry: PendingArea = { sheet: this.active(), ref: t.kind === "caption" ? rectRef(s) : g.ref, gridRef: g.ref, name, flags: [...this.chosenFlags(), ...parseExtraFlags(this.extraFlags())], wasNamed: g.name };
    this.pending.update((prev) => [...prev.filter((p) => !(p.sheet === entry.sheet && p.gridRef === entry.gridRef)), entry]);
    this.sel.set(null);
    this.name.set("");
    this.chosenFlags.set(new Set());
    this.extraFlags.set("");
  }

  remove(p: PendingArea): void {
    this.pending.update((list) => list.filter((x) => x !== p));
  }

  async submit(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.pending().length || this.submitting()) return;
    this.submitting.set(true);
    this.error.set("");
    this.outcome.set(null);
    try {
      const out = await this.api.applyNamedAreas(
        id,
        this.pending().map((p) => ({ sheet: p.sheet, ref: p.ref, name: p.name, flags: p.flags })),
        this.nameRest(),
      );
      this.store.applyAnalysis({ ...out, version: out.version });
      this.outcome.set(out);
      this.pending.set([]);
      this.sel.set(null);
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.submitting.set(false);
    }
  }

  async suggest(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.suggesting()) return;
    this.suggesting.set(true);
    this.error.set("");
    try {
      const out = await this.api.suggestGridNames(id, true);
      this.suggestions.set(out);
      if (out.plan) this.store.plan.set(out.plan);
      if (out.readiness) this.store.readiness.set(out.readiness);
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.suggesting.set(false);
    }
  }

  value(event: Event): string {
    return (event.target as HTMLInputElement).value;
  }

  failedCount = failedCount;
}
