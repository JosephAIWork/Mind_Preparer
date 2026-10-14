import { CommonModule } from "@angular/common";
import { Component, OnInit, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { CellWindow, MapCell, ModelMap } from "../../core/models";
import { SessionStore } from "../../core/session.store";

/** Where the model's inputs, loop driver, run settings and outputs are, read from the VBA, names and formulas. */
@Component({
  selector: "mr-map-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule],
  templateUrl: "./map-panel.html",
  styleUrl: "./panels.scss",
})
export class MapPanel implements OnInit {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly map = signal<ModelMap | null>(null);
  readonly loading = signal(false);
  readonly error = signal("");
  readonly peek = signal<{ key: string; window: CellWindow | null; error: string } | null>(null);
  readonly showAll = signal<Set<string>>(new Set());

  ngOnInit(): void {
    void this.load(false);
  }

  async load(refresh: boolean): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.loading.set(true);
    this.error.set("");
    try {
      const { map } = await this.api.modelMap(id, refresh);
      this.map.set(map);
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.loading.set(false);
    }
  }

  where(c: MapCell | { cell?: string | null; sheet?: string }): string {
    const cell = (c as MapCell).cell ?? "";
    const sheet = (c as { sheet?: string }).sheet;
    return sheet && cell && !cell.includes("!") ? `${sheet}!${cell}` : cell;
  }

  text(v: unknown): string {
    if (v === null || v === undefined) return "";
    return typeof v === "object" ? JSON.stringify(v) : String(v);
  }

  async show(ref: string): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !ref) return;
    if (this.peek()?.key === ref) {
      this.peek.set(null);
      return;
    }
    const bang = ref.lastIndexOf("!");
    const sheet = bang > 0 ? ref.slice(0, bang).replace(/^'|'$/g, "") : "";
    const cell = bang > 0 ? ref.slice(bang + 1) : ref;
    this.peek.set({ key: ref, window: null, error: "" });
    try {
      const w = await this.api.cellWindow(id, sheet, cell.split(":")[0].replace(/\$/g, ""), 3, 3);
      this.peek.set({ key: ref, window: w, error: w.error ?? "" });
    } catch (e) {
      this.peek.set({ key: ref, window: null, error: errorText(e) });
    }
  }

  more(section: string): boolean {
    return this.showAll().has(section);
  }

  toggleMore(section: string): void {
    const next = new Set(this.showAll());
    if (next.has(section)) next.delete(section);
    else next.add(section);
    this.showAll.set(next);
  }

  limit<T>(list: T[], section: string, n = 12): T[] {
    return this.more(section) ? list : list.slice(0, n);
  }
}
