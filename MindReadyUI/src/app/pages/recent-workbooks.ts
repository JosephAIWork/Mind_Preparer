import { CommonModule } from "@angular/common";
import { Component, inject, output, signal } from "@angular/core";
import { ApiService, errorText } from "../core/api.service";
import { PendingUpload, SavedWorkbook, SessionStart } from "../core/models";

/** 1.14.0: the main screen's "Recent workbooks" -- the workbooks the app kept in its data folder
 *  (app/session_store.py), under the upload box, only when there is at least one. Reopen loads one
 *  with everything it had (every version, the check result, the plans and results) and the page shows
 *  the same screens as after its last step; Remove asks once, in place, then deletes it from this
 *  computer. Nothing opens by itself: the list waits. */
@Component({
  selector: "mr-recent-workbooks",
  standalone: true,
  imports: [CommonModule],
  template: `
    @if (rows().length) {
      <section class="recent" data-show-me="recent-workbooks">
        <h2 class="recent__title">Recent workbooks</h2>
        <p class="recent__lead">Kept on this computer, with every version and result. Reopen one to go on where you left it.</p>
        <ul class="recent__list">
          @for (r of rows(); track r.id) {
            <li class="recent__row" [attr.data-id]="r.id">
              <div class="recent__main">
                <bdi class="recent__name">{{ r.name || "(no name)" }}</bdi>
                <span class="recent__meta">{{ metaText(r) }}</span>
              </div>
              @if (askingId() === r.id) {
                <span class="recent__ask" role="group" aria-label="Remove this workbook?">
                  <span>Remove <bdi>{{ r.name }}</bdi> from this computer?</span>
                  <button type="button" class="cds--btn cds--btn--danger cds--btn--sm recent__yes" (click)="confirmRemove(r)" [disabled]="busyId() !== null">{{ busyId() === r.id ? "Removing…" : "Yes" }}</button>
                  <button type="button" class="cds--btn cds--btn--ghost cds--btn--sm recent__no" (click)="keep()" [disabled]="busyId() !== null">No</button>
                </span>
              } @else {
                <span class="recent__buttons">
                  <button type="button" class="cds--btn cds--btn--tertiary cds--btn--sm recent__open" (click)="reopen(r)" [disabled]="busyId() !== null">{{ busyId() === r.id ? "Opening…" : "Reopen" }}</button>
                  <button type="button" class="cds--btn cds--btn--ghost cds--btn--sm recent__remove" (click)="askRemove(r)" [disabled]="busyId() !== null">Remove</button>
                </span>
              }
            </li>
          }
        </ul>
        @if (error()) {
          <p class="recent__error">{{ error() }}</p>
        }
      </section>
    }
  `,
  styles: [
    `
      .recent {
        margin: 1.5rem 0 1rem;
        padding: 1rem;
        background: var(--cds-layer-01, #f4f4f4);
        border-top: 1px solid var(--cds-border-subtle-01, #e0e0e0);
      }
      .recent__title {
        font-size: 1.125rem;
        font-weight: 600;
        margin: 0 0 0.25rem;
      }
      .recent__lead {
        color: var(--cds-text-secondary, #525252);
        margin: 0 0 0.75rem;
        font-size: 0.875rem;
      }
      .recent__list {
        list-style: none;
        margin: 0;
        padding: 0;
      }
      .recent__row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        padding: 0.5rem 0;
        border-top: 1px solid var(--cds-border-subtle-01, #e0e0e0);
      }
      .recent__main {
        display: flex;
        flex-direction: column;
        min-width: 0;
      }
      .recent__name {
        font-weight: 600;
        overflow-wrap: anywhere;
      }
      .recent__meta,
      .recent__ask {
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
      }
      .recent__buttons,
      .recent__ask {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        flex-shrink: 0;
      }
      .recent__error {
        color: var(--cds-text-error, #da1e28);
        font-size: 0.875rem;
        margin: 0.5rem 0 0;
      }
    `,
  ],
})
export class RecentWorkbooks {
  private api = inject(ApiService);
  readonly rows = signal<SavedWorkbook[]>([]);
  /** the row whose Remove asked "Remove <name> from this computer?" */
  readonly askingId = signal<string | null>(null);
  /** the row being opened or removed: every button waits meanwhile */
  readonly busyId = signal<string | null>(null);
  readonly error = signal("");
  /** The saved workbook, loaded again: the page shows it (a pending one goes to the sheet choice). */
  readonly opened = output<SessionStart | PendingUpload>();

  constructor() {
    void this.refresh();
  }

  /** The list as the server has it now (an older server without the route: no list). */
  async refresh(): Promise<void> {
    try {
      this.rows.set(await this.api.savedWorkbooks());
    } catch {
      this.rows.set([]);
    }
  }

  /** "10 Nov 2026, 14:05 · 3 versions", "· not checked yet" for an upload never scanned. */
  metaText(r: SavedWorkbook): string {
    const parts = [whenText(r.last_opened ?? r.uploaded_at), `${r.versions} version${r.versions === 1 ? "" : "s"}`];
    if (!r.has_result) parts.push("not checked yet");
    return parts.filter(Boolean).join(" · ");
  }

  async reopen(r: SavedWorkbook): Promise<void> {
    if (this.busyId()) return;
    this.error.set("");
    this.busyId.set(r.id);
    try {
      const started = await this.api.openSavedWorkbook(r.id);
      this.opened.emit(started);
    } catch (e) {
      this.error.set(`Could not open ${r.name}: ${errorText(e)}`);
      void this.refresh(); // gone meanwhile: the list says so
    } finally {
      this.busyId.set(null);
    }
  }

  askRemove(r: SavedWorkbook): void {
    if (this.busyId()) return;
    this.error.set("");
    this.askingId.set(r.id);
  }

  keep(): void {
    this.askingId.set(null);
  }

  async confirmRemove(r: SavedWorkbook): Promise<void> {
    if (this.busyId()) return;
    this.busyId.set(r.id);
    try {
      await this.api.removeSavedWorkbook(r.id);
      this.rows.set(this.rows().filter((x) => x.id !== r.id));
    } catch (e) {
      this.error.set(`Could not remove ${r.name}: ${errorText(e)}`);
      void this.refresh();
    } finally {
      this.busyId.set(null);
      this.askingId.set(null);
    }
  }
}

/** "10 Nov 2026, 14:05" from the server's seconds since 1970 ("" when unknown). */
export function whenText(seconds: number | null | undefined): string {
  if (!seconds) return "";
  const d = new Date(seconds * 1000);
  const day = d.getDate();
  const month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getMonth()];
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${day} ${month} ${d.getFullYear()}, ${hh}:${mm}`;
}
