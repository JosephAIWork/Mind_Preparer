import { CommonModule } from "@angular/common";
import { Component, inject, input, output, signal } from "@angular/core";
import { ApiService, errorText } from "../core/api.service";
import { PendingUpload, SavedWorkbook, SessionStart } from "../core/models";

/** 1.14.0:  the main  screen's "Recent workbooks" -- the workbooks the app kept in its data folder
 *  (app/session_store.py), under the upload box, only when there is at least one. Reopen loads one
 *  with everything it had (every  version, the check result, the plans and results) and the page  shows
 *  the same screens as after its  last step; Remove asks once, in place,  then deletes it from this
 *  computer. Nothing opens by itself: the list  waits. */
@Component({
  selector: "mr-recent-workbooks",
  standalone: true,
  imports: [CommonModule],
  template: `
    @if (rows().length) {
      <section class="recent" data-show-me="recent-workbooks">
        @if (collapsible()) {
          <!-- Basic screen: the list waits closed until it is clicked open -->
          <button type="button" class="recent__toggle" data-show-me="recent-workbooks-toggle" (click)="toggleOpen()" [attr.aria-expanded]="open()">
            {{ open() ? "Hide" : "Show" }} recent workbooks ({{ rows().length }})
          </button>
        } @else {
          <h2 class="recent__title">Recent workbooks</h2>
        }
        @if (!collapsible() || open()) {
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
        <div class="recent__all">
          @if (askingAll()) {
            <span class="recent__ask" role="group" aria-label="Remove all saved workbooks?">
              <span class="recent__all-question">Remove all {{ rows().length }} saved workbooks?</span>
              <button type="button" class="cds--btn cds--btn--danger cds--btn--sm recent__all-yes" (click)="confirmRemoveAll()" [disabled]="busyAll() || busyId() !== null">{{ busyAll() ? "Removing…" : "Yes" }}</button>
              <button type="button" class="cds--btn cds--btn--ghost cds--btn--sm recent__all-no" (click)="keepAll()" [disabled]="busyAll()">No</button>
            </span>
          } @else {
            <button type="button" class="cds--btn cds--btn--ghost cds--btn--sm recent__all-remove" (click)="askRemoveAll()" [disabled]="busyId() !== null">Remove all</button>
          }
        </div>
        }
        @if (note()) {
          <p class="recent__note">{{ note() }}</p>
        }
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
      .recent__toggle {
        background: none;
        border: 0;
        padding: 0;
        font: inherit;
        font-weight: 600;
        cursor: pointer;
        color: var(--cds-link-primary, #0f62fe);
      }
      .recent__all {
        display: flex;
        justify-content: flex-end;
        padding-top: 0.5rem;
        border-top: 1px solid var(--cds-border-subtle-01, #e0e0e0);
      }
      .recent__note {
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
        margin: 0.5rem 0 0;
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
  /** the  row whose  Remove asked "Remove <name> from this computer?" */
  readonly askingId = signal<string | null>(null);
  /** the row being opened or removed: every button waits meanwhile */
  readonly busyId = signal<string | null>(null);
  readonly error = signal("");
  /** Basic screen: the list is closed until clicked open (Custom shows it open, as before). */
  readonly collapsible = input(false);
  readonly open = signal(false);
  /** "Remove all N saved workbooks?" asked; the removal under way; what was kept, in words */
  readonly askingAll = signal(false);
  readonly busyAll = signal(false);
  readonly note = signal("");
  /** The saved workbook, loaded again: the page shows it (a pending one goes to the sheet  choice). */
  readonly opened = output<SessionStart | PendingUpload>();

  constructor() {
    void this.refresh();
  }

  /** The list as the server has it now (an older server  without the route: no list). */
  async refresh(): Promise<void> {
    try {
      this.rows.set(await this.api.savedWorkbooks());
    } catch {
      this.rows.set([]);
    }
  }

  /** "10 Nov  2026, 14:05 · 3 versions", "· not  checked yet" for an upload never scanned. */
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
      void this.refresh(); // gone meanwhile: the list says  so
    } finally {
      this.busyId.set(null);
    }
  }

  toggleOpen(): void {
    this.open.update((o) => !o);
  }

  askRemoveAll(): void {
    if (this.busyId() || this.busyAll()) return;
    this.error.set("");
    this.note.set("");
    this.askingId.set(null);
    this.askingAll.set(true);
  }

  keepAll(): void {
    this.askingAll.set(false);
  }

  /** Removes every saved workbook the server lets go of (not the ones another running copy of the app holds). */
  async confirmRemoveAll(): Promise<void> {
    if (this.busyAll() || this.busyId()) return;
    this.busyAll.set(true);
    try {
      const out = await this.api.removeAllSavedWorkbooks();
      this.note.set(
        out.kept
          ? `Removed ${out.removed}. ${out.kept} kept: ${out.kept === 1 ? "it is" : "they are"} open in another copy of the app or still running.`
          : "",
      );
    } catch (e) {
      this.error.set(`Could not remove the saved workbooks: ${errorText(e)}`);
    } finally {
      await this.refresh();
      this.busyAll.set(false);
      this.askingAll.set(false);
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

/** "10  Nov 2026,  14:05" from the server's seconds since 1970 ("" when unknown). */
export function whenText(seconds: number | null | undefined): string {
  if (!seconds) return "";
  const d = new Date(seconds * 1000);
  const day = d.getDate();
  const month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getMonth()];
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${day} ${month} ${d.getFullYear()}, ${hh}:${mm}`;
}
