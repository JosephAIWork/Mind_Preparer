import { CommonModule } from "@angular/common";
import { Component, computed, inject, signal } from "@angular/core";
import { ButtonModule, InlineLoadingModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { SessionStore } from "../../core/session.store";

/** Deliver: the report workbooks to download, and the numbers check (did the values change?). */
@Component({
  selector: "mr-deliver-panel",
  standalone: true,
  imports: [CommonModule, ButtonModule, InlineLoadingModule],
  templateUrl: "./deliver-panel.html",
  styleUrl: "./panels.scss",
})
export class DeliverPanel {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly building = signal(false);
  readonly checking = signal(false);
  readonly error = signal("");

  readonly reports = computed(() => this.store.reports());
  readonly numbers = computed(() => this.store.numbers());
  readonly numbersStale = computed(() => {
    const n = this.numbers();
    const v = this.store.currentVersion();
    return !!n?.version_id && !!v && n.version_id !== v.id;
  });
  readonly countRows = computed(() => Object.entries(this.numbers()?.counts ?? {}));
  readonly sampleRows = computed(() => {
    const s = this.numbers()?.samples ?? {};
    return Object.entries(s).flatMap(([kind, rows]) => rows.slice(0, 5).map((r) => ({ kind, ...r })));
  });

  url(name: string): string | null {
    const id = this.store.sessionId();
    return id ? this.api.downloadUrl(id, name) : null;
  }

  async build(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.building()) return;
    this.building.set(true);
    this.error.set("");
    try {
      this.store.reports.set(await this.api.generateReports(id));
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.building.set(false);
    }
  }

  async check(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.checking()) return;
    this.checking.set(true);
    this.error.set("");
    try {
      this.store.numbers.set(await this.api.numbersCheck(id));
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.checking.set(false);
    }
  }
}
