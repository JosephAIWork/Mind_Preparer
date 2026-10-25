import { Component, Input, OnInit, inject, signal } from "@angular/core";
import { ButtonModule } from "carbon-components-angular";
import { ApiService, errorText } from "../../core/api.service";
import { AssistantKey } from "../../core/models";

/** The AI gateway key: a box to paste it in while none is saved; once saved, its last 4 characters and Replace.
 *  `onlyWhenMissing` (main screen): nothing at all while a key is saved. */
@Component({
  selector: "mr-key-box",
  standalone: true,
  imports: [ButtonModule],
  templateUrl: "./key-box.html",
  styleUrl: "./key-box.scss",
})
export class KeyBox implements OnInit {
  private api = inject(ApiService);
  @Input() onlyWhenMissing = false;

  readonly status = signal<AssistantKey | null>(null);
  readonly editing = signal(false);
  readonly draft = signal("");
  readonly saving = signal(false);
  readonly error = signal("");
  /** after a save: what to tell the user (shown even with onlyWhenMissing) */
  readonly done = signal("");

  ngOnInit(): void {
    this.api
      .assistantKey()
      .then((s) => this.status.set(s))
      .catch(() => {});
  }

  value(event: Event): string {
    return (event.target as HTMLInputElement).value;
  }

  replace(): void {
    this.done.set("");
    this.error.set("");
    this.editing.set(true);
  }

  cancel(): void {
    this.draft.set("");
    this.error.set("");
    this.editing.set(false);
  }

  async save(): Promise<void> {
    const key = this.draft().trim();
    if (!key || this.saving()) return;
    this.saving.set(true);
    this.error.set("");
    try {
      const s = await this.api.saveAssistantKey(key);
      this.status.set(s);
      this.draft.set("");
      this.editing.set(false);
      this.done.set(s.message ? `Key saved, not checked: ${s.message}` : "Key saved. The assistant is on.");
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.saving.set(false);
    }
  }
}
