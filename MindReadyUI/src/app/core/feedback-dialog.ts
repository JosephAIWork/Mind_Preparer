import { Component, OnInit, inject, signal } from "@angular/core";
import { ButtonModule } from "carbon-components-angular";
import { ApiService, errorText } from "./api.service";
import { FeedbackService } from "./feedback.service";
import { FeedbackMessage } from "./models";
import { SessionStore } from "./session.store";

/** The dialog behind "Send feedback to the Tel Aviv team": an optional note, the message prepared and
 *  shown (codes only, nothing from the workbook), then one click that copies it and opens an email.
 *  A file is saved only when asked (Save as a file). */
@Component({
  selector: "mr-feedback-dialog",
  standalone: true,
  imports: [ButtonModule],
  templateUrl: "./feedback-dialog.html",
  styleUrl: "./feedback-dialog.scss",
})
export class FeedbackDialog implements OnInit {
  private api = inject(ApiService);
  private store = inject(SessionStore);
  readonly feedback = inject(FeedbackService);

  readonly note = signal("");
  readonly assistantOn = signal(false);
  readonly useAssistant = signal(true);
  readonly busy = signal(false);
  readonly error = signal("");
  readonly message = signal<FeedbackMessage | null>(null);
  readonly copied = signal(true);
  readonly sent = signal(false);

  ngOnInit(): void {
    this.api
      .assistantKey()
      .then((k) => this.assistantOn.set(k.saved))
      .catch(() => {});
  }

  value(event: Event): string {
    return (event.target as HTMLTextAreaElement).value;
  }

  /** Step 1: the message from the backend (the assistant may take a few seconds), shown to be read. */
  async prepare(): Promise<void> {
    if (this.busy()) return;
    this.busy.set(true);
    this.error.set("");
    try {
      this.message.set(
        await this.api.feedback({
          ref: this.feedback.dialog()?.ref ?? null,
          note: this.note(),
          useAssistant: this.assistantOn() && this.useAssistant(),
          sessionId: this.store.sessionId(),
        }),
      );
    } catch (e) {
      this.error.set(errorText(e));
    } finally {
      this.busy.set(false);
    }
  }

  /** Step 2, on its own click: the browser lets a page copy and open an email only right after a click. */
  async send(): Promise<void> {
    const m = this.message();
    if (!m) return;
    this.copied.set(await copyText(m.text)); // before the email takes the focus away
    openEmail(m.email, m.subject);
    this.sent.set(true);
    const ref = this.feedback.dialog()?.ref;
    if (ref && this.feedback.incident()?.ref === ref) this.feedback.dismiss();
  }

  /** Only when asked: the message as a .txt in the Downloads folder (to attach). */
  saveAsFile(): void {
    const m = this.message();
    if (m) saveFile(m.file_name, m.text);
  }

  back(): void {
    this.message.set(null);
  }

  async copyAgain(): Promise<void> {
    const m = this.message();
    if (m) this.copied.set(await copyText(m.text));
  }

  close(): void {
    this.message.set(null);
    this.sent.set(false);
    this.note.set("");
    this.error.set("");
    this.feedback.close();
  }
}

async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // older ways, for a browser that refuses the clipboard API
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch {
      ok = false;
    }
    area.remove();
    return ok;
  }
}

function saveFile(name: string, text: string): void {
  const url = URL.createObjectURL(new Blob([text], { type: "text/plain;charset=utf-8" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function openEmail(to: string, subject: string): void {
  const body = "Paste the Mind Ready message here (Ctrl+V), then send.";
  const a = document.createElement("a");
  a.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}
