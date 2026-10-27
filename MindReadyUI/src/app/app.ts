import { Component, OnInit, inject, signal } from "@angular/core";
import { ActivatedRoute, Router, RouterLink, RouterLinkActive, RouterOutlet } from "@angular/router";
import { UIShellModule } from "carbon-components-angular";
import { ApiService } from "./core/api.service";
import { FeedbackDialog } from "./core/feedback-dialog";
import { FeedbackService } from "./core/feedback.service";
import { SessionStore } from "./core/session.store";

/** The shell: Mind's dark Carbon header, the two places (main screen, Advanced), the version, and
 *  "Send feedback to the Tel Aviv team" (in the header, and in a box when a step failed). */
@Component({
  selector: "mr-root",
  standalone: true,
  imports: [RouterOutlet, RouterLink, RouterLinkActive, UIShellModule, FeedbackDialog],
  templateUrl: "./app.html",
  styleUrl: "./app.scss",
})
export class App implements OnInit {
  private api = inject(ApiService);
  private route = inject(ActivatedRoute);
  private router = inject(Router);
  readonly store = inject(SessionStore);
  readonly feedback = inject(FeedbackService);
  readonly version = signal("");
  readonly rules = signal<number | null>(null);

  ngOnInit(): void {
    this.api
      .health()
      .then((h) => {
        this.version.set(h.version);
        this.rules.set(h.rules ?? null);
        document.title = `Mind Ready Beta ${h.version}`;
      })
      .catch(() => {});
    // `?session=<id>` reopens a session the backend still holds (after a reload or a shared link).
    const wanted = new URLSearchParams(window.location.search).get("session");
    if (wanted && !this.store.sessionId()) {
      this.api
        .resume(wanted)
        .then((s) => this.store.setSession(s))
        .catch(() => void this.router.navigate([], { queryParams: {}, replaceUrl: true }));
    }
  }

  sessionQuery(): Record<string, string> {
    const id = this.store.sessionId();
    return id ? { session: id } : {};
  }
}
