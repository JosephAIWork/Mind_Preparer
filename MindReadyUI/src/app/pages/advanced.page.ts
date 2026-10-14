import { CommonModule } from "@angular/common";
import { Component, computed, inject, signal } from "@angular/core";
import { RouterLink } from "@angular/router";
import { ButtonModule, ContentSwitcherModule, TagModule } from "carbon-components-angular";
import { ApiService, errorText } from "../core/api.service";
import { Finding, Status } from "../core/models";
import { SessionStore } from "../core/session.store";
import { LoopPanel } from "./advanced/loop-panel";
import { PrepPanel } from "./advanced/prep-panel";
import { RecalcPanel } from "./advanced/recalc-panel";

const STATUS_TAG: Record<Status, { type: "red" | "magenta" | "purple" | "blue" | "cyan" | "teal" | "green" | "gray" | "cool-gray" | "warm-gray" | "high-contrast" | "outline"; label: string }> = {
  ERROR: { type: "red", label: "Error" },
  REQUIRES_USER_INPUT: { type: "magenta", label: "Needs input" },
  WARNING: { type: "warm-gray", label: "Warning" },
  NOT_SUPPORTED: { type: "gray", label: "Not supported" },
  PASS: { type: "green", label: "Pass" },
};

/** The lists and the manual work: findings, the Prep actions by hand, versions. Grows in later chunks. */
@Component({
  selector: "mr-advanced",
  standalone: true,
  imports: [CommonModule, RouterLink, ButtonModule, ContentSwitcherModule, TagModule, PrepPanel, RecalcPanel, LoopPanel],
  templateUrl: "./advanced.page.html",
  styleUrl: "./advanced.page.scss",
})
export class AdvancedPage {
  private api = inject(ApiService);
  readonly store = inject(SessionStore);
  readonly onlyOpen = signal(true);
  readonly tab = signal<"findings" | "prep" | "recalc" | "loop" | "versions">("findings");
  readonly restoreLine = signal("");

  readonly findings = computed<Finding[]>(() => {
    const all = this.store.findings();
    const list = this.onlyOpen() ? all.filter((f) => f.status !== "PASS") : all;
    const rank: Record<string, number> = { ERROR: 0, REQUIRES_USER_INPUT: 1, NOT_SUPPORTED: 2, WARNING: 3, PASS: 4 };
    return [...list].sort((a, b) => (rank[a.status] ?? 9) - (rank[b.status] ?? 9) || a.rule_id.localeCompare(b.rule_id));
  });
  readonly manualActions = computed(() => this.store.plan().filter((a) => a.level === "blocking" || a.level === "optional"));
  readonly versionsNewestFirst = computed(() => [...this.store.versions()].reverse());

  tag(status: Status) {
    return STATUS_TAG[status] ?? STATUS_TAG.WARNING;
  }

  where(f: Finding): string {
    const loc = (f.location ?? {}) as { sheet?: string; cell?: string; ref?: string };
    if (!loc.sheet && !loc.cell && !loc.ref) return "";
    return [loc.sheet, loc.cell ?? loc.ref].filter(Boolean).join("!");
  }

  downloadUrl(fileName: string): string | null {
    const id = this.store.sessionId();
    return id ? this.api.downloadUrl(id, fileName) : null;
  }

  async restore(versionId: string): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.restoreLine.set(`Restoring ${versionId}…`);
    try {
      const out = await this.api.reanalyze(id, versionId);
      this.store.applyAnalysis(out);
      const v = this.store.versions().find((x) => x.id === versionId);
      if (v) this.store.currentVersion.set(v);
      this.restoreLine.set(`${versionId} is the current version.`);
    } catch (e) {
      this.restoreLine.set(errorText(e));
    }
  }
}
