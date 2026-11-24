import { Injectable, computed, effect, inject, signal, untracked } from "@angular/core";
import { ApiService } from "./api.service";
import { WorkJob } from "./models";
import { SessionStore } from "./session.store";

/** What  a Cancel  on screen stops: one kind of work, a list of kinds, or (null) everything for the workbook. */
export type ServedKinds = string | readonly string[] | null;

/** What  runs on  the server for the open workbook, and whether a Cancel button is on screen for it
 *  (17 Nov 2026). The shell asks both: when a job runs and no Cancel for it  is drawn (the page that
 *  started it was left), it shows a  strip under the header -- "Still  working: Writing the changes."
 *  with its own Cancel  -- so there is always a way to stop the work.
 *
 *  `jobs`  is read from GET .../jobs every 3  s while a workbook is open. Only the shell starts that
 *  (`watch()`): a page or a test that merely draws a Cancel button starts no timer. */
@Injectable({ providedIn: "root" })
export class WorkService {
  private api = inject(ApiService);
  private store = inject(SessionStore);

  /** how  often the  list is read, in ms (the tests shorten it) */
  static everyMs = 3000;

  /** the jobs running for the open workbook, oldest first (empty: nothing runs) */
  readonly jobs = signal<WorkJob[]>([]);
  /** the Cancel buttons on screen (the strip's own is not among them):  each says which kind of work it stops
   *  (17 Nov 2026: or  which kinds: the Cancel beside Start  stops only Start's own steps) */
  private readonly buttons = signal<(() => ServedKinds)[]>([]);
  /** how many Cancel  buttons are on screen (the strip's own button is not counted) */
  readonly shown = computed(() => this.buttons().length);
  /** The  job the strip names; null: no strip.  A job is "left" when it runs and no Cancel on screen stops it:
   *  a Cancel for everything (kind null) serves every job, a Cancel for one kind of work serves that kind,
   *  a Cancel for a list of kinds serves each of them. */
  readonly left = computed<WorkJob | null>(() => {
    const served = this.buttons().map((kind) => kind());
    if (served.includes(null)) return null;
    const kinds = served.flatMap((k) => (k === null ? [] : typeof k === "string" ? [k] : [...k]));
    return this.jobs().find((j) => !kinds.includes(j.kind)) ?? null;
  });

  private readonly watching = signal(false);
  private timer: ReturnType<typeof setTimeout> | null = null;
  /** the workbook being followed; a look that answers for an older one is dropped */
  private following: string | null = null;
  /** bumped when the page says some work  has ended (`ended`): a look that was asked before that is dropped */
  private seq = 0;

  constructor() {
    effect(() => {
      // nothing is read from the store until the shell asked for the watch
      const id = this.watching() ? this.store.sessionId() : null;
      untracked(() => this.follow(id));
    });
  }

  /** The shell: start reading the list for whichever workbook is open (once; a second call does nothing). */
  watch(): void {
    this.watching.set(true);
  }

  /** Stop reading (the shell is gone); the list empties. */
  stop(): void {
    this.watching.set(false);
    this.follow(null);
  }

  /** A Cancel button came on screen; `kind` says what it stops now (null: everything for this workbook).
   *  The returned function: the button left the screen. */
  show(kind: () => ServedKinds = () => null): () => void {
    this.buttons.update((list) => [...list, kind]);
    return () => this.buttons.update((list) => list.filter((k) => k !== kind));
  }

  /** Read the list now (after a Cancel: the strip goes at  once, not at the next  look). */
  async refresh(): Promise<void> {
    const id = this.following;
    if (!id) return;
    const seq = this.seq;
    try {
      const list = await this.api.jobs(id);
      if (this.following === id && this.seq === seq) this.jobs.set(Array.isArray(list) ? list : []);
    } catch {
      /* an older server without the route, or  the workbook was  closed: keep  the last list */
    }
  }

  /** 17 Nov 2026 (live test L2b-3): the page's own work of these kinds has ended (Start's last step
   *  answered): they leave the list now, and the list is read again.  Without this the last list read
   *  (up to 3 s old) still  named the step, and the strip  "Still working: Recalculating." flashed beside
   *  "Everything ran and  every check passed" as the Cancel beside Start left the screen.  A look already on
   *  its way is dropped:  it was asked while the step still ran. */
  ended(kinds: readonly string[]): void {
    this.seq++;
    const kept = this.jobs().filter((j) => !kinds.includes(j.kind));
    if (kept.length !== this.jobs().length) this.jobs.set(kept);
    void this.refresh();
  }

  private follow(id: string | null): void {
    if (id === this.following) return;
    this.following = id;
    if (this.timer) clearTimeout(this.timer);
    this.timer = null;
    this.jobs.set([]);
    if (id) void this.tick(id);
  }

  private async tick(id: string): Promise<void> {
    if (this.following !== id) return;
    await this.refresh();
    if (this.following !== id) return;
    this.timer = setTimeout(() => void this.tick(id), WorkService.everyMs);
  }
}
