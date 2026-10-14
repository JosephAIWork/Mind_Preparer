import { ApiService } from "./api.service";

export interface ApplyTick {
  fraction: number | null;
  text: string;
}

/** Follows a running Apply through GET .../status (field `apply`) until `stop()` is called. */
export function followApply(api: ApiService, sessionId: string, onTick: (t: ApplyTick) => void, everyMs = 600): () => void {
  let live = true;
  let timer: ReturnType<typeof setTimeout> | null = null;
  const tick = async () => {
    if (!live) return;
    try {
      const s = await api.scanStatus(sessionId);
      const a = s.apply;
      if (live && a && a.state === "running") onTick({ fraction: a.fraction ?? null, text: [a.title, a.message].filter(Boolean).join(" · ") });
    } catch {
      /* keep the last tick */
    }
    if (live) timer = setTimeout(tick, everyMs);
  };
  void tick();
  return () => {
    live = false;
    if (timer) clearTimeout(timer);
  };
}
