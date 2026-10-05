import { useEffect, useState } from "react";
import type { ScanStatus } from "../types";

/**
 * 1.7.2: what an Apply is doing, from the backend's own status (never a
 * timer): the step it is on, how many changes are written, the time spent.
 * The last step is the re-analysis of the new file, read from the scan status.
 */
const STEPS: { id: string; label: string }[] = [
  { id: "apply_copy", label: "Copy" },
  { id: "apply_write", label: "Write changes" },
  { id: "apply_save", label: "Save" },
  { id: "apply_verify", label: "Verify it opens" },
  { id: "apply_log", label: "Change log" },
  { id: "reanalyze", label: "Re-analyze" },
];

function useElapsed(since: number | null): number {
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    if (since === null) return;
    const t = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(t);
  }, [since]);
  return since === null ? 0 : Math.max(0, Math.round((now - since) / 1000));
}

function clock(seconds: number): string {
  return seconds < 60 ? `${seconds} s` : `${Math.floor(seconds / 60)} min ${String(seconds % 60).padStart(2, "0")} s`;
}

interface Props {
  status: ScanStatus | null;
  /** Operations sent, known on the client before the backend's first answer. */
  total: number;
  /** When the Apply was sent (client clock, ms). */
  startedAt: number | null;
  compact?: boolean;
}

export default function ApplyProgress({ status, total, startedAt, compact = false }: Props) {
  const elapsed = useElapsed(startedAt);
  const apply = status?.apply ?? null;
  // Only an Apply sent after `startedAt` is ours: the status of the previous one stays on the server.
  const ours = apply != null && startedAt != null && apply.started_at != null && apply.started_at * 1000 >= startedAt - 2000;
  const a = ours ? apply : null;
  const reanalyzing = a?.state === "done" && a.reanalyze && status?.state === "running";
  const finishing = a?.state === "done" && !reanalyzing;
  const current = reanalyzing ? "reanalyze" : a?.state === "running" ? a.stage ?? "apply_copy" : finishing ? (a?.reanalyze ? "reanalyze" : "done") : "apply_copy";
  const steps = STEPS.filter((s) => s.id !== "reanalyze" || a == null || a.reanalyze);
  const index = Math.max(0, steps.findIndex((s) => s.id === current));
  const count = a?.total ?? total;
  const done = a ? Math.min(a.done, count) : 0;
  // 1.7.4: every step after "Write changes" has all of them in the file
  const writeIndex = steps.findIndex((s) => s.id === "apply_write");
  const implemented = index > writeIndex || (finishing && a != null) ? count : a?.stage === "apply_write" ? done : 0;

  let title = "Sending the changes";
  let detail = `${count} change${count !== 1 ? "s" : ""} to apply`;
  let fraction: number | null = null;
  if (reanalyzing && status) {
    title = `Re-analyzing the new version — ${status.title ?? "starting"}`;
    detail = status.message ?? "";
    fraction = Math.max(0, Math.min(1, status.overall ?? 0));
  } else if (a?.state === "running") {
    title = a.title ?? "Applying";
    detail = a.message ?? "";
    // the backend's fraction also counts the cells done inside a long operation (colours of thousands of cells)
    fraction = a.stage === "apply_write" && count > 0 ? Math.max(0, Math.min(1, a.fraction ?? done / count)) : null;
  } else if (a?.state === "error") {
    title = "Apply failed";
    detail = a.error ?? "";
  } else if (finishing && a) {
    title = a.reanalyze ? "Re-analyzing the new version" : "Done";
    detail = a.message ?? "";
    fraction = a.reanalyze ? null : 1;
  }
  const percent = fraction == null ? null : Math.round(fraction * 100);

  return (
    <div className={`w-full ${compact ? "" : "max-w-4xl"}`} role="status" aria-live="polite" aria-label="Applying changes">
      <div className="flex items-center gap-2 flex-wrap">
        <div className="w-3.5 h-3.5 border-2 border-[#1F3A5F] border-t-transparent rounded-full animate-spin flex-shrink-0" />
        <span className="text-[12px] font-semibold text-[#111827]">
          Step {index + 1} of {steps.length} · {title}
        </span>
        {a?.stage === "apply_write" && a.state === "running" && (
          <span className="text-[12px] font-mono text-[#1F3A5F] bg-[#EEF2FF] px-1.5 py-0.5 rounded">
            {done.toLocaleString()} / {count.toLocaleString()} written
          </span>
        )}
        <span className="ml-auto text-[11px] font-mono text-[#6B7280]">{clock(elapsed)}</span>
      </div>
      <div className="mt-0.5 text-[11px] text-[#1F3A5F]">
        Change <span className="font-mono font-semibold">{implemented.toLocaleString()}</span> of <span className="font-mono font-semibold">{count.toLocaleString()}</span> implemented
      </div>
      <div className="mt-1.5 h-2 w-full bg-[#E5E7EB] rounded-full overflow-hidden">
        {percent == null ? (
          <div className="h-full w-1/3 bg-[#1F3A5F]/60 rounded-full animate-pulse" />
        ) : (
          <div className="h-full bg-[#1F3A5F] rounded-full transition-all duration-300" style={{ width: `${percent}%` }} />
        )}
      </div>
      <div className="mt-1 flex items-center gap-3 text-[11px]">
        <span className="text-[#374151] truncate flex-1 min-w-0" title={detail}>{detail}</span>
        {percent != null && <span className="font-mono text-[#6B7280]">{percent}%</span>}
      </div>
      <div className={`mt-1.5 flex flex-wrap gap-x-3 gap-y-1 ${compact ? "text-[10px]" : "text-[11px]"}`}>
        {steps.map((s, i) => {
          const isDone = i < index;
          const isCurrent = i === index;
          return (
            <span key={s.id} className={`flex items-center gap-1 ${isDone ? "text-[#0F6E3A]" : isCurrent ? "text-[#1F3A5F] font-semibold" : "text-[#9CA3AF]"}`}>
              <span className="font-mono">{isDone ? "✓" : i + 1}</span>
              {s.label}
            </span>
          );
        })}
      </div>
    </div>
  );
}
