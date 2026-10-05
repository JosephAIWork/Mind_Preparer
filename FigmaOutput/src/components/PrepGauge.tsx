import type { ApplyRecord, PrepProgress, PrepProgressEntry, Readiness } from "../types";

/**
 * 1.7.2: the Prep gauge. Two questions, answered with the session's own
 * figures: where am I on the way to Mind (four steps, one is "you are here"),
 * and how much is done. The history underneath shows every analysed version,
 * so a run that stops going down is visible.
 * 1.7.3: the applied-changes gauge -- the changes every Apply really wrote
 * into the file (counted by the backend from what Excel saved), against what
 * is still planned.
 */
type StepState = "done" | "current" | "todo";

function steps(r: Readiness): { label: string; hint: string; state: StepState }[] {
  const viaPrep = r.blocking.filter((b) => b.fix === "prep").length;
  const byHand = r.blocking.length - viaPrep;
  const done = [viaPrep === 0, r.blocking.length === 0, r.state === "ready", r.state === "ready"];
  const first = done.findIndex((d) => !d);
  const state = (i: number): StepState => (done[i] ? "done" : i === first ? "current" : "todo");
  return [
    { label: "Apply the blocking repairs", hint: viaPrep ? `${viaPrep} problem${viaPrep !== 1 ? "s" : ""} Prep can repair` : "nothing left for Prep", state: state(0) },
    { label: "Fix the rest by hand", hint: byHand ? `${byHand} problem${byHand !== 1 ? "s" : ""} without automatic repair` : "nothing left by hand", state: state(1) },
    { label: "Recalculate in Excel", hint: r.recalc.ran ? r.recalc.message ?? "" : "not run for this version", state: state(2) },
    { label: "Ready for Mind", hint: r.state === "ready" ? "download or send" : "after a clean recalculation", state: state(3) },
  ];
}

function Bar({ label, left, most, note }: { label: string; left: number; most: number; note: string }) {
  const done = Math.max(0, most - left);
  const pct = most > 0 ? Math.round((done / most) * 100) : 100;
  const finished = left === 0;
  return (
    <div className="flex items-center gap-3">
      <div className="w-36 flex-shrink-0 text-[12px] font-medium text-[#374151]">{label}</div>
      <div
        className="flex-1 h-3 bg-[#E5E7EB] rounded-full overflow-hidden"
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={most}
        aria-valuenow={done}
        aria-label={`${label}: ${done} of ${most} done`}
      >
        <div className={`h-full rounded-full transition-all duration-500 ${finished ? "bg-[#0F6E3A]" : "bg-[#1F3A5F]"}`} style={{ width: `${pct}%` }} />
      </div>
      <div className="w-60 flex-shrink-0 text-[12px] text-[#374151]">
        {most === 0 ? (
          <span className="text-[#6B7280]">none found</span>
        ) : (
          <>
            <span className="font-mono font-semibold">{done.toLocaleString()}</span> of <span className="font-mono">{most.toLocaleString()}</span> done
            {" · "}
            <span className={`font-mono font-semibold ${finished ? "text-[#0F6E3A]" : "text-[#9F1D1D]"}`}>{left.toLocaleString()}</span> left
          </>
        )}
        <div className="text-[10px] text-[#9CA3AF]">{note}</div>
      </div>
    </div>
  );
}

/** Changes written by the Applies of this session against the ones still planned. */
function WrittenBar({ label, written, refused, left, note }: { label: string; written: number; refused: number; left: number; note: string }) {
  const total = written + left;
  const pct = total > 0 ? Math.round((written / total) * 100) : 0;
  const finished = total > 0 && left === 0;
  return (
    <div className="flex items-center gap-3">
      <div className="w-36 flex-shrink-0 text-[12px] font-medium text-[#374151]">{label}</div>
      <div
        className="flex-1 h-3 bg-[#E5E7EB] rounded-full overflow-hidden"
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={total}
        aria-valuenow={written}
        aria-label={`${label}: ${written} written, ${left} still planned`}
      >
        <div className={`h-full rounded-full transition-all duration-500 ${finished ? "bg-[#0F6E3A]" : "bg-[#0F766E]"}`} style={{ width: `${pct}%` }} />
      </div>
      <div className="w-60 flex-shrink-0 text-[12px] text-[#374151]">
        {total === 0 && refused === 0 ? (
          <span className="text-[#6B7280]">none planned, none written</span>
        ) : (
          <>
            <span className="font-mono font-semibold text-[#0F766E]">{written.toLocaleString()}</span> written
            {" · "}
            <span className={`font-mono font-semibold ${left === 0 ? "text-[#0F6E3A]" : "text-[#374151]"}`}>{left.toLocaleString()}</span> still planned
            {refused > 0 && <span className="text-[#9F1D1D]"> · <span className="font-mono font-semibold">{refused.toLocaleString()}</span> refused</span>}
          </>
        )}
        <div className="text-[10px] text-[#9CA3AF]">{note}</div>
      </div>
    </div>
  );
}

function short(v: string): string {
  return v.replace(/^ver-0*(\d)/, "v$1");
}

export default function PrepGauge({ readiness, progress }: { readiness: Readiness | null; progress: PrepProgress | null }) {
  const entries: PrepProgressEntry[] = progress?.entries ?? [];
  if (!readiness) return null;
  const most = (pick: (e: PrepProgressEntry) => number, now: number) => Math.max(now, ...entries.map(pick));
  const blocking = readiness.blocking_count;
  const blockingOps = readiness.prep.blocking_ops;
  const optionalOps = readiness.prep.optional_ops;
  const path = steps(readiness);
  const here = path.findIndex((s) => s.state === "current");
  const shown = entries.slice(-6);
  const applies: ApplyRecord[] = progress?.applies ?? [];
  const written = progress?.written ?? null;

  return (
    <div className="border border-[#E5E7EB] rounded-xl bg-white p-4 mb-3" aria-label="Preparation progress">
      <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF] mb-2">
        WHERE YOU ARE{here >= 0 ? ` · STEP ${here + 1} OF ${path.length}` : " · ALL STEPS DONE"}
      </div>
      <ol className="grid grid-cols-4 gap-2">
        {path.map((s, i) => (
          <li
            key={s.label}
            aria-current={s.state === "current" ? "step" : undefined}
            className={`rounded-lg border px-3 py-2 ${
              s.state === "done" ? "border-[#0F6E3A]/30 bg-[#F0FDF4]" : s.state === "current" ? "border-[#1F3A5F] bg-[#EEF2FF]" : "border-[#E5E7EB] bg-[#F9FAFB]"
            }`}
          >
            <div className={`text-[12px] font-semibold ${s.state === "done" ? "text-[#0F6E3A]" : s.state === "current" ? "text-[#1F3A5F]" : "text-[#9CA3AF]"}`}>
              <span className="font-mono mr-1">{s.state === "done" ? "✓" : i + 1}</span>
              {s.label}
            </div>
            <div className={`text-[11px] mt-0.5 ${s.state === "todo" ? "text-[#9CA3AF]" : "text-[#6B7280]"}`}>
              {s.state === "current" ? <strong className="text-[#1F3A5F]">You are here · </strong> : null}
              {s.hint}
            </div>
          </li>
        ))}
      </ol>

      <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF] mt-4 mb-2">HOW MUCH IS DONE</div>
      <div className="flex flex-col gap-2">
        <Bar label="Blocking problems" left={blocking} most={most((e) => e.blocking_findings ?? 0, blocking)} note="must reach 0 left for Mind to take the file" />
        {written ? (
          <>
            <WrittenBar label="Blocking changes" written={written.blocking_applied} refused={written.blocking_failed} left={blockingOps} note="written into the file by Apply, for the blocking problems" />
            <WrittenBar label="Optional changes" written={written.optional_applied} refused={written.optional_failed} left={optionalOps} note="Mind reads the file without them" />
          </>
        ) : (
          <>
            <Bar label="Blocking repairs" left={blockingOps} most={most((e) => e.blocking_ops ?? 0, blockingOps)} note="changes Prep can apply for the blocking problems" />
            <Bar label="Optional repairs" left={optionalOps} most={most((e) => e.optional_ops ?? 0, optionalOps)} note="Mind reads the file without them" />
          </>
        )}
      </div>

      {written && (
        <div className="mt-3 flex items-center gap-1.5 flex-wrap text-[11px]" aria-label="Changes written by every Apply">
          <span className="text-[10px] font-semibold tracking-wider text-[#9CA3AF] mr-1">CHANGES APPLIED</span>
          {applies.length === 0 ? (
            <span className="text-[#6B7280]">no Apply yet in this session — nothing has been written</span>
          ) : (
            <>
              <span className="rounded px-1.5 py-0.5 border border-[#0F766E]/30 bg-[#F0FDFA] text-[#0F766E]">
                <span className="font-mono font-semibold">{written.applied.toLocaleString()}</span> written in {applies.length} {applies.length !== 1 ? "Applies" : "Apply"}
                {written.failed > 0 && <span className="text-[#9F1D1D]"> · {written.failed.toLocaleString()} refused</span>}
              </span>
              {applies.slice(-6).map((a) => (
                <span
                  key={a.version_id}
                  className={`rounded px-1.5 py-0.5 border ${a.applied === 0 ? "border-[#9F1D1D]/30 bg-[#FFF7F7]" : "border-[#E5E7EB] bg-[#F9FAFB]"}`}
                  title={`${short(a.previous_version_id ?? "")} → ${short(a.version_id)}: ${a.blocking_applied} blocking and ${a.optional_applied} optional change(s) written, ${a.failed} refused`}
                >
                  <span className="font-mono font-semibold text-[#111827]">{short(a.version_id)}</span>
                  <span className={a.applied === 0 ? "text-[#9F1D1D]" : "text-[#0F766E]"}> +{a.applied.toLocaleString()}</span>
                  {a.failed > 0 && <span className="text-[#9F1D1D]"> · {a.failed.toLocaleString()} refused</span>}
                  {a.resolved_blocking.length > 0 && <span className="text-[#0F6E3A]"> · resolved {a.resolved_blocking.join(", ")}</span>}
                </span>
              ))}
            </>
          )}
        </div>
      )}

      {shown.length > 1 && (
        <div className="mt-2 flex items-center gap-1.5 flex-wrap text-[11px]" aria-label="Figures of every analysed version">
          <span className="text-[10px] font-semibold tracking-wider text-[#9CA3AF] mr-1">VERSION BY VERSION</span>
          {shown.map((e, i) => (
            <span key={e.version_id} className="flex items-center gap-1.5">
              {i > 0 && <span className="text-[#9CA3AF]">→</span>}
              <span className={`rounded px-1.5 py-0.5 border ${i === shown.length - 1 ? "border-[#1F3A5F] bg-[#EEF2FF]" : "border-[#E5E7EB] bg-[#F9FAFB]"}`}>
                <span className="font-mono font-semibold text-[#111827]">{short(e.version_id)}</span>
                <span className="text-[#9F1D1D]"> {e.blocking_findings} blocking</span>
                <span className="text-[#6B7280]"> · {(e.total_ops ?? 0).toLocaleString()} repair{e.total_ops !== 1 ? "s" : ""} planned</span>
              </span>
            </span>
          ))}
        </div>
      )}

      {progress?.stalled && (
        <div className="mt-3 text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded-lg px-3 py-2" role="alert">
          <strong>Not converging.</strong> {progress.message}
          {progress.repeating_actions.length > 0 && <span> Repeating: <span className="font-mono">{progress.repeating_actions.join(", ")}</span>.</span>}
        </div>
      )}
    </div>
  );
}
