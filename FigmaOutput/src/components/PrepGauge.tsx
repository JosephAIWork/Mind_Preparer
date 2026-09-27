import type { PrepProgress, Readiness } from "../types";

/**
 * 1.7.2: the Prep gauge. Three numbers that must go down -- blocking problems,
 * blocking operations planned, optional operations planned -- and the history
 * of planned operations per analysed version, so a stalled run is visible
 * instead of feeling like an endless "Apply" button.
 */
function Bars({ entries }: { entries: PrepProgress["entries"] }) {
  const max = Math.max(1, ...entries.map((e) => e.total_ops));
  return (
    <div className="flex items-end gap-1 h-10" aria-label="Planned operations per version">
      {entries.map((e) => {
        const h = Math.max(2, Math.round((e.total_ops / max) * 40));
        const blockingShare = e.total_ops ? e.blocking_ops / e.total_ops : 0;
        return (
          <div key={e.version_id} className="flex flex-col items-center gap-0.5" title={`${e.version_id}: ${e.total_ops} operation(s) planned (${e.blocking_ops} blocking, ${e.optional_ops} optional); ${e.blocking_findings} blocking problem(s)`}>
            <div className="w-4 rounded-sm overflow-hidden flex flex-col justify-end bg-[#E5E7EB]" style={{ height: h }}>
              <div className="w-full bg-[#9F1D1D]" style={{ height: `${Math.round(blockingShare * 100)}%` }} />
              <div className="w-full bg-[#1F3A5F] flex-1" />
            </div>
            <span className="text-[9px] font-mono text-[#9CA3AF]">{e.version_id.replace("ver-0", "v").replace("ver-", "v")}</span>
          </div>
        );
      })}
    </div>
  );
}

export default function PrepGauge({ readiness, progress }: { readiness: Readiness | null; progress: PrepProgress | null }) {
  const entries = progress?.entries ?? [];
  const last = entries[entries.length - 1];
  const prev = entries[entries.length - 2];
  const trend = last && prev ? last.total_ops - prev.total_ops : null;
  return (
    <div className="border border-[#E5E7EB] rounded-xl bg-white p-4 mb-3">
      <div className="flex items-start gap-6 flex-wrap">
        <div>
          <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">BLOCKING PROBLEMS</div>
          <div className={`text-2xl font-semibold ${readiness && readiness.blocking_count > 0 ? "text-[#9F1D1D]" : "text-[#0F6E3A]"}`}>{readiness?.blocking_count ?? "—"}</div>
          <div className="text-[11px] text-[#6B7280]">must reach 0{readiness && readiness.manual_blocking_count > 0 ? ` · ${readiness.manual_blocking_count} without automatic repair` : ""}</div>
        </div>
        <div>
          <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">BLOCKING REPAIRS PLANNED</div>
          <div className="text-2xl font-semibold text-[#111827]">{readiness?.prep.blocking_ops ?? "—"}</div>
          <div className="text-[11px] text-[#6B7280]">operations Prep can apply now</div>
        </div>
        <div>
          <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">OPTIONAL REPAIRS PLANNED</div>
          <div className="text-2xl font-semibold text-[#4B5563]">{readiness?.prep.optional_ops ?? "—"}</div>
          <div className="text-[11px] text-[#6B7280]">Mind reads the file without them</div>
        </div>
        {entries.length > 0 && (
          <div className="ml-auto">
            <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF] mb-1">
              PLANNED PER VERSION{trend != null && <span className={`ml-2 font-mono ${trend < 0 ? "text-[#0F6E3A]" : trend > 0 ? "text-[#9F1D1D]" : "text-[#8A5A00]"}`}>{trend < 0 ? "▼" : trend > 0 ? "▲" : "="} {Math.abs(trend)}</span>}
            </div>
            <Bars entries={entries.slice(-12)} />
          </div>
        )}
      </div>
      {progress?.stalled && (
        <div className="mt-3 text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded-lg px-3 py-2" role="alert">
          <strong>Not converging.</strong> {progress.message}
          {progress.repeating_actions.length > 0 && <span> Repeating: <span className="font-mono">{progress.repeating_actions.join(", ")}</span>.</span>}
        </div>
      )}
    </div>
  );
}
