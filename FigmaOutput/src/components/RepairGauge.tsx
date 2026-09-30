import { useStore } from "../store";
import type { PrepProgressEntry } from "../types";

/**
 * 1.7.2: the repair gauge, shared by the Fix panel and the Findings screen.
 * One column per analysed version: blocking problems (must reach 0 for Mind)
 * and findings still open (everything not PASS), so a run of fixes reads as
 * a curve going down instead of a list that never seems to end.
 */
function short(v: string): string {
  return v.replace("ver-0", "v").replace("ver-", "v");
}

function Trend({ now, before, goodWhenDown = true }: { now: number; before: number | null; goodWhenDown?: boolean }) {
  if (before == null || now === before) return null;
  const d = now - before;
  const good = goodWhenDown ? d < 0 : d > 0;
  return <span className={`ml-1 font-mono text-[11px] ${good ? "text-[#0F6E3A]" : "text-[#9F1D1D]"}`}>{d < 0 ? "▼" : "▲"}{Math.abs(d)}</span>;
}

export default function RepairGauge({ compact = false }: { compact?: boolean }) {
  const readiness = useStore((s) => s.readiness);
  const progress = useStore((s) => s.prepProgress);
  const entries: PrepProgressEntry[] = progress?.entries ?? [];
  if (entries.length === 0 && !readiness) return null;
  const last = entries[entries.length - 1];
  const prev = entries.length > 1 ? entries[entries.length - 2] : null;
  const blocking = readiness?.blocking_count ?? last?.blocking_findings ?? 0;
  const open = last?.open_findings ?? null;
  const maxOpen = Math.max(1, ...entries.map((e) => e.open_findings ?? 0));
  const shown = entries.slice(-10);
  return (
    <div className={`flex items-center gap-4 flex-wrap ${compact ? "text-[11px]" : "text-[12px]"}`} aria-label="Repair progress across versions">
      <div>
        <span className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">BLOCKING</span>{" "}
        <span className={`font-semibold ${blocking > 0 ? "text-[#9F1D1D]" : "text-[#0F6E3A]"}`}>{blocking}</span>
        <Trend now={blocking} before={prev ? prev.blocking_findings : null} />
        <span className="text-[#9CA3AF]"> must reach 0</span>
      </div>
      {open != null && (
        <div>
          <span className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">OPEN FINDINGS</span>{" "}
          <span className="font-semibold text-[#111827]">{open}</span>
          <Trend now={open} before={prev?.open_findings ?? null} />
        </div>
      )}
      {/* 1.7.3: the changes every Apply of this session really wrote into the file */}
      {progress?.written && (progress.applies?.length ?? 0) > 0 && (
        <div title={(progress.applies ?? []).map((a) => `${short(a.version_id)}: +${a.applied}${a.failed ? `, ${a.failed} refused` : ""}`).join(" · ")}>
          <span className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">CHANGES WRITTEN</span>{" "}
          <span className="font-semibold text-[#0F766E]">{progress.written.applied.toLocaleString()}</span>
          <span className="ml-1 font-mono text-[11px] text-[#0F766E]">+{(progress.applies ?? [])[(progress.applies ?? []).length - 1].applied.toLocaleString()} last</span>
          {progress.written.failed > 0 && <span className="ml-1 font-mono text-[11px] text-[#9F1D1D]">{progress.written.failed.toLocaleString()} refused</span>}
        </div>
      )}
      {shown.length > 1 && (
        <div className="flex items-end gap-1" title="Open findings per version (red part: blocking)">
          {shown.map((e) => {
            const o = e.open_findings ?? 0;
            const h = Math.max(2, Math.round((o / maxOpen) * (compact ? 18 : 28)));
            const share = o ? Math.min(1, e.blocking_findings / o) : 0;
            return (
              <div key={e.version_id} className="flex flex-col items-center gap-0.5" title={`${e.version_id}: ${o} open finding(s), ${e.blocking_findings} blocking`}>
                <div className="w-3 rounded-sm overflow-hidden flex flex-col justify-end bg-[#E5E7EB]" style={{ height: h }}>
                  <div className="w-full bg-[#9F1D1D]" style={{ height: `${Math.round(share * 100)}%` }} />
                  <div className="w-full bg-[#1F3A5F] flex-1" />
                </div>
                <span className="text-[9px] font-mono text-[#9CA3AF]">{short(e.version_id)}</span>
              </div>
            );
          })}
        </div>
      )}
      {readiness && (
        <span className={`ml-auto text-[11px] font-semibold rounded px-2 py-0.5 ${readiness.state === "ready" ? "text-[#0F6E3A] bg-[#E6F4EA]" : readiness.state === "unverified" ? "text-[#8A5A00] bg-[#FFF1CC]" : "text-[#9F1D1D] bg-[#FDE2E2]"}`}>
          {readiness.state === "ready" ? "Ready for Mind" : readiness.state === "unverified" ? "Recalculate to verify" : `Blocked · ${readiness.next_step.label}`}
        </span>
      )}
    </div>
  );
}
