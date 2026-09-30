import { useState } from "react";
import { useStore } from "../store";
import type { ApplyActionOutcome, ApplyReport, ApplyRuleMove, ApplyVerdict, Status } from "../types";
import LevelBadge from "./LevelBadge";

/**
 * 1.7.3: what an Apply did, repair by repair, from the backend's own
 * comparison of the analysis before and the analysis after: the changes
 * written, the rule before and now, what is still planned or left by hand,
 * and what the new analysis plans that the previous one did not. A resolved
 * problem is said to be resolved; one that is still there says who acts next.
 */
const VERDICT: Record<ApplyVerdict, { label: string; chip: string; row: string }> = {
  resolved: { label: "✓ RESOLVED", chip: "bg-[#DFF5E6] text-[#0F6E3A] border-[#0F6E3A]/20", row: "border-[#0F6E3A]/25 bg-[#F0FDF4]" },
  partial: { label: "◐ PARTLY — STILL OPEN", chip: "bg-[#FFF1CC] text-[#8A5A00] border-[#8A5A00]/20", row: "border-[#8A5A00]/25 bg-[#FFFBEB]" },
  unchanged: { label: "↻ WRITTEN — NO EFFECT", chip: "bg-[#FDE2E2] text-[#9F1D1D] border-[#9F1D1D]/20", row: "border-[#9F1D1D]/25 bg-[#FFF7F7]" },
  failed: { label: "✕ NOT WRITTEN", chip: "bg-[#FDE2E2] text-[#9F1D1D] border-[#9F1D1D]/20", row: "border-[#9F1D1D]/25 bg-[#FFF7F7]" },
  written: { label: "WRITTEN — NOT RE-ANALYZED", chip: "bg-[#F3F4F6] text-[#4B5563] border-[#E5E7EB]", row: "border-[#E5E7EB] bg-white" },
};

const STATUS_TEXT: Record<Status, string> = {
  PASS: "text-[#0F6E3A]",
  WARNING: "text-[#8A5A00]",
  ERROR: "text-[#9F1D1D]",
  REQUIRES_USER_INPUT: "text-[#9F1D1D]",
  NOT_SUPPORTED: "text-[#6B7280]",
};

export function shortVersion(id: string | null | undefined): string {
  return (id ?? "").replace("ver-0", "v").replace("ver-", "v").replace(/^v0+(\d)/, "v$1");
}

function arrows(text: string): string {
  return text.split(" -> ").join(" → ").split(" -- ").join(" — ");
}

export function RuleMove({ move }: { move: ApplyRuleMove }) {
  const same = move.from === move.to;
  return (
    <span className="inline-flex items-center gap-1 text-[11px] font-mono bg-white border border-[#E5E7EB] rounded px-1.5 py-0.5 whitespace-nowrap">
      <span className="font-semibold text-[#1F3A5F]">{move.rule_id}</span>
      {move.to == null ? (
        <span className={move.from ? STATUS_TEXT[move.from] : "text-[#6B7280]"}>{move.from ?? "—"}</span>
      ) : same ? (
        <span className={STATUS_TEXT[move.to]}>still {move.to}</span>
      ) : (
        <>
          <span className={move.from ? STATUS_TEXT[move.from] : "text-[#6B7280]"}>{move.from ?? "—"}</span>
          <span className="text-[#9CA3AF]">→</span>
          <span className={`font-semibold ${STATUS_TEXT[move.to]}`}>{move.to}</span>
        </>
      )}
      {move.sites_from != null && move.sites_to != null && move.sites_from !== move.sites_to && (
        <span className="text-[#6B7280]">· {move.sites_from.toLocaleString()} → {move.sites_to.toLocaleString()} cells</span>
      )}
    </span>
  );
}

function ActionRow({ action, compact }: { action: ApplyActionOutcome; compact: boolean }) {
  const v = VERDICT[action.verdict];
  return (
    <li className={`border rounded-lg px-3 py-2 ${v.row}`}>
      <div className="flex items-center gap-2 flex-wrap">
        <span className={`text-[10px] font-semibold tracking-wide rounded border px-1.5 py-0.5 whitespace-nowrap ${v.chip}`}>{v.label}</span>
        {!compact && <LevelBadge level={action.level} />}
        <span className="text-[12px] font-semibold text-[#111827]">{action.title}</span>
        <span className="text-[11px] font-mono text-[#374151] bg-white border border-[#E5E7EB] rounded px-1.5 py-0.5 whitespace-nowrap">
          {action.applied.toLocaleString()} of {action.sent.toLocaleString()} written
        </span>
      </div>
      {action.rules.length > 0 && (
        <div className="flex items-center gap-1.5 flex-wrap mt-1.5">
          {action.rules.map((r) => <RuleMove key={r.rule_id} move={r} />)}
        </div>
      )}
      {/* the chips above already say what was written and how each rule moved: only the conclusion is spelled out */}
      <p className="text-[12px] text-[#374151] mt-1.5">{arrows(action.next ?? action.summary)}</p>
      {action.errors.length > 0 && action.verdict !== "failed" && (
        <div className="mt-1 flex flex-col gap-0.5">
          {action.errors.map((e, i) => (
            <div key={i} className="text-[11px] font-mono text-[#9F1D1D]">refused — {e}</div>
          ))}
        </div>
      )}
    </li>
  );
}

interface Props {
  report: ApplyReport;
  /** "verified in Excel" / download link of the file the Apply produced */
  verified?: boolean | null;
  downloadHref?: string;
  downloadName?: string;
  compact?: boolean;
}

export default function ApplyReportCard({ report, verified, downloadHref, downloadName, compact = false }: Props) {
  const openFix = useStore((s) => s.openFix);
  const [showAll, setShowAll] = useState(false);
  const counts = report.actions.reduce<Record<string, number>>((n, a) => ({ ...n, [a.verdict]: (n[a.verdict] ?? 0) + 1 }), {});
  const tone = report.applied === 0 ? "border-[#9F1D1D]/30" : report.failed > 0 || (counts.unchanged ?? 0) > 0 ? "border-[#8A5A00]/40" : "border-[#0F766E]/40";
  const most = compact ? 4 : 8;
  const shown = showAll ? report.actions : report.actions.slice(0, most);
  const byHand = report.remaining_blocking.filter((r) => r.fix !== "prep");
  const viaPrep = report.remaining_blocking.filter((r) => r.fix === "prep");

  return (
    <div className={`border-2 rounded-xl bg-white ${compact ? "p-3" : "p-4"} ${tone}`} aria-live="polite" aria-label="What the last Apply did">
      <div className="flex items-center gap-2 flex-wrap">
        <span className="text-[10px] font-semibold tracking-wider text-[#6B7280]">
          WHAT THE LAST APPLY DID · {shortVersion(report.previous_version_id)} → {shortVersion(report.version_id)}
        </span>
        {verified && <span className="text-[11px] text-[#0F766E]">· written by Excel, verified to open</span>}
        {verified === false && <span className="text-[11px] font-semibold text-[#9F1D1D]">· the new file FAILED to open in Excel</span>}
        {downloadHref && (
          <a href={downloadHref} className="ml-auto px-3 py-1 bg-[#1F3A5F] text-white rounded-md text-[12px] font-medium" title={downloadName}>
            Download {shortVersion(report.version_id)}
          </a>
        )}
      </div>

      <div className={`flex items-stretch gap-2 flex-wrap mt-2 ${compact ? "text-[11px]" : "text-[12px]"}`}>
        <div className="rounded-lg border border-[#E5E7EB] bg-[#F9FAFB] px-3 py-1.5">
          <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">CHANGES WRITTEN</div>
          <div className="font-mono font-semibold text-[#111827]">
            {report.applied.toLocaleString()} <span className="font-normal text-[#6B7280]">of {report.sent.toLocaleString()}</span>
            {report.failed > 0 && <span className="text-[#9F1D1D]"> · {report.failed.toLocaleString()} refused</span>}
          </div>
        </div>
        {report.blocking_after != null && (
          <div className={`rounded-lg border px-3 py-1.5 ${report.blocking_after === 0 ? "border-[#0F6E3A]/30 bg-[#F0FDF4]" : "border-[#9F1D1D]/25 bg-[#FFF7F7]"}`}>
            <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">BLOCKING PROBLEMS</div>
            <div className="font-mono font-semibold text-[#111827]">
              {report.blocking_before} <span className="text-[#9CA3AF]">→</span>{" "}
              <span className={report.blocking_after === 0 ? "text-[#0F6E3A]" : "text-[#9F1D1D]"}>{report.blocking_after}</span>
              {report.resolved_blocking.length > 0 && <span className="font-normal text-[#0F6E3A]"> · resolved: {report.resolved_blocking.join(", ")}</span>}
            </div>
          </div>
        )}
        <div className="rounded-lg border border-[#E5E7EB] bg-[#F9FAFB] px-3 py-1.5">
          <div className="text-[10px] font-semibold tracking-wider text-[#9CA3AF]">REPAIRS</div>
          <div className="text-[#374151]">
            {(["resolved", "partial", "unchanged", "failed", "written"] as ApplyVerdict[])
              .filter((k) => counts[k])
              .map((k) => `${counts[k]} ${k === "partial" ? "partly" : k === "unchanged" ? "without effect" : k === "failed" ? "not written" : k}`)
              .join(" · ") || "—"}
          </div>
        </div>
      </div>

      <ul className="flex flex-col gap-1.5 mt-2">
        {shown.map((a) => <ActionRow key={a.id} action={a} compact={compact} />)}
      </ul>
      {report.actions.length > most && (
        <button onClick={() => setShowAll(!showAll)} className="mt-1 text-[11px] text-[#1F3A5F] underline underline-offset-2">
          {showAll ? "Show fewer" : `Show all ${report.actions.length} repairs`}
        </button>
      )}

      {(report.appeared.length > 0 || report.regressed.length > 0) && (
        <div className="mt-2 text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded-lg px-3 py-2">
          <div className="text-[10px] font-semibold tracking-wider mb-0.5">NEW IN {shortVersion(report.version_id).toUpperCase()} — NOT THERE BEFORE THIS APPLY</div>
          {report.appeared.map((a) => (
            <div key={a.id}>
              <span className="font-semibold">{a.title}</span>: {a.planned_after.toLocaleString()} repair{a.planned_after !== 1 ? "s" : ""} planned
              {a.planned_before > 0 ? ` (was ${a.planned_before.toLocaleString()})` : ""} — found by the analysis of the file as it is now.
            </div>
          ))}
          {report.regressed.map((r) => (
            <div key={r.rule_id}>
              <span className="font-mono font-semibold">{r.rule_id}</span> went from {r.from} to {r.to}.
            </div>
          ))}
        </div>
      )}

      {report.remaining_blocking.length > 0 && (
        <div className="mt-2 text-[12px] text-[#9F1D1D] bg-[#FFF7F7] border border-[#9F1D1D]/20 rounded-lg px-3 py-2">
          <div className="text-[10px] font-semibold tracking-wider mb-1">STILL BLOCKING — WHO ACTS NEXT</div>
          {viaPrep.length > 0 && (
            <div className="text-[#374151]">
              <span className="font-mono font-semibold text-[#9F1D1D]">{viaPrep.map((r) => r.rule_id).join(", ")}</span> — Prep has a repair planned: tick it below and Apply.
            </div>
          )}
          {byHand.length > 0 && (
            <div className="flex items-center gap-1.5 flex-wrap text-[#374151]">
              <span>No automatic repair, fix by hand or with the assistant:</span>
              {byHand.slice(0, compact ? 4 : 10).map((r) => (
                <button
                  key={r.rule_id}
                  onClick={() => openFix({ rule_id: r.rule_id })}
                  className="px-2 py-0.5 rounded-md border border-[#9F1D1D]/40 bg-white text-[#9F1D1D] text-[11px] font-mono font-semibold hover:bg-[#FDE2E2]"
                  title={`${r.touched ? "This Apply changed cells for it, the rule still fails" : "Not touched by this Apply"}${r.sites != null ? ` · ${r.sites} cell(s)` : ""}`}
                >
                  {r.rule_id}{r.new ? " (new)" : ""} →
                </button>
              ))}
              {byHand.length > (compact ? 4 : 10) && <span className="text-[#6B7280]">+{byHand.length - (compact ? 4 : 10)} more</span>}
            </div>
          )}
        </div>
      )}
      {report.reanalyzed && report.blocking_after === 0 && report.blocking_before > 0 && (
        <div className="mt-2 text-[12px] text-[#0F6E3A] bg-[#DFF5E6] border border-[#0F6E3A]/20 rounded-lg px-3 py-2">
          ✓ No blocking problem is left in {shortVersion(report.version_id)}.
        </div>
      )}
    </div>
  );
}
