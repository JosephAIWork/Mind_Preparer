import type { ActionLevel } from "../types";

/**
 * 1.7.2: one vocabulary everywhere. "Blocking" = Mind refuses the file (or
 * computes it wrong) without this; "Optional" = Mind reads the file as it is,
 * the change only improves it. The level comes from the rule's priority.
 */
export default function LevelBadge({ level, size = "sm", title }: { level: ActionLevel | "pass" | "verify"; size?: "sm" | "md"; title?: string }) {
  const pad = size === "md" ? "px-2 py-0.5 text-[11px]" : "px-1.5 py-0.5 text-[10px]";
  if (level === "blocking") {
    return (
      <span className={`inline-flex items-center gap-1 rounded font-semibold tracking-wide bg-[#FDE2E2] text-[#9F1D1D] border border-[#9F1D1D]/20 ${pad}`} title={title ?? "Mind refuses the workbook (or computes it wrong) without this"}>
        ● BLOCKING
      </span>
    );
  }
  if (level === "optional") {
    return (
      <span className={`inline-flex items-center gap-1 rounded font-semibold tracking-wide bg-[#F3F4F6] text-[#4B5563] border border-[#E5E7EB] ${pad}`} title={title ?? "Mind reads the workbook as it is; this only improves it"}>
        ○ OPTIONAL
      </span>
    );
  }
  if (level === "verify") {
    return (
      <span className={`inline-flex items-center gap-1 rounded font-semibold tracking-wide bg-[#FFF1CC] text-[#8A5A00] border border-[#8A5A00]/20 ${pad}`} title={title ?? "Cleared by a real Excel recalculation"}>
        ▶ VERIFY
      </span>
    );
  }
  return null;
}
