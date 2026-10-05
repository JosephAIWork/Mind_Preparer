import type { Version } from "../types";

interface Props {
  version: Version;
}

const sourceLabel: Record<string, string> = {
  upload: "original",
  prep: "written by Excel, verified",
  assistant: "assistant change, verified",
  formula: "formula fix, verified",
  convert: "converted",
  recalculate: "recalculated",
};

export default function VersionChip({ version }: Props) {
  const vNum = version.label.match(/v(\d+(?:\.\d+)?)/)?.[1] ?? "?";
  const desc = sourceLabel[version.source] ?? version.source;
  const isVerified = version.verified_opens_in_excel === true;
  return (
    <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md border border-[#E5E7EB] bg-white font-mono text-xs text-[#374151]">
      <span className="font-semibold">v{vNum}</span>
      <span className="text-[#9CA3AF]">—</span>
      <span>{desc}</span>
      {isVerified && (
        <span className="ml-1 inline-flex items-center gap-0.5 text-[#0F766E] font-sans text-[10px] font-semibold">
          <svg width="10" height="10" viewBox="0 0 10 10" fill="currentColor">
            <path d="M4 7.5 L1.5 5 L2.5 4 L4 5.5 L7.5 2 L8.5 3 Z" />
          </svg>
          verified
        </span>
      )}
    </span>
  );
}
