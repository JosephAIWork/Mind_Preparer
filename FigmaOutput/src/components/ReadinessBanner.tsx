import { Link, useLocation } from "react-router-dom";
import { useStore } from "../store";
import type { BlockingItem, Readiness } from "../types";

/**
 * 1.7.2: the one question, answered on every screen -- is this workbook
 * acceptable by Mind? Three states (blocked / not verified / ready), what
 * stands in the way, and the single next step. Everything optional stays
 * out of the colour; it is counted in grey underneath.
 */
const STEPS: { key: string; label: string; to: string }[] = [
  { key: "upload", label: "Upload", to: "/" },
  { key: "findings", label: "Findings", to: "/findings" },
  { key: "fix", label: "Fix", to: "/prep" },
  { key: "verify", label: "Verify", to: "/recalculate" },
  { key: "deliver", label: "Deliver", to: "/reports" },
];

function stepStates(r: Readiness): Record<string, "done" | "current" | "todo"> {
  const blocked = r.state === "blocked";
  const verified = r.state === "ready";
  return {
    upload: "done",
    findings: "done",
    fix: blocked ? "current" : "done",
    verify: blocked ? "todo" : verified ? "done" : "current",
    deliver: verified ? "current" : "todo",
  };
}

function BlockingLine({ item }: { item: BlockingItem }) {
  const openFix = useStore((s) => s.openFix);
  const loc = [item.location?.sheet, item.location?.cell].filter(Boolean).join("!");
  return (
    <li className="flex items-start gap-2 text-[12px] text-[#7F1D1D]">
      <span className="font-mono font-semibold whitespace-nowrap">{item.rule_id}</span>
      <span className="flex-1 min-w-0 truncate" title={item.message}>
        {item.message}
        {item.sites != null && item.sites > 1 && <span className="text-[#9F1D1D]/70"> · {item.sites} cells</span>}
        {loc && <span className="font-mono text-[#9F1D1D]/70"> · {loc}</span>}
      </span>
      {item.fix === "prep" ? (
        <Link to="/prep" className="whitespace-nowrap px-2 py-0.5 rounded bg-[#9F1D1D] text-white font-semibold hover:bg-[#7F1D1D]">Repair in Prep</Link>
      ) : (
        <button
          onClick={() => openFix({ rule_id: item.rule_id, sheet: item.location?.sheet, cell: item.location?.cell })}
          className="whitespace-nowrap px-2 py-0.5 rounded border border-[#9F1D1D] text-[#9F1D1D] font-semibold hover:bg-[#9F1D1D]/10"
          title={item.fix === "prep_skipped" ? "Prep looked at it and left every cell for review: fix it with the assistant" : "No automatic repair: fix it with the assistant"}
        >
          Fix with assistant
        </button>
      )}
    </li>
  );
}

export default function ReadinessBanner() {
  const readiness = useStore((s) => s.readiness);
  const report = useStore((s) => s.report);
  const { pathname } = useLocation();
  if (!report || !readiness) return null;

  const r = readiness;
  const tone =
    r.state === "blocked"
      ? { bg: "bg-[#FDE2E2]", border: "border-[#9F1D1D]/30", text: "text-[#7F1D1D]", dot: "bg-[#DC2626]", word: "Not acceptable by Mind" }
      : r.state === "unverified"
      ? { bg: "bg-[#FFF1CC]", border: "border-[#8A5A00]/30", text: "text-[#7A3E00]", dot: "bg-[#F59E0B]", word: "Acceptable by Mind — not verified" }
      : { bg: "bg-[#DFF5E6]", border: "border-[#0F6E3A]/30", text: "text-[#0F6E3A]", dot: "bg-[#16A34A]", word: "Acceptable by Mind — verified" };
  const states = stepStates(r);
  const next = r.next_step;
  const nextTo = next.screen === "prep" ? "/prep" : next.screen === "findings" ? "/findings" : next.screen === "recalculate" ? "/recalculate" : "/reports";
  const onNext = pathname === nextTo;

  return (
    <div className={`border-b ${tone.border} ${tone.bg} px-4 py-2.5 flex-shrink-0`} role="status" aria-live="polite">
      <div className="flex items-center gap-3 flex-wrap">
        <span className={`w-2.5 h-2.5 rounded-full ${tone.dot} flex-shrink-0`} aria-hidden />
        <span className={`font-semibold text-[13px] ${tone.text}`}>{tone.word}</span>
        <span className={`text-[12px] ${tone.text} opacity-80 truncate`}>{r.headline}</span>
        <nav className="ml-auto flex items-center gap-1 text-[11px]" aria-label="Workflow">
          {STEPS.map((s, i) => {
            const st = states[s.key];
            return (
              <Link
                key={s.key}
                to={s.to}
                className={`flex items-center gap-1 px-2 py-0.5 rounded-full ${
                  st === "done" ? "text-[#0F6E3A]" : st === "current" ? "bg-white/70 text-[#111827] font-semibold border border-black/10" : "text-[#6B7280]"
                }`}
                title={st === "done" ? "done" : st === "current" ? "you are here" : "later"}
              >
                <span className="font-mono">{st === "done" ? "✓" : i + 1}</span>
                {s.label}
              </Link>
            );
          })}
        </nav>
        {!onNext && (
          <Link to={nextTo} className={`px-3 py-1 rounded-md text-[12px] font-semibold text-white ${r.state === "blocked" ? "bg-[#9F1D1D] hover:bg-[#7F1D1D]" : r.state === "unverified" ? "bg-[#8A5A00] hover:bg-[#6B4600]" : "bg-[#0F6E3A] hover:bg-[#0B5A2E]"}`}>
            {next.label} →
          </Link>
        )}
      </div>
      {r.state === "blocked" && r.blocking.length > 0 && (
        <ul className="mt-2 flex flex-col gap-1 max-w-4xl">
          {r.blocking.slice(0, 6).map((b) => (
            <BlockingLine key={b.rule_id} item={b} />
          ))}
          {r.blocking.length > 6 && <li className="text-[11px] text-[#7F1D1D]">… {r.blocking.length - 6} more in Findings</li>}
        </ul>
      )}
      <div className={`mt-1.5 text-[11px] ${tone.text} opacity-75 flex gap-3 flex-wrap`}>
        {r.state !== "blocked" && r.recalc.ran && !r.recalc.clean && (
          <span>
            {r.recalc.new_errors} new formula error(s) after recalculation, not in error in the original
            {r.recalc.preexisting_errors ? ` · ${r.recalc.preexisting_errors} already in the original (the model's own)` : ""}
            {r.recalc.addin_gap_errors ? ` · ${r.recalc.addin_gap_errors} add-in gap(s), not a defect` : ""}
          </span>
        )}
        {r.state === "ready" && r.recalc.preexisting_errors > 0 && <span>{r.recalc.preexisting_errors} error cell(s) were already errors in the original upload — the model's own, Mind takes them as they are</span>}
        {r.state === "ready" && r.recalc.addin_gap_errors > 0 && <span>{r.recalc.addin_gap_errors} #NAME? on MM_ functions come from the missing add-in on this machine, not from the workbook</span>}
        <span>{r.optional_count} optional improvement{r.optional_count !== 1 ? "s" : ""} available{r.prep.optional_ops ? ` (${r.prep.optional_ops} planned in Prep)` : ""} — Mind reads the file without them</span>
        {next.detail && !onNext && <span>· next: {next.detail}</span>}
      </div>
    </div>
  );
}
