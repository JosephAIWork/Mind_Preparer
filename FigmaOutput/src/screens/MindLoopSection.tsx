import { useEffect, useState } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { LoopPlan, LoopPlanOptions } from "../services/api";

/**
 * Mind loop (1.10.0): the macro's run loop written as Mind formulas. The plan comes from
 * the Model map: every driver cell gets MM_LOOP over its column of the table (or the list
 * the macro walks, or the count), and the block the macro writes on every pass gets
 * MM_RESULT row by row. Nothing is written until "Write it into a new version".
 */

const chip = (status: string) => {
  const map: Record<string, string> = {
    ready: "bg-[#DFF5E6] text-[#0F6E3A]",
    needs_decision: "bg-[#FEF3C7] text-[#92400E]",
    not_built: "bg-[#F3F4F6] text-[#6B7280]",
    nothing_to_do: "bg-[#F3F4F6] text-[#6B7280]",
  };
  const text: Record<string, string> = { ready: "ready", needs_decision: "needs a decision", not_built: "not built", nothing_to_do: "nothing to do" };
  return <span className={`text-[11px] px-1.5 py-0.5 rounded ${map[status] ?? map.not_built}`}>{text[status] ?? status}</span>;
};

export default function MindLoopSection({ sessionId }: { sessionId: string }) {
  const addVersion = useStore((s) => s.addVersion);
  const applyAnalysis = useStore((s) => s.applyAnalysis);
  const currentVersion = useStore((s) => s.currentVersion);
  const [plan, setPlan] = useState<LoopPlan | null>(null);
  const [opts, setOpts] = useState<LoopPlanOptions>({});
  const [busy, setBusy] = useState(false);
  const [writing, setWriting] = useState(false);
  const [err, setErr] = useState<string | null>(null);
  const [done, setDone] = useState<string | null>(null);
  const [showOps, setShowOps] = useState(false);

  const load = (o: LoopPlanOptions) => {
    setBusy(true);
    setErr(null);
    api.mindLoopPlan(sessionId, o).then((r) => setPlan(r.plan)).catch((e) => setErr(String(e?.message ?? e))).finally(() => setBusy(false));
  };
  useEffect(() => { setDone(null); load(opts); }, [sessionId, currentVersion?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  async function write() {
    if (!plan || plan.operation_count === 0) return;
    setWriting(true);
    setErr(null);
    setDone(null);
    try {
      const out = await api.applyMindLoopPlan(sessionId, opts);
      addVersion(out.version);
      applyAnalysis({ summary: out.summary, report: out.report, plan: out.plan, delta: out.delta, readiness: out.readiness, prep_progress: out.prep_progress });
      setDone(`${out.version.label} — ${out.result.applied.length} written${out.result.failed.length ? `, ${out.result.failed.length} failed` : ""}.`);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setWriting(false);
    }
  }

  const needsCount = plan?.loops.some((l) => l.status === "needs_decision" && /count cell/.test(l.why ?? ""));
  const field = "text-[12px] px-2 py-1 rounded border border-[#E5E7EB] bg-white w-[120px]";

  return (
    <section className="border border-[#E5E7EB] rounded-xl bg-white">
      <div className="px-4 py-3 border-b border-[#E5E7EB] flex items-baseline gap-3">
        <h2 className="font-semibold text-[14px] text-[#111827]">Mind loop</h2>
        <span className="text-[12px] text-[#6B7280]">the same run without the macro: MM_LOOP in the driver cells, MM_RESULT where each pass's results go</span>
      </div>
      <div className="px-4 py-3 space-y-3">
        <div className="flex items-end gap-3 flex-wrap text-[12px] text-[#374151]">
          <label className="flex flex-col gap-1">loop name<input className={field} value={opts.name ?? ""} placeholder={plan?.loops[0]?.name ?? ""} onChange={(e) => setOpts({ ...opts, name: e.target.value })} /></label>
          <label className="flex flex-col gap-1">first value<input className={field} value={opts.first ?? ""} placeholder={plan?.loops[0]?.first != null ? String(plan.loops[0].first) : ""} onChange={(e) => setOpts({ ...opts, first: e.target.value })} /></label>
          <label className="flex flex-col gap-1">last value<input className={field} value={opts.last ?? ""} placeholder={plan?.loops[0]?.last != null ? String(plan.loops[0].last) : ""} onChange={(e) => setOpts({ ...opts, last: e.target.value })} /></label>
          {needsCount && <label className="flex flex-col gap-1">count cell (Sheet!Cell)<input className={field} value={opts.count_cell ?? ""} onChange={(e) => setOpts({ ...opts, count_cell: e.target.value })} /></label>}
          <label className="flex items-center gap-1.5 pb-1.5"><input type="checkbox" checked={!!opts.clear_old} onChange={(e) => setOpts({ ...opts, clear_old: e.target.checked })} />also empty the areas the macro clears</label>
          <button onClick={() => load(opts)} disabled={busy} className="text-[12px] px-3 py-1.5 rounded border border-[#E5E7EB] bg-white hover:bg-[#F3F4F6] disabled:opacity-50">{busy ? "Planning…" : "Preview"}</button>
        </div>
        {err && <div className="text-[12px] text-[#9F1D1D] border border-[#9F1D1D]/20 bg-[#FDE2E2] rounded-lg px-3 py-2">{err}</div>}
        {plan && !plan.found && <div className="text-[12px] text-[#6B7280]">{plan.notes.join(" ")}</div>}
        {plan && plan.found && (
          <>
            {plan.loops.map((l) => (
              <div key={l.driver} className="border border-[#F3F4F6] rounded-lg px-3 py-2">
                <div className="text-[13px] text-[#111827] flex items-center gap-2 flex-wrap">
                  <span className="font-mono font-semibold">{l.name}</span>
                  {chip(l.status)}
                  <span className="text-[#374151]">driver <span className="font-mono">{l.driver}</span>{l.label ? ` (${l.label})` : ""}</span>
                  {l.first != null && l.last != null && <span className="text-[#6B7280]">{l.first} → {l.last} · {l.count} values</span>}
                  {l.cells.length > 0 && <span className="text-[#6B7280]">· {l.cells.length} cell{l.cells.length !== 1 ? "s" : ""}</span>}
                </div>
                {l.how && l.status === "ready" && <div className="text-[12px] text-[#6B7280] mt-0.5">{l.how}</div>}
                {l.why && <div className="text-[12px] text-[#92400E] mt-0.5">{l.why}</div>}
                {l.notes.map((n) => <div key={n} className="text-[12px] text-[#374151] mt-0.5">{n}</div>)}
                {l.cells.slice(0, 4).map((c) => <div key={c.cell} className="font-mono text-[11px] text-[#374151] mt-0.5"><span className="text-[#9CA3AF]">{c.cell}</span> {c.formula}</div>)}
                {l.cells.length > 4 && <div className="text-[11px] text-[#9CA3AF] mt-0.5">… {l.cells.length - 4} more cells, the same way</div>}
              </div>
            ))}
            {plan.outputs.map((o) => (
              <div key={o.target + (o.from ?? "")} className="border border-[#F3F4F6] rounded-lg px-3 py-2">
                <div className="text-[13px] text-[#111827] flex items-center gap-2 flex-wrap">
                  <span className="font-mono font-semibold">{o.target}</span>
                  {o.from && <span className="text-[#374151]">← <span className="font-mono">{o.from}</span></span>}
                  {chip(o.status)}
                  {o.cells > 0 && <span className="text-[#6B7280]">{o.cells} MM_RESULT cells in {o.items} blocks</span>}
                </div>
                {o.example && <div className="font-mono text-[11px] text-[#374151] mt-0.5">e.g. {o.example}</div>}
                {o.why && <div className="text-[12px] text-[#6B7280] mt-0.5">{o.why}</div>}
              </div>
            ))}
            {plan.clears.map((c) => (
              <div key={c.target} className="text-[12px] text-[#6B7280] px-3">
                cleared before a run: <span className="font-mono">{c.resolved ?? c.target}</span> · {c.cells ?? 0} filled cells {c.included ? "— emptied" : "— listed only (tick the box to empty them)"}{c.why ? ` · ${c.why}` : ""}
              </div>
            ))}
            {plan.how_it_runs && <div className="text-[12px] text-[#374151] bg-[#EFF6FF] border border-[#BFDBFE] rounded-lg px-3 py-2">{plan.how_it_runs}</div>}
            {plan.notes.map((n) => <div key={n} className="text-[12px] text-[#92400E]">{n}</div>)}
            <div className="flex items-center gap-3 flex-wrap">
              <button onClick={write} disabled={writing || busy || plan.operation_count === 0} className="text-[12px] px-3 py-1.5 rounded bg-[#1F3A5F] text-white hover:bg-[#2a4c7a] disabled:opacity-50">
                {writing ? "Writing through Excel…" : `Write it into a new version (${plan.operation_count} formulas)`}
              </button>
              {plan.operation_count > 0 && <button onClick={() => setShowOps(!showOps)} className="text-[12px] text-[#2563EB]">{showOps ? "Hide the formulas" : "Show the formulas"}</button>}
            </div>
            {showOps && (
              <div className="max-h-[260px] overflow-auto border border-[#E5E7EB] rounded-lg p-2 font-mono text-[11px] text-[#374151]">
                {plan.operations.map((op, i) => <div key={i}><span className="text-[#9CA3AF]">{op.sheet}!{op.cell}</span> {String(op.after ?? "(cleared)")}</div>)}
                {plan.operation_count > plan.operations.length && <div className="text-[#9CA3AF]">… {plan.operation_count - plan.operations.length} more</div>}
              </div>
            )}
            {done && (
              <div className="text-[12px] text-[#0F6E3A] bg-[#DFF5E6] border border-[#0F6E3A]/20 rounded-lg px-3 py-2">
                {done} Recalculate shows Excel's values for the first loop value (MM for Excel evaluates one pass; without the add-in, MM_ cells show #NAME? and are reported as the add-in gap, not as errors). Upload the new version to Mind to run the whole loop.
              </div>
            )}
          </>
        )}
      </div>
    </section>
  );
}
