import { useEffect, useMemo, useState, type ReactNode } from "react";
import { Link } from "react-router-dom";
import { useStore } from "../store";
import type { PrepAction, ApplyResult, BlockingItem } from "../types";
import * as api from "../services/api";
import OperationsTable from "../components/OperationsTable";
import LevelBadge from "../components/LevelBadge";
import PrepGauge from "../components/PrepGauge";

/**
 * Prep (1.7.2): three blocks, one vocabulary.
 *   Blocking  -- Mind refuses the file (or computes it wrong) without these; on by default.
 *   Optional  -- Mind reads the file as it is; these improve it.
 *   By hand   -- blocking problems Prep has no repair for: straight to the Fix panel.
 * Every action shows what it left for review up front, with a Fix button, so
 * nothing is discovered only after the next scan.
 */
function ActionCard({ action, checked, onChange }: { action: PrepAction; checked: boolean; onChange: (id: string, checked: boolean) => void }) {
  const [expanded, setExpanded] = useState(false);
  const openFix = useStore((s) => s.openFix);
  const disabled = action.count === 0;
  const leftover = action.skipped.length;
  return (
    <div className={`border rounded-xl bg-white transition-colors ${checked && !disabled ? (action.level === "blocking" ? "border-[#9F1D1D]/60" : "border-[#1F3A5F]") : "border-[#E5E7EB]"} ${disabled ? "opacity-80" : ""}`}>
      <div className="flex items-start gap-3 p-4">
        <input
          type="checkbox"
          checked={checked && !disabled}
          disabled={disabled}
          onChange={(e) => onChange(action.id, e.target.checked)}
          className="mt-0.5 w-4 h-4 accent-[#1F3A5F] cursor-pointer flex-shrink-0 disabled:cursor-not-allowed"
          aria-label={action.title}
        />
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <LevelBadge level={action.level} />
            <span className="font-semibold text-[13px] text-[#111827]">{action.title}</span>
            <span className={`text-[11px] font-mono px-1.5 py-0.5 rounded ${action.count ? "text-[#111827] bg-[#F3F4F6]" : "text-[#9CA3AF] bg-[#F9FAFB]"}`}>
              {action.count} change{action.count !== 1 ? "s" : ""}
            </span>
            {leftover > 0 && (
              <button
                onClick={() => setExpanded(true)}
                className="text-[11px] font-semibold text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 px-1.5 py-0.5 rounded hover:bg-[#FFE9A8]"
                title="Prep looked at these and did not touch them -- see why"
              >
                {leftover} left for review
              </button>
            )}
            {action.rule_ids.map((r) => (
              <span key={r} className="text-[11px] font-mono text-[#1F3A5F] bg-[#EEF2FF] px-1.5 py-0.5 rounded">{r}</span>
            ))}
          </div>
          {action.level_note && <div className="text-[12px] text-[#6B7280] mt-1">{action.level_note}</div>}
          {action.caution && action.count > 0 && (
            <div className="text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded px-2 py-0.5 inline-block mt-1.5">⚠ {action.caution}</div>
          )}
          {!action.default_on && action.count > 0 && (
            <div className="text-[12px] text-[#6B7280] mt-1">Off by default — tick it if you want it.</div>
          )}
        </div>
        <button onClick={() => setExpanded(!expanded)} className="text-[12px] text-[#6B7280] hover:text-[#374151] flex-shrink-0">
          {expanded ? "Hide" : "Show"} details ↓
        </button>
      </div>
      {expanded && (
        <div className="px-4 pb-4 flex flex-col gap-3 border-t border-[#F3F4F6]">
          {action.operations.length > 0 ? (
            <div className="pt-3">
              <OperationsTable operations={action.operations} compact />
            </div>
          ) : (
            <div className="pt-3 text-[12px] text-[#9CA3AF]">No operation planned{leftover ? " — everything this action looked at is listed below" : ""}.</div>
          )}
          {leftover > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-2">
                <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider">LEFT FOR REVIEW — PREP WILL NOT TOUCH THESE</div>
                <button
                  onClick={() => openFix({ rule_id: action.rule_ids[0] })}
                  className="ml-auto px-2.5 py-1 rounded-md border border-[#1F3A5F] text-[#1F3A5F] text-[11px] font-semibold hover:bg-[#EEF2FF]"
                  title="Open the Fix panel for this rule: it lists every cell and can ask the assistant for one fix covering them"
                >
                  Fix with assistant →
                </button>
              </div>
              <div className="flex flex-col gap-1">
                {action.skipped.map((s, i) => (
                  <div key={i} className="text-[12px] text-[#8A5A00] font-mono bg-[#FFF1CC] border border-[#8A5A00]/15 rounded px-2 py-1">{s}</div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function ManualCard({ item }: { item: BlockingItem }) {
  const openFix = useStore((s) => s.openFix);
  const loc = [item.location?.sheet, item.location?.cell].filter(Boolean).join("!");
  return (
    <div className="border border-[#9F1D1D]/40 rounded-xl bg-[#FFF7F7] p-4 flex items-start gap-3">
      <LevelBadge level="blocking" />
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="font-mono font-semibold text-[13px] text-[#9F1D1D]">{item.rule_id}</span>
          {item.sites != null && <span className="text-[11px] font-mono text-[#6B7280] bg-[#F3F4F6] px-1.5 py-0.5 rounded">{item.sites} cell{item.sites !== 1 ? "s" : ""}</span>}
          {loc && <code className="text-[11px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded">{loc}</code>}
        </div>
        <p className="text-[12px] text-[#374151] mt-1">{item.message}</p>
        <p className="text-[11px] text-[#6B7280] mt-1">
          {item.fix === "prep_skipped" ? "Prep looked at every cell and left them for review (see the action above). " : "No automatic repair exists for this. "}
          The Fix panel lists the cells and can ask the assistant for one fix covering all of them.
        </p>
      </div>
      <button
        onClick={() => openFix({ rule_id: item.rule_id, sheet: item.location?.sheet, cell: item.location?.cell })}
        className="px-3 py-1.5 rounded-md bg-[#9F1D1D] text-white text-[12px] font-semibold hover:bg-[#7F1D1D] flex-shrink-0"
      >
        Fix with assistant
      </button>
    </div>
  );
}

function ResultCard({ result, sessionId }: { result: ApplyResult; sessionId: string }) {
  return (
    <div className={`border rounded-xl p-4 ${
      result.status === "APPLIED" ? "border-[#0F766E]/30 bg-[#F0FDFA]" : result.status === "PARTIAL" ? "border-[#8A5A00]/30 bg-[#FFF1CC]" : "border-[#9F1D1D]/30 bg-[#FDE2E2]"
    }`}>
      <div className="flex items-center gap-2 mb-1">
        <span className={`font-semibold text-sm ${result.status === "APPLIED" ? "text-[#0F766E]" : result.status === "PARTIAL" ? "text-[#8A5A00]" : "text-[#9F1D1D]"}`}>{result.status}</span>
        {result.verified_opens_in_excel && <span className="text-[12px] text-[#0F766E]">· written by Excel and verified to open</span>}
        <a href={api.downloadUrl(sessionId, result.output_name)} className="ml-auto px-3 py-1 bg-[#1F3A5F] text-white rounded-md text-[12px] font-medium">Download {result.output_name}</a>
      </div>
      <p className="text-[13px] text-[#374151]">{result.message}</p>
      {result.failed.length > 0 && (
        <div className="mt-2">
          <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1">NOT APPLIED — AND WHY</div>
          <div className="flex flex-col gap-1">
            {result.failed.map((op, i) => (
              <div key={i} className="text-[12px] font-mono text-[#9F1D1D]">{op.sheet}{op.cell ? `!${op.cell}` : ""} — {op.error}</div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

/** Context-aware grid names from the assistant -- only changes what the title action proposes. */
function GridNamesPanel({ sessionId, onPlan }: { sessionId: string; onPlan: (r: api.GridNamesResult) => void }) {
  const [loading, setLoading] = useState(false);
  const [res, setRes] = useState<api.GridNamesResult | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [open, setOpen] = useState(false);

  async function run() {
    setLoading(true);
    setErr(null);
    try {
      const out = await api.suggestGridNames(sessionId, true);
      setRes(out);
      setOpen(true);
      onPlan(out);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  const changed = (res?.names ?? []).filter((n) => n.suggested && n.suggested !== n.deterministic);
  return (
    <div className="border border-[#E5E7EB] rounded-xl bg-white">
      <div className="flex items-start gap-3 p-4">
        <div className="flex-1 min-w-0">
          <div className="font-semibold text-[13px] text-[#111827]">Name grids from their context <span className="text-[11px] font-normal text-[#6B7280]">(optional, uses the assistant)</span></div>
          <p className="text-[12px] text-[#6B7280] mt-0.5">
            Grids with no heading nearby can only be called <span className="font-mono">&lt;Sheet&gt; &lt;Cell&gt;</span>. The assistant reads the cells around each one and proposes a real name; nothing is written until you apply the grid-titles action.
          </p>
          {res && !res.available && <div className="text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded px-2 py-1 mt-2">Assistant unavailable — the deterministic names stand. {res.message}</div>}
          {res?.available && <div className="text-[12px] text-[#0F766E] mt-2">{changed.length} of {res.names.length} grid{res.names.length !== 1 ? "s" : ""} renamed{res.applied && " · folded into the plan"}</div>}
        </div>
        <div className="flex gap-2 flex-shrink-0">
          {res && res.names.length > 0 && <button onClick={() => setOpen(!open)} className="text-[12px] text-[#6B7280] hover:text-[#374151]">{open ? "Hide" : "Show"} names ↓</button>}
          <button onClick={run} disabled={loading} className="px-3 py-1.5 border border-[#1F3A5F] text-[#1F3A5F] rounded-md text-[12px] font-medium disabled:opacity-50">
            {loading ? "Naming…" : res ? "Name again" : "Suggest names"}
          </button>
        </div>
      </div>
      {err && <div className="px-4 pb-3 text-[12px] text-[#9F1D1D]" role="alert">{err}</div>}
      {open && res?.names?.length ? (
        <div className="px-4 pb-4 border-t border-[#F3F4F6]">
          <div className="overflow-x-auto pt-3">
            <table className="w-full text-[12px]">
              <thead>
                <tr className="text-left text-[11px] text-[#9CA3AF] tracking-wider">
                  <th className="pb-2 pr-3 font-semibold">GRID</th>
                  <th className="pb-2 pr-3 font-semibold">FROM THE CELLS AROUND IT</th>
                  <th className="pb-2 font-semibold">ASSISTANT</th>
                </tr>
              </thead>
              <tbody>
                {res.names.map((n) => (
                  <tr key={n.grid} className="border-t border-[#F3F4F6] align-top">
                    <td className="py-1.5 pr-3 font-mono text-[#6B7280] whitespace-nowrap">{n.grid} <span className="text-[#C4C8CE]">{n.size}</span></td>
                    <td className="py-1.5 pr-3 text-[#9CA3AF]">{n.deterministic || "—"}<div className="text-[10px] text-[#C4C8CE]">{n.deterministic_source}</div></td>
                    <td className={`py-1.5 ${n.suggested && n.suggested !== n.deterministic ? "text-[#0F766E] font-medium" : "text-[#6B7280]"}`}>{n.suggested || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function Section({ title, hint, children, tone }: { title: string; hint: string; children: ReactNode; tone: "blocking" | "optional" | "manual" }) {
  const color = tone === "blocking" ? "text-[#9F1D1D]" : tone === "manual" ? "text-[#9F1D1D]" : "text-[#4B5563]";
  return (
    <section className="flex flex-col gap-2">
      <div className="flex items-baseline gap-2 mt-2">
        <h2 className={`text-[12px] font-semibold tracking-wider ${color}`}>{title}</h2>
        <span className="text-[11px] text-[#9CA3AF]">{hint}</span>
      </div>
      {children}
    </section>
  );
}

export default function PrepScreen() {
  const plan = useStore((s) => s.plan);
  const sessionId = useStore((s) => s.sessionId);
  const readiness = useStore((s) => s.readiness);
  const prepProgress = useStore((s) => s.prepProgress);
  const addVersion = useStore((s) => s.addVersion);
  const applyAnalysis = useStore((s) => s.applyAnalysis);
  const lastDelta = useStore((s) => s.lastDelta);
  const lastApply = useStore((s) => s.lastApply);
  const setLastApply = useStore((s) => s.setLastApply);

  const defaults = (p: PrepAction[]) => new Set(p.filter((a) => a.default_on && a.count > 0).map((a) => a.id));
  const [selected, setSelected] = useState<Set<string>>(defaults(plan));
  const [applying, setApplying] = useState(false);
  const [applyPhase, setApplyPhase] = useState("");
  const [result, setResult] = useState<ApplyResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  // A new plan (after a re-analysis) resets the selection to its defaults.
  useEffect(() => {
    setSelected(defaults(plan));
  }, [plan]);

  function toggle(id: string, checked: boolean) {
    setSelected((prev) => {
      const next = new Set(prev);
      checked ? next.add(id) : next.delete(id);
      return next;
    });
  }

  const blockingActions = useMemo(() => plan.filter((a) => a.level === "blocking"), [plan]);
  const optionalActions = useMemo(() => plan.filter((a) => a.level !== "blocking"), [plan]);
  const manual = useMemo(() => (readiness?.blocking ?? []).filter((b) => b.fix !== "prep"), [readiness]);
  const selectedActions = plan.filter((a) => selected.has(a.id) && a.count > 0);
  const allOps = selectedActions.flatMap((a) => a.operations);
  const selectedBlocking = selectedActions.filter((a) => a.level === "blocking").reduce((n, a) => n + a.count, 0);
  const selectedOptional = allOps.length - selectedBlocking;
  const nothingPlanned = plan.every((a) => a.count === 0);

  async function applyChanges() {
    if (!sessionId || allOps.length === 0) return;
    setApplying(true);
    setResult(null);
    setError(null);
    const phases = ["Copy", "Excel", "Verify", "Change log"];
    for (const p of phases) {
      setApplyPhase(p);
      await new Promise((r) => setTimeout(r, 500));
    }
    try {
      const out = await api.applyOperations(sessionId, allOps, { reanalyze: true });
      setResult(out.result);
      addVersion(out.version);
      applyAnalysis({ summary: out.summary, report: out.report, plan: out.plan, delta: out.delta, readiness: out.readiness, prep_progress: out.prep_progress });
      setLastApply({
        label: `${selectedActions.length} action${selectedActions.length !== 1 ? "s" : ""}`,
        applied: out.result.applied.length,
        failed: out.result.failed.length,
        versionLabel: out.version.label.split(" — ")[0],
        verified: out.result.verified_opens_in_excel,
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setApplying(false);
      setApplyPhase("");
    }
  }

  if (plan.length === 0) {
    return (
      <div className="p-8 flex flex-col items-center justify-center h-64 text-[#9CA3AF] text-sm gap-2">
        <span className="text-2xl">✓</span>
        Nothing to prepare. Upload and analyze a workbook first.
      </div>
    );
  }

  const remainingBlockingOps = readiness?.prep.blocking_ops ?? 0;
  const remainingOptionalOps = readiness?.prep.optional_ops ?? 0;

  return (
    <div className="flex flex-col flex-1 min-h-0">
      <div className="px-6 pt-6 pb-4 border-b border-[#E5E7EB] bg-white flex-shrink-0">
        <h1 className="text-lg font-semibold text-[#111827] mb-1">Prep — step 3: fix</h1>
        <p className="text-[13px] text-[#6B7280]">
          Repairs are planned from the findings and listed cell by cell. <strong>Blocking</strong> ones are what Mind needs; <strong>optional</strong> ones improve a file Mind already reads.
          Tick what you approve: Excel writes a fresh copy, the copy is verified to open, the file is re-analyzed.
        </p>
      </div>

      <div className="flex-1 overflow-auto p-6 pb-28">
        <div className="max-w-4xl">
          <PrepGauge readiness={readiness} progress={prepProgress} />

          {error && <div className="mb-3 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded px-3 py-2" role="alert">{error}</div>}

          {lastApply && (
            <div className="mb-3 text-[12px] text-[#0F766E] bg-[#F0FDFA] border border-[#0F766E]/20 rounded-lg px-3 py-2 flex items-center gap-2 flex-wrap" aria-live="polite">
              <span>
                <strong>{lastApply.versionLabel} created:</strong> {lastApply.applied} change{lastApply.applied !== 1 ? "s" : ""} applied
                {lastApply.failed ? <span className="text-[#9F1D1D]"> · {lastApply.failed} not applied</span> : null}
                {lastApply.verified ? " · verified in Excel" : ""}.
              </span>
              {lastDelta && lastDelta.fixed.length > 0 && <span>Fixed: <span className="font-mono">{lastDelta.fixed.map((d) => d.rule_id).join(", ")}</span>.</span>}
              <span>
                Now {readiness?.blocking_count ?? 0} blocking problem{(readiness?.blocking_count ?? 0) !== 1 ? "s" : ""} left, {remainingBlockingOps + remainingOptionalOps} repair{remainingBlockingOps + remainingOptionalOps !== 1 ? "s" : ""} still planned
                {remainingBlockingOps + remainingOptionalOps > 0 ? ` (${remainingBlockingOps} blocking, ${remainingOptionalOps} optional)` : ""}.
              </span>
              <Link to="/findings" className="underline underline-offset-2">Findings →</Link>
              {readiness?.state === "unverified" && <Link to="/recalculate" className="underline underline-offset-2 font-semibold">Recalculate →</Link>}
            </div>
          )}
          {result && sessionId && <div className="mb-4"><ResultCard result={result} sessionId={sessionId} /></div>}

          <div className="flex flex-col gap-3">
            <Section title="BLOCKING — MIND NEEDS THESE" hint={blockingActions.some((a) => a.count > 0) ? "on by default; read the caution before Apply" : "nothing blocking is left for Prep to do"} tone="blocking">
              {blockingActions.filter((a) => a.count > 0 || a.skipped.length > 0).length === 0 ? (
                <div className="text-[12px] text-[#0F6E3A] bg-[#DFF5E6] border border-[#0F6E3A]/20 rounded-lg px-3 py-2">✓ No blocking repair to apply.</div>
              ) : (
                blockingActions.filter((a) => a.count > 0 || a.skipped.length > 0).map((a) => <ActionCard key={a.id} action={a} checked={selected.has(a.id)} onChange={toggle} />)
              )}
            </Section>

            {manual.length > 0 && (
              <Section title="BY HAND — BLOCKING, NO AUTOMATIC REPAIR" hint="Prep cannot resolve these; the Fix panel lists the cells and asks the assistant" tone="manual">
                {manual.map((m) => <ManualCard key={m.rule_id} item={m} />)}
              </Section>
            )}

            <Section title="OPTIONAL — MIND READS THE FILE WITHOUT THESE" hint="quality and readability; apply them before delivery, not before testing" tone="optional">
              {sessionId && <GridNamesPanel sessionId={sessionId} onPlan={(r) => applyAnalysis({ plan: r.plan ?? undefined, readiness: r.readiness ?? undefined })} />}
              {optionalActions.map((a) => <ActionCard key={a.id} action={a} checked={selected.has(a.id)} onChange={toggle} />)}
            </Section>
          </div>
        </div>
      </div>

      <div className="fixed bottom-0 left-52 right-0 bg-white border-t border-[#E5E7EB] px-6 py-3 flex items-center gap-4 z-20">
        {applying ? (
          <div className="flex items-center gap-3 text-[13px] text-[#6B7280]" aria-live="assertive" aria-label="Applying changes">
            <div className="w-4 h-4 border-2 border-[#1F3A5F] border-t-transparent rounded-full animate-spin" />
            <span>Phase: <strong className="text-[#111827]">{applyPhase}</strong></span>
            <span>Copy · Excel · Verify · Change log · Re-analyze</span>
          </div>
        ) : (
          <>
            <span className="text-[13px] text-[#6B7280]">
              {nothingPlanned ? (
                <>Nothing left for Prep to apply{readiness?.state === "blocked" ? ` — ${readiness.manual_blocking_count} blocking problem(s) need the assistant` : readiness?.state === "unverified" ? " — recalculate to verify" : ""}.</>
              ) : (
                <>
                  {selectedActions.length} action{selectedActions.length !== 1 ? "s" : ""} · {allOps.length} operation{allOps.length !== 1 ? "s" : ""}
                  {allOps.length > 0 && <span className="text-[#9CA3AF]"> ({selectedBlocking} blocking, {selectedOptional} optional)</span>}
                </>
              )}
            </span>
            {readiness?.state === "unverified" && nothingPlanned ? (
              <Link to="/recalculate" className="ml-auto px-5 py-2 bg-[#8A5A00] text-white rounded-lg text-sm font-semibold hover:bg-[#6B4600]">Recalculate in Excel →</Link>
            ) : (
              <button
                onClick={applyChanges}
                disabled={allOps.length === 0}
                className="ml-auto px-5 py-2 bg-[#1F3A5F] text-white rounded-lg text-sm font-semibold hover:bg-[#162d4a] transition-colors disabled:opacity-40"
              >
                Apply {allOps.length} change{allOps.length !== 1 ? "s" : ""}{selectedBlocking > 0 && selectedOptional === 0 ? " (blocking only)" : ""}
              </button>
            )}
          </>
        )}
      </div>
    </div>
  );
}
