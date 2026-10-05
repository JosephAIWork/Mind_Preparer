import { useEffect, useRef, useState } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { AutoFixLeft, AutoFixStatus } from "../services/api";

/**
 * 1.8.0: "Fix all automatically" on the Recalculate step.
 *
 * One click and the backend goes through the formula errors by itself
 * (app/autofix.py): it finds the cells where errors start, fixes them,
 * recalculates, and undoes any fix that changed a value that was good.
 * The manual way -- a Fix button on every error -- stays right below.
 */
function clock(seconds: number): string {
  const s = Math.max(0, Math.round(seconds));
  return s < 60 ? `${s} s` : `${Math.floor(s / 60)} min ${String(s % 60).padStart(2, "0")} s`;
}

function show(v: unknown): string {
  if (v === "" || v === null || v === undefined) return "(empty)";
  return typeof v === "string" ? v : String(v);
}

const N = (n: number | null | undefined) => (n ?? 0).toLocaleString();

interface Props {
  /** something else is working on the workbook (a recalculation): the fixer waits */
  blocked?: boolean;
  /** told whenever the fixer starts or stops running, so the screen can hold its own buttons */
  onRunningChange?: (running: boolean) => void;
}

export default function AutoFixPanel({ blocked = false, onRunningChange }: Props) {
  const sessionId = useStore((s) => s.sessionId);
  const currentVersion = useStore((s) => s.currentVersion);
  const recalcResult = useStore((s) => s.recalcResult);
  const addVersion = useStore((s) => s.addVersion);
  const applyAnalysis = useStore((s) => s.applyAnalysis);
  const setRecalcResult = useStore((s) => s.setRecalcResult);
  const openFix = useStore((s) => s.openFix);

  const [status, setStatus] = useState<AutoFixStatus | null>(null);
  const [useAssistant, setUseAssistant] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [showAllFixes, setShowAllFixes] = useState(false);
  const [showAllLeft, setShowAllLeft] = useState(false);
  const [showMoved, setShowMoved] = useState(false);
  const starting = useRef(false);

  /** Put a finished run into the app's state -- once: its version, its recalculation, its analysis. */
  function fold(full: AutoFixStatus) {
    if (!full.version) return;
    if (useStore.getState().versions.some((v) => v.id === full.version!.id)) return;
    addVersion(full.version);
    if (full.analysis) applyAnalysis(full.analysis);
    if (full.recalc) setRecalcResult(full.recalc);
  }

  // Coming (back) to the screen: is the fixer running, or done, for this session?
  useEffect(() => {
    if (!sessionId) return;
    let cancelled = false;
    api
      .autoFixStatus(sessionId)
      .then(async (st) => {
        if (cancelled) return;
        if (st.state === "done") {
          const full = await api.autoFixStatus(sessionId, true);
          if (cancelled) return;
          fold(full);
          setStatus(full);
        } else {
          setStatus(st);
        }
      })
      .catch(() => undefined); // an older backend without the route: the panel stays a plain button
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sessionId]);

  // While it runs: ask every 1.5 s; when it is done, fetch the whole result once.
  useEffect(() => {
    if (!sessionId || status?.state !== "running") return;
    let cancelled = false;
    const timer = window.setInterval(async () => {
      try {
        const st = await api.autoFixStatus(sessionId);
        if (cancelled) return;
        if (st.state === "done") {
          const full = await api.autoFixStatus(sessionId, true);
          if (cancelled) return;
          fold(full);
          setStatus(full);
        } else {
          setStatus(st);
        }
      } catch {
        /* one missed poll is nothing: the next one will answer */
      }
    }, 1500);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sessionId, status?.state]);

  const isRunning = status?.state === "running";
  useEffect(() => {
    onRunningChange?.(isRunning);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isRunning]);

  if (!sessionId) return null;

  async function start(keepGoodValues: boolean) {
    if (starting.current) return;
    starting.current = true;
    setError(null);
    setShowAllFixes(false);
    setShowAllLeft(false);
    setShowMoved(false);
    try {
      setStatus(await api.startAutoFix(sessionId!, { use_assistant: useAssistant, keep_good_values: keepGoodValues }));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      starting.current = false;
    }
  }

  async function stop() {
    try {
      setStatus(await api.stopAutoFix(sessionId!));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  function fixByHand(item: AutoFixLeft) {
    openFix({ kind: "recalc", rule_id: "READY-001", sheet: item.sheet, cell: item.cell, error: item.error, formula: item.formula });
  }

  const running = status?.state === "running";
  const result = status?.state === "done" ? status.result : undefined;
  // A result belongs on screen while the workbook is still where the run left it.
  const resultIsCurrent = !!result && !!currentVersion && (status?.version ? status.version.id === currentVersion.id : status?.version_id === currentVersion.id);
  const knownClean = !!recalcResult && recalcResult.ran !== false && recalcResult.version_id === currentVersion?.id && recalcResult.formula_errors.length === 0;
  const errorsNow = recalcResult && recalcResult.version_id === currentVersion?.id ? recalcResult.formula_errors.length : null;

  if (!running && !resultIsCurrent && knownClean) return null; // nothing to fix, nothing to report

  const fixes = result?.fixes ?? [];
  const left = result?.left ?? [];
  const moved = result?.good_values_changed_list ?? [];
  const clean = result?.status === "clean";
  const nothingToDo = result?.status === "unchanged";
  const tone = clean || nothingToDo ? "border-[#0F766E]/30 bg-[#F0FDFA]" : "border-[#8A5A00]/30 bg-[#FFF9E6]";

  return (
    <div className="mt-6 border border-[#1F3A5F]/30 rounded-xl bg-white overflow-hidden" aria-label="Fix all automatically">
      <div className="px-5 py-4">
        <div className="text-[13px] font-semibold text-[#111827]">Fix all automatically</div>
        <p className="text-[12px] text-[#6B7280] mt-1 max-w-2xl">
          The app goes through the errors by itself: it finds the cells where errors start, fixes them, recalculates,
          and undoes any fix that changes a value that is good today. You can still fix errors one by one below.
        </p>

        {!running && (
          <div className="mt-3 flex items-center gap-4 flex-wrap">
            <button
              onClick={() => start(true)}
              disabled={blocked}
              title={blocked ? "Wait for the recalculation to finish" : undefined}
              className="px-5 py-2.5 bg-[#0F766E] text-white rounded-lg font-semibold text-sm hover:bg-[#0B5E58] transition-colors disabled:opacity-50"
            >
              {resultIsCurrent && !clean && !nothingToDo ? "Try again" : errorsNow ? `Fix all ${N(errorsNow)} errors automatically` : "Fix all automatically"}
            </button>
            <label className="flex items-center gap-2 text-[12px] text-[#374151] cursor-pointer select-none">
              <input type="checkbox" checked={useAssistant} onChange={(e) => setUseAssistant(e.target.checked)} className="accent-[#0F766E]" />
              Ask the assistant which value fits each error
            </label>
          </div>
        )}

        {running && status && (
          <div className="mt-3" role="status" aria-live="polite">
            <div className="flex items-center gap-2 flex-wrap">
              <div className="w-3.5 h-3.5 border-2 border-[#0F766E] border-t-transparent rounded-full animate-spin flex-shrink-0" />
              <span className="text-[12px] font-semibold text-[#111827]">
                {status.pass_no ? `Pass ${status.pass_no} · ` : ""}
                {status.title ?? "Working"}
              </span>
              {status.errors != null && (
                <span className="text-[12px] font-mono text-[#0F766E] bg-[#F0FDFA] px-1.5 py-0.5 rounded">
                  {N(status.errors)} error{status.errors !== 1 ? "s" : ""} left
                  {status.errors_before != null && status.errors_before !== status.errors ? ` of ${N(status.errors_before)}` : ""}
                </span>
              )}
              <span className="ml-auto text-[11px] font-mono text-[#6B7280]">{clock(status.elapsed_s ?? 0)}</span>
            </div>
            <div className="mt-1.5 h-2 w-full bg-[#E5E7EB] rounded-full overflow-hidden">
              {status.errors != null && status.errors_before ? (
                <div className="h-full bg-[#0F766E] rounded-full transition-all duration-500" style={{ width: `${Math.round(100 * (1 - status.errors / status.errors_before))}%` }} />
              ) : (
                <div className="h-full w-1/3 bg-[#0F766E]/60 rounded-full animate-pulse" />
              )}
            </div>
            <div className="mt-1 flex items-center gap-3 text-[11px]">
              <span className="text-[#374151] truncate flex-1 min-w-0" title={status.message ?? ""}>{status.message}</span>
              <button onClick={stop} className="px-2.5 py-1 border border-[#E5E7EB] rounded-md text-[#6B7280] hover:border-[#9F1D1D] hover:text-[#9F1D1D]">
                Stop (keeps what is fixed)
              </button>
            </div>
          </div>
        )}

        {(error || status?.state === "error") && (
          <div className="mt-3 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded px-3 py-2" role="alert">
            {error ?? status?.error}
          </div>
        )}
      </div>

      {resultIsCurrent && result && (
        <div className={`border-t px-5 py-4 ${tone}`} aria-live="polite">
          <div className={`text-[13px] font-semibold ${clean || nothingToDo ? "text-[#0F766E]" : "text-[#8A5A00]"}`}>
            {clean ? "✓ " : ""}
            {result.summary}
          </div>
          {!nothingToDo && (
            <div className="mt-1 text-[11px] font-mono text-[#6B7280]">
              {N(result.errors_before)} → {N(result.errors_after)} errors · {N(result.cells_rewritten)} cells rewritten · {result.passes} pass{result.passes !== 1 ? "es" : ""}
              {result.rolled_back ? ` · ${N(result.rolled_back)} fix${result.rolled_back !== 1 ? "es" : ""} undone` : ""}
              {status?.elapsed_s ? ` · ${clock(status.elapsed_s)}` : ""}
              {status?.version ? ` · saved as ${status.version.label.split(" — ")[0]}` : ""}
            </div>
          )}

          {result.numbers_check?.ran && result.numbers_check.clean === false && (
            <div className="mt-3 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded-lg px-3 py-2" role="alert">
              <div className="font-semibold">Before the fixer started: {result.numbers_check.verdict}</div>
              <div className="mt-1">
                That happened in an earlier step (Prep), not here.
                {result.regression_cells ? ` The ${N(result.regression_cells)} cell${result.regression_cells !== 1 ? "s" : ""} that became errors are left as they are: a value put over them would hide it.` : ""}
                {" "}Go back to the version before that step in History, or untick the Prep action that inserts rows.
              </div>
              {(result.numbers_check.samples?.good_to_error ?? result.numbers_check.samples?.good_to_other ?? []).slice(0, 3).map((x, i) => (
                <div key={i} className="mt-1 font-mono text-[11px]">
                  {x.now}: {x.was} → {x.is}
                </div>
              ))}
            </div>
          )}

          {fixes.length > 0 && (
            <div className="mt-4">
              <div className="text-[11px] font-semibold text-[#6B7280] tracking-wider mb-1.5">WHAT WAS FIXED ({N(result.fix_groups)} kind{result.fix_groups !== 1 ? "s" : ""} of error)</div>
              <div className="flex flex-col gap-1.5">
                {(showAllFixes ? fixes : fixes.slice(0, 5)).map((g, i) => (
                  <div key={i} className="bg-white border border-[#E5E7EB] rounded-lg px-3 py-2">
                    <div className="flex items-baseline gap-2 flex-wrap">
                      <span className="font-mono text-[12px] font-semibold text-[#111827]">{N(g.cells)} cell{g.cells !== 1 ? "s" : ""}</span>
                      <code className="text-[11px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded">
                        {g.sheet}!{g.ranges[0]}
                        {g.ranges.length + g.more_ranges > 1 ? ` +${g.ranges.length + g.more_ranges - 1} more` : ""}
                      </code>
                      <span className="font-mono text-[11px] text-[#9F1D1D] bg-[#FDE2E2] px-1 py-0.5 rounded">{g.error}</span>
                    </div>
                    <div className="text-[12px] text-[#374151] mt-1">{g.reason}</div>
                    <div className="text-[11px] font-mono text-[#6B7280] mt-1 break-all">
                      <span className="text-[#9CA3AF]">{g.cell}:</span> {show(g.before).slice(0, 160)} <span className="text-[#0F766E]">→</span> {show(g.after).slice(0, 180)}
                    </div>
                  </div>
                ))}
              </div>
              {fixes.length > 5 && (
                <button onClick={() => setShowAllFixes((v) => !v)} className="mt-1.5 text-[12px] text-[#1F3A5F] underline">
                  {showAllFixes ? "Show fewer" : `Show all ${N(fixes.length)}${result.fix_groups > fixes.length ? ` (of ${N(result.fix_groups)})` : ""}`}
                </button>
              )}
            </div>
          )}

          {left.length > 0 && (
            <div className="mt-4">
              <div className="text-[11px] font-semibold text-[#8A5A00] tracking-wider mb-1.5">LEFT FOR A PERSON ({N(result.left_count)} cell{result.left_count !== 1 ? "s" : ""})</div>
              <div className="flex flex-col gap-1.5">
                {(showAllLeft ? left : left.slice(0, 5)).map((g, i) => (
                  <div key={i} className="bg-white border border-[#8A5A00]/20 rounded-lg px-3 py-2 flex items-start gap-3">
                    <div className="flex-1 min-w-0">
                      <div className="flex items-baseline gap-2 flex-wrap">
                        <code className="text-[11px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded">{g.sheet}!{g.cell}</code>
                        {g.count > 1 && <span className="text-[11px] text-[#6B7280]">and {N(g.count - 1)} more like it</span>}
                        <span className="font-mono text-[11px] text-[#9F1D1D] bg-[#FDE2E2] px-1 py-0.5 rounded">{g.error}</span>
                      </div>
                      <div className="text-[11px] font-mono text-[#6B7280] mt-1 truncate" title={g.formula}>{g.formula}</div>
                      <div className="text-[12px] text-[#374151] mt-1">{g.why}</div>
                    </div>
                    <button onClick={() => fixByHand(g)} className="px-2.5 py-1 rounded-md text-[12px] font-medium bg-[#1F3A5F] text-white hover:bg-[#162d4a] whitespace-nowrap">
                      Fix by hand
                    </button>
                  </div>
                ))}
              </div>
              {left.length > 5 && (
                <button onClick={() => setShowAllLeft((v) => !v)} className="mt-1.5 text-[12px] text-[#1F3A5F] underline">
                  {showAllLeft ? "Show fewer" : `Show all ${N(left.length)}${result.left_groups > left.length ? ` (of ${N(result.left_groups)})` : ""}`}
                </button>
              )}
              {result.keep_good_values && (
                <div className="mt-3 text-[12px] text-[#374151] bg-white border border-[#E5E7EB] rounded-lg px-3 py-2">
                  Some of these are errors a formula is hiding on purpose (it shows 0 or a text instead). Fixing them changes what that formula shows.
                  <button onClick={() => start(false)} disabled={running} className="ml-2 px-2.5 py-1 rounded-md text-[12px] font-medium border border-[#8A5A00]/50 text-[#8A5A00] hover:bg-[#FFF1CC] disabled:opacity-40">
                    Fix them too, and list every value that changes
                  </button>
                </div>
              )}
            </div>
          )}

          {result.good_values_changed > 0 && (
            <div className="mt-4">
              <button onClick={() => setShowMoved((v) => !v)} className="text-[11px] font-semibold text-[#8A5A00] tracking-wider underline">
                VALUES THAT CHANGED ({N(result.good_values_changed)}) {showMoved ? "▲" : "▼"}
              </button>
              {showMoved && (
                <div className="mt-1.5 max-h-64 overflow-auto bg-white border border-[#E5E7EB] rounded-lg">
                  <table className="w-full border-collapse text-[11px] font-mono">
                    <tbody>
                      {moved.map((m, i) => (
                        <tr key={i} className="border-b border-[#E5E7EB] last:border-0">
                          <td className="px-2 py-1 text-[#374151] whitespace-nowrap">{m.sheet}!{m.cell}</td>
                          <td className="px-2 py-1 text-[#6B7280]">{show(m.before)} <span className="text-[#8A5A00]">→</span> {show(m.after)}</td>
                          <td className="px-2 py-1 text-[#9CA3AF] truncate max-w-[280px]" title={m.formula}>{m.formula}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
