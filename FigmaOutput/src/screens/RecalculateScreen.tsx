import { useMemo, useState } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { FixTarget, RecalcGroup } from "../types";
import StatusPill from "../components/StatusPill";

export default function RecalculateScreen() {
  const sessionId = useStore((s) => s.sessionId);
  const currentVersion = useStore((s) => s.currentVersion);
  const result = useStore((s) => s.recalcResult);
  const setRecalcResult = useStore((s) => s.setRecalcResult);
  const addVersion = useStore((s) => s.addVersion);
  const setReport = useStore((s) => s.setReport);
  const openFix = useStore((s) => s.openFix);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const groups: RecalcGroup[] = result?.groups ?? [];
  const groupOfCell = useMemo(() => {
    const m = new Map<string, RecalcGroup>();
    for (const g of groups) for (const c of g.cells) m.set(`${c.sheet}!${c.cell}`, g);
    return m;
  }, [groups]);
  const shared = groups.filter((g) => g.count > 1);

  if (!sessionId) {
    return (
      <div className="p-8 text-[13px] text-[#9CA3AF]">Upload and analyze a workbook first.</div>
    );
  }

  async function recalc() {
    setRunning(true);
    setError(null);
    try {
      const res = await api.recalculate(sessionId!);
      if (res.version) addVersion(res.version);
      if (res.report) setReport(res.report, res.plan ?? useStore.getState().plan);
      setRecalcResult(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setRunning(false);
    }
  }

  function singleTarget(sheet: string, cell: string, err: string, formula: string, addinGap = false, array?: string): FixTarget {
    const g = groupOfCell.get(`${sheet}!${cell}`);
    const siblings = g ? g.cells.filter((c) => !(c.sheet === sheet && c.cell === cell)) : [];
    return { kind: "recalc", rule_id: "READY-001", sheet, cell, error: err, formula, addin_gap: addinGap, cells: siblings, cause: g?.cause, group_id: g?.id, array };
  }

  function groupTarget(g: RecalcGroup): FixTarget {
    return { kind: "recalc-group", rule_id: "READY-001", error: g.error, cause: g.cause, cells: g.cells, group_id: g.id, addin_gap: g.kind === "addin_gap" };
  }

  const stale = result && currentVersion && result.version_id && result.version_id !== currentVersion.id;

  return (
    <div className="p-8 max-w-4xl">
      <h1 className="text-lg font-semibold text-[#111827] mb-1">Recalculate</h1>
      <p className="text-[13px] text-[#6B7280] mb-6 max-w-xl">
        Drives the installed Excel (hidden, macros disabled) to recalculate a fresh copy and scan for
        formula errors. This is the only path to a confirmed PASS status. Every error has a Fix button;
        errors that share one root cause can be fixed together.
      </p>

      <button
        onClick={recalc}
        disabled={running}
        className="px-6 py-3 bg-[#1F3A5F] text-white rounded-lg font-semibold text-sm hover:bg-[#162d4a] transition-colors disabled:opacity-50 flex items-center gap-2"
        aria-label="Recalculate now"
      >
        {running && (
          <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
        )}
        {running ? "Recalculating…" : result ? "Recalculate again" : "Recalculate now"}
      </button>

      {running && (
        <div className="mt-4 text-[13px] text-[#6B7280]" aria-live="assertive" aria-label="Recalculation in progress">
          Opening Excel in the background · running full recalculation · scanning for formula errors…
        </div>
      )}
      {error && (
        <div className="mt-4 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded px-3 py-2" role="alert">
          {error}
        </div>
      )}

      {result && (
        <div className="mt-6 border border-[#E5E7EB] rounded-xl bg-white overflow-hidden" aria-live="polite">
          <div className="px-5 py-4 border-b border-[#E5E7EB] flex items-center gap-3 flex-wrap">
            <StatusPill status={result.status} />
            <span className="text-[13px] text-[#374151]">{result.message}</span>
            {result.version_id && (
              <span className="ml-auto text-[11px] font-mono text-[#9CA3AF]">
                recalculated {result.version_id.replace("ver-00", "v").replace("ver-0", "v")}
              </span>
            )}
          </div>
          {result.ran === false && (
            <div className="px-5 py-2 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border-b border-[#9F1D1D]/20" role="alert">
              Excel could not run the recalculation, so nothing below is a result — try again (close other Excel automation first).
            </div>
          )}
          {stale && (
            <div className="px-5 py-2 text-[12px] text-[#8A5A00] bg-[#FFF1CC] border-b border-[#8A5A00]/20">
              The workbook changed since this recalculation ({currentVersion?.label.split(" — ")[0]} is current) — run it again to confirm.
            </div>
          )}

          {shared.length > 0 && (
            <div className="px-5 pt-5">
              <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-2">SHARED ROOT CAUSES — FIX ALL AT ONCE</div>
              <div className="flex flex-col gap-2">
                {shared.map((g) => (
                  <div key={g.id} className={`border rounded-lg p-3 flex items-start gap-3 ${g.kind === "addin_gap" ? "border-[#8A5A00]/30 bg-[#FFF1CC]/40" : "border-[#1F3A5F]/30 bg-[#EEF2FF]"}`}>
                    <div className="flex-1 min-w-0">
                      <div className="text-[13px] text-[#111827]">
                        <span className="font-mono font-semibold">{g.count}</span> errors share one cause:{" "}
                        <span className="font-mono text-[12px] text-[#9F1D1D] bg-[#FDE2E2] px-1 py-0.5 rounded">{g.error}</span>{" "}
                        <span className="text-[#374151]">{g.cause}</span>
                      </div>
                      <div className="text-[11px] font-mono text-[#6B7280] mt-1 truncate" title={g.cells.map((c) => `${c.sheet}!${c.cell}`).join(", ")}>
                        {g.cells.slice(0, 8).map((c) => `${c.sheet}!${c.cell}`).join(", ")}{g.cells.length > 8 ? ` … +${g.cells.length - 8}` : ""}
                      </div>
                    </div>
                    <button
                      onClick={() => openFix(groupTarget(g))}
                      className="px-3 py-1.5 rounded-md text-[12px] font-semibold bg-[#1F3A5F] text-white hover:bg-[#162d4a] whitespace-nowrap"
                    >
                      Fix all {g.count}
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {result.formula_errors.length > 0 && (
            <div className="p-5">
              <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-3">FORMULA ERRORS ({result.formula_errors.length})</div>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-[13px]">
                  <thead className="bg-[#F9FAFB] border-b border-[#E5E7EB]">
                    <tr>
                      {["Sheet", "Cell", "Error", "Formula", "Cause", "Fix"].map((h) => (
                        <th key={h} className="px-4 py-2 text-left text-[11px] font-semibold text-[#6B7280] tracking-wide">
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {result.formula_errors.map((e, i) => {
                      const g = groupOfCell.get(`${e.sheet}!${e.cell}`);
                      return (
                        <tr key={i} className="border-b border-[#E5E7EB] last:border-0 hover:bg-[#F9FAFB]">
                          <td className="px-4 py-2.5 font-mono text-[13px] text-[#374151]">{e.sheet}</td>
                          <td className="px-4 py-2.5 font-mono text-[13px] text-[#374151]">{e.cell}</td>
                          <td className="px-4 py-2.5">
                            <button
                              onClick={() => openFix(singleTarget(e.sheet, e.cell, e.error, e.formula, false, e.array))}
                              className="font-mono text-[12px] text-[#9F1D1D] bg-[#FDE2E2] px-1.5 py-0.5 rounded hover:ring-1 hover:ring-[#9F1D1D]/40"
                              title="Open the fix panel"
                            >
                              {e.error}
                            </button>
                          </td>
                          <td className="px-4 py-2.5 font-mono text-[12px] text-[#374151] max-w-[360px] truncate" title={e.formula}>{e.formula}</td>
                          <td className="px-4 py-2.5 text-[11px] text-[#6B7280] whitespace-nowrap">
                            {e.array && (
                              <span className="font-mono text-[#8A5A00] bg-[#FFF1CC] px-1.5 py-0.5 rounded mr-1" title="Array formula (Ctrl+Shift+Enter) — changed as a whole">
                                array {e.array}
                              </span>
                            )}
                            {g && g.count > 1 ? (
                              <button onClick={() => openFix(groupTarget(g))} className="font-mono text-[#1F3A5F] bg-[#EEF2FF] px-1.5 py-0.5 rounded hover:ring-1 hover:ring-[#1F3A5F]/40" title={g.cause}>
                                shared ×{g.count}
                              </button>
                            ) : (
                              <span className="text-[#9CA3AF]">unique</span>
                            )}
                          </td>
                          <td className="px-4 py-2.5 whitespace-nowrap">
                            <button
                              onClick={() => openFix(singleTarget(e.sheet, e.cell, e.error, e.formula, false, e.array))}
                              className="px-2.5 py-1 rounded-md text-[12px] font-medium bg-[#1F3A5F] text-white hover:bg-[#162d4a]"
                            >
                              Fix
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {result.addin_gap_errors.length > 0 && (
            <div className="px-5 pb-5">
              <div className="bg-[#FFF1CC] border border-[#8A5A00]/20 rounded-lg p-3">
                <div className="text-[12px] font-semibold text-[#8A5A00] mb-1">
                  MM_ function errors (#NAME?) — not a workbook defect
                </div>
                <p className="text-[12px] text-[#8A5A00]">
                  {result.addin_gap_errors.length} cell{result.addin_gap_errors.length !== 1 ? "s" : ""} use MM_ functions that return #NAME? because the MMForExcel add-in is not installed in this environment.
                  Mind evaluates MM_ functions itself; open a cell to see what, if anything, needs to change.
                </p>
                <div className="mt-2 flex flex-col gap-1">
                  {result.addin_gap_errors.map((e, i) => (
                    <div key={i} className="flex items-center gap-2">
                      <code className="text-[11px] text-[#8A5A00] font-mono flex-1 truncate" title={e.formula}>
                        {e.sheet}!{e.cell}: {e.formula}
                      </code>
                      <button
                        onClick={() => openFix(singleTarget(e.sheet, e.cell, "#NAME?", e.formula, true))}
                        className="px-2 py-0.5 rounded text-[11px] font-medium border border-[#8A5A00]/40 text-[#8A5A00] hover:bg-white"
                      >
                        Ask / Fix
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {result.ran !== false && result.formula_errors.length === 0 && result.addin_gap_errors.length === 0 && result.status === "PASS" && (
            <div className="px-5 py-4 flex items-center gap-2 text-[#0F766E]">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                <circle cx="8" cy="8" r="7" stroke="currentColor" strokeWidth="1.5" fill="none" />
                <path d="M5 8l2.5 2.5L11 5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" fill="none" />
              </svg>
              <span className="text-[13px] font-medium">No formula errors. Workbook recalculates cleanly.</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
