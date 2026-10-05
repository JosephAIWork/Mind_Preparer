import { useMemo, useState } from "react";
import { Link } from "react-router-dom";
import { useStore } from "../store";
import { findingLevel, type Finding, type Grid, type PrepAction, type Status } from "../types";
import StatusPill from "../components/StatusPill";
import LevelBadge from "../components/LevelBadge";
import KpiTiles from "../components/KpiTile";
import EvidenceTag from "../components/EvidenceTag";
import GridCard, { worstStatus } from "../components/GridCard";
import RepairGauge from "../components/RepairGauge";

const severityOrder: Record<string, number> = { BLOCKER: 0, HIGH: 1, MEDIUM: 2, LOW: 3, INFO: 4 };
const statusFilterOrder: Status[] = ["ERROR", "REQUIRES_USER_INPUT", "NOT_SUPPORTED", "WARNING", "PASS"];

function colToNum(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function cellRowCol(cell: string): { row: number; col: number } | null {
  const m = cell.replace(/\$/g, "").match(/^([A-Za-z]{1,3})(\d+)/);
  return m ? { row: parseInt(m[2], 10), col: colToNum(m[1]) } : null;
}

/** Findings whose location falls inside this grid (or on its title cell). */
export function findingsForGrid(grid: Grid, findings: Finding[]): Finding[] {
  return findings.filter((f) => {
    if (f.location.sheet !== grid.sheet || !f.location.cell) return false;
    if (grid.title_cell && f.location.cell === grid.title_cell) return true;
    const rc = cellRowCol(f.location.cell);
    return !!rc && rc.row >= grid.first_row && rc.row <= grid.last_row && rc.col >= grid.first_col && rc.col <= grid.last_col;
  });
}

/** 1.7.2: how a finding gets fixed -- the same routing as the readiness banner.
 *  prep: a Prep action holds operations for the rule; prep_skipped: Prep looked
 *  and left every site for review (by hand or the assistant); assistant: no
 *  automatic repair exists for the rule. */
export type FixRoute = { kind: "prep" | "prep_skipped" | "assistant"; action?: PrepAction };

export function fixRouteFor(ruleId: string, plan: PrepAction[]): FixRoute {
  const withOps = plan.find((a) => a.rule_ids.includes(ruleId) && a.count > 0);
  if (withOps) return { kind: "prep", action: withOps };
  const skipped = plan.find((a) => a.rule_ids.includes(ruleId) && a.skipped.length > 0);
  if (skipped) return { kind: "prep_skipped", action: skipped };
  return { kind: "assistant" };
}

function FindingRow({ finding, fixedBadge, route, onFix }: { finding: Finding; fixedBadge: boolean; route: FixRoute; onFix: () => void }) {
  const [expanded, setExpanded] = useState(false);
  const loc = [finding.location.sheet, finding.location.cell].filter(Boolean).join("!");
  const actionable = finding.status !== "PASS";
  return (
    <>
      <tr
        className="border-b border-[#E5E7EB] hover:bg-[#F9FAFB] cursor-pointer transition-colors"
        onClick={() => setExpanded(!expanded)}
      >
        <td className="px-4 py-2.5">
          <span className="font-mono text-[#1F3A5F] text-[12px] font-semibold">{finding.rule_id}</span>
          {fixedBadge && (
            <span className="ml-2 text-[10px] font-semibold text-[#0F766E] bg-[#F0FDFA] border border-[#0F766E]/20 rounded px-1.5 py-0.5">FIXED</span>
          )}
        </td>
        <td className="px-4 py-2.5 whitespace-nowrap">
          {/* 1.7.2: blocking (Mind needs it) vs optional (Mind reads the file as it is) -- from the rule's priority */}
          <LevelBadge level={findingLevel(finding)} />
        </td>
        <td className="px-4 py-2.5">
          <span className="text-[12px] font-mono text-[#6B7280]">{finding.severity ?? "—"}</span>
        </td>
        <td className="px-4 py-2.5">
          <button
            onClick={(e) => { e.stopPropagation(); if (actionable) onFix(); }}
            className={actionable ? "cursor-pointer" : "cursor-default"}
            title={actionable ? "Open the fix panel" : undefined}
          >
            <StatusPill status={finding.status} size="sm" />
          </button>
        </td>
        <td className="px-4 py-2.5">
          {loc ? (
            <code className="text-[12px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded">{loc}</code>
          ) : (
            <span className="text-[#9CA3AF] text-[12px]">—</span>
          )}
        </td>
        <td className="px-4 py-2.5 max-w-[340px]">
          <span className="text-[13px] text-[#374151] line-clamp-2">{finding.message}</span>
        </td>
        <td className="px-4 py-2.5">
          <EvidenceTag evidence={finding.evidence} />
        </td>
        <td className="px-4 py-2.5 whitespace-nowrap">
          {!actionable ? (
            <span className="text-[12px] text-[#9CA3AF]">—</span>
          ) : route.kind === "prep" && route.action ? (
            <span className="inline-flex items-center gap-1.5">
              <Link
                to="/prep"
                onClick={(e) => e.stopPropagation()}
                title={`Prep "${route.action.title}" holds ${route.action.count} change(s) for this rule`}
                className="inline-block px-2.5 py-1 rounded-md text-[12px] font-medium bg-[#1F3A5F] text-white hover:bg-[#162d4a]"
              >
                Repair in Prep ({route.action.count})
              </Link>
              {/* the automatic repair is a proposal, never the only way */}
              <button
                onClick={(e) => { e.stopPropagation(); onFix(); }}
                title="Open the Fix panel: every cell, the formula behind it, and a fix you write or ask the assistant for"
                className="px-2.5 py-1 rounded-md text-[12px] font-medium border border-[#1F3A5F] text-[#1F3A5F] hover:bg-[#EEF2FF]"
              >
                By hand
              </button>
            </span>
          ) : (
            <span className="inline-flex items-center gap-1.5">
              {route.kind === "prep_skipped" && (
                <span title="Prep looked at it and left every cell for review" className="text-[10px] font-semibold text-[#92400E] bg-[#FEF3C7] border border-[#F59E0B]/30 rounded px-1.5 py-0.5">BY HAND</span>
              )}
              <button
                onClick={(e) => { e.stopPropagation(); onFix(); }}
                title={route.kind === "prep_skipped" ? "Ask the assistant for a cell-by-cell fix" : "No automatic repair for this rule: ask the assistant"}
                className={`px-2.5 py-1 rounded-md text-[12px] font-medium transition-colors ${
                  finding.correction_available
                    ? "bg-[#1F3A5F] text-white hover:bg-[#162d4a]"
                    : "border border-[#1F3A5F] text-[#1F3A5F] hover:bg-[#EEF2FF]"
                }`}
              >
                {finding.correction_available ? "Fix with assistant" : "Ask assistant"}
              </button>
            </span>
          )}
        </td>
      </tr>
      {expanded && (
        <tr className="border-b border-[#E5E7EB] bg-[#F9FAFB]">
          <td colSpan={8} className="px-6 py-4">
            <div className="flex flex-col gap-3 text-[13px]">
              <p className="text-[#374151]">{finding.message}</p>
              {finding.source.document && (
                <div className="text-[12px] text-[#6B7280] font-mono">
                  Source: {finding.source.document}
                  {finding.source.article && ` · ${finding.source.article}`}
                </div>
              )}
              {finding.observed != null && (
                <div>
                  <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1">OBSERVED</div>
                  <pre className="text-[12px] font-mono bg-white border border-[#E5E7EB] rounded p-3 overflow-x-auto text-[#374151] max-h-64">
                    {JSON.stringify(finding.observed, null, 2)}
                  </pre>
                </div>
              )}
              {actionable && (
                <div className="flex gap-2">
                  <button
                    onClick={(e) => { e.stopPropagation(); onFix(); }}
                    className="text-[12px] text-[#1F3A5F] font-medium underline underline-offset-2"
                  >
                    Open the fix panel →
                  </button>
                </div>
              )}
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

export default function FindingsScreen() {
  const report = useStore((s) => s.report);
  const summary = useStore((s) => s.summary);
  const lastDelta = useStore((s) => s.lastDelta);
  const openFix = useStore((s) => s.openFix);
  const plan = useStore((s) => s.plan);
  const currentVersion = useStore((s) => s.currentVersion);
  const [activeStatuses, setActiveStatuses] = useState<Set<Status>>(
    new Set(["ERROR", "REQUIRES_USER_INPUT", "WARNING"] as Status[])
  );
  const [search, setSearch] = useState("");
  const [tab, setTab] = useState<"table" | "map">("table");
  const [kpiFilter, setKpiFilter] = useState<Status | null>(null);
  const [showFixed, setShowFixed] = useState(true);
  // 1.7.2: the shortest path -- only what Mind needs fixed
  const [blockingOnly, setBlockingOnly] = useState(false);

  const fixedIds = useMemo(() => new Set((lastDelta?.fixed ?? []).map((d) => d.rule_id)), [lastDelta]);

  if (!report) {
    return (
      <div className="p-8 flex flex-col items-center justify-center h-64 text-[#9CA3AF] text-sm">
        No analysis loaded. Upload a workbook to see findings.
      </div>
    );
  }

  const counts = report.summary.status_counts;

  function toggleStatus(s: Status) {
    setActiveStatuses((prev) => {
      const next = new Set(prev);
      if (next.has(s)) next.delete(s); else next.add(s);
      return next;
    });
    setKpiFilter(null);
  }

  function handleKpiClick(s: Status) {
    setKpiFilter(s === kpiFilter ? null : s);
  }

  function fixFinding(f: Finding) {
    openFix({ rule_id: f.rule_id, sheet: f.location.sheet, cell: f.location.cell });
  }

  const filtered = report.findings
    .filter((f) => {
      const isFixedRow = showFixed && fixedIds.has(f.rule_id) && f.status === "PASS";
      if (blockingOnly && findingLevel(f) !== "blocking" && !isFixedRow) return false;
      if (kpiFilter) return f.status === kpiFilter || isFixedRow;
      if (!activeStatuses.has(f.status) && !isFixedRow) return false;
      if (search) {
        const q = search.toLowerCase();
        return (
          f.rule_id.toLowerCase().includes(q) ||
          f.message.toLowerCase().includes(q) ||
          (f.location.sheet ?? "").toLowerCase().includes(q) ||
          (f.location.cell ?? "").toLowerCase().includes(q)
        );
      }
      return true;
    })
    .sort((a, b) => {
      // blocking first, then verify, then optional; severity inside each band
      const band = (f: Finding) => ({ blocking: 0, verify: 1, optional: 2, pass: 3 }[findingLevel(f)]);
      return band(a) - band(b) || (severityOrder[a.severity ?? "INFO"] ?? 4) - (severityOrder[b.severity ?? "INFO"] ?? 4);
    });

  const deltaHasChanges = lastDelta && (lastDelta.fixed.length + lastDelta.improved.length + lastDelta.regressed.length > 0);

  return (
    <div className="flex flex-col flex-1 min-h-0">
      <div className="px-6 pt-6 pb-4 border-b border-[#E5E7EB] bg-white flex-shrink-0">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-3">
            <h1 className="text-lg font-semibold text-[#111827]">Findings</h1>
            <StatusPill status={report.status} />
            {currentVersion && <span className="text-[12px] font-mono text-[#9CA3AF]">{currentVersion.label.split(" — ")[0]}</span>}
          </div>
          <div className="flex gap-1 text-[12px] border border-[#E5E7EB] rounded-lg overflow-hidden">
            <button onClick={() => setTab("table")} className={`px-4 py-1.5 ${tab === "table" ? "bg-[#1F3A5F] text-white" : "text-[#374151] hover:bg-[#F9FAFB]"}`}>Table</button>
            <button onClick={() => setTab("map")} className={`px-4 py-1.5 ${tab === "map" ? "bg-[#1F3A5F] text-white" : "text-[#374151] hover:bg-[#F9FAFB]"}`}>Workbook Map</button>
          </div>
        </div>
        <KpiTiles counts={counts} onFilter={handleKpiClick} activeFilter={kpiFilter} />
        {/* 1.7.2: the repair gauge -- blocking problems and open findings per version */}
        <div className="mt-3 border border-[#E5E7EB] rounded-lg bg-white px-4 py-2"><RepairGauge /></div>
        {report.status !== "PASS" && (
          <p className="text-[12px] text-[#6B7280] mt-2">
            PASS requires a clean recalculation. "Repair in Prep" means Prep already holds the change; "By hand" opens the Fix panel on the same cells, to write the fix yourself or ask the assistant. Grids in the Workbook Map are clickable too.
          </p>
        )}
        {lastDelta && (
          <div
            className={`mt-3 rounded-lg border px-3 py-2 text-[12px] flex items-center gap-3 flex-wrap ${
              deltaHasChanges ? "border-[#0F766E]/30 bg-[#F0FDFA] text-[#0F766E]" : "border-[#E5E7EB] bg-[#F9FAFB] text-[#6B7280]"
            }`}
            aria-live="polite"
          >
            <span className="font-semibold">Since the previous analysis:</span>
            {deltaHasChanges ? (
              <>
                {lastDelta.fixed.length > 0 && (
                  <span>
                    fixed {lastDelta.fixed.length} — <span className="font-mono">{lastDelta.fixed.map((d) => d.rule_id).join(", ")}</span>
                  </span>
                )}
                {lastDelta.improved.length > 0 && (
                  <span>
                    · improved {lastDelta.improved.length} — <span className="font-mono">{lastDelta.improved.map((d) => `${d.rule_id} (${d.from}→${d.to})`).join(", ")}</span>
                  </span>
                )}
                {lastDelta.regressed.length > 0 && (
                  <span className="text-[#9F1D1D]">
                    · regressed {lastDelta.regressed.length} — <span className="font-mono">{lastDelta.regressed.map((d) => d.rule_id).join(", ")}</span>
                  </span>
                )}
                <span className="font-mono text-[#6B7280]">
                  · errors {lastDelta.previous_counts.ERROR ?? 0} → {lastDelta.counts.ERROR ?? 0}
                </span>
                {lastDelta.fixed.length > 0 && (
                  <label className="ml-auto flex items-center gap-1.5 cursor-pointer">
                    <input type="checkbox" checked={showFixed} onChange={(e) => setShowFixed(e.target.checked)} className="accent-[#1F3A5F]" />
                    show fixed rows
                  </label>
                )}
              </>
            ) : (
              <span>no finding changed status (same file re-analyzed, or the change did not affect any rule).</span>
            )}
          </div>
        )}
      </div>

      {tab === "table" ? (
        <>
          <div className="px-6 py-3 border-b border-[#E5E7EB] bg-white flex items-center gap-3 flex-shrink-0">
            <button
              onClick={() => setBlockingOnly((v) => !v)}
              className={`px-2.5 py-1 rounded-full text-[11px] font-semibold border transition-colors ${
                blockingOnly ? "border-[#9F1D1D] bg-[#9F1D1D] text-white" : "border-[#9F1D1D]/40 text-[#9F1D1D] hover:bg-[#FDE2E2]"
              }`}
              title="Only what Mind needs fixed (REQUIRED rules in error)"
            >
              ● Blocking only
            </button>
            <div className="w-px h-5 bg-[#E5E7EB]" />
            <div className="flex gap-1.5 flex-wrap">
              {statusFilterOrder.map((s) => {
                const on = activeStatuses.has(s);
                return (
                  <button
                    key={s}
                    onClick={() => toggleStatus(s)}
                    className={`px-2.5 py-1 rounded-full text-[11px] font-medium border transition-colors ${
                      on ? "border-[#1F3A5F] bg-[#1F3A5F] text-white" : "border-[#E5E7EB] text-[#6B7280] hover:border-[#1F3A5F]"
                    }`}
                  >
                    {s === "REQUIRES_USER_INPUT" ? "NEEDS INPUT" : s}
                  </button>
                );
              })}
            </div>
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search rule, sheet, text…"
              className="ml-auto w-56 px-3 py-1.5 rounded-md border border-[#E5E7EB] text-[13px] text-[#374151] placeholder-[#9CA3AF] focus:outline-none focus:border-[#1F3A5F]"
            />
          </div>
          <div className="flex-1 overflow-auto">
            {filtered.length === 0 ? (
              <div className="flex flex-col items-center justify-center h-48 text-[#9CA3AF] text-sm gap-2">
                <span className="text-2xl">✓</span>
                No findings match the current filter.
              </div>
            ) : (
              <table className="w-full border-collapse text-[13px]">
                <thead className="bg-[#F9FAFB] border-b border-[#E5E7EB] sticky top-0 z-10">
                  <tr>
                    {["Rule", "Level", "Severity", "Status", "Location", "Finding", "Evidence", "Fix"].map((h) => (
                      <th key={h} className="px-4 py-2.5 text-left text-[11px] font-semibold text-[#6B7280] tracking-wide whitespace-nowrap">
                        {h.toUpperCase()}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((f, i) => (
                    <FindingRow
                      key={`${f.rule_id}-${i}`}
                      finding={f}
                      fixedBadge={fixedIds.has(f.rule_id) && f.status === "PASS"}
                      route={fixRouteFor(f.rule_id, plan)}
                      onFix={() => fixFinding(f)}
                    />
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      ) : (
        <div className="flex-1 overflow-auto p-6">
          {!summary ? (
            <div className="text-[#9CA3AF] text-sm">No workbook data available.</div>
          ) : (
            <div className="flex gap-6 overflow-x-auto pb-4">
              {summary.sheets.map((sheet) => {
                const sheetLevel = report.findings.filter((f) => f.location.sheet === sheet.name && !f.location.cell && f.status !== "PASS");
                return (
                  <div key={sheet.name} className="flex-shrink-0 w-64">
                    <div className={`flex items-center gap-2 mb-3 ${sheet.state !== "visible" ? "opacity-50" : ""}`}>
                      <span className="font-mono font-semibold text-[13px] text-[#111827]">{sheet.name}</span>
                      {sheet.state !== "visible" && (
                        <span className="text-[11px] text-[#9CA3AF] font-mono">{sheet.state}</span>
                      )}
                      {sheet.protected && (
                        <span className="text-[10px] text-[#7A3E00] bg-[#FFE4C7] px-1 py-0.5 rounded font-semibold">PROTECTED</span>
                      )}
                    </div>
                    {sheetLevel.length > 0 && (
                      <div className="flex flex-wrap gap-1 mb-2">
                        {sheetLevel.map((f) => (
                          <button
                            key={f.rule_id}
                            onClick={() => fixFinding(f)}
                            className="font-mono text-[10px] text-[#9F1D1D] bg-[#FDE2E2] px-1.5 py-0.5 rounded hover:ring-1 hover:ring-[#9F1D1D]/40"
                            title={f.message}
                          >
                            {f.rule_id}
                          </button>
                        ))}
                      </div>
                    )}
                    <div className="flex flex-col gap-2">
                      {sheet.grids.map((grid) => {
                        const gridFindings = findingsForGrid(grid, report.findings);
                        const worst = worstStatus(gridFindings);
                        const first = gridFindings.filter((f) => f.status === worst)[0];
                        return (
                          <GridCard
                            key={grid.anchor}
                            grid={grid}
                            findings={gridFindings}
                            onClick={first ? () => openFix({ rule_id: first.rule_id, sheet: first.location.sheet, cell: first.location.cell, grid: `${grid.display_name} ${grid.ref}` }) : undefined}
                          />
                        );
                      })}
                      {sheet.standalone_text_cells.map((c) => (
                        <div key={c.cell} className="border border-dashed border-[#E5E7EB] rounded-md px-2.5 py-1.5 text-[11px] text-[#9CA3AF] font-mono">
                          {c.cell}: {c.text}
                        </div>
                      ))}
                      {sheet.grids.length === 0 && sheet.standalone_text_cells.length === 0 && (
                        <div className="border border-dashed border-[#E5E7EB] rounded-md px-3 py-4 text-[11px] text-[#D1D5DB] text-center font-mono">
                          no grids
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
