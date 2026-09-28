import { useEffect, useMemo, useRef, useState } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { CellWindow, ChatMessage, Finding, FixTarget, Operation, Proposal } from "../types";
import StatusPill from "./StatusPill";
import OperationsTable from "./OperationsTable";
import ChatMarkdown from "./ChatMarkdown";
import RepairGauge from "./RepairGauge";

interface Site { sheet: string; cell: string; detail: string }

/** Every Sheet!Cell mentioned in a finding's observed data (call sites, problems, ...). */
export function extractSites(observed: unknown, depth = 0): Site[] {
  const out: Site[] = [];
  const seen = new Set<string>();
  const push = (sheet: string, cell: string, detail: string) => {
    const key = `${sheet}!${cell}`;
    if (!seen.has(key)) {
      seen.add(key);
      out.push({ sheet, cell, detail });
    }
  };
  const walk = (v: unknown, d: number) => {
    if (d > 4 || v == null) return;
    if (Array.isArray(v)) {
      for (const item of v) {
        if (typeof item === "string") {
          const m = item.match(/^(.+?)!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$/);
          if (m) push(m[1], m[2], "");
        } else walk(item, d + 1);
      }
      return;
    }
    if (typeof v === "object") {
      const o = v as Record<string, unknown>;
      if (typeof o.sheet === "string" && typeof o.cell === "string") {
        const detail = [o.formula, o.functions, o.issues, o.issue, o.error, o.reference, o.header, o.name]
          .filter((x) => x != null)
          .map((x) => (Array.isArray(x) ? x.join(", ") : String(x)))
          .join(" · ");
        push(o.sheet, o.cell, detail);
        return;
      }
      for (const val of Object.values(o)) walk(val, d + 1);
    }
  };
  walk(observed, depth);
  return out.slice(0, 40);
}

interface SiteBlock { sheet: string; range: string; cells: number; formula?: string }

/** `observed.site_ranges` (same-formula blocks) when a finding provides it. */
export function extractBlocks(observed: unknown): SiteBlock[] {
  const raw = (observed as { site_ranges?: unknown } | null | undefined)?.site_ranges;
  if (!Array.isArray(raw)) return [];
  return raw.filter(
    (b): b is SiteBlock => typeof b === "object" && b != null && typeof (b as SiteBlock).sheet === "string" && typeof (b as SiteBlock).range === "string" && typeof (b as SiteBlock).cells === "number"
  );
}

const ERROR_MEANING: Record<string, string> = {
  "#NAME?": "Excel does not recognise a name in the formula — an unknown function (not installed add-in, misspelt or unsupported function) or an undefined name.",
  "#REF!": "A reference points to cells that no longer exist (deleted rows/columns/sheets) or is out of range.",
  "#VALUE!": "An operand has the wrong type (text where a number is expected, a range where a single value is expected).",
  "#DIV/0!": "Division by zero or by an empty cell.",
  "#N/A": "A lookup found no match (VLOOKUP/MATCH/INDEX), a value is not available, or an array spilled outside its grid.",
  "#NUM!": "A numeric calculation is invalid (out of range, iteration did not converge).",
  "#NULL!": "Two ranges in the formula do not intersect.",
};

function cellKey(c: { sheet?: string; cell?: string }) {
  return `${c.sheet ?? ""}!${c.cell ?? ""}`;
}

function cellText(v: unknown): string {
  if (v == null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(Number(v.toPrecision(10)));
  return String(v);
}

/** What is going on in the workbook around the cell: contents, or formulas on request. */
function WorkbookView({ win, showFormulas, loading, onToggle }: { win: CellWindow | null; showFormulas: boolean; loading: boolean; onToggle: () => void }) {
  return (
    <div>
      <div className="flex items-center gap-2 mb-1.5">
        <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider">
          IN THE WORKBOOK{win?.range ? ` · ${win.sheet}!${win.range}` : ""}
        </div>
        <button
          onClick={onToggle}
          className={`ml-auto px-2 py-0.5 rounded-full border text-[11px] ${showFormulas ? "border-[#1F3A5F] text-[#1F3A5F] bg-[#EEF2FF]" : "border-[#E5E7EB] text-[#374151] hover:border-[#1F3A5F]"}`}
          title={showFormulas ? "Show the cells' contents" : "Show the formulas behind the cells"}
        >
          {showFormulas ? "Showing formulas" : "Show formulas"}
        </button>
      </div>
      {loading && !win && <div className="text-[11px] text-[#9CA3AF]">Reading the workbook…</div>}
      {win?.error && <div className="text-[11px] text-[#9F1D1D]">{win.error}</div>}
      {win && !win.error && (
        <div className="overflow-auto border border-[#E5E7EB] rounded-lg max-h-56">
          <table className="border-collapse text-[11px] font-mono">
            <thead>
              <tr>
                <th className="sticky top-0 left-0 z-10 bg-[#F3F4F6] border-b border-r border-[#E5E7EB] px-1.5 py-1 text-[#9CA3AF] font-semibold text-left" />
                {win.columns.map((c) => (
                  <th key={c} className="sticky top-0 bg-[#F3F4F6] border-b border-r border-[#E5E7EB] px-2 py-1 text-[#6B7280] font-semibold text-center min-w-[64px]">{c}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {win.rows.map((r) => (
                <tr key={r.row}>
                  <th className="sticky left-0 bg-[#F3F4F6] border-b border-r border-[#E5E7EB] px-1.5 py-1 text-[#6B7280] font-semibold text-right">{r.row}</th>
                  {r.cells.map((c) => {
                    const text = showFormulas ? (c.formula ?? cellText(c.value)) : cellText(c.value);
                    const title = `${win.sheet}!${c.ref}${c.formula ? `\n${c.formula}` : ""}${c.array ? `\narray ${c.array}` : ""}${c.value != null ? `\n= ${cellText(c.value)}` : ""}`;
                    return (
                      <td
                        key={c.ref}
                        title={title}
                        className={`border-b border-r border-[#E5E7EB] px-2 py-1 whitespace-nowrap max-w-[220px] truncate ${
                          c.focus ? "bg-[#EEF2FF] outline outline-1 outline-[#1F3A5F]/50 font-semibold" : ""
                        } ${c.error ? "text-[#9F1D1D]" : showFormulas && c.formula ? "text-[#1F3A5F]" : "text-[#374151]"}`}
                      >
                        {text}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {win && !win.error && (
        <div className="text-[10px] text-[#9CA3AF] mt-1">
          {showFormulas ? "Formulas as entered ({…} = array formula); constants show their value." : "Cell contents as Excel last stored them"}
          {win.values_from ? ` · values from the ${win.values_from}` : ""} · hover a cell for details
        </div>
      )}
    </div>
  );
}

export default function FixPanel() {
  const target = useStore((s) => s.fixTarget);
  const openFix = useStore((s) => s.openFix);
  const closeFix = useStore((s) => s.closeFix);
  const sessionId = useStore((s) => s.sessionId);
  const report = useStore((s) => s.report);
  const plan = useStore((s) => s.plan);
  const recalcResult = useStore((s) => s.recalcResult);
  const currentVersion = useStore((s) => s.currentVersion);
  const setRecalcResult = useStore((s) => s.setRecalcResult);
  const addVersion = useStore((s) => s.addVersion);
  const applyAnalysis = useStore((s) => s.applyAnalysis);
  const appendChatMessage = useStore((s) => s.appendChatMessage);

  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState<"chat" | "apply" | "recalc" | null>(null);
  const [proposal, setProposal] = useState<Proposal | null>(null);
  const [outcome, setOutcome] = useState<string | null>(null);
  const [recalcNote, setRecalcNote] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const bottomRef = useRef<HTMLDivElement>(null);
  // the workbook view: which cell is shown, its window, and whether formulas are on
  const [view, setView] = useState<{ sheet: string; cell: string } | null>(null);
  const [win, setWin] = useState<CellWindow | null>(null);
  const [winLoading, setWinLoading] = useState(false);
  const [showFormulas, setShowFormulas] = useState(false);

  const isRecalc = target?.kind === "recalc";
  const isGroup = target?.kind === "recalc-group";
  const isRecalcAny = isRecalc || isGroup;

  // A new target resets the mini conversation.
  useEffect(() => {
    setMessages([]);
    setInput("");
    setProposal(null);
    setOutcome(null);
    setRecalcNote(null);
    setError(null);
    setView(null);
    setWin(null);
  }, [target?.kind, target?.rule_id, target?.sheet, target?.cell, target?.error, target?.group_id]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, proposal, busy]);

  const finding: Finding | undefined = useMemo(
    () => (target && !isRecalcAny ? report?.findings.find((f) => f.rule_id === target.rule_id) : undefined),
    [report, target, isRecalcAny]
  );
  const sites = useMemo(() => (finding ? extractSites(finding.observed) : []), [finding]);
  // Same-formula blocks (FRM-002 site_ranges): every affected cell, grouped so one range op can fix each block.
  const blocks = useMemo(() => extractBlocks(finding?.observed), [finding]);
  const blockTotal = useMemo(() => {
    const o = finding?.observed as { total_call_sites?: unknown } | undefined;
    return typeof o?.total_call_sites === "number" ? o.total_call_sites : null;
  }, [finding]);
  // Cells this panel is about: the group's cells, or the single cell.
  const targetCells = useMemo(() => {
    if (!target) return [] as { sheet: string; cell: string; formula?: string; error?: string; array?: string }[];
    if (isGroup) return target.cells ?? [];
    if (isRecalc && target.sheet && target.cell) return [{ sheet: target.sheet, cell: target.cell, formula: target.formula, error: target.error }];
    return [];
  }, [target, isGroup, isRecalc]);
  const quickFixes = useMemo(() => {
    if (!target) return [];
    if (isRecalcAny) {
      const keys = new Set(targetCells.map(cellKey));
      return plan
        .map((a) => ({ ...a, operations: a.operations.filter((o) => keys.has(cellKey(o)) || (o.cells ?? []).some((c) => keys.has(`${o.sheet}!${c}`))) }))
        .filter((a) => a.operations.length > 0)
        .map((a) => ({ ...a, count: a.operations.length }));
    }
    return plan.filter((a) => a.count > 0 && a.rule_ids.includes(target.rule_id));
  }, [plan, target, isRecalcAny, targetCells]);

  // Latest recalculation: which of the target cells still fail?
  const stillFailing = useMemo(() => {
    if (!isRecalcAny || !target || !recalcResult || recalcResult.ran === false) return null;
    const failing = new Set([...recalcResult.formula_errors, ...recalcResult.addin_gap_errors].map(cellKey));
    return targetCells.filter((c) => failing.has(cellKey(c)));
  }, [isRecalcAny, target, recalcResult, targetCells]);

  // default cell for the workbook view: the error cell, the group's first cell, or the finding's first site
  useEffect(() => {
    if (!target || view) return;
    const first = targetCells[0] ?? sites[0];
    if (first) setView({ sheet: first.sheet, cell: first.cell });
  }, [target, targetCells, sites, view]);

  // (re)load the window when the viewed cell or the workbook version changes
  useEffect(() => {
    if (!sessionId || !view) return;
    let cancelled = false;
    setWinLoading(true);
    api
      .cellWindow(sessionId, view.sheet, view.cell)
      .then((w) => { if (!cancelled) setWin(w); })
      .catch((e) => { if (!cancelled) setWin({ sheet: view.sheet, focus: view.cell, columns: [], rows: [], error: e instanceof Error ? e.message : String(e) }); })
      .finally(() => { if (!cancelled) setWinLoading(false); });
    return () => { cancelled = true; };
  }, [sessionId, view, currentVersion?.id, recalcResult?.version_id]);

  if (!target || !sessionId) return null;

  const siblings = isRecalc ? target.cells ?? [] : [];
  const total = isGroup ? targetCells.length : siblings.length + 1;

  async function applyOps(ops: Operation[], label: string) {
    if (!sessionId || ops.length === 0) return;
    setBusy("apply");
    setError(null);
    try {
      const out = await api.applyOperations(sessionId, ops, { reanalyze: true });
      addVersion(out.version);
      applyAnalysis({ summary: out.summary, report: out.report, plan: out.plan, delta: out.delta, readiness: out.readiness, prep_progress: out.prep_progress });
      const verified = out.result.verified_opens_in_excel ? "verified" : "not verified";
      const vlabel = out.version.label.split(" — ")[0];
      const failed = out.result.failed ?? [];
      const failedNote = failed.length ? ` · ${failed.length} FAILED` : "";
      const base = `${label}: ${out.result.applied.length} change${out.result.applied.length !== 1 ? "s" : ""} applied via ${
        out.result.method === "excel_com" ? "Excel" : "openpyxl"
      }${failedNote} · ${verified} · re-analyzed as ${vlabel}`;
      if (failed.length) {
        setError(
          `${failed.length} change${failed.length !== 1 ? "s" : ""} could not be applied: ` +
            failed.slice(0, 5).map((f) => `${f.sheet}!${f.cell ?? f.row ?? f.column ?? ""} — ${f.error}`).join("; ") +
            (failed.length > 5 ? "; …" : "")
        );
      }
      if (out.result.applied.length === 0) {
        setOutcome(`${label}: nothing was applied${failed.length ? ` — ${failed.length} change${failed.length !== 1 ? "s" : ""} failed` : ""}; the workbook is unchanged (${vlabel} is a copy of it)`);
        setRecalcNote(null);
      } else if (isRecalcAny) {
        setOutcome(`${base} → recalculate to confirm the ${targetCells.length > 1 ? `${targetCells.length} cells are` : "cell is"} clean`);
        setRecalcNote(null);
      } else {
        const after = out.report?.findings.find((f) => f.rule_id === target?.rule_id);
        setOutcome(`${base} → ${target?.rule_id} is now ${after ? after.status : "?"}`);
      }
      appendChatMessage({
        role: "assistant",
        content: `Fix panel (${isGroup ? `${targetCells.length} cells · ${target?.cause ?? target?.error}` : isRecalc ? `${target?.error} at ${target?.sheet}!${target?.cell}` : target?.rule_id}): ${out.result.applied.length} change(s) applied via ${out.result.method === "excel_com" ? "Excel" : "openpyxl"}${failed.length ? `, ${failed.length} failed` : ""}, re-analyzed (${vlabel}).`,
        note: true,
      });
      setProposal(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  }

  async function recalcNow() {
    if (!sessionId) return;
    setBusy("recalc");
    setError(null);
    try {
      const res = await api.recalculate(sessionId);
      setRecalcResult(res);
      if (res.ran === false) {
        setRecalcNote(`Recalculation could not run (${res.message}) — nothing can be concluded; try again.`);
        return;
      }
      const failing = new Set([...res.formula_errors, ...res.addin_gap_errors].map(cellKey));
      const left = targetCells.filter((c) => failing.has(cellKey(c)));
      setRecalcNote(
        left.length === 0
          ? `Recalculated: ${targetCells.length > 1 ? `none of the ${targetCells.length} cells fail` : `no error at ${targetCells[0]?.sheet}!${targetCells[0]?.cell}`} any more (${res.status}: ${res.message})`
          : `Recalculated: ${left.length} of ${targetCells.length} still fail — ${left.slice(0, 6).map(cellKey).join(", ")} (${res.status}: ${res.message})`
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  }

  async function ask(question: string) {
    if (!sessionId || !question.trim() || busy) return;
    const q = question.trim();
    setInput("");
    setError(null);
    const next = [...messages, { role: "user" as const, content: q }];
    setMessages(next);
    setBusy("chat");
    try {
      const reply = await api.chat(sessionId, messages, q, target ?? undefined);
      setMessages([...next, { role: "assistant", content: reply.text, provenance: reply.provenance }]);
      if (reply.proposal && (reply.proposal.operations.length > 0 || reply.proposal.errors.length > 0)) {
        setProposal(reply.proposal);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  }

  function fixAll() {
    if (!target) return;
    const all = [...(isRecalc && target.sheet && target.cell ? [{ sheet: target.sheet, cell: target.cell, formula: target.formula ?? "", error: target.error ?? "" }] : []), ...siblings];
    const groupTarget: FixTarget = { kind: "recalc-group", rule_id: "READY-001", error: target.error, cause: target.cause, cells: all, group_id: target.group_id, addin_gap: target.addin_gap };
    openFix(groupTarget);
  }

  const loc = [target.sheet, target.cell].filter(Boolean).join("!");
  const isFixed = !isRecalcAny ? finding?.status === "PASS" : stillFailing !== null && stillFailing.length === 0;
  const title = isGroup ? `${targetCells.length} × ${target.error ?? "error"}` : isRecalc ? `${target.error ?? "error"} at ${loc}` : target.rule_id;
  const proposePrompt = isGroup
    ? `Propose one fix for all ${targetCells.length} cells`
    : isRecalc
    ? target.addin_gap
      ? "Propose a fix (or confirm none is needed)"
      : "Propose a fix"
    : sites.length > 1
    ? `Propose a fix for all ${sites.length} sites`
    : "Propose a fix";
  const prompts = isGroup
    ? [proposePrompt, "Explain the shared root cause"]
    : isRecalc
    ? target.addin_gap
      ? ["Why does this show #NAME?", proposePrompt]
      : ["Explain this error", proposePrompt, "Which cells does this formula depend on?"]
    : ["Explain this finding", proposePrompt, "Why is this a problem for Mind?"];

  return (
    <aside
      className="fixed top-0 right-0 h-screen w-[480px] max-w-full bg-white border-l border-[#E5E7EB] shadow-2xl z-30 flex flex-col"
      role="dialog"
      aria-label={`Fix ${title}`}
    >
      <div className="px-5 py-4 border-b border-[#E5E7EB] flex items-start gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            {isRecalcAny ? (
              <>
                <span className="text-[10px] font-semibold tracking-wider text-[#6B7280]">{isGroup ? "SHARED ROOT CAUSE" : "RECALCULATION"}</span>
                <span className="font-mono text-[12px] text-[#9F1D1D] bg-[#FDE2E2] px-1.5 py-0.5 rounded">{isGroup ? `${targetCells.length} × ${target.error}` : target.error}</span>
                {target.addin_gap && <span className="text-[10px] font-semibold text-[#8A5A00] bg-[#FFF1CC] rounded px-1.5 py-0.5">ADD-IN GAP</span>}
              </>
            ) : (
              <>
                <span className="font-mono font-semibold text-[#1F3A5F] text-[13px]">{target.rule_id}</span>
                {finding && <StatusPill status={finding.status} size="sm" />}
              </>
            )}
            {isFixed && <span className="text-[11px] font-semibold text-[#0F766E] bg-[#F0FDFA] border border-[#0F766E]/20 rounded px-1.5 py-0.5">FIXED</span>}
          </div>
          {loc && !isGroup && <code className="text-[12px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded mt-1 inline-block">{loc}</code>}
          {isGroup && target.cause && <div className="text-[12px] text-[#374151] mt-1">{target.cause}</div>}
          {target.grid && <div className="text-[11px] text-[#9CA3AF] font-mono mt-1">grid {target.grid}</div>}
          {/* 1.7.2: the repair gauge -- how the run of fixes is going, version by version */}
          <div className="mt-2"><RepairGauge compact /></div>
        </div>
        <button onClick={closeFix} className="text-[#9CA3AF] hover:text-[#374151] text-lg leading-none" aria-label="Close">×</button>
      </div>

      <div className="flex-1 overflow-auto px-5 py-4 flex flex-col gap-4 text-[13px]">
        {isGroup ? (
          <div>
            <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1.5">CELLS ({targetCells.length})</div>
            <div className="flex flex-col gap-1 max-h-44 overflow-auto">
              {targetCells.map((c) => (
                <div
                  key={cellKey(c)}
                  onClick={() => setView({ sheet: c.sheet, cell: c.cell })}
                  className={`flex gap-2 items-baseline cursor-pointer rounded px-1 -mx-1 ${view?.sheet === c.sheet && view?.cell === c.cell ? "bg-[#EEF2FF]" : "hover:bg-[#F9FAFB]"}`}
                  title="Show this cell in the workbook view"
                >
                  <code className="text-[11px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded whitespace-nowrap">{c.sheet}!{c.cell}</code>
                  {c.formula && <span className="text-[11px] text-[#6B7280] truncate font-mono" title={c.formula}>{c.formula}</span>}
                  {c.array && (
                    <span className="text-[10px] font-mono text-[#8A5A00] bg-[#FFF1CC] px-1 py-0.5 rounded whitespace-nowrap" title="Array formula (Ctrl+Shift+Enter) — Excel only changes it as a whole">
                      array {c.array}
                    </span>
                  )}
                </div>
              ))}
            </div>
            {targetCells.some((c) => c.array) && (
              <p className="text-[11px] text-[#8A5A00] mt-2">
                These cells belong to an array formula — the fix must cover the whole array (the assistant is told this).
              </p>
            )}
            <p className="text-[#374151] leading-relaxed mt-3">
              {target.addin_gap
                ? "These cells call MM_ functions; Excel on this machine has no MMForExcel add-in, so they show #NAME? here. Mind evaluates MM_ functions itself — an environment gap, not a workbook defect. Ask the assistant to confirm nothing needs to change."
                : "These errors share one root cause, so one fix can resolve all of them — ask the assistant to propose it, review the operations, and apply them together."}
            </p>
          </div>
        ) : isRecalc ? (
          <>
            {target.formula && (
              <div>
                <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1">FORMULA{target.array ? ` · ARRAY ${target.array}` : ""}</div>
                <pre className="text-[12px] font-mono bg-[#F9FAFB] border border-[#E5E7EB] rounded p-2 whitespace-pre-wrap break-all text-[#374151]">{target.formula}</pre>
                {target.array && (
                  <p className="text-[11px] text-[#8A5A00] mt-1">Part of the array formula {target.array} (Ctrl+Shift+Enter) — it can only be changed as a whole, so fix all its cells together.</p>
                )}
              </div>
            )}
            <p className="text-[#374151] leading-relaxed">
              {target.addin_gap
                ? "This cell calls an MM_ function; Excel on this machine has no MMForExcel add-in, so it shows #NAME? here. Mind evaluates MM_ functions itself — this is an environment gap, not a workbook defect."
                : ERROR_MEANING[target.error ?? ""] ?? "Excel produced an error for this cell after a full recalculation."}
            </p>
            {siblings.length > 0 && (
              <div className="border border-[#1F3A5F]/30 rounded-lg bg-[#EEF2FF] px-3 py-2 text-[12px] text-[#1F3A5F] flex items-center gap-3 flex-wrap">
                <span>
                  Same root cause in <span className="font-mono font-semibold">{siblings.length}</span> other cell{siblings.length !== 1 ? "s" : ""}
                  {target.cause ? ` — ${target.cause}` : ""}: <span className="font-mono">{siblings.slice(0, 6).map(cellKey).join(", ")}{siblings.length > 6 ? " …" : ""}</span>
                </span>
                <button onClick={fixAll} className="ml-auto px-3 py-1 rounded-md bg-[#1F3A5F] text-white text-[12px] font-semibold hover:bg-[#162d4a] whitespace-nowrap">
                  Fix all {total}
                </button>
              </div>
            )}
          </>
        ) : finding ? (
          <p className="text-[#374151] leading-relaxed">{finding.message}</p>
        ) : (
          <p className="text-[#9CA3AF]">This finding is no longer in the current analysis.</p>
        )}

        {view && <WorkbookView win={win} showFormulas={showFormulas} loading={winLoading} onToggle={() => setShowFormulas((v) => !v)} />}

        {isRecalcAny && recalcResult && stillFailing !== null && (
          <div className={`text-[12px] rounded-lg px-3 py-2 border ${stillFailing.length ? "text-[#9F1D1D] bg-[#FDE2E2] border-[#9F1D1D]/20" : "text-[#0F766E] bg-[#F0FDFA] border-[#0F766E]/20"}`}>
            {stillFailing.length
              ? `Still failing in the latest recalculation (${recalcResult.version_id ?? "current"}): ${stillFailing.length} of ${targetCells.length}.`
              : "Not failing in the latest recalculation."}
          </div>
        )}

        {blocks.length > 0 && (
          <div>
            <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1.5">
              BLOCKS ({blocks.length}{blockTotal != null ? ` · ${blockTotal} cells` : ""})
            </div>
            <div className="flex flex-col gap-1 max-h-40 overflow-auto">
              {blocks.slice(0, 12).map((b) => (
                <div
                  key={`${b.sheet}!${b.range}`}
                  onClick={() => setView({ sheet: b.sheet, cell: b.range.split(":")[0] })}
                  className="flex gap-2 items-baseline cursor-pointer rounded px-1 -mx-1 hover:bg-[#F9FAFB]"
                  title="Show the block's first cell in the workbook view"
                >
                  <code className="text-[11px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded whitespace-nowrap">{b.sheet}!{b.range}</code>
                  <span className="text-[11px] text-[#6B7280] whitespace-nowrap">{b.cells} cell{b.cells !== 1 ? "s" : ""}{b.cells > 1 ? ", same formula" : ""}</span>
                  {b.formula && <span className="text-[11px] text-[#6B7280] truncate font-mono" title={b.formula}>{b.formula}</span>}
                </div>
              ))}
              {blocks.length > 12 && <div className="text-[11px] text-[#9CA3AF]">… {blocks.length - 12} more block{blocks.length - 12 !== 1 ? "s" : ""}</div>}
            </div>
            <p className="text-[11px] text-[#6B7280] mt-1">Each block shares one formula pattern — the assistant can fix a whole block with a single range operation.</p>
          </div>
        )}

        {sites.length > 0 && (
          <div>
            <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1.5">SITES ({sites.length}{blockTotal != null && blockTotal > sites.length ? ` of ${blockTotal}` : ""})</div>
            <div className="flex flex-col gap-1 max-h-40 overflow-auto">
              {sites.map((s) => (
                <div
                  key={`${s.sheet}!${s.cell}`}
                  onClick={() => setView({ sheet: s.sheet, cell: s.cell })}
                  className={`flex gap-2 items-baseline cursor-pointer rounded px-1 -mx-1 ${view?.sheet === s.sheet && view?.cell === s.cell ? "bg-[#EEF2FF]" : "hover:bg-[#F9FAFB]"}`}
                  title="Show this cell in the workbook view"
                >
                  <code className="text-[11px] font-mono text-[#374151] bg-[#F3F4F6] px-1.5 py-0.5 rounded whitespace-nowrap">{s.sheet}!{s.cell}</code>
                  {s.detail && <span className="text-[11px] text-[#6B7280] truncate" title={s.detail}>{s.detail}</span>}
                </div>
              ))}
            </div>
            {sites.length > 1 && <p className="text-[11px] text-[#6B7280] mt-1">All {sites.length} sites share this rule — ask for one fix covering all of them.</p>}
          </div>
        )}

        {outcome && (
          <div className="text-[12px] text-[#0F766E] bg-[#F0FDFA] border border-[#0F766E]/20 rounded-lg px-3 py-2 flex flex-col gap-2" aria-live="polite">
            <span>{outcome}</span>
            {isRecalcAny && !outcome.includes(": nothing was applied") && (
              <button
                onClick={recalcNow}
                disabled={busy !== null}
                className="self-start px-3 py-1.5 bg-[#1F3A5F] text-white rounded-md text-[12px] font-semibold hover:bg-[#162d4a] disabled:opacity-40"
              >
                {busy === "recalc" ? "Recalculating…" : "Recalculate now"}
              </button>
            )}
            {recalcNote && <span>{recalcNote}</span>}
          </div>
        )}
        {error && (
          <div className="text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded-lg px-3 py-2" role="alert">
            {error}
          </div>
        )}

        {!isFixed && (
          <div>
            <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1.5">FIX</div>
            <div className="flex flex-col gap-2">
              {quickFixes.map((a) => (
                <div key={a.id} className="border border-[#E5E7EB] rounded-lg p-3 bg-[#F9FAFB]">
                  <div className="font-medium text-[#111827] text-[12px]">{a.title}</div>
                  <div className="text-[11px] text-[#6B7280] mt-0.5">{a.count} change{a.count !== 1 ? "s" : ""} · rules {a.rule_ids.join(", ")}</div>
                  <button
                    onClick={() => applyOps(a.operations, a.title)}
                    disabled={busy !== null}
                    className="mt-2 px-3 py-1.5 bg-[#1F3A5F] text-white rounded-md text-[12px] font-semibold hover:bg-[#162d4a] disabled:opacity-40"
                  >
                    {busy === "apply" ? "Applying…" : `Apply ${a.count} change${a.count !== 1 ? "s" : ""}`}
                  </button>
                </div>
              ))}
              <div className="border border-[#1F3A5F]/30 rounded-lg p-3 bg-white">
                <div className="text-[12px] text-[#374151]">
                  {quickFixes.length > 0 ? "Or let the assistant propose a fix:" : isRecalcAny ? "Let the assistant propose the fix — it sees the formula, the cells it depends on and the recalculation result." : "Let the assistant propose the fix for this finding."}
                </div>
                <button
                  onClick={() => ask(proposePrompt)}
                  disabled={busy !== null}
                  className="mt-2 px-3 py-1.5 bg-[#1F3A5F] text-white rounded-md text-[12px] font-semibold hover:bg-[#162d4a] disabled:opacity-40"
                >
                  {busy === "chat" ? "Thinking…" : proposePrompt}
                </button>
              </div>
            </div>
          </div>
        )}

        <div>
          <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-1.5">{isRecalcAny ? "ASK ABOUT THIS ERROR" : "ASK ABOUT THIS FINDING"}</div>
          {messages.length === 0 && (
            <div className="flex flex-wrap gap-1.5 mb-2">
              {prompts.map((q) => (
                <button
                  key={q}
                  onClick={() => ask(q)}
                  disabled={busy !== null}
                  className="px-2.5 py-1 rounded-full border border-[#E5E7EB] text-[11px] text-[#374151] hover:border-[#1F3A5F] disabled:opacity-40"
                >
                  {q}
                </button>
              ))}
            </div>
          )}
          <div className="flex flex-col gap-2">
            {messages.map((m, i) => (
              <div key={i} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                <div
                  className={`max-w-[92%] rounded-xl px-3 py-2 text-[12px] leading-relaxed ${
                    m.role === "user" ? "bg-[#1F3A5F] text-white" : "bg-[#F9FAFB] border border-[#E5E7EB] text-[#374151]"
                  }`}
                >
                  <ChatMarkdown content={m.content} />
                </div>
              </div>
            ))}
            {busy === "chat" && (
              <div className="flex gap-1 px-2 py-1">
                {[0, 1, 2].map((i) => (
                  <div key={i} className="w-1.5 h-1.5 bg-[#9CA3AF] rounded-full animate-bounce" style={{ animationDelay: `${i * 150}ms` }} />
                ))}
              </div>
            )}
            {proposal && (
              <div className="border border-[#1F3A5F]/30 rounded-xl bg-[#EEF2FF] p-3">
                {proposal.summary && <p className="text-[12px] font-semibold text-[#1F3A5F] mb-2">{proposal.summary}</p>}
                {/* 1.7.2: a proposal that covers fewer cells than the finding counts is said so, before Apply -- not discovered after the rescan */}
                {(() => {
                  const expected = isGroup ? targetCells.length : sites.length;
                  const covered = new Set(proposal.operations.map((o) => `${o.sheet}!${String(o.cell ?? o.range ?? "").split(":")[0]}`)).size;
                  const rangeOps = proposal.operations.some((o) => String(o.cell ?? o.range ?? "").includes(":"));
                  return expected > 1 && covered < expected && !rangeOps ? (
                    <div className="mb-2 text-[11px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded px-2 py-1">
                      ⚠ Covers {covered} of {expected} cells. Apply it, then ask again for the rest (or ask for one range operation covering all of them).
                    </div>
                  ) : null;
                })()}
                {proposal.operations.length > 0 && <OperationsTable operations={proposal.operations} compact />}
                {proposal.errors.map((e, i) => (
                  <div key={i} className="mt-1 text-[11px] text-[#8A5A00] bg-[#FFF1CC] rounded px-2 py-1">⚠ {e}</div>
                ))}
                <div className="flex gap-2 mt-2">
                  {proposal.operations.length > 0 && (
                    <button
                      onClick={() => applyOps(proposal.operations, "Assistant fix")}
                      disabled={busy !== null}
                      className="px-3 py-1.5 bg-[#1F3A5F] text-white rounded-md text-[12px] font-semibold hover:bg-[#162d4a] disabled:opacity-40"
                    >
                      {busy === "apply" ? "Applying…" : `Apply ${proposal.operations.length} change${proposal.operations.length !== 1 ? "s" : ""}`}
                    </button>
                  )}
                  <button onClick={() => setProposal(null)} className="px-3 py-1.5 border border-[#E5E7EB] text-[#6B7280] rounded-md text-[12px]">Discard</button>
                </div>
              </div>
            )}
            <div ref={bottomRef} />
          </div>
        </div>
      </div>

      <div className="px-5 py-3 border-t border-[#E5E7EB] flex gap-2 items-end">
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); ask(input); } }}
          placeholder={isGroup ? `Ask about these ${targetCells.length} cells, or tell me how to fix them…` : isRecalc ? `Ask about ${target.error} at ${loc}, or tell me how to fix it…` : `Ask about ${target.rule_id}, or tell me how to fix it…`}
          rows={2}
          className="flex-1 resize-none px-3 py-2 border border-[#E5E7EB] rounded-lg text-[12px] text-[#374151] placeholder-[#9CA3AF] focus:outline-none focus:border-[#1F3A5F]"
        />
        <button
          onClick={() => ask(input)}
          disabled={busy !== null || !input.trim()}
          className="px-3 py-2 bg-[#1F3A5F] text-white rounded-lg text-[12px] font-semibold hover:bg-[#162d4a] disabled:opacity-40"
        >
          Send
        </button>
      </div>
    </aside>
  );
}
