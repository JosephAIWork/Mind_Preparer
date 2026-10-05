import { useMemo, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useStore } from "../store";
import * as api from "../services/api";
import type { Mode, SheetPart } from "../types";
import VersionChip from "../components/VersionChip";
import ScanProgress, { useScanStatus } from "../components/ScanProgress";

const modes: { value: Mode; label: string; desc: string }[] = [
  { value: "plan", label: "Plan", desc: "Analyze everything — every active readiness rule" },
  { value: "prep_loops", label: "Prep Mind Loops", desc: "Focus on MM_LOOP and dimension rules" },
  { value: "fix_formulas", label: "Fix Incompatible Formulas", desc: "Target FRM-* rules only" },
  { value: "structure_fix", label: "Structure Fix", desc: "STR-* and FMT-* rules only" },
];

const trustStatements = [
  "Original never modified",
  "Outputs written by Excel and verified",
  "Rules mined from the Mind knowledge base",
];

const MB = 1024 * 1024;
const mb = (bytes: number, digits = 1) => `${(bytes / MB).toFixed(digits)} MB`;

type Phase = "idle" | "uploading" | "choose" | "analyzing";

export default function WorkbookScreen() {
  const navigate = useNavigate();
  const summary = useStore((s) => s.summary);
  const mode = useStore((s) => s.mode);
  const currentVersion = useStore((s) => s.currentVersion);
  const versions = useStore((s) => s.versions);
  const setMode = useStore((s) => s.setMode);
  const setSession = useStore((s) => s.setSession);
  const applyAnalysis = useStore((s) => s.applyAnalysis);
  const setVersions = useStore((s) => s.setVersions);
  const setChatHistory = useStore((s) => s.setChatHistory);
  const resetSession = useStore((s) => s.resetSession);
  const setIsUploading = useStore((s) => s.setIsUploading);
  const isUploading = useStore((s) => s.isUploading);

  const [dragOver, setDragOver] = useState(false);
  const [phase, setPhase] = useState<Phase>("idle");
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [pendingUpload, setPendingUpload] = useState<api.PendingUpload | null>(null);
  const [ignore, setIgnore] = useState<string[]>([]);
  const [scanSessionId, setScanSessionId] = useState<string | null>(null);
  const [scanStartedAt, setScanStartedAt] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const scan = useScanStatus(scanSessionId, phase === "analyzing");

  function chooseFile(file: File) {
    if (!file.name.match(/\.(xlsx|xlsm|xlsb)$/i)) {
      setError("Please choose an .xlsx, .xlsm or .xlsb file.");
      return;
    }
    setError(null);
    setPendingFile(file);
    setPendingUpload(null);
    setIgnore([]);
  }

  function finish(res: api.SessionStart) {
    setSession(res.sessionId, res.summary, res.report, res.plan, res.version);
    applyAnalysis({ readiness: res.readiness ?? null, prep_progress: res.prep_progress ?? null });
    if (res.versions && res.versions.length) setVersions(res.versions);
    setChatHistory([]);
    navigate("/findings");
  }

  /** Step 1: upload + inspect. Large workbooks stop here and ask about sheets; the rest scan at once. */
  async function runAnalysis() {
    if (!pendingFile) {
      fileRef.current?.click();
      return;
    }
    setIsUploading(true);
    setPhase("uploading");
    setError(null);
    try {
      const res = await api.uploadWorkbook(pendingFile, mode, { defer: true });
      if (api.isPending(res)) {
        setPendingUpload(res);
        const askable = res.size.above_threshold && res.size.sheets.length > 1;
        if (askable) {
          setPhase("choose");
          setIsUploading(false);
          return;
        }
        await startScan(res.sessionId, []);
        return;
      }
      finish(res); // an older backend scanned at once
      setPhase("idle");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      setPhase("idle");
    } finally {
      setIsUploading(false);
    }
  }

  /** Step 2: the scan itself, with the sheets to skip; the status indicator polls meanwhile. */
  async function startScan(sessionId: string, ignoreSheets: string[]) {
    setIsUploading(true);
    setError(null);
    setScanSessionId(sessionId);
    setScanStartedAt(Date.now());
    setPhase("analyzing");
    try {
      const res = await api.analyzeSession(sessionId, ignoreSheets);
      finish(res);
      setPhase("idle");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      setPhase(pendingUpload ? "choose" : "idle");
    } finally {
      setIsUploading(false);
      setScanSessionId(null);
      setScanStartedAt(null);
    }
  }

  function onDrop(e: React.DragEvent) {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) chooseFile(file);
  }

  function toggleIgnore(name: string) {
    setIgnore((cur) => (cur.includes(name) ? cur.filter((n) => n !== name) : [...cur, name]));
  }

  const sheetsBySize: SheetPart[] = useMemo(
    () => [...(pendingUpload?.size.sheets ?? [])].sort((a, b) => b.bytes - a.bytes),
    [pendingUpload]
  );
  const allSelected = pendingUpload ? ignore.length >= pendingUpload.size.sheets.length : false;
  const skippedBytes = sheetsBySize.filter((s) => ignore.includes(s.name)).reduce((n, s) => n + s.bytes, 0);

  if (summary) {
    const ignored = summary.ignored_sheets ?? [];
    return (
      <div className="p-8 max-w-3xl">
        <h1 className="text-lg font-semibold text-[#111827] mb-1">Workbook</h1>
        <p className="text-sm text-[#6B7280] mb-6">Current file and version lineage.</p>

        <div className="border border-[#E5E7EB] rounded-xl bg-white p-6 mb-6">
          <div className="flex items-start justify-between gap-4">
            <div>
              <div className="font-mono font-semibold text-[#111827] text-base">{summary.file_name}</div>
              <div className="text-sm text-[#6B7280] mt-0.5 font-mono">{summary.sha256} · {summary.file_type.toUpperCase()}</div>
            </div>
            {currentVersion && <VersionChip version={currentVersion} />}
          </div>
          <div className="mt-4 grid grid-cols-2 gap-x-8 gap-y-2 text-[13px]">
            {[
              ["Application", summary.application ?? "—"],
              ["Has VBA", summary.has_vba ? "Yes" : "No"],
              [
                "Sheets",
                `${summary.sheet_count} scanned (${summary.hidden_sheet_count} hidden)` + (summary.ignored_sheet_count ? ` · ${summary.ignored_sheet_count} skipped` : ""),
              ],
              ["Grids", `${summary.grid_count} (${summary.flagged_grid_count} flagged)`],
              ["Formulas", `${summary.formula_count.toLocaleString()} (${summary.array_formula_count} array)`],
              ["External links", String(summary.external_link_part_count)],
              ...(summary.size ? [["Size", `${mb(summary.size.file_bytes)} on disk` + (summary.size.format === "xlsb" ? ` · ${mb(summary.size.decompressed_bytes)} unpacked` : "")]] : []),
            ].map(([k, v]) => (
              <div key={k} className="flex justify-between gap-4 py-1 border-b border-[#F3F4F6]">
                <span className="text-[#6B7280]">{k}</span>
                <span className="font-mono text-[#111827] text-right">{v}</span>
              </div>
            ))}
          </div>
          {ignored.length > 0 && (
            <div className="mt-4 rounded-lg bg-[#FFF1CC] border border-[#8A5A00]/20 px-3 py-2 text-[12px] text-[#8A5A00]">
              <span className="font-semibold">Skipped sheets (not scanned, checked by no rule):</span>{" "}
              <span className="font-mono">{ignored.map((s) => s.name).join(", ")}</span>. They stay in every output untouched. Load the workbook again to scan them.
            </div>
          )}
          {Object.keys(summary.mm_functions_used).length > 0 && (
            <div className="mt-4">
              <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-2">MM_ FUNCTIONS</div>
              <div className="flex flex-wrap gap-2">
                {Object.entries(summary.mm_functions_used).map(([fn, count]) => (
                  <span key={fn} className="font-mono text-[12px] px-2 py-1 rounded-md bg-[#EEF2FF] text-[#3730A3] border border-[#C7D2FE]">
                    {fn} ×{count}
                  </span>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="border border-[#E5E7EB] rounded-xl bg-white p-6 mb-6">
          <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-3">VERSION LINEAGE</div>
          <div className="flex flex-col gap-0">
            {versions.map((v, i) => (
              <div key={v.id} className="flex items-start gap-3">
                <div className="flex flex-col items-center">
                  <div className="w-2 h-2 rounded-full bg-[#1F3A5F] mt-1.5 flex-shrink-0" />
                  {i < versions.length - 1 && <div className="w-px flex-1 bg-[#E5E7EB] my-1" style={{ minHeight: 20 }} />}
                </div>
                <div className="pb-4">
                  <VersionChip version={v} />
                  <div className="text-[11px] text-[#9CA3AF] mt-1 font-mono">
                    {new Date(v.created_at).toLocaleString()} · {v.sha256}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        <button
          onClick={() => { if (window.confirm("Start over with another workbook? This session's versions stay downloadable until the server restarts.")) resetSession(); }}
          className="px-4 py-2 border border-[#E5E7EB] rounded-lg text-[13px] text-[#374151] hover:border-[#1F3A5F] transition-colors"
        >
          Load another workbook
        </button>
      </div>
    );
  }

  return (
    <div className="p-8 max-w-2xl">
      <h1 className="text-lg font-semibold text-[#111827] mb-1">Upload workbook</h1>
      <p className="text-sm text-[#6B7280] mb-6">
        Drop an Excel file to analyze it against every Mind readiness rule.
      </p>

      {phase === "uploading" ? (
        <div className="border border-[#E5E7EB] rounded-xl bg-white p-10 flex flex-col items-center gap-4" role="status" aria-live="polite">
          <div className="w-10 h-10 border-2 border-[#1F3A5F] border-t-transparent rounded-full animate-spin" />
          <div className="text-sm font-medium text-[#374151]">Uploading {pendingFile?.name}…</div>
          <div className="text-[12px] text-[#9CA3AF]">Checking the size and reading the sheet list (nothing is scanned yet)</div>
        </div>
      ) : phase === "analyzing" ? (
        <ScanProgress status={scan} fileName={pendingFile?.name} ignored={ignore} startedAt={scanStartedAt} />
      ) : phase === "choose" && pendingUpload ? (
        <div className="border border-[#E5E7EB] rounded-xl bg-white p-6">
          <div className="flex items-start gap-3">
            <span className="text-xl" aria-hidden>⚖️</span>
            <div>
              <div className="text-sm font-semibold text-[#111827]">This workbook is large — skip some sheets?</div>
              <div className="text-[12px] text-[#6B7280] mt-1">
                <span className="font-mono">{pendingFile?.name}</span>
                {pendingUpload.size.format === "xlsb" ? (
                  <>
                    {" "}unpacks to <span className="font-mono font-semibold text-[#111827]">{mb(pendingUpload.size.decompressed_bytes)}</span>
                    {" "}(the binary .xlsb is measured unpacked; limit {pendingUpload.size.threshold_mb} MB) and
                  </>
                ) : null}
                {" "}will take <span className="font-semibold text-[#111827]">{pendingUpload.size.estimate_text ?? "a minute or more"}</span> to scan
                {" "}({mb(pendingUpload.size.file_bytes)} on disk).
                {" "}Tick the sheets the app should <span className="font-semibold">ignore</span>: they are neither read nor checked, and stay in every output untouched.
              </div>
            </div>
          </div>

          <div className="mt-4 border border-[#E5E7EB] rounded-lg overflow-hidden">
            <div className="grid grid-cols-[28px_1fr_90px_110px] gap-2 px-3 py-2 bg-[#F9FAFB] text-[11px] font-semibold text-[#9CA3AF] tracking-wider">
              <span>SKIP</span>
              <span>SHEET</span>
              <span className="text-right">UNPACKED</span>
              <span>SHARE</span>
            </div>
            <div className="max-h-72 overflow-auto">
              {sheetsBySize.map((sh) => {
                const on = ignore.includes(sh.name);
                return (
                  <label
                    key={sh.name}
                    className={`grid grid-cols-[28px_1fr_90px_110px] gap-2 items-center px-3 py-2 border-t border-[#F3F4F6] text-[13px] cursor-pointer ${on ? "bg-[#FFF8E6]" : "hover:bg-[#F9FAFB]"}`}
                  >
                    <input type="checkbox" checked={on} onChange={() => toggleIgnore(sh.name)} className="accent-[#1F3A5F]" aria-label={`Skip ${sh.name}`} />
                    <span className="truncate">
                      <span className={`font-mono ${on ? "line-through text-[#9CA3AF]" : "text-[#111827]"}`}>{sh.name}</span>
                      {sh.state !== "visible" && <span className="ml-2 text-[10px] px-1.5 py-0.5 rounded bg-[#E9ECEF] text-[#4B5563]">{sh.state}</span>}
                    </span>
                    <span className="font-mono text-right text-[#374151]">{mb(sh.bytes, sh.bytes < MB ? 2 : 1)}</span>
                    <span className="flex items-center gap-2">
                      <span className="h-1.5 flex-1 bg-[#E5E7EB] rounded-full overflow-hidden">
                        <span className="block h-full bg-[#1F3A5F] rounded-full" style={{ width: `${Math.max(2, Math.round(sh.share * 100))}%` }} />
                      </span>
                      <span className="font-mono text-[11px] text-[#6B7280] w-9 text-right">{Math.round(sh.share * 100)}%</span>
                    </span>
                  </label>
                );
              })}
            </div>
          </div>

          {allSelected && (
            <div className="mt-3 text-[12px] text-[#9F1D1D]" role="alert">At least one sheet must be scanned.</div>
          )}
          {error && (
            <div className="mt-3 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded px-3 py-2" role="alert">{error}</div>
          )}

          <div className="mt-4 flex flex-wrap items-center gap-2">
            <button
              onClick={() => startScan(pendingUpload.sessionId, ignore)}
              disabled={isUploading || ignore.length === 0 || allSelected}
              className="px-4 py-2.5 bg-[#1F3A5F] text-white rounded-lg font-semibold text-sm hover:bg-[#162d4a] transition-colors disabled:opacity-50"
            >
              Scan without {ignore.length} sheet{ignore.length === 1 ? "" : "s"}{ignore.length ? ` (−${mb(skippedBytes)})` : ""}
            </button>
            <button
              onClick={() => startScan(pendingUpload.sessionId, [])}
              disabled={isUploading}
              className="px-4 py-2.5 border border-[#1F3A5F] text-[#1F3A5F] rounded-lg font-semibold text-sm hover:bg-[#EEF2FF] transition-colors disabled:opacity-50"
            >
              Scan everything
            </button>
            <button
              onClick={() => { setPendingUpload(null); setPendingFile(null); setIgnore([]); setPhase("idle"); }}
              className="ml-auto text-[12px] text-[#6B7280] hover:text-[#1F3A5F] underline-offset-2 hover:underline"
            >
              Choose another file
            </button>
          </div>
        </div>
      ) : (
        <>
          <div
            onClick={() => fileRef.current?.click()}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
            className={`border-2 border-dashed rounded-xl p-12 flex flex-col items-center gap-3 cursor-pointer transition-colors ${
              dragOver
                ? "border-[#1F3A5F] bg-[#EEF2FF]"
                : pendingFile
                ? "border-[#1F3A5F] bg-white"
                : "border-[#D1D5DB] bg-[#F9FAFB] hover:border-[#1F3A5F] hover:bg-white"
            }`}
          >
            <svg width="36" height="36" viewBox="0 0 36 36" fill="none" className="text-[#9CA3AF]">
              <rect x="4" y="4" width="20" height="28" rx="2" stroke="currentColor" strokeWidth="1.5" fill="none" />
              <path d="M12 4v8h12" stroke="currentColor" strokeWidth="1.5" fill="none" />
              <path d="M25 21v8m0 0l-3-3m3 3l3-3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
            </svg>
            {pendingFile ? (
              <>
                <div className="text-sm font-medium text-[#111827] font-mono">{pendingFile.name}</div>
                <div className="text-[12px] text-[#9CA3AF]">{(pendingFile.size / 1024 / 1024).toFixed(1)} MB on disk · click to choose another file</div>
              </>
            ) : (
              <>
                <div className="text-sm font-medium text-[#374151]">Drop .xlsx / .xlsm / .xlsb here</div>
                <div className="text-[12px] text-[#9CA3AF]">or click to browse</div>
              </>
            )}
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xlsm,.xlsb"
              className="hidden"
              onChange={(e) => { const f = e.target.files?.[0]; if (f) chooseFile(f); e.target.value = ""; }}
            />
          </div>

          <div className="mt-5">
            <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-3">ANALYSIS MODE</div>
            <div className="grid grid-cols-2 gap-2">
              {modes.map((m) => (
                <button
                  key={m.value}
                  onClick={() => setMode(m.value)}
                  className={`text-left px-4 py-3 rounded-lg border text-[13px] transition-colors ${
                    mode === m.value
                      ? "border-[#1F3A5F] bg-[#EEF2FF] text-[#1F3A5F]"
                      : "border-[#E5E7EB] bg-white text-[#374151] hover:border-[#1F3A5F]"
                  }`}
                >
                  <div className="font-semibold">{m.label}</div>
                  <div className="text-[11px] mt-0.5 text-[#6B7280]">{m.desc}</div>
                </button>
              ))}
            </div>
          </div>

          {error && (
            <div className="mt-4 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded px-3 py-2" role="alert">
              {error}
            </div>
          )}

          <button
            onClick={runAnalysis}
            disabled={isUploading}
            className="mt-5 w-full px-5 py-3 bg-[#1F3A5F] text-white rounded-lg font-semibold text-sm hover:bg-[#162d4a] transition-colors disabled:opacity-50"
          >
            {pendingFile ? "Run analysis" : "Choose a workbook"}
          </button>

          <div className="mt-5 flex flex-col gap-1.5">
            {trustStatements.map((s) => (
              <div key={s} className="flex items-center gap-2 text-[12px] text-[#6B7280]">
                <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                  <circle cx="7" cy="7" r="6" stroke="#0F766E" strokeWidth="1.2" />
                  <path d="M4.5 7l2 2 3-3" stroke="#0F766E" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                {s}
              </div>
            ))}
            <div className="flex items-center gap-2 text-[12px] text-[#6B7280]">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                <circle cx="7" cy="7" r="6" stroke="#0F766E" strokeWidth="1.2" />
                <path d="M4.5 7l2 2 3-3" stroke="#0F766E" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Large workbooks (measured unpacked, so an .xlsb counts honestly) let you skip sheets before the scan
            </div>
          </div>
        </>
      )}
    </div>
  );
}
