import { useStore } from "../store";
import StatusPill from "./StatusPill";
import VersionChip from "./VersionChip";
import ScanProgress, { useScanStatus } from "./ScanProgress";
import * as api from "../services/api";

export default function TopBar() {
  const summary = useStore((s) => s.summary);
  const report = useStore((s) => s.report);
  const currentVersion = useStore((s) => s.currentVersion);
  const sessionId = useStore((s) => s.sessionId);
  const applyAnalysis = useStore((s) => s.applyAnalysis);
  const setIsAnalyzing = useStore((s) => s.setIsAnalyzing);
  const isAnalyzing = useStore((s) => s.isAnalyzing);
  // 1.6.6: the re-analysis shows the backend's live stage (sheet / rule being worked on)
  const scan = useScanStatus(sessionId, isAnalyzing);

  async function handleReanalyze() {
    if (!sessionId || !currentVersion) return;
    setIsAnalyzing(true);
    try {
      // Always the current (latest) version; the response carries the delta vs the previous analysis.
      const res = await api.reanalyze(sessionId, currentVersion.id);
      applyAnalysis({ summary: res.summary, report: res.report, plan: res.plan, delta: res.delta, versions: res.versions, readiness: res.readiness, prep_progress: res.prep_progress });
    } catch (e) {
      window.alert(`Re-analysis failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setIsAnalyzing(false);
    }
  }

  return (
    <header className="h-12 border-b border-[#E5E7EB] bg-white flex items-center px-4 gap-3 flex-shrink-0 sticky top-0 z-10">
      {summary ? (
        <>
          <span className="font-mono font-semibold text-sm text-[#111827] truncate max-w-[240px]">
            {summary.file_name}
          </span>
          {currentVersion && <VersionChip version={currentVersion} />}
          {report && <StatusPill status={report.status} />}
          <div className="ml-auto flex items-center gap-3">
            {isAnalyzing && scan && scan.state === "running" && <ScanProgress status={scan} compact />}
            <button
              onClick={handleReanalyze}
              disabled={isAnalyzing}
              className="px-3 py-1.5 rounded-md border border-[#E5E7EB] text-sm text-[#374151] hover:bg-[#F9FAFB] transition-colors font-medium disabled:opacity-50"
            >
              {isAnalyzing ? "Re-analyzing…" : "Re-analyze"}
            </button>
          </div>
        </>
      ) : (
        <span className="text-sm text-[#9CA3AF] font-medium">Mind Ready — No workbook loaded</span>
      )}
    </header>
  );
}
