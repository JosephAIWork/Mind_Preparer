import { useEffect, useState } from "react";
import { BrowserRouter, Routes, Route } from "react-router-dom";
import Sidebar from "./components/Sidebar";
import TopBar from "./components/TopBar";
import WorkbookScreen from "./screens/WorkbookScreen";
import FindingsScreen from "./screens/FindingsScreen";
import PrepScreen from "./screens/PrepScreen";
import AssistantScreen from "./screens/AssistantScreen";
import RecalculateScreen from "./screens/RecalculateScreen";
import ReportsScreen from "./screens/ReportsScreen";
import HistoryScreen from "./screens/HistoryScreen";
import MindScreen from "./screens/MindScreen";
import GridNamerScreen from "./screens/GridNamerScreen";
import ModelMapScreen from "./screens/ModelMapScreen";
import FixPanel from "./components/FixPanel";
import ReadinessBanner from "./components/ReadinessBanner";
import { health, resumeSession } from "./services/api";
import { useStore } from "./store";

export default function App() {
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const sessionId = useStore((s) => s.sessionId);
  const setSession = useStore((s) => s.setSession);
  const setMode = useStore((s) => s.setMode);
  const applyAnalysis = useStore((s) => s.applyAnalysis);

  // 1.10.0: "/model-map?session=<id>" opens a session the backend still holds (after a reload).
  useEffect(() => {
    if (sessionId) return;
    const wanted = new URLSearchParams(window.location.search).get("session");
    if (!wanted) return;
    resumeSession(wanted)
      .then((r) => {
        setMode(r.mode);
        setSession(r.sessionId, r.summary, r.report, r.plan, r.version);
        applyAnalysis({ versions: r.versions, readiness: r.readiness, prep_progress: r.prep_progress });
      })
      .catch(() => {});
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Browser tab: "MindPrep v1.7.0" -- the version comes from the running backend.
  useEffect(() => {
    document.title = "MindPrep";
    health()
      .then((h) => { document.title = `MindPrep v${h.version}`; })
      .catch(() => {});
  }, []);

  return (
    <BrowserRouter>
      <div className="flex h-screen overflow-hidden bg-[#F9FAFB]">
        <FixPanel />
        <Sidebar collapsed={sidebarCollapsed} onToggle={() => setSidebarCollapsed((v) => !v)} />
        <div className="flex flex-col flex-1 min-w-0 overflow-hidden">
          <TopBar />
          <ReadinessBanner />
          <main className="flex-1 overflow-auto flex flex-col">
            <Routes>
              <Route path="/" element={<WorkbookScreen />} />
              <Route path="/findings" element={<FindingsScreen />} />
              <Route path="/prep" element={<PrepScreen />} />
              <Route path="/grid-namer" element={<GridNamerScreen />} />
              <Route path="/assistant" element={<AssistantScreen />} />
              <Route path="/recalculate" element={<RecalculateScreen />} />
              <Route path="/reports" element={<ReportsScreen />} />
              <Route path="/history" element={<HistoryScreen />} />
              <Route path="/mind" element={<MindScreen />} />
              <Route path="/model-map" element={<ModelMapScreen />} />
            </Routes>
          </main>
        </div>
      </div>
    </BrowserRouter>
  );
}
