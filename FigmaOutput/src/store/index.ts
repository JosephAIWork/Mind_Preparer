import { create } from "zustand";
import type {
  WorkbookSummary,
  ValidationReport,
  PrepAction,
  Version,
  ChatMessage,
  ApplyReport,
  Mode,
  Delta,
  FixTarget,
  RecalcResult,
  Readiness,
  PrepProgress,
} from "../types";

interface SessionState {
  sessionId: string | null;
  mode: Mode;
  summary: WorkbookSummary | null;
  report: ValidationReport | null;
  plan: PrepAction[];
  versions: Version[];
  currentVersion: Version | null;
  chatHistory: ChatMessage[];
  isUploading: boolean;
  isAnalyzing: boolean;
  /** Result of the most recent re-analysis in this session (null until one ran). */
  lastDelta: Delta | null;
  /** Finding the Fix panel is open for (null = closed). */
  fixTarget: FixTarget | null;
  /** Last real Excel recalculation in this session (shared by the Recalculate screen and the Fix panel). */
  recalcResult: RecalcResult | null;
  setRecalcResult: (r: RecalcResult | null) => void;
  /** 1.7.2: the one verdict for the current version, and the Prep gauge history. */
  readiness: Readiness | null;
  prepProgress: PrepProgress | null;
  /** 1.7.2: what the last Apply did, shown on the Prep screen until the next one. */
  lastApply: {
    label: string;
    applied: number;
    failed: number;
    versionLabel: string;
    verified: boolean | null;
    /** the file the Apply produced, by its download name */
    outputName?: string;
    /** 1.7.3: the Apply repair by repair (absent from an older backend) */
    outcome?: ApplyReport | null;
  } | null;
  setLastApply: (a: SessionState["lastApply"]) => void;

  setMode: (mode: Mode) => void;
  setSession: (
    sessionId: string,
    summary: WorkbookSummary,
    report: ValidationReport,
    plan: PrepAction[],
    version: Version
  ) => void;
  setReport: (report: ValidationReport, plan: PrepAction[]) => void;
  setSummary: (summary: WorkbookSummary) => void;
  addVersion: (version: Version) => void;
  setVersions: (versions: Version[]) => void;
  setCurrentVersion: (version: Version) => void;
  setChatHistory: (messages: ChatMessage[]) => void;
  appendChatMessage: (message: ChatMessage) => void;
  setIsUploading: (v: boolean) => void;
  setIsAnalyzing: (v: boolean) => void;
  setLastDelta: (delta: Delta | null) => void;
  openFix: (target: FixTarget) => void;
  closeFix: () => void;
  /** Apply a fresh analysis (summary/report/plan, optional delta, versions, verdict and gauge) in one go. */
  applyAnalysis: (a: {
    summary?: WorkbookSummary;
    report?: ValidationReport;
    plan?: PrepAction[];
    delta?: Delta;
    versions?: Version[];
    readiness?: Readiness | null;
    prep_progress?: PrepProgress | null;
  }) => void;
  resetSession: () => void;
}

export const useStore = create<SessionState>((set) => ({
  sessionId: null,
  mode: "plan",
  summary: null,
  report: null,
  plan: [],
  versions: [],
  currentVersion: null,
  chatHistory: [],
  isUploading: false,
  isAnalyzing: false,
  lastDelta: null,
  fixTarget: null,
  recalcResult: null,
  readiness: null,
  prepProgress: null,
  lastApply: null,
  setLastApply: (lastApply) => set({ lastApply }),
  // A recalculation carries the verdict it produces (a clean one is what turns it green).
  setRecalcResult: (recalcResult) => set((s) => ({ recalcResult, readiness: recalcResult?.readiness ?? s.readiness })),

  setMode: (mode) => set({ mode }),
  setSession: (sessionId, summary, report, plan, version) =>
    set({ sessionId, summary, report, plan, versions: [version], currentVersion: version, lastDelta: null, fixTarget: null, recalcResult: null, readiness: null, prepProgress: null, lastApply: null }),
  setReport: (report, plan) => set({ report, plan }),
  setSummary: (summary) => set({ summary }),
  addVersion: (version) =>
    set((s) => ({ versions: [...s.versions.filter((v) => v.id !== version.id), version], currentVersion: version })),
  setVersions: (versions) => set({ versions }),
  setCurrentVersion: (version) => set({ currentVersion: version }),
  setChatHistory: (chatHistory) => set({ chatHistory }),
  appendChatMessage: (message) =>
    set((s) => ({ chatHistory: [...s.chatHistory, message] })),
  setIsUploading: (isUploading) => set({ isUploading }),
  setIsAnalyzing: (isAnalyzing) => set({ isAnalyzing }),
  setLastDelta: (lastDelta) => set({ lastDelta }),
  openFix: (fixTarget) => set({ fixTarget }),
  closeFix: () => set({ fixTarget: null }),
  applyAnalysis: (a) =>
    set((s) => ({
      summary: a.summary ?? s.summary,
      report: a.report ?? s.report,
      plan: a.plan ?? s.plan,
      lastDelta: a.delta ?? s.lastDelta,
      versions: a.versions ?? s.versions,
      readiness: a.readiness === undefined ? s.readiness : a.readiness,
      prepProgress: a.prep_progress === undefined ? s.prepProgress : a.prep_progress,
    })),
  resetSession: () =>
    set({
      sessionId: null,
      summary: null,
      report: null,
      plan: [],
      versions: [],
      currentVersion: null,
      chatHistory: [],
      isUploading: false,
      isAnalyzing: false,
      lastDelta: null,
      fixTarget: null,
      recalcResult: null,
      readiness: null,
      prepProgress: null,
      lastApply: null,
    }),
}));
