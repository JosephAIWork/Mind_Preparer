/** Types shared with the FastAPI backend (app/web/server.py). Subset of FigmaOutput/src/types.ts. */

export type Status = "PASS" | "WARNING" | "ERROR" | "REQUIRES_USER_INPUT" | "NOT_SUPPORTED";
export type Mode = "plan" | "prep_loops" | "fix_formulas" | "structure_fix";

export interface Finding {
  rule_id: string;
  rule_name?: string;
  category?: string;
  status: Status;
  severity?: string;
  message: string;
  evidence?: string;
  location?: { sheet?: string; cell?: string; ref?: string } | Record<string, unknown>;
  observed?: unknown;
  fix?: string;
  action_id?: string | null;
}

export interface ValidationReport {
  status: Status;
  findings: Finding[];
  counts?: Record<string, number>;
  [k: string]: unknown;
}

export interface WorkbookSummary {
  file_name: string;
  file_type: string;
  sha256?: string;
  sheet_count: number;
  grid_count: number;
  formula_count: number;
  has_vba?: boolean;
  mm_functions_used?: Record<string, number>;
  sheets?: SheetSummary[];
  [k: string]: unknown;
}

export interface Operation {
  op: string;
  sheet: string;
  cell?: string;
  action_id?: string;
  rule_id?: string;
  [k: string]: unknown;
}

export interface PrepAction {
  id: string;
  title: string;
  rule_ids: string[];
  default_on: boolean;
  level: "blocking" | "optional";
  level_note?: string;
  caution?: string;
  operations: Operation[];
  count: number;
  skipped?: unknown[];
}

export interface Version {
  id: string;
  parent?: string | null;
  major: number;
  minor: number;
  label: string;
  file_name: string;
  sha256?: string;
  created_at?: string;
  source: string;
  change_log?: { op?: string; sheet?: string; cell?: string; before?: unknown; after?: unknown; note?: string; action_id?: string; rule_id?: string }[];
  verified_opens_in_excel?: boolean;
}

export interface ReadinessBlocking {
  rule_id: string;
  status: Status;
  message: string;
  location?: Record<string, unknown>;
  fix: "prep" | "prep_skipped" | "assistant";
  action_id?: string | null;
  sites?: unknown;
}

export interface Readiness {
  state: "blocked" | "unverified" | "ready";
  headline: string;
  version_id?: string;
  blocking: ReadinessBlocking[];
  blocking_count: number;
  manual_blocking_count: number;
  optional_count: number;
  prep?: { blocking_ops: number; optional_ops: number };
  recalc?: { ran: boolean; clean: boolean; version_id?: string; formula_errors?: number; new_errors?: number; preexisting_errors?: number; addin_gap_errors?: number; message?: string } | null;
  next_step?: { screen: string; label: string; detail?: string; rule_id?: string } | null;
}

export interface PrepProgress {
  entries?: unknown[];
  applies?: unknown[];
  written?: Record<string, number>;
  stalled?: boolean;
  message?: string;
}

export interface ScanStatus {
  state: "pending" | "idle" | "running" | "ready" | "error";
  /** the stage running now (app/progress.py: convert, copy, load, inventory, names, rules, report, plan) */
  stage?: string | null;
  title?: string | null;
  message?: string | null;
  fraction?: number | null;
  /** 0..1 over the whole scan */
  overall?: number;
  /** the stages already finished, with their seconds */
  stages?: { stage: string; seconds: number }[];
  elapsed_s?: number;
  error?: string | null;
  apply?: { state: string; stage?: string; title?: string; message?: string; fraction?: number; done?: number; total?: number };
}

export interface SessionStart {
  sessionId: string;
  summary: WorkbookSummary;
  report: ValidationReport;
  plan: PrepAction[];
  version: Version;
  versions?: Version[];
  size?: unknown;
  scan?: ScanStatus;
  readiness?: Readiness;
  prep_progress?: PrepProgress;
  delta?: Delta | null;
  mode?: Mode;
  /** 1.11.7: what Prep (and Do everything) does with the pivot tables, and the choices */
  pivot_mode?: PivotMode;
  pivot_modes?: Record<PivotMode, string>;
}

export type PivotMode = "formulas" | "values";

export interface PendingUpload {
  pending: true;
  sessionId: string;
  mode: Mode;
  size: SizeInfo;
  version: Version;
  versions: Version[];
  scan: ScanStatus;
}

export interface ApplyResult {
  status: string;
  method?: string;
  output_name?: string;
  /** The operations written (the backend sends the list; some routes send a count). */
  applied: Operation[] | number;
  /** The operations Excel refused, each with its `error`. */
  failed: (Operation & { error?: string })[] | number;
  verified_opens_in_excel?: boolean | null;
  warnings?: string[];
  message?: string;
}

export function appliedCount(r: ApplyResult | null | undefined): number {
  const a = r?.applied;
  return Array.isArray(a) ? a.length : Number(a) || 0;
}

export function failedCount(r: ApplyResult | null | undefined): number {
  const f = r?.failed;
  return Array.isArray(f) ? f.length : Number(f) || 0;
}

/** The first refusals, as "Sheet!A1: reason". */
export function failedText(r: ApplyResult | null | undefined, max = 3): string {
  const f = r?.failed;
  if (!Array.isArray(f) || !f.length) return "";
  const parts = f.slice(0, max).map((o) => `${o.sheet}!${o["cell"] ?? ""}${o.error ? ": " + String(o.error).slice(0, 80) : ""}`);
  return parts.join("; ") + (f.length > max ? `; … ${f.length - max} more` : "");
}

export interface ApplyOutcome {
  result: ApplyResult;
  version: Version;
  summary?: WorkbookSummary;
  report?: ValidationReport;
  plan?: PrepAction[];
  versions?: Version[];
  readiness?: Readiness;
  prep_progress?: PrepProgress;
}

export interface AutoFixStatus {
  state: "idle" | "running" | "done" | "error";
  stage?: string;
  title?: string;
  message?: string;
  fraction?: number;
  pass_no?: number;
  errors?: number;
  errors_before?: number;
  kept?: number;
  error?: string | null;
  elapsed_s?: number;
  version?: Version | null;
  result?: { status: string; summary?: string; errors_before?: number; errors_after?: number; cells_fixed?: number; left?: AutoFixLeft[]; regression_cells?: number; stopped?: string | null; numbers_check?: { ran?: boolean } | null } | null;
  analysis?: Partial<SessionStart> | null;
}

export interface RecalcResult {
  status: Status;
  message: string;
  version_id?: string;
  ran: boolean;
  formula_errors: { sheet: string; cell: string; error: string; formula?: string }[];
  addin_gap_errors: { sheet: string; cell: string; error: string; formula?: string }[];
  new_errors?: number;
  preexisting_errors?: number;
  readiness?: Readiness;
  report?: ValidationReport;
  plan?: PrepAction[];
  version?: Version;
}

export interface LoopPlanLoop {
  name: string;
  driver?: string;
  driver_name?: string;
  label?: string;
  macro?: string;
  source?: string;
  first?: number;
  last?: number;
  count?: number;
  cells: unknown[];
  notes?: string[];
  status: "ready" | "needs_decision" | "not_built";
  kind?: string;
  why?: string;
}

export interface LoopPlanOutput {
  target?: string;
  from?: string;
  written_by?: string;
  when?: string;
  status: "ready" | "not_built" | "nothing_to_do" | string;
  cells?: number;
  items?: number;
  example?: string;
  loop?: string;
  why?: string;
}

export interface LoopPlanClear {
  target?: string;
  resolved?: string;
  cells?: number;
  included?: boolean;
  why?: string;
}

export interface LoopPlanOperation {
  op?: string;
  sheet: string;
  cell?: string;
  before?: unknown;
  after?: unknown;
  note?: string;
}

export interface LoopPlan {
  found: boolean;
  loops: LoopPlanLoop[];
  outputs: LoopPlanOutput[];
  clears: LoopPlanClear[];
  operations?: LoopPlanOperation[];
  operation_count: number;
  notes?: string[];
  counts?: { loop_cells?: number; result_cells?: number; cleared?: number; extra?: number };
  how_it_runs?: string;
}

export interface LoopChoices {
  name?: string;
  first?: string;
  last?: string;
  count_cell?: string;
  clear_old?: boolean;
  /** the old X2:X6 ranges instead of MM_COLUMN / MM_ROW that follow the table */
  fixed_ranges?: boolean;
}

export interface AppHealth {
  version: string;
  rules?: number;
  excel?: boolean | string;
  assistant?: boolean | string;
  frontend?: string;
}

/** GET/POST /api/assistant-key: never the key itself. `message`: saved but not checked (gateway unreachable). */
export interface AssistantKey {
  saved: boolean;
  last4: string | null;
  message?: string | null;
}

export function isPending(x: SessionStart | PendingUpload): x is PendingUpload {
  return (x as PendingUpload).pending === true;
}

/** The Prep actions that rewrite formulas (FRM-*, REF-*): "Fix the formulas" applies exactly these. */
export const FORMULA_ACTION_IDS = new Set([
  "fix_broken_refs",
  "fix_broken_refs_whole",
  "expand_3d_references",
  "freeze_external_links",
  "direct_references",
  "freeze_spill_refs",
  "guard_blank_arithmetic",
  "freeze_environment_formulas",
  "line_breaks_as_char",
  "keep_error_cells",
  "empty_cells_as_zero",
]);

// ---- chunks 3 + 4 --------------------------------------------------------------------------
export interface MapCell {
  cell: string | null;
  name?: string | null;
  label?: string | null;
  value?: unknown;
  evidence?: string;
  how?: string;
}

export interface ModelMap {
  workbook: string;
  has_macros: boolean;
  notes: string[];
  sheets: { name: string; role: string; dims: string | null; state: string | null; grids: number }[];
  loop: {
    found: boolean;
    how_to_run: string | null;
    drivers: (MapCell & { macro: string; loop_var: string; set_from?: string | null; set_from_name?: string | null; nested_depth: number })[];
    bounds: (MapCell & { role: "first" | "last"; formula?: string | null; macro: string })[];
    item_cells: (MapCell & { macro: string; template_sheet: string | null; template_present: boolean })[];
    key_cells: (MapCell & { readers: number; cells: number })[];
    entry_points: string[];
    macros: { procedure: string; module: string | null; var: string; from: string; to: string; depth: number; drivers: string[]; outputs: number }[];
  };
  inputs: {
    per_item: { cell: string; what: string; macro: string }[];
    sheets: { sheet: string; formulas: number; state: string | null; why: string }[];
    tables: { name: string; cell: string; readers: number; role: string }[];
  };
  run_settings: {
    switches: (MapCell & { options: string[]; how: string })[];
    named_settings: (MapCell & { readers: number; switch_like: boolean })[];
    labelled: { sheet: string; cell: string; label: string; value: unknown }[];
    status: (MapCell & { written_in: string[] })[];
    tables: { name: string; cell: string; readers: number; role: string }[];
  };
  outputs: {
    cells: (MapCell & { from: string | null; from_name: string | null; written_by: string; when: string })[];
    sheets: { sheet: string; formulas: number; why: string }[];
    saves: { in: string; path_from: string | null; name: string | null }[];
  };
  macros: { procedures: { name: string; module: string | null; statements: number; calls: string[] }[]; entry_points: string[] };
}

export interface WindowCell {
  ref: string;
  value: unknown;
  formula: string | null;
  array: string | null;
  error: boolean;
  focus: boolean;
}

export interface CellWindow {
  sheet: string;
  focus: string;
  range?: string;
  columns: string[];
  rows: { row: number; cells: WindowCell[] }[];
  values_from?: string;
  error?: string;
}

export interface RecalcCell {
  sheet: string;
  cell: string;
  formula: string;
  error: string;
  array?: string;
  preexisting?: boolean;
}

/** What the assistant is asked about: a finding, or one recalculation error. */
export interface FixTarget {
  kind?: "finding" | "recalc" | "recalc-group";
  rule_id: string;
  sheet?: string;
  cell?: string;
  error?: string;
  formula?: string;
  addin_gap?: boolean;
  cells?: RecalcCell[];
  cause?: string;
}

export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  note?: boolean;
  tone?: "success" | "warning";
}

export interface Proposal {
  summary: string | null;
  operations: Operation[];
  errors: string[];
}

export interface ChatReply {
  text: string;
  proposal: Proposal | null;
  provenance?: { context_chars: number; detail_chars: number; lookups: number };
}

export type ChatStreamEvent = { type: "phase"; phase: string; n?: number } | { type: "delta"; text: string };

export interface MindLoopOptions {
  maxIterations?: number;
  useAssistant?: boolean;
  runModel?: boolean;
  checkNumbers?: boolean;
  skipMind?: boolean;
  enable?: string[];
  disable?: string[];
}

export interface MindLoopEvent {
  t: number;
  event: string;
  message?: string;
  n?: number;
  verdict?: string;
}

export interface MindLoopIteration {
  n: number;
  enabled: string[];
  disabled: string[];
  actions: { id: string; count: number }[];
  apply?: { status: string; applied: number; failed: number; verified: boolean | null; message: string; output: string };
  names?: { grids: number; baseline_grids: number; fragmented: boolean; unnamed: number; ok: boolean };
  numbers?: { ran: boolean; match?: boolean; compared?: number; differences?: number; message?: string };
  mind?: {
    skipped: boolean;
    reason?: string;
    error?: string;
    convert?: { success: boolean; steps: Record<string, string>; errors: string[] };
    counts?: { mind_untitled: number | null; app_untitled: number };
    run?: { completed: boolean; audit_consistent: boolean; seconds: number | null };
  };
  decision?: { verdict: "converged" | "retry" | "stuck"; reason: string; enable: string[]; disable: string[] };
}

export interface MindLoopReport {
  source: string;
  started: string;
  finished?: string;
  iterations: MindLoopIteration[];
  projects: { name: string; status: string; ran: boolean; delete?: string }[];
  verdict: "converged" | "stuck" | "exhausted" | null;
  reason: string | null;
  final_workbook?: string;
  leftover_projects?: string[];
}

export interface MindLoopStatus {
  state: "idle" | "running" | "done" | "error";
  events: MindLoopEvent[];
  next: number;
  report: MindLoopReport | null;
  error?: string | null;
  summary?: string | null;
  work_dir?: string;
}

export interface NumbersCheck {
  ran: boolean;
  clean: boolean | null;
  verdict: string;
  version_id?: string;
  counts: Record<string, number>;
  by_sheet?: Record<string, Record<string, number>>;
  samples?: Record<string, { original: string; now: string; was: string; is?: string }[]>;
  moves?: { row_inserts: number; column_inserts: number; sheet_renames: number };
  regressions?: number;
}

export interface ReportBuild {
  standalone_name: string;
  workbook_name: string;
  method: "excel_com" | "openpyxl";
  verified_opens_in_excel: boolean | null;
  warnings: string[];
}

export interface SheetPart {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  index: number;
  bytes: number;
  compressed_bytes: number;
  share: number;
}

export interface SizeInfo {
  format: string;
  file_bytes: number;
  decompressed_bytes: number;
  file_mb: number;
  decompressed_mb: number;
  threshold_mb: number;
  above_threshold: boolean;
  estimate_text?: string;
  message: string;
  sheets: SheetPart[];
  warnings: string[];
}

// ---- Grid Namer ------------------------------------------------------------------------------
export interface GridFlag {
  raw: string;
  name: string;
  args: string[];
}

export interface Grid {
  sheet: string;
  title_cell: string | null;
  title: string | null;
  name: string | null;
  display_name: string;
  flags: GridFlag[];
  flag_names?: string[];
  anchor: string;
  ref: string;
  first_row: number;
  last_row: number;
  first_col: number;
  last_col: number;
  n_rows: number;
  n_cols: number;
  header_values: unknown[];
  formula_count?: number;
}

export interface SheetSummary {
  name: string;
  state: "visible" | "hidden" | "veryHidden" | string;
  dimensions?: string | null;
  grids: Grid[];
}

export interface SheetCells {
  sheet: string;
  rows: unknown[][];
  n_rows: number;
  n_cols: number;
  total_rows: number;
  total_cols: number;
  truncated: boolean;
  values_from?: string;
}

export interface MindFlag {
  name: string;
  meaning: string;
}

export interface NamedArea {
  sheet: string;
  ref: string;
  name: string;
  flags: string[];
}

export interface GridNamerOutcome {
  result: ApplyResult;
  version: Version;
  summary?: WorkbookSummary;
  report?: ValidationReport;
  plan?: PrepAction[];
  versions?: Version[];
  readiness?: Readiness;
  prep_progress?: PrepProgress;
  named: string[];
  skipped: string[];
  manual_ops: number;
  auto_ops: number;
}

export interface GridNameSuggestion {
  grid: string;
  sheet: string;
  ref: string;
  size: string;
  current: string | null;
  deterministic: string;
  deterministic_source: string;
  suggested: string | null;
}

export interface GridNamesResult {
  available: boolean;
  names: GridNameSuggestion[];
  message: string | null;
  applied: boolean;
  plan?: PrepAction[] | null;
  readiness?: Readiness | null;
}

// ---- final chunk ------------------------------------------------------------------------------
export interface DeltaEntry {
  rule_id: string;
  from: Status;
  to: Status;
  note?: string;
}

export interface Delta {
  version_id: string;
  previous_version_id: string | null;
  fixed: DeltaEntry[];
  improved: DeltaEntry[];
  regressed: DeltaEntry[];
  previous_counts: Partial<Record<Status, number>>;
  counts: Partial<Record<Status, number>>;
}

export interface AutoFixLeft {
  sheet: string;
  cell: string;
  error: string;
  formula: string;
  count: number;
  why: string;
  tried: boolean;
}

/** 1.11.6: GET .../estimate -- how long "Do everything" may take, per step (seconds). */
export interface ChainEstimateStep {
  id: "fixAll" | "prepare" | "loop" | "recalc";
  low_s: number;
  mid_s: number;
  high_s: number;
  why: string;
}
export interface ChainEstimate {
  steps: ChainEstimateStep[];
  total: { low_s: number; mid_s: number; high_s: number };
  caveat: string;
  basis: { file_mb: number; scan_s: number; scan_measured: boolean; stored_errors: number | null; blocking_ops: number };
}
