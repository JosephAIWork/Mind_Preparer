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
  change_log?: unknown[];
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
  overall?: string;
  stages?: { id: string; label: string; state: string; detail?: string }[];
  elapsed_s?: number;
  error?: string;
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
  mode?: Mode;
}

export interface PendingUpload {
  pending: true;
  sessionId: string;
  mode: Mode;
  size: unknown;
  version: Version;
  versions: Version[];
  scan: ScanStatus;
}

export interface ApplyResult {
  status: string;
  method?: string;
  output_name?: string;
  applied: number;
  failed: number;
  verified_opens_in_excel?: boolean;
  warnings?: string[];
  message?: string;
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
  result?: { status: string; summary?: string; errors_before?: number; errors_after?: number; cells_fixed?: number } | null;
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
}

export interface AppHealth {
  version: string;
  rules?: number;
  excel?: boolean | string;
  assistant?: boolean | string;
  frontend?: string;
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
