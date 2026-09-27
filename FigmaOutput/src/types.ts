export type Status = "PASS" | "WARNING" | "ERROR" | "REQUIRES_USER_INPUT" | "NOT_SUPPORTED";
export type Evidence = "DOCUMENTED_RULE" | "DETERMINISTIC_FINDING" | "USER_PROVIDED" | "INFERENCE" | "RECOMMENDATION";
export type Mode = "plan" | "prep_loops" | "fix_formulas" | "structure_fix";
/** The rule's own priority (rules/*.yaml). REQUIRED = Mind needs it; the rest is quality. */
export type Priority = "REQUIRED" | "RECOMMENDED" | "INFORMATIONAL";
/** blocking: Mind refuses the file (or computes it wrong) without the change; optional: Mind reads the file as it is. */
export type ActionLevel = "blocking" | "optional";

/** Is a finding one Mind needs fixed? (Same rule as the backend's readiness.level_of.) */
export function findingLevel(f: Finding): ActionLevel | "pass" | "verify" {
  if (f.rule_id === "READY-001") return "verify";
  if (f.status === "PASS") return "pass";
  if ((f.status === "ERROR" || f.status === "REQUIRES_USER_INPUT") && f.priority === "REQUIRED") return "blocking";
  return "optional";
}

export interface Finding {
  rule_id: string;
  status: Status;
  priority?: Priority;
  severity?: "HIGH" | "MEDIUM" | "LOW" | "BLOCKER" | "INFO";
  confidence: "HIGH" | "MEDIUM" | "LOW";
  evidence: Evidence;
  location: { sheet?: string; cell?: string; [k: string]: unknown };
  observed?: unknown;
  expected?: unknown;
  message: string;
  readiness_impact: Status;
  correction_available: boolean;
  source: { document?: string; article?: string };
}

export interface ValidationReport {
  schema_version: string;
  report_id: string;
  status: Status;
  findings: Finding[];
  summary: {
    finding_count: number;
    status_counts: Partial<Record<Status, number>>;
    not_supported_rule_ids: string[];
    blocking_rule_ids: string[];
  };
}

export interface GridFlag { name: string; raw: string; args: string[] }

export interface Grid {
  sheet: string;
  title_cell: string | null;
  title: string | null;
  name: string | null;
  display_name: string;
  flags: GridFlag[];
  flag_names: string[];
  anchor: string;
  ref: string;
  first_row: number;
  last_row: number;
  first_col: number;
  last_col: number;
  n_rows: number;
  n_cols: number;
  header_values: (string | number | boolean | null)[];
  header_is_all_text: boolean;
  inner_title_cells: string[];
  formula_count: number;
}

export interface SheetSummary {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  dimensions: string;
  grids: Grid[];
  standalone_text_cells: { cell: string; text: string }[];
  formula_count: number;
  array_formula_count: number;
  protected: boolean;
}

/** One sheet as read from the package (before any scan): name, visibility, decompressed size of its part. */
export interface SheetPart {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  index: number;
  part: string | null;
  bytes: number;
  compressed_bytes: number;
  /** share of all sheet bytes, 0..1 */
  share: number;
}

/** Upload size gate (1.6.6): container facts read from the zip alone, and the threshold verdict. */
export interface SizeInfo {
  format: string;
  file_bytes: number;
  decompressed_bytes: number;
  sheet_bytes: number;
  model_bytes?: number;
  file_mb: number;
  decompressed_mb: number;
  threshold_bytes: number;
  threshold_mb: number;
  measure: "decompressed";
  above_threshold: boolean;
  /** rough scan duration from the unpacked size (1.6.7) */
  estimated_seconds?: number;
  estimate_text?: string;
  message: string;
  sheets: SheetPart[];
  largest_parts: { part: string; bytes: number }[];
  sheet_list_source: string | null;
  warnings: string[];
}

/** Live scan status (GET /api/sessions/{id}/status). */
export interface ScanStatus {
  state: "idle" | "pending" | "running" | "ready" | "error";
  stage: string | null;
  title: string | null;
  message: string | null;
  /** progress within the current stage, 0..1, or null when the stage cannot say */
  fraction: number | null;
  /** progress across the whole scan, 0..1 */
  overall: number;
  sheet?: string | null;
  rule_id?: string | null;
  elapsed_s: number;
  stages: { stage: string; seconds: number }[];
  error: string | null;
  version_id?: string;
  needs_convert?: boolean;
  ignore_sheets?: string[];
}

export interface WorkbookSummary {
  file_name: string;
  file_type: "xlsx" | "xlsm";
  sha256: string;
  application: string | null;
  app_version: string | null;
  has_vba: boolean;
  sheet_count: number;
  /** 1.6.6: sheets the user chose to skip (present in the file, not scanned, checked by no rule) */
  ignored_sheet_count?: number;
  total_sheet_count?: number;
  ignored_sheets?: { name: string; state: string }[];
  size?: SizeInfo | null;
  hidden_sheet_count: number;
  grid_count: number;
  flagged_grid_count: number;
  formula_count: number;
  array_formula_count: number;
  defined_name_count: number;
  external_link_part_count: number;
  mm_functions_used: Record<string, number>;
  native_function_usage: Record<string, number>;
  sheets: SheetSummary[];
}

export type OpKind =
  | "set_value"
  | "set_formula"
  | "clear_cell"
  | "set_array_formula"
  | "rename_sheet"
  | "insert_row"
  | "insert_column"
  | "set_sheet_visibility"
  | "unprotect_sheet"
  | "explicit_colors";

export interface Operation {
  op: OpKind;
  action_id: string;
  rule_id: string;
  sheet: string;
  cell?: string;
  row?: number;
  column?: string;
  cells?: string[];
  range?: string;
  before: unknown;
  after: unknown;
  note?: string;
  after_inserts?: boolean;
}

export interface PrepAction {
  id: string;
  title: string;
  rule_ids: string[];
  default_on: boolean;
  /** 1.7.2: blocking (Mind needs it) or optional (Mind reads the file as it is) */
  level: ActionLevel;
  /** why an optional action still matters (e.g. untitled grids show as "Untitled" in Mind) */
  level_note?: string | null;
  /** why a blocking action deserves a look before Apply (e.g. it rewrites formulas) */
  caution?: string | null;
  operations: Operation[];
  count: number;
  skipped: string[];
}

/** One thing that stands between the workbook and Mind (readiness.blocking[]). */
export interface BlockingItem {
  rule_id: string;
  status: Status;
  message: string;
  location: { sheet?: string; cell?: string; [k: string]: unknown };
  /** prep: a Prep action holds operations for it; prep_skipped: the action exists but left every site for review; assistant: no automatic fix */
  fix: "prep" | "prep_skipped" | "assistant";
  action_id: string | null;
  sites: number | null;
}

/** The one verdict (1.7.2): is this version acceptable by Mind, and if not, what stands in the way. */
export interface Readiness {
  state: "blocked" | "unverified" | "ready";
  headline: string;
  version_id: string;
  blocking: BlockingItem[];
  blocking_count: number;
  manual_blocking_count: number;
  optional_count: number;
  prep: { blocking_ops: number; optional_ops: number };
  /** new_errors: error cells that were NOT errors in the original upload (the preparation's doing); preexisting_errors: the model's own */
  recalc: { ran: boolean; clean: boolean; version_id: string | null; formula_errors: number; new_errors: number; preexisting_errors: number; addin_gap_errors: number; message: string | null };
  next_step: { screen: "prep" | "findings" | "recalculate" | "reports"; label: string; detail: string | null; rule_id?: string };
}

/** The Prep gauge: how much prep work each analysed version still needed. */
export interface PrepProgressEntry {
  version_id: string;
  blocking_ops: number;
  optional_ops: number;
  total_ops: number;
  blocking_findings: number;
  by_action: Record<string, number>;
}
export interface PrepProgress {
  entries: PrepProgressEntry[];
  /** three analyses in a row with work planned and no net decrease: Prep cannot resolve what is left */
  stalled: boolean;
  repeating_actions: string[];
  message: string | null;
}

export interface ApplyResult {
  status: "APPLIED" | "PARTIAL" | "ERROR" | "NOT_APPLICABLE";
  method: "excel_com" | "openpyxl";
  output_path: string;
  output_name: string;
  applied: Operation[];
  failed: (Operation & { error: string })[];
  verified_opens_in_excel: boolean | null;
  warnings: string[];
  message: string;
}

export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  note?: boolean;
  provenance?: { context_chars: number; detail_chars: number; lookups: number };
}

export interface Proposal {
  summary: string | null;
  operations: Operation[];
  errors: string[];
}

export interface ChatReply {
  text: string;
  proposal: Proposal | null;
  provenance: { context_chars: number; detail_chars: number; lookups: number };
}

export interface RecalcCell { sheet: string; cell: string; formula: string; error: string; array?: string; /** 1.7.2: this cell was already an error in the original upload */ preexisting?: boolean }

/** Recalculation errors that share one root cause (same unknown function, same formula shape, same missing add-in function). */
export interface RecalcGroup {
  id: string;
  error: string;
  cause: string;
  kind: "formula" | "addin_gap";
  signature: string;
  cells: RecalcCell[];
  count: number;
}

export interface RecalcResult {
  status: Status;
  message: string;
  /** Version the recalculation ran on (set by the backend). */
  version_id?: string;
  /** false when Excel could not run at all (COM failure / not available) -- then the error lists mean nothing. */
  ran?: boolean;
  formula_errors: RecalcCell[];
  addin_gap_errors: { sheet: string; cell: string; formula: string }[];
  groups?: RecalcGroup[];
  /** 1.7.2: errors split against the original upload -- the model's own vs the preparation's */
  preexisting_errors?: number;
  new_errors?: number;
  compared_with_original?: boolean;
  /** 1.7.2: the verdict after this recalculation (a clean one turns it green) */
  readiness?: Readiness | null;
}

export interface ReportBuild {
  standalone_name: string;
  workbook_name: string;
  method: "excel_com" | "openpyxl";
  verified_opens_in_excel: boolean | null;
  warnings: string[];
}

export interface Version {
  id: string;
  label: string;
  file_name: string;
  sha256: string;
  created_at: string;
  source: "upload" | "prep" | "assistant" | "formula" | "convert" | "grid_namer";
  change_log: Operation[];
  verified_opens_in_excel: boolean | null;
}

/** What changed between the previous analysis and this one (same session). */
export interface DeltaEntry { rule_id: string; from: Status; to: Status; note?: string }
export interface Delta {
  version_id: string;
  previous_version_id: string | null;
  fixed: DeltaEntry[];
  improved: DeltaEntry[];
  regressed: DeltaEntry[];
  previous_counts: Partial<Record<Status, number>>;
  counts: Partial<Record<Status, number>>;
}

/** What the Fix panel is opened for: a rule finding, a recalculation error at a cell, or a group of errors with one root cause. */
export interface WindowCell { ref: string; value: unknown; formula: string | null; array: string | null; error: boolean; focus: boolean }
export interface CellWindow { sheet: string; focus: string; range?: string; columns: string[]; rows: { row: number; cells: WindowCell[] }[]; values_from?: string; error?: string }

export interface FixTarget {
  kind?: "finding" | "recalc" | "recalc-group";
  rule_id: string;
  sheet?: string;
  cell?: string;
  grid?: string;
  /** recalc only: the error Excel produced (#NAME?, #REF!, ...) and the formula it evaluated */
  error?: string;
  formula?: string;
  /** recalc only: the cell only fails because the MMForExcel add-in is missing here */
  addin_gap?: boolean;
  /** recalc-group: every cell sharing the cause; recalc: the other cells sharing this cell's cause */
  cells?: RecalcCell[];
  cause?: string;
  group_id?: string; array?: string;
}
