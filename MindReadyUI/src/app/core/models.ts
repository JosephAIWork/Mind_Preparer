/** Types shared with the FastAPI backend (app/web/server.py). Subset of FigmaOutput/src/types.ts. */

export type Status = "PASS" | "WARNING" | "ERROR" | "REQUIRES_USER_INPUT" | "NOT_SUPPORTED";
export type Mode = "plan" | "prep_loops" | "fix_formulas" | "structure_fix";

export interface Finding {
  rule_id: string;
  /** the rule's plain name ("an empty cell is compared with a number"), shown next to its code */
  rule_name?: string | null;
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
  /** each rule's plain name, in the order of `rule_ids` */
  rule_names?: (string | null)[];
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
  rule_name?: string | null;
  status: Status;
  message: string;
  location?: Record<string, unknown>;
  fix: "prep" | "prep_skipped" | "assistant" | "outside_inputs";  // outside_inputs: the "Outside files" card brings them in
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
  next_step?: { screen: string; label: string; detail?: string; rule_id?: string; rule_name?: string | null } | null;
}

/** GET /api/plain-codes: the plain words for every code the screens show -- each rule's name
 *  (rules/plain_names.yaml) and what each of Excel's error values means (app/plain_codes.py). */
export interface PlainCodesGlossary {
  rules: Record<string, string>;
  errors: Record<string, string>;
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
  /** the outside files the workbook's macros open (null: they open none); from the scan of the upload */
  outsideInputs?: OutsideInputs | null;
  /** the upload's model workbook and the input files picked with it (several files picked at once) */
  main?: string;
  inputs?: InputFile[];
  /** 1.14.0: the last recalculation of the current version, when a saved workbook is opened again */
  recalc?: RecalcResult | null;
}

/** 1.14.0: one line of the main screen's "Recent workbooks" list (GET /api/saved-workbooks): a session
 *  kept in the data folder, to open again after a restart of the app. */
export interface SavedWorkbook {
  id: string;
  name: string;
  /** seconds since 1970, as the server keeps them */
  uploaded_at: number | null;
  last_opened: number | null;
  versions: number;
  current_version: string | null;
  bytes: number;
  has_result: boolean;
  pending: boolean;
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
  main?: string;
  inputs?: InputFile[];
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

/** Room for the reason of one refused change: the server says it as a plain sentence with Excel's own
 *  words in brackets ('Excel refused the action (a protected sheet, ...) (Excel said: "...")'), up to
 *  about 170 characters for the common refusals -- all of it is shown. */
const REASON_MAX = 200;

/** `text` kept to about `max` characters, cut after a whole word with "…", never in the middle of one. */
export function clipAtWord(text: string, max: number): string {
  const t = (text ?? "").replace(/\s+/g, " ").trim();
  if (t.length <= max) return t;
  const cut = t.slice(0, max);
  const space = cut.lastIndexOf(" ");
  return (space > max / 2 ? cut.slice(0, space) : cut).replace(/[\s,;:(]+$/, "") + "…";
}

/** The first refusals, as "Sheet!A1: reason". Refusals with the same reason say it once ("S!B5, S!B6:
 *  reason": a protected sheet refuses every change the same way), and a longer reason is cut after a
 *  whole word. */
export function failedText(r: ApplyResult | null | undefined, max = 3): string {
  const f = r?.failed;
  if (!Array.isArray(f) || !f.length) return "";
  const groups: { cells: string[]; reason: string }[] = [];
  for (const o of f.slice(0, max)) {
    const cell = `${o.sheet}!${o["cell"] ?? ""}`;
    const reason = o.error ? clipAtWord(String(o.error), REASON_MAX) : "";
    const same = groups.find((g) => g.reason === reason);
    if (same) same.cells.push(cell);
    else groups.push({ cells: [cell], reason });
  }
  const parts = groups.map((g) => g.cells.join(", ") + (g.reason ? ": " + g.reason : ""));
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
  /** an empty cell the plan offers for the loop's count, when it needs one */
  suggested_count_cell?: string | null;
  /** the choices this loop was built with */
  chosen?: LoopChoice;
  /** a loop of another macro: listed, not written (one macro at a time) */
  other_macro?: boolean;
  /** the macro a person runs for this loop */
  run?: string | null;
}

/** One loop's own choices; first / last are a number or a cell holding one (Sheet!A1). */
export interface LoopChoice {
  name?: string;
  first?: string;
  last?: string;
  count_cell?: string;
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
  /** Prep's blocking items that are not formulas (a lone text cell, a pivot table), still open for this version */
  prep_blockers?: { id: string; title: string; count: number; rule_ids?: string[] }[];
  counts?: { loop_cells?: number; result_cells?: number; cleared?: number; extra?: number; repairs?: number };
  how_it_runs?: string;
  /** the formula repairs Prep still had for this version: written with the loop (written = false: only named) */
  repairs?: { id: string; title: string; count: number; written: boolean }[];
  /** the macro whose loops this plan builds */
  macro?: string | null;
  /** every macro with loops, when there are two or more: Excel runs each on its own, the plan builds one at a time */
  macros?: { name: string; loops: string[]; chosen: boolean }[];
}

export interface LoopChoices {
  name?: string;
  first?: string;
  last?: string;
  count_cell?: string;
  clear_old?: boolean;
  /** the old X2:X6 ranges instead of MM_COLUMN / MM_ROW that follow the table */
  fixed_ranges?: boolean;
  /** each loop's own choices, by driver cell */
  loops?: Record<string, LoopChoice>;
  /** the macro whose loops are built (default: the first macro with a loop) */
  macro?: string;
}

// ---- Mind instances (app/instance_builder.py) ------------------------------------------------
export type InstanceStatus = "ready" | "needs_decision" | "not_built";

/** The run loop every copy shares: MM_LOOPINSTANCE in the driver, COUNTA in the count cell, the biggest count in the max cell. */
export interface InstancePlanLoop {
  name?: string;
  driver?: string;
  count_cell?: string;
  max_cell?: string;
  formula?: string;
  macro?: string;
  kind?: string;
}

/** The model-point table the loop reads through (INDEX / OFFSET on the driver). */
export interface InstancePlanTable {
  sheet?: string;
  ref?: string;
  header_row?: number;
  first_row?: number;
  last_row?: number;
  rows?: number;
  columns?: number;
}

/** A text column of the table that could be the key: its letter, header and number of distinct values. */
export interface InstanceKeyCandidate {
  column: string;
  header: string;
  distinct: number;
  /** true when the count comes from the first rows of a long table only ("10+ values") */
  sampled?: boolean;
}

/** The macro's paste area: emptied, because the copies' results live in the block. */
export interface InstancePlanClear {
  area: string;
  cells: number;
  status: string;
  over?: number;
  why?: string | null;
}

/** The key column: one copy of the model per distinct value. */
export interface InstancePlanKey {
  column?: string;
  header?: string;
  distinct?: number;
  /** key value -> number of rows */
  counts?: Record<string, number>;
  candidates?: InstanceKeyCandidate[];
  chosen_by?: "macro" | "heuristic" | "user";
}

/** The table as a split grid: the title "#<Name> /InstanceSplit" above it, the key column first. */
export interface InstancePlanSplitGrid {
  title_cell?: string;
  title?: string;
  header_cell?: string;
  inserted_rows?: number;
  inserted_column?: boolean | string | null;
}

/** The block the macro copies out on every pass: every formula cell becomes MM_RESULT(...). */
export interface InstancePlanOutput {
  block?: string;
  source_sheet?: string;
  cells?: number;
  example?: string;
  status: InstanceStatus | "nothing_to_do" | string;
  why?: string;
}

/** One operation as the backend shows it (a set_range_values op comes with at most 20 values and `values_count`). */
export interface InstancePlanOperation extends LoopPlanOperation {
  ref?: string;
  row?: number;
  column?: string;
  values?: unknown[][];
  values_count?: number;
}

export interface InstancePlanMainGrid {
  title: string;
  rows?: number;
  cols?: number;
  what?: string;
  /** set on the per-copy grids entry: how many grids of that shape (one per key) */
  count?: number;
}

/** The second, small workbook: the keys and the results, saved next to the model. */
export interface InstancePlanMain {
  file: string;
  model_file: string;
  keys: (string | number)[];
  grids: InstancePlanMainGrid[];
}

/** A loop of the Mind loop plan, and whether copies can be made from its values (a counter, or a list). */
export interface InstanceLoopOption {
  name: string;
  driver?: string;
  kind?: string | null;
  count?: number | null;
  first?: number | null;
  last?: number | null;
  macro?: string | null;
  usable: boolean;
  why?: string | null;
}

/** A loop the copies are made from: its driver now reads the copy's value from the split grid. */
export interface InstanceCopiesFrom {
  name: string;
  driver?: string;
  kind?: string;
  values?: unknown[];
  count?: number;
  formula?: string;
}

/** GET .../mind-instances-plan: the macro's run loop as Mind instances, one copy of the model per key value. */
export interface InstancePlan {
  found: boolean;
  status: InstanceStatus;
  /** "table": one copy per key of the model-point table; "loops": one copy per combination of the chosen loops' values */
  mode?: "table" | "loops";
  /** every loop of the Mind loop plan, usable or not for copies */
  loop_options?: InstanceLoopOption[];
  copies_from?: InstanceCopiesFrom[];
  /** the macro's other loops: they stay MM_LOOPs inside every copy */
  kept_loops?: { name: string; driver?: string; count?: number; written?: boolean; already?: boolean }[];
  /** the block the main workbook shows for every copy */
  results?: { block: string; rows: number; cols: number; chosen_by: string } | null;
  notes?: string[];
  /** Prep's blocking items that are not formulas, still open for this version (as on the loop plan) */
  prep_blockers?: { id: string; title: string; count: number; rule_ids?: string[] }[];
  loop?: InstancePlanLoop | null;
  table?: InstancePlanTable | null;
  key?: InstancePlanKey | null;
  split_grid?: InstancePlanSplitGrid | null;
  outputs?: InstancePlanOutput[];
  clears?: InstancePlanClear[];
  operations?: InstancePlanOperation[];
  operation_count: number;
  counts?: { key_values?: number; result_cells?: number; cleared?: number; inserted_rows?: number; repairs?: number };
  /** the formula repairs Prep still had for this version: written with the instances, as with the Mind loop */
  repairs?: { id: string; title: string; count: number; written: boolean }[];
  main?: InstancePlanMain | null;
  how_it_runs?: string;
}

/** The user's choices: the key column (a header text or a column letter) and which keys the main workbook lists. */
export interface InstanceChoices {
  key?: string;
  /** "present" = only the key values that have rows (default); "all" = every value of the list the macro reads */
  keys?: "present" | "all";
  /** the loops the copies are made from, comma separated (names); empty = by the key column */
  loops?: string;
  /** each loop's own choices, as set on the Mind loop panel (by driver cell) */
  loop_choices?: Record<string, LoopChoice>;
  /** the macro whose loop the copies by the key column use (the Mind loop panel's choice) */
  macro?: string;
}

/** POST .../mind-instances-plan/apply: the Apply outcome, the plan (without operations), the zip's URL and the two file names. */
/** POST .../mind-instances-plan/apply: `download` is null when the zip was not made (a change refused, the main workbook failed); `warnings` say why. */
export type InstancesOutcome = Omit<ApplyOutcome, "plan"> & { plan: InstancePlan; download: string | null; files: string[]; warnings?: string[] };

// ---- Where instances could help (app/instance_finder.py) ---------------------------------------
export type FinderFits = "yes" | "partly" | "no";

/** One line of evidence: what was seen, where, how many. */
export interface FinderEvidence {
  kind: "macro" | "lookups" | "branches" | "stepping" | "keyed" | "list" | "bands" | "sheets" | "counts" | string;
  text: string;
  cells?: string[];
  count?: number;
  per_copy?: number;
  vba?: string[];
}

/** The grouping column of a P1 dimension: "Use as the key" fills the builder with `column`. */
export interface FinderKey {
  sheet: string;
  column: string;
  header?: string | null;
  table?: string;
  distinct: number;
  rows: number;
  biggest?: { value: string | number; rows: number };
  smallest?: { value: string | number; rows: number };
  balance?: number;
  /** key value -> number of rows (at most 12 keys, then "…": how many more) */
  counts?: Record<string, number>;
  type?: string;
  keyed_formulas?: number;
  keyed_tables?: string[];
  mapped_through?: { cell: string; name?: string | null; table?: string | null }[];
  branch_formulas?: number;
  branch_literals_count?: number;
  literals_are_values?: number;
}

/** A repeated block: its copies along rows, columns or sheets, and how alike they are. */
export interface FinderBlock {
  sheet: string;
  axis: "rows" | "columns" | "sheets";
  where: string;
  copies: number;
  per_copy: number;
  computed?: string | null;
  diff?: Record<string, number>;
  chain_share?: number;
}

/** One cell the copies disagree on, and what to do with it. */
export interface FinderEdit {
  cell: string;
  values_per_copy?: (string | number | null)[];
  literals_per_copy?: string[];
  formulas_per_copy?: string[];
  action: string;
}

export interface FinderReaderClass {
  count: number;
  example?: string | null;
  text: string;
}

/** Cells that read the macro's paste area, by what they need in Mind. */
export interface FinderReaders {
  all_slots: FinderReaderClass;
  fixed_slot: FinderReaderClass;
  computed_slot: FinderReaderClass;
  keyed_slot: FinderReaderClass;
  follow_adders: FinderReaderClass;
}

export interface FinderFit {
  fits: FinderFits;
  reasons: string[];
}

/** The cells to unify before the copies can be one model (each list capped at 40; the counts are complete). */
export interface FinderEdits {
  numbers: FinderEdit[];
  formula_numbers: FinderEdit[];
  logic: FinderEdit[];
  size: FinderEdit[];
  numbers_count: number;
  formula_numbers_count: number;
  logic_count: number;
  size_count: number;
  branches?: { count: number; action: string } | null;
}

/** One dimension: a thing the model repeats over (a key column, a list, copies of a block or a sheet). */
export interface FinderDimension {
  id: string;
  rank: number;
  score: number;
  label: string;
  label_from?: string;
  headline: string;
  kinds: string[];
  key?: FinderKey | null;
  copies: number;
  biggest_copy?: number;
  formulas_saved: number;
  cells_changed: number;
  evidence: FinderEvidence[];
  blocks?: FinderBlock[];
  readers?: FinderReaders | null;
  cross_reads?: { from: string; to: string; cells: number }[];
  edits?: FinderEdits;
  fit: { instances: FinderFit; loop: FinderFit; leave: FinderFit };
  assistant_rank?: number | null;
  assistant_text?: string | null;
  assistant_edits?: string[];
}

/** Copies set aside: each one reads the one before it, so they cannot run side by side. */
export interface FinderChain {
  label: string;
  copies: number;
  per_copy: number;
  chain_share: number;
  why: string;
  cells?: string[];
}

/** The assistant's view on top of the deterministic result; `asked` is false when the tick was off. */
export interface FinderAssistant {
  asked: boolean;
  available: boolean;
  text: string | null;
  ranking: number[] | null;
  explanations: Record<string, string>;
  edits: Record<string, string[]>;
  extra: { headline: string; why: string; cells: string[]; flag: string }[];
  message: string | null;
}

/** The builder's own key next to the finder's top dimension (`source` "plan" = the builder's real plan, "rule" = the finder's re-run of its rule). */
export interface FinderBuilder {
  column?: string | null;
  header?: string | null;
  chosen_by?: string | null;
  candidates?: string[];
  matches_top?: boolean | null;
  source?: "plan" | "rule" | string;
}

/** GET .../instance-finder: where instances or a loop would help, ranked; nothing written. */
export interface FinderResult {
  schema: string;
  version: string;
  file: string;
  counts: { sheets: number; formulas: number; dimensions: number; chains: number; time_axes: number; seconds: number; reload_seconds?: number };
  dimensions: FinderDimension[];
  chains: FinderChain[];
  macro_loops: { name: string; driver: string; kind: string; status: string; handled_by: string; first?: number | string | null; last?: number | string | null }[];
  switches?: { cell: string; name?: string | null; readers: number }[];
  builder?: FinderBuilder | null;
  ignored?: Record<string, number>;
  ignored_detail?: { filled_areas?: { sheet: string; where: string; cells: number }[] };
  skipped: string[];
  truncated?: Record<string, boolean>;
  assistant: FinderAssistant;
}

export interface AppHealth {
  version: string;
  rules?: number;
  excel?: boolean | string;
  assistant?: boolean | string;
  frontend?: string;
  /** a shared copy's last working day, "2026-12-15" (null: no end date, the developer version) */
  expires?: string | null;
  /** 0 on the last day */
  days_left?: number | null;
  /** past the date: every other call answers 410 with `expired_message` */
  expired?: boolean;
  expired_message?: string | null;
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
  "offset_ranges",
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
  rule_name?: string | null;
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
  /** The changes to make before uploading; [] uploads the version as it is. */
  actions?: string[];
}

export interface MindLoopEvent {
  t: number;
  event: string;
  message?: string;
  n?: number;
  verdict?: string;
  stage?: string;
}

export interface MindRunEstimate {
  stages: { id: string; title: string; seconds: number }[];
  first_s: number;
  retry_s: number;
  caveat: string;
}

export interface MindRunVersion {
  id: string;
  label: string | null;
  file_name: string;
}

export interface MindRunChoice {
  id: string;
  title: string;
  count: number;
  default_on: boolean;
  level?: string;
}

export interface MindRunOptionsInfo {
  version: MindRunVersion;
  newest_version_id: string | null;
  busy: string | null;
  choices: MindRunChoice[];
  estimate: MindRunEstimate;
}

/** The Mind sign-in the app's own browser profile keeps (GET/POST /api/mind/sign-in). */
export interface MindSignIn {
  state: "idle" | "waiting" | "checking" | "done" | "failed";
  message: string | null;
  last_login: string | null;
  checked: { signed_in: boolean; at: string | null; error?: string } | null;
  elapsed_s: number | null;
}

export interface MindRunTiming {
  elapsed_s: number;
  remaining_s: number;
  estimate_s: number | null;
  retry_s: number | null;
  over_estimate: boolean;
  stage: string | null;
  stage_title: string;
  stage_elapsed_s: number;
  iteration: number;
  max_iterations: number | null;
  stages: { id: string; title: string; seconds: number; state: "done" | "now" | "todo" }[];
  last_news_s: number;
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
  version?: MindRunVersion | null;
  actions?: string[] | null;
  timing?: MindRunTiming;
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
  rule_name?: string | null;
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

/** The steps of "Do everything", in order ("outside" only when the model's macros open outside files). */
export type ChainStepId = "outside" | "fixAll" | "prepare" | "loop" | "recalc";

/** 1.11.6: GET .../estimate -- how long "Do everything" may take, per step (seconds). */
export interface ChainEstimateStep {
  id: ChainStepId;
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

/** POST /api/feedback: the email for "Send feedback to the Tel Aviv team" (codes only, nothing from the workbook). */
export interface FeedbackMessage {
  ref: string;
  text: string;
  subject: string;
  email: string;
  file_name: string;
  assistant: boolean;
}

// ---- outside input files: a model whose macros open other files (app/outside_inputs.py) --------------
/** One row of the model's import table (a cell that holds an outside file's path), with the picked file. */
export interface OutsideInputRow {
  /** "Sheet!A1": the cell that holds the path */
  cell: string;
  /** the exact file name the model expects, when Excel's stored value shows it */
  expected: string | null;
  /** the row's key values (e.g. "BBA_01") */
  keys: string[];
  /** the sheet the macro fills from this file */
  target_sheet: string | null;
  /** the picked file that goes with this row (null: missing) */
  file: string | null;
  how: "exact" | "key" | null;
  /** that file is a stand-in the assistant made from the numbers the model holds, not the user's */
  stand_in?: boolean;
  /** the picked files that could fit, when more than one does */
  candidates: string[];
}

/** The outside files the workbook's macros open, matched with the files picked with it
 *  (in the scan answer and GET /api/sessions/{id}; null when the macros open none). */
export interface OutsideInputs {
  imports: boolean;
  macro: string | null;
  calls: { procedure: string; call: string; line: string }[];
  rows: OutsideInputRow[];
  matched: number;
  /** "Sheet!A1" of the rows without a file */
  missing: string[];
  /** picked files no row wants */
  extra: string[];
  /** one plain sentence: "6 of 6 input files matched." */
  summary: string;
  /** the current version has the files brought in: by a bring-in that finished, with the files picked now */
  brought_in: boolean;
  /** why not, when the current line went through a bring-in that no longer counts ("the input files
   *  changed since they were brought in", "the last bring-in did not finish"); null otherwise */
  brought_in_reason?: string | null;
  /** the Mind inputs (/Input grids) the bring-in made for the current line */
  mind_inputs?: string[];
  /** the matched files the assistant made (stand-ins from the model's own numbers): best results come
   *  with the actual files */
  stand_ins?: string[];
  /** the assistant is set up: it can make the missing files ("Make them with AI") */
  assistant?: boolean;
}

export interface InputFile {
  name: string;
  bytes: number;
  /** made by the assistant, not picked by the user */
  stand_in?: boolean;
}

export interface OutsideImportFile {
  cell: string;
  file: string | null;
  expected: string | null;
  exact: boolean;
  opened: boolean;
  areas: { sheet: string; ref: string; rows?: number; cols?: number }[];
  cleared?: unknown[];
}

/** What the model's own import macro did with the picked files (run_import). */
export interface OutsideImportReport {
  macro: string;
  seconds: number | null;
  files: OutsideImportFile[];
  /** areas put back from the upload: the macro cleared them for a file nobody picked */
  kept: { sheet: string; ref: string }[];
  /** the message boxes the macro showed, answered by the app */
  dialogs: { title?: string; message?: string; buttons?: string[]; pressed?: string | null }[];
  stopped_early: boolean;
  error: string | null;
  saved?: string | null;
}

export interface MindInputMade {
  name: string;
  sheet: string;
  ref: string;
  file?: string | null;
  rows: number;
  cols: number;
  /** the CSV's file name (a download of the session) */
  csv: string;
  separator?: string;
  readers?: number;
}

/** The Mind inputs made from the import areas, and the areas that stay inside the model (apply_mind_inputs). */
export interface MindInputsReport {
  inputs: MindInputMade[];
  kept: { sheet: string; ref: string; name?: string; why: string }[];
  formulas_changed: number;
  names_changed: number;
  sums_split?: number;
  saved?: string | null;
}

export interface OutsideRunResult {
  import?: OutsideImportReport | null;
  mindInputs?: MindInputsReport | null;
  versions: string[];
  download?: { zip: string; files: string[] } | null;
  /** one plain line: "6 outside files brought in; 6 Mind inputs made (IN_BBA_01, ...); no computed value changed." */
  summary: string;
  /** the comparison of every computed value before and after the Mind inputs: `ran` false when it
   *  could not run (the values were NOT checked); null when it was not asked for */
  numbers?: { ran?: boolean; clean?: boolean | null; verdict?: string; counts?: Record<string, number> } | null;
  /** what the user should know, in plain words (a file not picked, an area that stays inside the model) */
  notes?: string[];
  /** the stand-in files the assistant made first (when the user had none): how the macro reads a file,
   *  the rows left without one, and each attempt's outcome */
  standIns?: { ok: boolean; explain?: string | null; files?: string[]; left_out?: string[]; attempts?: { n: number; problem: string | null; details?: string[] }[] } | null;
}

/** GET .../outside-inputs: the background job that brings the outside files in. */
export interface OutsideRunStatus {
  outsideInputs: OutsideInputs | null;
  running: boolean;
  /** idle | running | done | error */
  state?: string;
  /** the versions the job made (also when it stopped half-way) */
  versions?: string[];
  /** the stage in words ("Running the workbook's import macro") */
  stage: string | null;
  result: OutsideRunResult | null;
  error: string | null;
  /** the latest progress line of the stage, and the seconds since the job started */
  message?: string | null;
  elapsed_s?: number;
  /** the stand-in files the assistant made (or tried to): each try's outcome, also when it failed */
  standIns?: OutsideRunResult["standIns"];
}

/** The model workbook itself (the main file of an upload). */
export const WORKBOOK_FILE_RE = /\.(xlsx|xlsm|xlsb)$/i;
/** What a model's input file may be (INPUT_FILE_EXTENSIONS of app/outside_inputs.py). */
export const INPUT_FILE_RE = /\.(xlsx|xlsm|xlsb|xls|csv|txt|prn)$/i;
/** The File Explorer filter: the model and its input files. */
export const PICK_ACCEPT = ".xlsx,.xlsm,.xlsb,.xls,.csv,.txt";

/** "a, b, c and 2 more" */
export function nameList(names: string[], max = 3): string {
  return names.slice(0, max).join(", ") + (names.length > max ? ` and ${names.length - max} more` : "");
}

/** How the server compares two file names (NFKC: full-width （１） = (1); case ignored). */
function sameName(name: string): string {
  return name.normalize("NFKC").toLowerCase().trim();
}

/** Why these input files cannot be sent (not an Excel or text file, the same name twice), or "". */
export function inputFilesProblem(names: string[]): string {
  const bad = names.filter((n) => !INPUT_FILE_RE.test(n));
  if (bad.length) return `${nameList(bad)} ${bad.length === 1 ? "is not an Excel or text file" : "are not Excel or text files"}: pick .xlsx, .xlsm, .xlsb, .xls, .csv or .txt files.`;
  const seen = new Set<string>();
  for (const n of names) {
    const k = sameName(n);
    if (seen.has(k)) return `Two files are named ${n}: pick only one of them.`;
    seen.add(k);
  }
  return "";
}

/** Why the files picked on the upload screen cannot be sent, or "": one workbook (as before), or the
 *  model workbook together with its input files (the server finds which one is the model). */
export function pickedFilesProblem(names: string[]): string {
  if (names.length <= 1) return names.length && !WORKBOOK_FILE_RE.test(names[0]) ? "Choose an .xlsx, .xlsm or .xlsb workbook." : "";
  if (!names.some((n) => WORKBOOK_FILE_RE.test(n))) return "Pick the model workbook (.xlsx, .xlsm or .xlsb) together with its input files.";
  return inputFilesProblem(names);
}

/** A row of the import table in words: the file name the model expects, else the row's key. */
export function outsideRowLabel(r: OutsideInputRow): string {
  return r.expected || (r.keys ?? []).filter(Boolean).join(" / ") || r.cell;
}

/** The outcome of "Bring in the outside files" in plain words, and whether Do everything may go on.
 *  It stops when the model's macro did not read every picked file (a file it did not open, or a
 *  question the app answered No), or when a computed value changed (values must never change) --
 *  in red -- and, in amber, when the comparison of the computed values could not run (`unchecked`:
 *  the files are in, the user goes on on purpose). An area that stays inside the model is reported,
 *  not a stop. */
export function outsideOutcome(r: OutsideRunResult): { ok: boolean; tone: "ok" | "warn" | "error"; line: string; detail: string; unchecked: boolean } {
  const imp = r.import;
  const made = r.mindInputs;
  const picked = (imp?.files ?? []).filter((f) => f.file);
  const notRead = picked.filter((f) => !f.opened).map((f) => f.file as string);
  const dialogs = imp?.dialogs ?? [];
  const answer = (p: string | null | undefined) => (p ?? "").replace(/&/g, "").split("(")[0].trim();
  const count = (n: number | null | undefined) => Number(n ?? 0).toLocaleString("en-US");
  const declined = dialogs.find((d) => d.pressed && answer(d.pressed).toUpperCase() !== "OK");
  const stopped = !!imp?.stopped_early;
  const changed = !!r.numbers?.ran && r.numbers.clean === false;
  // the comparison was asked for and could not run (Excel failed it): no value was checked
  const unchecked = !stopped && !changed && !!r.numbers && r.numbers.ran === false;
  const kept = made?.kept ?? [];
  const inputs = made?.inputs ?? [];
  const notes = (r.notes ?? []).filter(Boolean);
  const line = !stopped
    ? unchecked
      ? "The outside files are brought in, but the computed values were not checked"
      : r.summary || `${inputs.length} Mind input${inputs.length === 1 ? "" : "s"} made`
    : notRead.length
      ? `The import macro stopped before it read every file (not read: ${nameList(notRead)})`
      : declined
        ? `The import macro asked a question and the app answered ${answer(declined.pressed)}, so it may have stopped early`
        : "The import macro stopped early";
  // the server's notes say in words what the reports hold; without them, the reports are read here
  const own = [
    imp?.error ? `the macro said: ${imp.error.slice(0, 120)}` : "",
    imp?.kept?.length ? `the model keeps its own data in ${nameList(imp.kept.map((k) => `${k.sheet}!${k.ref}`))} (no file picked)` : "",
    kept.length ? `stays inside the model: ${kept.slice(0, 2).map((k) => `${k.sheet}!${k.ref} (${k.why})`).join("; ")}${kept.length > 2 ? `; … ${kept.length - 2} more` : ""}` : "",
  ];
  const detail = [
    unchecked ? `${(r.numbers?.verdict || "The comparison of the computed values could not run.").trim()} Check them under Advanced › Deliver (Check the numbers).` : "",
    stopped || unchecked ? r.summary : "",
    imp ? `${picked.length - notRead.length} of ${picked.length} file${picked.length === 1 ? "" : "s"} read by the macro ${imp.macro}${imp.seconds != null ? ` (${Math.round(imp.seconds)} s)` : ""}` : "",
    dialogs.length ? `${dialogs.length} message box${dialogs.length === 1 ? "" : "es"} answered${dialogs[0].message ? `: "${dialogs[0].message.slice(0, 80)}"` : ""}` : "",
    made ? `${count(made.formulas_changed)} formula${made.formulas_changed === 1 ? "" : "s"} now read the Mind inputs${made.sums_split ? ` (${count(made.sums_split)} sums split in two)` : ""}` : "",
    ...(notes.length ? notes : own),
  ]
    .filter(Boolean)
    .join(" · ");
  const tone = stopped || changed ? "error" : unchecked || kept.length || !inputs.length ? "warn" : "ok";
  return { ok: !stopped && !changed && !unchecked, tone, line, detail, unchecked };
}

/** The first clause of a message from the server, for a one-line status (the whole message goes in the
 *  details): it ends at the first ". " or "; " outside quotes, and is cut at about 160 characters. */
export function firstClause(text: string, max = 160): string {
  const t = (text ?? "").trim();
  let quoted = false;
  let end = t.length;
  for (let i = 0; i < t.length - 1; i++) {
    const c = t[i];
    if (c === '"' || c === "“" || c === "”") quoted = c === "“" ? true : c === "”" ? false : !quoted;
    else if (!quoted && (c === "." || c === ";") && t[i + 1] === " ") {
      end = i;
      break;
    }
  }
  const out = t.slice(0, end).replace(/[.;]+$/, "");
  return out.length > max ? out.slice(0, max - 1).trimEnd() + "…" : out;
}

/** The label of the card's button that brings the outside files in: "Try again" after a run that
 *  failed, the matched files only when some are missing, "Bring in again" when the current line was
 *  brought in before (files changed since, or that bring-in did not finish). */
export function outsideButtonLabel(o: OutsideInputs, failed: boolean): string {
  if (failed) return "Try again";
  if (o.matched && o.missing.length) return `Bring in the ${o.matched} matched file${o.matched === 1 ? "" : "s"}`;
  if (o.brought_in || o.brought_in_reason) return "Bring in again";
  return "Bring in the outside files";
}

/** Why "Do everything" skips the outside step ("" when it runs it): only when the current line has the
 *  files brought in -- the server says so for a bring-in that finished, with the files picked now --
 *  or when the app cannot bring them in. It runs after a failure too: the server then says not brought in. */
export function outsideSkipReason(o: OutsideInputs | null, unchecked = false): string {
  if (!o) return "The model's macros open no outside files";
  if (o.brought_in) return unchecked ? "Already brought in (the computed values were not checked)" : "Already brought in";
  if (!o.rows.length) return "No cell names the files the macro opens: their data stays inside the model";
  if (!o.macro) return "Which macro brings the outside files in is not clear: their data stays inside the model";
  return "";
}

/** The warning while stand-ins the assistant made are in use, or "": they reproduce the numbers the
 *  model holds, but the actual files are what the model was built for. */
export function outsideStandInLine(o: OutsideInputs | null): string {
  const n = o?.stand_ins?.length ?? 0;
  if (!o || !n) return "";
  return `${n === o.rows.length ? "All" : `${n} of ${o.rows.length}`} input files are stand-ins the AI assistant made from the numbers the model already holds. Best results come with the actual input files: add them, then bring them in again.`;
}

/** What to do when files are missing, in words: add them, or let the assistant make them. */
export function outsideAddLine(o: OutsideInputs | null): string {
  return o?.assistant
    ? `Add the model's input files on the Outside files card (Add files), or press "Make them with AI" there: the assistant makes stand-ins from the numbers the model holds (best results come with the actual files).`
    : "Add the model's input files on the Outside files card (Add files).";
}

/** The card's line when the current line no longer counts as brought in (the server's reason), or "". */
export function outsideReasonLine(o: OutsideInputs | null): string {
  const why = (o && !o.brought_in && o.brought_in_reason ? o.brought_in_reason : "").trim().replace(/\.$/, "");
  return why ? `${why[0].toUpperCase()}${why.slice(1)}: bring them in again.` : "";
}

/** "Bring in the outside files" when the server's estimate has no such step, with the measures of
 *  app/estimate.py: the model's own import macro and the making of the Mind inputs each start Excel,
 *  open and save the workbook; each new version is scanned again; the numbers check opens both. */
export function outsideStepEstimate(basis: Partial<ChainEstimate["basis"]> | undefined, standIns = false): ChainEstimateStep {
  const mb = Math.max(0, basis?.file_mb ?? 0);
  const scan = basis?.scan_s && basis.scan_s > 0 ? basis.scan_s : 30;
  const open = 5 + 8 + 0.3 * mb;
  const save = 10 + 0.18 * mb;
  const fixed = 2 * (open + save) + 10 + 2 * (open + 10 + save) + 0.2 * scan; // import, Mind inputs, numbers check
  // stand-ins: the assistant reads the macro (1-3 min a try, up to 3 tries), the files are built, the
  // macro runs once more on them and every computed value is compared
  const made = standIns ? { low: 90 + 2 * (open + save), mid: 240 + 3 * (open + save), high: 3 * 300 + 6 * (open + save) } : { low: 0, mid: 0, high: 0 };
  return {
    id: "outside",
    low_s: Math.round(fixed + 10 + 2 * 0.7 * scan + made.low),
    mid_s: Math.round(fixed + 60 + 2 * 1.2 * scan + made.mid),
    high_s: Math.round(fixed + 300 + 2 * 1.5 * scan + made.high),
    why:
      (standIns ? "the AI assistant making the missing input files (checked with the model's own macro), then " : "") +
      "the model's own import macro (seconds to minutes, by the number of files), the Mind inputs, two scans and the numbers check",
  };
}

/** The chain's estimate with the outside step in it when that step will run and the server did not count it. */
export function withOutsideStep(e: ChainEstimate | null, include: boolean, standIns = false): ChainEstimate | null {
  if (!e || !include || e.steps.some((s) => s.id === "outside")) return e;
  const step = outsideStepEstimate(e.basis, standIns);
  return {
    ...e,
    steps: [step, ...e.steps],
    total: { low_s: e.total.low_s + step.low_s, mid_s: e.total.mid_s + step.mid_s, high_s: e.total.high_s + step.high_s },
  };
}


// ---- next quarter's Mind input files (POST / GET .../next-quarter-inputs) ---------------------------
/** One file of the zip: the grid it fills in Mind and the picked file it came from. */
export interface NextQuarterFile {
  /** the name Mind's Inputs window matches: the grid's name (IN_BBA_01.csv, or .txt when TAB-separated) */
  file: string;
  grid: string;
  /** the picked file it was made from */
  from: string;
  /** "macro": made from a raw file by the model's own import macro; "as_given": already shaped like the grid */
  source: "macro" | "as_given";
  rows: number;
  cols: number;
  separator?: string;
  /** values that differ from the files made when the outside files were brought in (null: nothing to compare with) */
  changed?: number | null;
}

export interface NextQuarterResult {
  ok: boolean;
  /** one plain line: "6 Mind input files made (IN_BBA_01.csv, ...): put them in Mind's Inputs window, then Run." */
  summary: string;
  files: NextQuarterFile[];
  /** the grids without a new file (Mind keeps the data it has), or whose new file does not fit */
  not_in_zip: { grid: string; why: string }[];
  notes: string[];
  /** picked files nothing takes, with why */
  unused: { file: string; why: string }[];
  /** "Next quarter Mind inputs.zip": README.md and the folder "Mind inputs" */
  zip: string | null;
  readme?: string | null;
  error: string | null;
}

/** GET .../next-quarter-inputs: the grids the files fill and the last job making them. */
export interface NextQuarterStatus {
  grids: { name: string; sheet: string; ref: string; rows: number; cols: number; file: string | null }[];
  /** the outside files were brought in: there are Mind input grids to fill */
  available: boolean;
  why_not?: string | null;
  running: boolean;
  /** idle | running | done | error */
  state?: string;
  stage: string | null;
  message?: string | null;
  elapsed_s?: number;
  result: NextQuarterResult | null;
  error: string | null;
}

/** What the Next quarter part of the Outside files card says before anything is picked. */
export const NEXT_QUARTER_LEAD =
  "Raw input files cannot go into Mind directly. Mind takes one file per input grid, shaped like the grid. " +
  "Pick the new quarter's raw files: the app runs the model's own import macro on them and gives you the files for Mind's Inputs window. No new model upload.";

/** One line per file of the zip: which grid it fills and what it was made from. */
export function nextQuarterLines(r: NextQuarterResult): string[] {
  const count = (n: number) => n.toLocaleString("en-US");
  return (r.files ?? []).map((f) => {
    const from = f.source === "as_given" ? `${f.from} (as it is)` : f.from;
    const changed = f.changed == null ? "" : f.changed === 0 ? "; same values as before" : `; ${count(f.changed)} value${f.changed === 1 ? "" : "s"} changed`;
    return `${f.file} → grid ${f.grid}, from ${from}${changed}`;
  });
}

/** The outcome in words: the server's summary, and the notes (a raw file recognised, a grid left out). */
export function nextQuarterOutcome(r: NextQuarterResult): { tone: "ok" | "warn" | "error"; line: string; detail: string } {
  if (!r.ok) return { tone: "error", line: firstClause(r.error ?? "") || "No Mind input file was made", detail: r.error ?? "" };
  const left = (r.not_in_zip ?? []).map((x) => `${x.grid}: ${x.why}`);
  const unused = (r.unused ?? []).map((u) => `${u.file}: ${u.why}`);
  const detail = [...(r.notes ?? []), ...left, ...unused].filter(Boolean).join(" · ");
  const tooBig = (r.not_in_zip ?? []).some((x) => !/Mind keeps the data it has/.test(x.why));
  return { tone: tooBig || unused.length ? "warn" : "ok", line: r.summary, detail };
}
