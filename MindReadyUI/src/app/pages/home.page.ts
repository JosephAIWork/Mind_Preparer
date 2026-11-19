import { CommonModule } from "@angular/common";
import { Component, OnDestroy, computed, effect, inject, signal, untracked } from "@angular/core";
import { ActivatedRoute, Router, RouterLink } from "@angular/router";
import { ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule } from "carbon-components-angular";
import { ApiService, errorText, wasCancelled } from "../core/api.service";
import { CancelButton } from "../core/cancel-button";
import {
  ChainEstimate, ChainStepId, FORMULA_ACTION_IDS, InputFile, LoopPlan, NEXT_QUARTER_LEAD, NextQuarterResult, NextQuarterStatus, Operation, OutsideInputRow, OutsideInputs, OutsideRunResult, OutsideRunStatus, PICK_ACCEPT, PendingUpload, PrepAction, PrepAsk, PrepRunStatus, ReadinessBlocking, START_KINDS, STOPPED_NOT_CHECKED, STOPPED_NOTHING_CHANGED, ScanStatus, SessionStart, WorkJob,
  appliedCount, assistantFixedWords, blockingLeftLine, chainStoppedLine, failedCount, failedText, firstClause, inputFilesProblem, isPending, nameList, nextQuarterLines, nextQuarterOutcome, outsideAddLine, outsideButtonLabel, outsideOutcome, outsideReasonLine, outsideRowLabel, outsideSkipReason, outsideStandInLine,
  outsideWaitingLine, pickedFilesProblem, prepDoneWords, prepNoKeyAsk, withOutsideStep,
} from "../core/models";
import { WorkService } from "../core/work.service";
import { FeedbackService } from "../core/feedback.service";
import { FinishBanner } from "../core/finish-banner";
import { RecalcNow } from "../core/recalc-now";
import { Thinking } from "../core/thinking";
import { PlainCodes, meaningsText } from "../core/plain-codes";
import { ReadinessBanner } from "../core/readiness-banner";
import { RunBar } from "../core/run-bar";
import { SessionStore } from "../core/session.store";
import { SHOW_ME_PLACES, ShowMeId, ShowMeRequest, showMeFor } from "../core/show-me";
import { ShowMeLink } from "../core/show-me-link";
import { KeyBox } from "./advanced/key-box";
import { RecentWorkbooks } from "./recent-workbooks";

type Phase = "idle" | "uploading" | "choose" | "scanning" | "loaded";
type Tone = "idle" | "ok" | "warn" | "error";
interface ActionState {
  running: boolean;
  fraction: number | null;
  line: string;
  detail: string;
  tone: Tone;
  /** the  result text  is shown only on "View details" */
  open?: boolean;
  /** set after Stop: the fixer finishes the round it is in */
  stopping?: boolean;
  /** Fix all issues: the errors it left could not be compared with the original (the comparison  did not run) */
  unchecked?: boolean;
  /** 17 Nov 2026: the step was stopped with Cancel;  the line is "Stopped. Nothing was  changed." (no error, no "Try again") */
  cancelled?: boolean;
}

const idle = (): ActionState => ({ running: false, fraction: null, line: "", detail: "", tone: "idle" });
/** 17  Nov 2026:  a step stopped with Cancel: one plain sentence, in no colour (it is not a failure), shown at once. */
const stoppedState = (): ActionState => ({ running: false, fraction: null, line: STOPPED_NOTHING_CHANGED, detail: "", tone: "idle", open: true, cancelled: true });

/** The  kind of  work the server runs for each button of the main screen: what its Cancel stops. */
const CANCEL_KINDS: Record<StepId | "fixFormulas" | "nextQuarter", string> = {
  outside: "outside",
  fixAll: "autofix",
  fixFormulas: "apply",
  prepare: "prep", // 17  Nov 2026:  the Prep step is one job of its own on the server (it was up to three Applies)
  loop: "apply",
  recalc: "recalc",
  nextQuarter: "next_quarter",
};

/** The  "Do everything  right after the upload" box, remembered in this browser. */
const AUTO_RUN_KEY = "mindReady.autoRun";

function readAutoRun(): boolean {
  try {
    return localStorage.getItem(AUTO_RUN_KEY) === "1";
  } catch {
    return false; // storage  blocked: the  box starts unticked
  }
}

/** "under  a minute",  "12 min", "1 h 25 min" */
function minutes(seconds: number): string {
  if (seconds < 60) return "under a minute";
  const m = Math.round(seconds / 60);
  return m < 90 ? `${m} min` : `${Math.floor(m / 60)} h ${m % 60} min`;
}

/** 17  Nov 2026  (live test L2b-5): no step is called "taking longer than usual" in its first minute. */
const LATE_AFTER_S = 60;

/** What  is left,  for a sentence: "about 7 min", or "under a minute" (never "about under a minute"). */
function aboutLeft(seconds: number): string {
  return seconds < 60 ? "under a minute" : `about ${minutes(seconds)}`;
}

/** The  scan's stages,  in order, as the backend names them (app/progress.py). */
const SCAN_STAGES: { id: string; title: string }[] = [
  { id: "convert", title: "Converting .xlsb through Excel" },
  { id: "copy", title: "Copying the workbook" },
  { id: "load", title: "Reading the workbook" },
  { id: "inventory", title: "Scanning cells and grids" },
  { id: "names", title: "Reading defined names" },
  { id: "rules", title: "Checking the rules" },
  { id: "report", title: "Building the report" },
  { id: "plan", title: "Planning the fixes" },
];

interface ScanRow {
  id: string;
  title: string;
  state: "done" | "current" | "waiting";
  seconds?: number;
  message?: string;
}

type StepId = ChainStepId;
interface Step {
  id: StepId;
  label: string;
  check: string;
  /** stopped:  amber --  the step did its work, but a check could not run (the user goes on on purpose) */
  state: "waiting" | "running" | "done" | "failed" | "stopped" | "skipped";
  line: string;
  detail: string;
  open?: boolean;
  /** 17 Nov 2026 (live test L2b-6): skipped after a step had begun: it keeps its  number in "Step 3 of 4",
   *  so the count does not change  in the middle of a run */
  counted?: boolean;
}

/** What  each step  is doing, for Basic's one line (plain words, no codes). */
const BASIC_WORDS: Record<StepId, string> = {
  outside: "bringing in the input files the model's macros read",
  fixAll: "fixing the formula errors, so the workbook calculates cleanly",
  prepare: "changing what Mind cannot read the way Excel does: table titles, formulas, empty cells",
  loop: "building the Mind loop, so Mind repeats the run for every case, then fixing anything the loop uncovered",
  recalc: "recalculating in Excel and checking that no value changed",
};

/** The  chain's steps  in order; "outside" only when the model's macros open outside files. */
const STEPS: { id: StepId; label: string; check: string }[] = [
  { id: "outside", label: "Bring in the outside files", check: "every input file matched and read by the model's own import macro" },
  { id: "fixAll", label: "Fix all issues", check: "the fixer ends clean, or every error left was already in the workbook (none caused by the preparation)" },
  { id: "prepare", label: "Fix the formulas and the structure", check: "every blocking action applied, nothing refused, none left" },
  { id: "loop", label: "Build the Mind loop", check: "nothing refused, and nothing newly blocking after it (skipped when the model has no loop)" },
  { id: "recalc", label: "Recalculate", check: "no formula error that was not already in the original" },
];

/** Where  to fix  what stopped a step of Do everything: the Advanced tab where that step's work is done
 *  by hand. The fixer has no tab of its own (stoppedWhere); the outside step's card is  on this screen. */
const FIX_UNDER: Record<Exclude<StepId, "outside" | "fixAll">, string> = { prepare: "Prep actions", loop: "Mind loop", recalc: "Recalculate" };

/** What  to do  about the step that stopped Do everything, and where. Fix all issues depends on how it
 *  ended: the errors it left are in the recalculation's list; when its comparison with the original  did
 *  not run, the numbers are checked under Deliver; when it did  not start, was lost or stopped  with an
 *  error, nothing under Advanced fixes that  -- it just runs again.
 *  17 Nov 2026: `basic`: the sentence is  for the Basic screen, which has no  Outside files card of its
 *  own (the card is on Custom). The button's name is swapped for Basic's afterwards (`basicWords`). */
function stoppedWhere(id: StepId, r: ActionState, basic = false): string {
  const again = "press Do everything again.";
  if (id === "outside") return `See the Outside files card ${basic ? "on the Custom screen" : "above"}, then ${again}`;
  if (id === "fixAll") {
    if (r.unchecked) return `Check the numbers under Advanced › Deliver, then ${again}`;
    if (r.tone === "error") return "Press Do everything again.";
    return `Fix that under Advanced › Recalculate, then ${again}`;
  }
  return `Fix that under Advanced › ${FIX_UNDER[id]}, then ${again}`;
}

/** 17  Nov 2026  (live test L2b-2): the server refused to start a step because other work of this workbook
 *  still runs: its own sentence, "Wait a moment: a recalculation is still running." (also the older  "a
 *  recalculation is running for this workbook: wait for it to finish"),  with a capital and a full  stop.
 *  "" when the words say something else. */
function waitWords(detail: string): string {
  const said = (detail ?? "").trim().replace(/\.+$/, "");
  if (!/^wait a moment\b/i.test(said) && !/\bwait for it to finish$/i.test(said)) return "";
  return `${said[0].toUpperCase()}${said.slice(1)}.`;
}

/** A  sentence of  the chain in Basic's words: its button is called Start. Capitals stay as they are
 *  (17 Nov 2026: "... Not started. Press Do everything again." came out as "press Start again.", a
 *  sentence  that began with a small letter). */
function basicWords(text: string): string {
  return text.replace(/\bDo everything\b/g, "Start");
}

/** 17  Nov 2026:  the kinds of work this page runs on the server itself: Start's steps, the tiles (the same
 *  kinds) and next quarter's files */
const PAGE_KINDS: readonly string[] = [...START_KINDS, "next_quarter"];

/** 17  Nov 2026  (repair round): the work of this workbook that Start's steps cannot begin beside, so Start /
 *  Do everything is greyed while it runs: Start's own kinds, next quarter's files, and the suggested  names
 *  and the report files (they write at their end, and the  server lets no change begin then).  An
 *  assistant's answer, a Run in Mind, the  numbers check and the instance finder do not hold Start back:
 *  the  server starts Start's steps beside them, and  Start's Cancel leaves them running. */
const HOLDS_START: readonly string[] = [...START_KINDS, "next_quarter", "names", "reports"];

/** The  kinds of  work each step of Start runs on the server: the loop step writes the loop, then may fix
 *  what it uncovered (a Prep job). */
const STEP_KINDS: Record<StepId, readonly string[]> = { outside: ["outside"], fixAll: ["autofix"], prepare: ["prep"], loop: ["apply", "prep"], recalc: ["recalc"] };

/** 17  Nov 2026  (live test L2b-4): this browser tab remembers which workbook a Start was running for, until
 *  it ends in the page (sessionStorage outlives a reload, not the tab). A page reloaded or left  while it
 *  ran finds it there and says that the rest of  Start did not run. */
const START_MARK = "mindReady.startRunning";

/** What  the tab  remembers of a Start that runs (repair round, 17 Nov 2026): the workbook; the kinds of
 *  work of the step that runs ([] between two steps); whether it is Start's last step  (nothing would be
 *  left undone once it ends); and when the page  was left or reloaded (seconds since  1970, 0 while the
 *  page is open). Only  a job of those kinds that began before the page was  left is that step: a
 *  Recalculate pressed under  Advanced afterwards is not. */
interface StartMark {
  id: string;
  kinds: readonly string[];
  last: boolean;
  left: number;
}

/** A  job that  began this many seconds after the page was left is still the step's own (the server's
 *  clock and the page's are one machine's, but the step's request may reach the server a  moment late). */
const LEFT_SLACK_S = 2;

function markStart(mark: StartMark | null): void {
  try {
    if (mark) sessionStorage.setItem(START_MARK, JSON.stringify(mark));
    else sessionStorage.removeItem(START_MARK);
  } catch {
    /* storage  blocked: nothing  is remembered, and nothing is said after a reload (as before) */
  }
}

/** The  mark this  tab holds (null: none). A mark that is only the workbook's id (the first form of it):
 *  any of Start's kinds is its step. */
function startMarked(): StartMark | null {
  let text: string | null = null;
  try {
    text = sessionStorage.getItem(START_MARK);
  } catch {
    return null;
  }
  if (!text) return null;
  try {
    const m = JSON.parse(text) as Partial<StartMark> | null;
    if (m && typeof m === "object" && typeof m.id === "string") {
      return { id: m.id, kinds: Array.isArray(m.kinds) ? m.kinds : START_KINDS, last: m.last === true, left: typeof m.left === "number" ? m.left : 0 };
    }
  } catch {
    /* not  written as  an object: the id alone */
  }
  return { id: text, kinds: START_KINDS, last: false, left: 0 };
}

/** The  job is  the step that ran when the page was left: one of its kinds, begun before the page was left. */
function isMarkedStep(mark: StartMark, j: WorkJob): boolean {
  if (!mark.kinds.includes(j.kind)) return false;
  return !mark.left || !(typeof j.started === "number" && j.started > mark.left + LEFT_SLACK_S);
}

/** Repair  round, 17  Nov 2026: one blocking problem, the same before and after the loop (its rule and
 *  where it is). */
function problemKey(b: ReadinessBlocking): string {
  return `${b.rule_id}|${JSON.stringify(b.location ?? {})}`;
}

/** 17  Nov 2026  (review R2-3): where "Show me" leads for a blocking problem left at the end of Start: the
 *  Prep action that repairs its rule (its "By hand" entry when Prep skipped the cell),  else the rule's
 *  own row under Findings, ending on the cell (as  the server's `ask.show_me` does for the  question). */
/** 19 Nov 2026 (Joseph): after a Cancel,  "Show me" leads to the step that was stopped -- the  place where that
 *  step runs by itself: the  Outside files card, Fix all issues on Custom, the Prep actions, the Mind loop's
 *  Build it, Recalculate. */
const PLACE_OF_STEP: Record<StepId, ShowMeId> = { outside: "outside-card", fixAll: "fix-all", prepare: "prep-apply", loop: "loop-build", recalc: "recalc-run" };

function placeOfProblem(b: ReadinessBlocking): ShowMeRequest {
  if (b.fix === "outside_inputs") return { id: "outside-bring-in" };
  const loc = b.location ?? {};
  const sheet = typeof loc["sheet"] === "string" ? loc["sheet"].trim() : "";
  const cell = typeof loc["cell"] === "string" ? loc["cell"].trim() : "";
  const where = sheet && cell ? { where: { sheet, cell } } : {};
  return { id: b.fix === "prep" || b.fix === "prep_skipped" ? "prep-action" : "assistant-ask", ruleId: b.rule_id, ...where, open: true };
}

/** A  blocking Prep  action that is on by default still has changes to write (the Prep step's own work). */
function prepWorkLeft(plan: PrepAction[]): boolean {
  return plan.some((a) => a.level === "blocking" && a.default_on && a.count > 0);
}

/** The  main screen:  one workbook, the buttons that do the work on their own, one download. */
@Component({
  selector: "mr-home",
  standalone: true,
  imports: [CommonModule, RouterLink, ButtonModule, FileUploaderModule, InlineLoadingModule, NotificationModule, ProgressBarModule, TagModule, TilesModule, KeyBox, ShowMeLink, ReadinessBanner, FinishBanner, RecentWorkbooks, RecalcNow, Thinking, RunBar, CancelButton],
  templateUrl: "./home.page.html",
  styleUrl: "./home.page.scss",
})
export class HomePage implements OnDestroy {
  private api = inject(ApiService);
  private router = inject(Router);
  private route = inject(ActivatedRoute, { optional: true });
  private codes = inject(PlainCodes);
  readonly store = inject(SessionStore);
  /** 17  Nov 2026:  the question's "Send feedback" opens the dialog with the problem already written */
  readonly feedback = inject(FeedbackService);
  /** 17 Nov 2026: what runs on the server for this workbook (the shell reads it every 3 s) */
  private work = inject(WorkService);
  /** 17 Nov 2026  (review R1-5): what the Cancel beside Start / Do everything stops:  Start's own steps only */
  readonly startKinds = START_KINDS;

  /** 1.15.0: Basic (route  "", the default screen: upload, one tick  box, Go) or Custom (route "custom": everything the main screen had). */
  readonly basic: boolean = this.route?.snapshot?.data?.["mode"] === "basic";
  /** Basic's  one tick box: the loop step runs  only when this is on (and the workbook has a loop). */
  readonly buildLoop = signal(true);

  readonly phase = signal<Phase>("idle");
  readonly scan = signal<ScanStatus | null>(null);
  /** 17 Nov 2026: the workbook being checked for the first time (the store does not hold it yet): its Cancel needs the id */
  readonly scanningId = signal<string | null>(null);
  /** 17 Nov 2026: the first check was stopped with Cancel: one calm sentence on the upload screen (not an error) */
  readonly scanNote = signal("");
  readonly uploadError = signal("");
  /** Words for the uploading screen: one workbook, or the model  with its input files. */
  readonly sending = signal("");
  readonly dragOver = signal(false);
  readonly pending = signal<PendingUpload | null>(null);
  readonly skip = signal<Set<string>>(new Set());
  readonly fixAll = signal<ActionState>(idle());
  readonly fixFormulas = signal<ActionState>(idle());
  readonly loop = signal<ActionState>(idle());
  readonly recalc = signal<ActionState>(idle());
  readonly prepare = signal<ActionState>(idle());
  /** Bring in the outside files: the model's own import macro, then one Mind input per file */
  readonly outside = signal<ActionState>(idle());
  readonly outsideResult = signal<OutsideRunResult | null>(null);
  /** the input files sent with the workbook (the upload's answer), for the note when its macros open none */
  readonly pickedInputs = signal<InputFile[]>([]);
  readonly addingInputs = signal(false);
  readonly inputsError = signal("");
  readonly cardDragOver = signal(false);
  /** the last run on this page failed (the job stopped, its check failed, it could not start): the card offers "Try again" */
  readonly outsideFailed = signal(false);
  /** the last run brought the files in but could not compare the computed values: amber, and Do everything stops there */
  readonly outsideUnchecked = signal(false);
  /** 1.14.0: Do everything stopped at its  first step and waits for  the input files (the server's job is  in
   *  state "waiting_for_inputs"): one  box under  Do everything -- pick them, Go on, or Make them with AI.
   *  Nothing runs meanwhile; the chain goes on from the bring-in itself (nothing done before is done again). */
  readonly outsideWaiting = signal(false);
  /** the files the model's  macro opens, as the import table has them now (a file  picked meanwhile shows as present) */
  readonly waitingFiles = computed(() => (this.store.outsideInputs()?.rows ?? []).map((r) => ({ name: outsideRowLabel(r), present: !!r.file, fills: r.target_sheet })));
  readonly waitingMissing = computed(() => this.waitingFiles().filter((f) => !f.present));
  /** (17 Nov  2026, review R2-5: in Basic's words on  Basic, whose button is Start: "Start is waiting for ...") */
  readonly waitingLine = computed(() => this.basicSays(outsideWaitingLine(this.waitingFiles(), !!this.store.outsideInputs()?.assistant)));
  /** Go on:  every file is there (and none is  being added) */
  readonly canGoOn = computed(() => this.outsideWaiting() && this.waitingMissing().length === 0 && !this.addingInputs());
  /** Make them with AI: only with an AI key saved (the server says so), and only while a file is missing */
  readonly canMakeStandIns = computed(() => this.outsideWaiting() && this.waitingMissing().length > 0 && !!this.store.outsideInputs()?.assistant && !this.addingInputs());
  /** Go on / Make them with AI pressed: its request is on its way (a second click would be refused by the server) */
  readonly goingOn = signal(false);
  /** "Make them with AI" pressed after a reload (no chain in this page): the next Do everything takes  the stand-in path */
  private chainStandIns = false;
  readonly pickAccept = PICK_ACCEPT;
  /** Make next quarter's Mind input files: the new quarter's raw files through the model's own import
   *  macro, one file per Mind input grid (raw files cannot go into Mind's Inputs window as they are) */
  readonly nextQuarter = signal<ActionState>(idle());
  readonly nextQuarterResult = signal<NextQuarterResult | null>(null);
  /** the session the result belongs to: never offered for another workbook */
  private readonly nextQuarterSession = signal<string | null>(null);
  /** bumped when the card is reset (another workbook, Load another): an older job's polling stops */
  private nextQuarterSeq = 0;
  readonly nextQuarterLead = NEXT_QUARTER_LEAD;
  /** "Next quarter Mind inputs.zip" of the last job that made files, in this session */
  readonly nextQuarterZipUrl = computed(() => {
    const id = this.store.sessionId();
    const r = this.nextQuarterResult();
    return id && r?.ok && r.zip && this.nextQuarterSession() === id ? this.api.downloadUrl(id, r.zip) : null;
  });
  /** which file goes into which grid */
  readonly nextQuarterFiles = computed(() => {
    const r = this.nextQuarterResult();
    return r?.ok ? nextQuarterLines(r) : [];
  });
  readonly loopPlan = signal<LoopPlan | null>(null);
  readonly steps = signal<Step[]>([]);
  /** Ticked: once an uploaded workbook  is read, Do everything starts  by itself -- and a big workbook  is
   *  read whole, without  stopping to  ask which sheets to leave out. */
  readonly autoRun = signal(this.basic ? false : readAutoRun());
  readonly chainRunning = signal(false);
  readonly chainLine = signal("");
  /** 17 Nov 2026: Cancel was clicked while Start / Do everything runs, and the chain has not ended yet: the
   *  screen says "Stopping…" and the head's Cancel  stays drawn until the end (the chain waits to hear from  it) */
  readonly chainStopping = signal(false);
  /** 17 Nov 2026: the chain's line  is the sentence after a Cancel ("Stopped.  Your workbook is as it was
   *  ..."): the person stopped it, there  is no place to fix, so no  "Show me" goes with it */
  /** 19 Nov 2026 (Joseph): a "Show me" does go with that sentence now: to the step that was stopped (`endPlace`, set
   *  in `endCancelled`). The sentence's own words still name no place. */
  readonly chainStoppedByCancel = signal(false);
  /** 17 Nov 2026: set at the click on Cancel (the button's `stopping`), reset in `runAll`; read wherever the
   *  chain would start something NEW (the next step, the next Prep  round, the fixes after the loop). No
   *  runner returns at the click: each one ends when its own call or its own job ends. */
  private chainCancelled = false;
  /** 17 Nov 2026 (never stop for a person when the app can fix it): the chain's line is the question the
   *  Prep step ended with -- something is left that the app could not fix by itself, or no AI key is
   *  saved. Its two sentences are drawn one per line on Basic, with "Show me" (to the problem's own cell)
   *  and  "Send feedback" (the problem already  written). Null: the line is an ordinary  one. */
  readonly chainAsk = signal<PrepAsk | null>(null);
  /** Custom: the question's  "View details"  is open (why the assistant could not fix it) */
  readonly askOpen = signal(false);
  /** 17 Nov 2026 (live test L2b-9): Basic's second sentence after "Everything ran and every check passed":
   *  "The assistant fixed 1: see Advanced › Versions."  ("" when it fixed nothing). Custom says it on the
   *  step's own  line. */
  readonly chainNote = signal("");
  /** 17 Nov 2026 (review R2-3): the  place a line names that the map  of words cannot find: the first
   *  blocking problem left when every step  ran ("Everything ran, but 2 blocking problems  are left.") */
  readonly endPlace = signal<ShowMeRequest | null>(null);
  /** 17 Nov 2026 (live test L2b-4): Start was running when this page was reloaded or left (this browser
   *  tab remembers it), so the rest of its steps never ran: one plain line says so */
  readonly interrupted = signal(false);
  /** what the tab remembered of that Start (which step ran, when the page was left): only its step counts */
  private interruptedMark: StartMark | null = null;
  /** a step of that Start still runs on the server:  from the check after the reload, then from the shell's list */
  private readonly startStepOnServer = signal(false);
  /** The line about it: while a step of that Start still runs on the server, and after it ended. */
  readonly interruptedLine = computed(() => {
    if (!this.interrupted() || this.chainRunning()) return "";
    const still = this.busy() || this.startStepOnServer();
    return `Do everything was interrupted when the page was reloaded or left. Press Do everything again ${still ? "when this step ends" : "to finish"}.`;
  });
  /** The chain's last line as the screen shows it (in Custom's words): the chain's own, else the one above. */
  readonly endLine = computed(() => this.chainLine() || this.interruptedLine());
  /** Basic: the place the chain's line names (null: none, and then no "Show me" row is drawn under it).
   *  The line is read in Custom's words: they are the ones the map of places knows ("press  Do everything").
   *  17 Nov 2026: a  place given with the line (endPlace) first;  then the place  the line  names; then the
   *  one its second sentence names (the assistant's fixes: Versions). */
  readonly stopPlace = computed<ShowMeRequest | null>(() => this.endPlace() ?? showMeFor(this.endLine()) ?? showMeFor(this.chainNote()));
  /** a question that does not stop Start (a blocking problem no Prep action repairs, which the assistant
   *  could not fix either): Start  goes on, and when it ends without another stop the question  is its last
   *  line, in place of  "Everything ran and every check passed". The  Prep run after the loop replaces it. */
  private pendingAsk: PrepAsk | null = null;
  /** the question the step that  just ran stopped Start with (read once  by `runAll`: `takeStepAsk`) */
  private stepAsk: PrepAsk | null = null;
  /** what the last Prep run wrote: for the loop step's line when its fixes follow the loop */
  private prepDid = { applied: 0, fixed: 0 };
  /** 17 Nov 2026 (live test L2b-9): what every Prep run of this Start wrote (the Prep step, and the fixes
   *  after the loop): for Basic's sentence on the assistant's fixes at the end */
  private prepTotal = { applied: 0, fixed: 0 };
  private prepFollowing = false;
  /** the loop's build asked for the workbook again (the Prep plan of  the new version): on its way */
  private afterLoop: Promise<void> | null = null;
  /** how the running step ended when it was stopped with Cancel ("": it was not). "after_loop": the Mind loop
   *  was built and only its fixes afterwards were stopped, so the loop counts as finished. */
  private stepEnd: "" | "cancelled" | "after_loop" = "";
  /** the head's Cancel has heard from the server (stopped, or it could not ask): the chain waits for this
   *  before it writes its sentence, so that the store already holds the workbook as it is now */
  private cancelHeard: Promise<void> | null = null;
  private cancelHeardNow: (() => void) | null = null;
  /** 1.11.6: the estimate for "Do everything" (shown before the click)  and the countdown while it  runs */
  readonly estimate = signal<ChainEstimate | null>(null);
  /** 16 Nov 2026 (night): this run asked  for its estimate  and the  answer is in (there may be none). Until
   *  then Basic says "Working out how long this will take." and its bar only moves. */
  private readonly estimateSettled = signal(false);
  /** the highest percentage Basic's bar has shown in this run:  the bar never moves backwards */
  private barHigh = 0;
  private readonly clock = signal(Date.now());
  private chainT0 = 0;
  private stepT0 = 0;
  /** when the wait for the input files  began: the wait is not the  step's time (stepT0 moves past it when  it ends) */
  private waitT0 = 0;
  private stepSeconds: Partial<Record<StepId, number>> = {};
  private ticker: ReturnType<typeof setInterval> | null = null;
  /** Do everything starts by bringing in the outside files: the  macros open some, the app knows which
   *  cells  name them, and they are not in the current version yet. */
  readonly outsideWillRun = computed(() => {
    const o = this.store.outsideInputs();
    return !!o && !o.brought_in && this.canBringIn(o);
  });
  /** The server's estimate, with the outside step in it when that step runs and the server left it out. */
  readonly fullEstimate = computed(() => {
    const include = this.chainRunning()
      ? this.steps().some((s) => s.id === "outside" && (s.state === "waiting" ? this.outsideWillRun() : s.state !== "skipped"))
      : this.outsideWillRun();
    // 1.14.0: Do everything no longer makes stand-ins by itself (it stops and asks): the plain step's time
    return withOutsideStep(this.estimate(), include, false);
  });
  readonly estimateLine = computed(() => {
    const e = this.fullEstimate();
    if (!e) return "";
    const t = e.total;
    return `Rough estimate: about ${minutes(t.mid_s)} (between ${minutes(t.low_s)} and ${minutes(t.high_s)}).`;
  });
  /** Elapsed and what is left, recomputed every 10 s: each step's likely time, scaled by how the
   *  finished  steps compared with their own estimate, minus what the running step has already taken. */
  readonly timerLine = computed(() => {
    const now = this.clock();
    if (!this.chainRunning() || !this.chainT0) return "";
    const elapsed = `${minutes((now - this.chainT0) / 1000)} so far`;
    if (this.chainStopping()) return `${elapsed} · stopping…`; // 17 Nov 2026: Cancel was clicked: no countdown any more
    if (this.outsideWaiting()) return `${elapsed} · paused: waiting for the input files`; // nothing runs: no countdown
    const t = this.chainLeft();
    if (!t) return elapsed;
    // 16 Nov 2026 (night): a late step says one thing, the same words as on Basic
    if (t.slow) return `${elapsed} · this step is taking longer than usual; ${aboutLeft(t.left)} left, maybe more`;
    return `${elapsed} · ${aboutLeft(t.left)} left`;
  });
  /** What is left of Do everything, in seconds (null without an estimate): each step's likely time, scaled by
   *  how the finished steps compared with their own estimate, minus what the running step has already taken. */
  readonly chainLeft = computed<{ left: number; slow: boolean } | null>(() => {
    const now = this.clock();
    const e = this.fullEstimate();
    if (!this.chainRunning() || !this.chainT0 || !e) return null;
    const est = new Map(e.steps.map((s) => [s.id, s]));
    let planned = 0;
    let took = 0;
    for (const s of this.steps()) {
      const actual = this.stepSeconds[s.id];
      const x = est.get(s.id);
      // a step without an estimate would only push  the ratio up
      if (s.state === "done" && actual !== undefined && x) {
        planned += x.mid_s;
        took += actual;
      }
    }
    const k = planned >= 60 ? Math.min(3, Math.max(0.5, took / planned)) : 1;
    let left = 0;
    let slow = false;
    for (const s of this.steps()) {
      const x = est.get(s.id);
      if (!x) continue;
      if (s.state === "waiting") left += x.mid_s * k;
      else if (s.state === "running") {
        const inStep = (now - this.stepT0) / 1000;
        left += Math.max(x.mid_s * k - inStep, 0.1 * x.mid_s * k, 30);
        // 17 Nov 2026  (live test L2b-5): late only when this  step's own time  is over:  past its own longest
        // estimate, and never in its first minute. A step whose estimate is a few seconds (the loop's one
        // round through Excel, before its fixes) or nothing (estimated as skipped)  was "taking longer than
        // usual" 3 seconds after it began.
        slow = x.high_s > 0 && inStep > Math.max(x.high_s * k, LATE_AFTER_S);
      }
    }
    return { left, slow };
  });
  readonly busy = computed(
    () => this.chainRunning() || this.outside().running || this.nextQuarter().running || this.addingInputs() || this.fixAll().running || this.fixFormulas().running || this.prepare().running || this.loop().running || this.recalc().running,
  );
  /** Which step holds  the buttons, for the hint on  a disabled button (17 Nov 2026, review  R2-5: in Basic's
   *  words on Basic: "Start is running"). */
  readonly busyHint = computed(() => this.basicSays(this.busyWhy()));
  private readonly busyWhy = computed(() => {
    if (this.chainRunning()) return this.outsideWaiting() ? "Do everything is waiting for the input files" : "Do everything is running";
    if (this.outside().running) return "Bring in the outside files is running";
    if (this.nextQuarter().running) return "Next quarter's Mind input files are being made";
    if (this.addingInputs()) return "The input files are being added";
    if (this.fixAll().running) return this.fixAll().stopping ? "Fix all issues is stopping after its round" : "Fix all issues is running";
    if (this.fixFormulas().running) return "Fix the formulas is running";
    if (this.prepare().running) return "Fix the formulas and the structure is running";
    if (this.loop().running) return "Build the Mind loop is running";
    if (this.recalc().running) return "Recalculate is running";
    return "";
  });
  /** 17 Nov 2026 (live  test L2b-2): work of this workbook that  this page is not doing itself runs on the
   *  server (a Recalculate pressed under Advanced, a step of Start picked up by the shell's strip): Start /
   *  Do everything would be refused, so it is greyed with the reason (it was blue, and its click ended in
   *  'Stopped at "Fix all issues": Not started.'). Only work that holds Start back counts (HOLDS_START,
   *  repair round): an  assistant's answer or a Run in Mind does not grey it for its whole length. While
   *  Start waits for the input files after a reload, the bring-in waits too: its box is the way on, so
   *  that job does not count. */
  readonly otherWork = computed<WorkJob | null>(() => {
    if (this.busy()) return null; // the page's own work: busyHint says which
    return this.work.jobs().find((j) => HOLDS_START.includes(j.kind) && !(j.kind === "outside" && this.outsideWaiting())) ?? null;
  });
  /** Start / Do everything cannot be pressed: the page's own work, or other work of this workbook, runs */
  readonly startBlocked = computed(() => this.busy() || !!this.otherWork());
  /** Why Start / Do everything is greyed, in plain words ("" when it is not). */
  readonly startHint = computed(() => {
    const j = this.otherWork();
    return j ? this.basicSays(`Still working: ${j.title}. Press Do everything once it has finished.`) : this.busyHint();
  });

  readonly readinessKind = computed(() => {
    const r = this.store.readiness();
    if (!r) return "info";
    return r.state === "ready" ? "success" : r.state === "unverified" ? (r.recalc?.ran ? "warning" : "info") : "error";
  });
  readonly readinessTitle = computed(() => this.store.readiness()?.headline ?? "");
  readonly readinessDetail = computed(() => {
    const r = this.store.readiness();
    if (!r) return "";
    const n = r.next_step;
    const step = n ? [n.label, n.detail].filter(Boolean).join(": ") : "";
    const counts = [r.blocking_count ? `${r.blocking_count} blocking` : "", r.manual_blocking_count ? `${r.manual_blocking_count} by hand` : "", r.optional_count ? `${r.optional_count} optional` : ""].filter(Boolean).join(", ");
    return [step, counts].filter(Boolean).join(" · ");
  });
  readonly formulaOps = computed<Operation[]>(() => this.store.plan().filter((a) => FORMULA_ACTION_IDS.has(a.id)).flatMap((a) => a.operations));
  readonly formulaCells = computed(() => this.store.plan().filter((a) => FORMULA_ACTION_IDS.has(a.id)).reduce((n, a) => n + a.count, 0));
  readonly errorCount = computed(() => this.store.statusCounts()["ERROR"] ?? 0);
  readonly recalcFact = computed(() => {
    const rc = this.store.readiness()?.recalc;
    if (!rc?.ran) return "Not recalculated yet";
    const current = this.store.currentVersion()?.id;
    const stale = rc.version_id && current && rc.version_id !== current ? " (an earlier version)" : "";
    return (rc.clean ? "Last recalculation clean" : `${rc.formula_errors ?? "?"} error cells at the last recalculation`) + stale;
  });
  readonly downloadUrl = computed(() => {
    const id = this.store.sessionId();
    const v = this.store.currentVersion();
    return id && v ? this.api.downloadUrl(id, v.file_name) : null;
  });
  /** What the Outside files card shows:  the job's own state --  or, once the current version no longer
   *  has  the files brought  in by  that job (the server says why: files replaced or added since, a
   *  bring-in that did not finish; or another version was chosen), that, not the old result. The
   *  server is the one source:  the same after a reload. */
  readonly outsideCard = computed<ActionState>(() => {
    const s = this.outside();
    const o = this.store.outsideInputs();
    // (17 Nov 2026: a bring-in stopped with  Cancel says so, whatever the server  says about the files)
    if (!o || o.brought_in || s.running || s.tone === "error" || s.cancelled) return s;
    const why = outsideReasonLine(o);
    if (why) return { running: false, fraction: null, line: why, detail: "", tone: "warn" };
    return this.outsideResult() ? idle() : s;
  });
  /** "Mind inputs.zip": one CSV  per Mind input, made by the last Bring in the outside  files -- offered
   *  only while the current version  has them brought in (not the CSVs of files replaced since) */
  readonly zipUrl = computed(() => {
    const id = this.store.sessionId();
    const r = this.outsideResult();
    return id && this.store.outsideInputs()?.brought_in && r?.download?.zip && r.mindInputs?.inputs?.length ? this.api.downloadUrl(id, r.download.zip) : null;
  });

  // ---- the scan, as rows -------------------------------------------------------------------------
  readonly scanRows = computed<ScanRow[]>(() => {
    const s = this.scan();
    const done = new Map((s?.stages ?? []).map((x) => [x.stage, x.seconds]));
    const current = s?.stage ?? null;
    const needsConvert = (s?.stages ?? []).some((x) => x.stage === "convert") || current === "convert";
    return SCAN_STAGES.filter((st) => st.id !== "convert" || needsConvert).map((st) => ({
      id: st.id,
      title: st.title,
      state: done.has(st.id) ? "done" : st.id === current ? "current" : "waiting",
      seconds: done.get(st.id),
      message: st.id === current ? s?.message || undefined : undefined,
    }));
  });
  readonly scanPercent = computed(() => Math.round(Math.max(0, Math.min(1, this.scan()?.overall ?? 0)) * 100));
  readonly scanElapsed = computed(() => {
    const e = this.scan()?.elapsed_s ?? 0;
    return e >= 60 ? `${Math.floor(e / 60)} min ${Math.round(e % 60)} s` : `${Math.round(e)} s`;
  });

  private timers: ReturnType<typeof setTimeout>[] = [];
  /** 1.14.0: the scan that follows a Reopen (a saved workbook never checked): "Do everything right after
   *  the upload" is for uploads, so that scan's end starts nothing */
  private reopening = false;
  private fixAllFollowing = false;
  private outsideFollowing = false;
  private estimateSeq = 0;
  /** Set when the page is left (Advanced): a running Do everything starts no further step unseen. */
  private destroyed = false;
  /** takes the page's "pagehide" listener away  (the page is gone) */
  private unhide: (() => void) | null = null;
  /** the blocking problems the readiness listed when the Prep step of this Start ended (repair round, 17
   *  Nov 2026): the Prep job had them, and its assistant tried what Prep cannot repair. null: no Prep step
   *  ran in this Start. */
  private prepSaw: Set<string> | null = null;
  /** how many reads of the loop plan are on their way (Start pressed before the plan came cannot know yet
   *  whether the loop step runs) */
  private loopPlanReads = 0;

  constructor() {
    // The session may arrive after this page is built (a `?session=<id>` link is resumed by the
    // shell asynchronously): follow  the store rather than reading  it once.
    effect(() => {
      const loaded = this.store.loaded();
      const phase = this.phase();
      if (loaded && phase === "idle") {
        this.phase.set("loaded");
        void this.loadLoopPlan();
        void this.resumeFixer();
        void this.resumePrepRun();
        void this.resumeOutside();
        void this.resumeNextQuarter();
        void this.checkInterrupted();
      } else if (!loaded && phase === "loaded") {
        this.phase.set("idle");
        this.loopPlan.set(null);
        this.resetNextQuarter();
        this.interrupted.set(false);
      }
    });
    // 17 Nov 2026 (live tests L2b-2,  L2b-3): when the  page's own  work ends, the list of what runs on the
    // server drops it at once and is read again: a list up to 3 s old would grey Start, and the shell's
    // strip would  say "Still working: …" beside the line that says it ended
    let wasBusy = false;
    effect(() => {
      const now = this.busy();
      if (wasBusy && !now) untracked(() => this.work.ended(PAGE_KINDS));
      wasBusy = now;
    });
    // 17  Nov 2026 (live test L2b-4): the  interrupted Start's step, as the shell's list  says (every 3 s);
    // once none of Start's steps runs any more  the line says "to finish" and the  tab forgets that Start
    effect(() => {
      const jobs = this.work.jobs();
      untracked(() => {
        const mark = this.interruptedMark;
        if (!this.interrupted() || !mark) return;
        const still = jobs.some((j) => isMarkedStep(mark, j));
        this.startStepOnServer.set(still);
        if (!still) markStart(null);
      });
    });
    // repair round, 17 Nov 2026: a reload (or the tab closed) calls no ngOnDestroy: the mark learns here
    // when the page was left
    const hide = () => this.leaving();
    window.addEventListener("pagehide", hide);
    this.unhide = () => window.removeEventListener("pagehide", hide);
    // the estimate follows the analysed version and the loop plan (refreshed once a step is done)
    effect(() => {
      const id = this.store.sessionId();
      const version = this.store.currentVersion()?.id;
      this.loopPlan(); // whether the loop step runs changes the estimate
      this.buildLoop();
      if (id && version && !this.chainRunning()) void this.loadEstimate(id);
    });
  }

  private async loadEstimate(id: string): Promise<void> {
    // only the latest request counts: an older answer (made before the loop plan arrived) may  come last
    const seq = ++this.estimateSeq;
    try {
      const e = await this.api.estimate(id, this.canBuildLoop() && this.loopWanted());
      if (seq === this.estimateSeq) this.estimate.set(e);
    } catch {
      if (seq === this.estimateSeq) this.estimate.set(null); // an older server without the estimate: nothing shown
    }
  }

  ngOnDestroy(): void {
    this.leaving();
    this.unhide?.();
    this.destroyed = true;
    this.timers.forEach(clearTimeout);
    if (this.ticker) clearInterval(this.ticker);
    this.fixAllFollowing = false;
    this.cancelHeardNow?.(); // the Cancel button goes with the page: nothing waits for it any more
  }

  // ---- Cancel (17 Nov 2026) --------------------------------------------------------------------
  /** What the Cancel of a tile or a card stops on the server: that button's own kind of work. */
  cancelKind(which: StepId | "fixFormulas" | "nextQuarter"): string {
    return CANCEL_KINDS[which];
  }

  /** The click on the Cancel beside Start / Do everything, before the server is asked: from now on the chain
   *  starts nothing new. The step that runs ends by itself (its request answers "cancelled", or its job
   *  says so); then  the chain waits for the  button to hear from the server and  writes its one  sentence. */
  onChainStopping(): void {
    if (!this.chainRunning()) return;
    this.chainCancelled = true;
    this.chainStopping.set(true);
    this.cancelHeardNow?.(); // a second  try after a call that failed: the first wait is over
    this.cancelHeard = new Promise<void>((done) => (this.cancelHeardNow = done));
  }

  /** The button heard from the server: everything is stopped and the store holds the workbook as it is now. */
  onChainStopped(): void {
    this.cancelHeardNow?.();
    this.cancelHeardNow = null;
  }

  /** The server could  not be asked to stop (it did not answer): the button  is back to "Cancel" for another
   *  try.  The chain still starts nothing new --  the step that runs goes to its own end -- and  must not
   *  wait for ever for a button  that has nothing to say. */
  onChainStopFailed(): void {
    this.chainStopping.set(false);
    this.onChainStopped();
  }

  /** How the step that just ran was stopped with Cancel ("": it was not), read once: the mark is cleared. */
  private takeStepEnd(): "" | "cancelled" | "after_loop" {
    const end = this.stepEnd;
    this.stepEnd = "";
    return end;
  }

  /** The chain's end after a Cancel: the one sentence, built from the steps that really finished. */
  private async endCancelled(finished: StepId | null, stopped: StepId | null = null): Promise<void> {
    if (this.cancelHeard) await this.cancelHeard;
    this.chainStoppedByCancel.set(true);
    this.say(chainStoppedLine(finished));
    // 19 Nov 2026: "Show me" beside the sentence leads to the step that was stopped (say() has just cleared the place)
    this.endPlace.set(stopped ? { id: PLACE_OF_STEP[stopped] } : null);
  }

  /** The chain's line:  an ordinary sentence, or (17 Nov 2026) the question the Prep step ended with --
   *  then the line is its two sentences and the question itself is kept for its buttons. The job wrote
   *  the sentences for the screen that started it; a job picked up on the other screen (Start pressed
   *  on Basic, the page left for Custom) still names the button this screen has. */
  private say(line: string, ask: PrepAsk | null = null): void {
    const mine = ask && !this.basic ? { ...ask, lines: ask.lines.map((l) => l.replace(/\bpress Start again\b/g, "press Do everything again")) } : ask;
    this.chainAsk.set(mine);
    this.askOpen.set(false);
    this.chainLine.set(mine ? mine.lines.join(" ") : line);
    // 17 Nov 2026: a new line has none of the last one's extras (the assistant's sentence, its own place)
    this.chainNote.set("");
    this.endPlace.set(null);
  }

  /** A  sentence written in Custom's words,  in Basic's words on Basic (its button  is called Start). */
  private basicSays(text: string): string {
    return this.basic ? basicWords(text) : text;
  }

  /** 17  Nov 2026  (live test L2b-4): after a reload (or a visit to another screen) in the middle of Start:
   *  when this tab remembers a Start for this workbook and one of Start's steps still  runs on the server,
   *  the line says that the rest of Start  did not run, and to press  Start again when the step ends (the
   *  step  itself is picked up as before: the strip, a greyed Start,  Custom's list). When nothing runs any
   *  more it  says so once, "Press Start again to finish.", and the tab forgets it. */
  private async checkInterrupted(): Promise<void> {
    const id = this.store.sessionId();
    const mark = startMarked();
    if (!id || mark?.id !== id || this.chainRunning()) return;
    if (mark.last) {
      // repair round: left during Start's last step, which ends by itself on the server: nothing of
      // Start is left undone, so nothing is said (it said "Press Start again to finish.")
      markStart(null);
      return;
    }
    let jobs: WorkJob[] = [];
    try {
      jobs = (await this.api.jobs(id)) ?? [];
    } catch {
      jobs = []; // the server did not say: the line says "to finish"
    }
    if (this.destroyed || this.chainRunning() || this.store.sessionId() !== id) return;
    // only the step that ran when the page was left: a  Recalculate or an Apply pressed under Advanced
    // afterwards is not Start's (repair round)
    const still = jobs.some((j) => isMarkedStep(mark, j));
    if (!still) markStart(null);
    this.interruptedMark = mark;
    this.startStepOnServer.set(still);
    this.interrupted.set(true);
  }

  /** Repair round, 17 Nov 2026: the page is left or reloaded while Start runs here. In its last step
   *  nothing is left undone once that step ends on the server, so the tab forgets the run; else the mark
   *  says when the page was left (a job begun after that is not Start's). */
  private leaving(): void {
    if (this.destroyed || !this.chainRunning()) return;
    const mark = startMarked();
    if (!mark || mark.id !== this.store.sessionId()) return;
    markStart(mark.last ? null : { ...mark, left: Date.now() / 1000 });
  }

  /** The mark of the run, at the start of each step: which work it runs, and whether it is the  last one. */
  private markStep(id: string, step: StepId | null): void {
    const steps = this.steps();
    const at = step ? steps.findIndex((s) => s.id === step) : -1;
    const last = at >= 0 && !steps.slice(at + 1).some((s) => s.state === "waiting");
    markStart({ id, kinds: step ? STEP_KINDS[step] : [], last, left: 0 });
  }

  /** The question the step  that just ran stopped Start with (null:  none), read once. */
  private takeStepAsk(): PrepAsk | null {
    const ask = this.stepAsk;
    this.stepAsk = null;
    return ask;
  }

  /** Where  the question's  "Show me" leads: the place the server names (the Prep action that repairs the
   *  rule, the rule's own row under Findings, the AI key box). A place this screen does not know:  the
   *  Findings list, where every problem is. */
  askPlace(ask: PrepAsk): ShowMeId {
    const id = ask.show_me?.id as ShowMeId | undefined;
    return id && SHOW_ME_PLACES[id] ? id : "assistant-ask";
  }

  // ---- upload -----------------------------------------------------------------------------
  onFileInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    const files = Array.from(input.files ?? []);
    input.value = "";
    if (files.length) void this.uploadFiles(files);
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.dragOver.set(false);
    const files = Array.from(event.dataTransfer?.files ?? []);
    if (files.length) void this.uploadFiles(files);
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.dragOver.set(true);
  }

  /** One workbook, or the  model workbook together with its input  files (the server finds the model). */
  async uploadFiles(files: File[]): Promise<void> {
    const problem = pickedFilesProblem(files.map((f) => f.name));
    if (problem) {
      this.uploadError.set(problem);
      return;
    }
    this.reopening = false;
    this.uploadError.set("");
    this.scanNote.set("");
    this.sending.set(files.length === 1 ? "Sending the file to the app…" : `Sending the ${files.length} files to the app…`);
    this.phase.set("uploading");
    this.scan.set(null);
    this.pickedInputs.set([]);
    try {
      const started = await this.api.upload(files.length === 1 ? files[0] : files, "plan", true);
      this.pickedInputs.set(started.inputs ?? []);
      if (isPending(started)) {
        // A big  workbook with several sheets: let the user leave sheets out before  the scan
        // (not when Do everything is to  follow on its own: nothing stops the run).
        if (started.size?.above_threshold && started.size.sheets?.length > 1 && !this.autoRun()) {
          this.pending.set(started);
          this.skip.set(new Set());
          this.phase.set("choose");
          return;
        }
        await this.scanSession(started.sessionId, []);
      } else {
        this.store.setSession(started);
        this.afterLoad();
      }
    } catch (e) {
      this.uploadError.set(errorText(e));
      this.phase.set("idle");
    }
  }

  async scanSession(id: string, ignoreSheets: string[]): Promise<void> {
    this.scanNote.set("");
    this.scanningId.set(id);
    this.phase.set("scanning");
    const poll = this.pollScan(id);
    try {
      const full = await this.api.analyze(id, ignoreSheets);
      this.store.setSession(full);
      this.afterLoad();
    } catch (e) {
      if (wasCancelled(e)) {
        // 17 Nov 2026: the first check was stopped with Cancel: back to the upload screen, with one calm
        // sentence. The workbook is saved (uploaded, never checked): the list below offers it again.
        this.pending.set(null);
        this.scanNote.set(STOPPED_NOT_CHECKED);
      } else this.uploadError.set(errorText(e));
      this.phase.set("idle");
    } finally {
      poll.stop();
      this.scanningId.set(null);
    }
  }

  private afterLoad(): void {
    this.phase.set("loaded");
    this.pending.set(null);
    this.resetOutside();
    this.fixAll.set(idle());
    this.fixFormulas.set(idle());
    this.loop.set(idle());
    this.recalc.set(idle());
    this.prepare.set(idle());
    this.steps.set([]);
    this.say("");
    void this.router.navigate([], { queryParams: { session: this.store.sessionId() }, replaceUrl: true });
    const loop = this.loadLoopPlan();
    // the chain needs the loop plan first: it decides whether the loop step runs (never after a Reopen)
    const auto = this.autoRun() && !this.reopening;
    this.reopening = false;
    if (auto) void loop.then(() => this.runAll());
  }

  setAutoRun(on: boolean): void {
    this.autoRun.set(on);
    try {
      localStorage.setItem(AUTO_RUN_KEY, on ? "1" : "0");
    } catch {
      /* storage blocked: the box holds for this  visit only */
    }
  }

  private pollScan(id: string): { stop: () => void } {
    let live = true;
    const tick = async () => {
      if (!live) return;
      try {
        const s = await this.api.scanStatus(id);
        if (live) this.scan.set(s);
      } catch {
        /* keep the last status */
      }
      if (live) this.timers.push(setTimeout(tick, 700));
    };
    void tick();
    return { stop: () => (live = false) };
  }

  toggleSkip(name: string): void {
    const next = new Set(this.skip());
    if (next.has(name)) next.delete(name);
    else next.add(name);
    this.skip.set(next);
  }

  mb(bytes: number): string {
    return (bytes / 1048576).toFixed(1) + " MB";
  }

  cancelChoose(): void {
    this.pending.set(null);
    this.phase.set("idle");
  }

  /** 1.14.0: a saved workbook opened again from the Recent workbooks list, with everything it had. The
   *  same screens as after its last step: the constructor's effect shows the loaded screen and picks up
   *  the results of Fix all issues, the outside files and next quarter from the server. A workbook that
   *  was uploaded but never checked goes where an upload goes: the sheet choice, or straight to the
   *  scan. Nothing runs by itself: Do everything is for uploads only. */
  openSaved(started: SessionStart | PendingUpload): void {
    this.uploadError.set("");
    this.scanNote.set("");
    this.scan.set(null);
    this.pickedInputs.set(started.inputs ?? []);
    this.reopening = true;
    if (isPending(started)) {
      if (started.size?.above_threshold && started.size.sheets?.length > 1) {
        this.pending.set(started);
        this.skip.set(new Set());
        this.phase.set("choose");
      } else {
        void this.scanSession(started.sessionId, []);
      }
      return;
    }
    this.pending.set(null);
    this.resetOutside();
    this.fixAll.set(idle());
    this.fixFormulas.set(idle());
    this.loop.set(idle());
    this.recalc.set(idle());
    this.prepare.set(idle());
    this.steps.set([]);
    this.say("");
    this.store.setSession(started);
    void this.router.navigate([], { queryParams: { session: started.sessionId }, replaceUrl: true });
  }

  loadAnother(): void {
    if (this.busy()) return;
    this.store.reset();
    this.loopPlan.set(null);
    this.resetOutside();
    this.pickedInputs.set([]);
    this.phase.set("idle");
    void this.router.navigate([], { queryParams: {}, replaceUrl: true });
  }

  // ---- bring in  the outside files (the model's  macros open other files) ---------------------------
  private resetOutside(): void {
    this.outside.set(idle());
    this.outsideResult.set(null);
    this.inputsError.set("");
    this.cardDragOver.set(false);
    this.outsideFailed.set(false);
    this.outsideUnchecked.set(false);
    this.outsideWaiting.set(false);
    this.chainStandIns = false;
    this.resetNextQuarter();
  }

  /** Another workbook (or  none): the last  next-quarter job's  line, files and zip go, and its polling stops. */
  private resetNextQuarter(): void {
    this.nextQuarterSeq++;
    this.nextQuarter.set(idle());
    this.nextQuarterResult.set(null);
    this.nextQuarterSession.set(null);
  }

  rowLabel(r: OutsideInputRow): string {
    return outsideRowLabel(r);
  }

  /** The warning while stand-ins the assistant made are in use ("" when none). */
  standInLine(o: OutsideInputs): string {
    return outsideStandInLine(o);
  }

  /** The app can run the model's own import: it knows the  macro and the cells that name the files. */
  canBringIn(o: OutsideInputs): boolean {
    return o.rows.length > 0 && !!o.macro;
  }

  /** The card's sentence: how  the picked files match, or why  the app cannot bring them in. */
  outsideLead(o: OutsideInputs): string {
    const cannot = "the app cannot bring them in, so their data stays inside the model.";
    if (!o.rows.length) return `${o.macro ? `The macro ${o.macro} opens` : "The workbook's macros open"} outside files, but no cell of the workbook names them: ${cannot}`;
    if (!o.macro) return `The workbook's macros open outside files, but which macro brings them in is not clear: ${cannot}`;
    return o.summary;
  }

  /** "Try again"  after a failed run; with files missing, it says it brings  in only the matched ones;
   *  "Bring in again"  when the current line was brought in before (files changed since). */
  outsideButton(o: OutsideInputs): string {
    return outsideButtonLabel(o, this.outsideFailed());
  }

  /** Why Do everything skips the outside step ("" when it runs it): only when the server says the
   *  current line has the files brought in (a bring-in that finished, with the files picked now). It
   *  runs again after a failed or half-way bring-in, and after input files were replaced or added. */
  private outsideSkip(): string {
    return outsideSkipReason(this.store.outsideInputs(), this.outsideUnchecked());
  }

  /** Bring in the  outside files: the model's own import macro runs on a copy with the matched files, then
   *  each file's area becomes a Mind input (a background job, followed every second). `standIns`: the
   *  assistant first makes the missing files from the numbers the model holds ("Make them with AI").
   *  From Do everything (`strict`) missing files stop the chain at this step: the server's job waits for
   *  them (1.14.0, `waitForInputs`) and the box under Do everything asks -- pick them and Go on, or Make
   *  them with AI (the  stand-ins, on purpose); nothing runs  meanwhile and the chain goes on from  here. */
  async runOutside(strict = false, standIns = false): Promise<boolean> {
    const id = this.store.sessionId();
    const o = this.store.outsideInputs();
    if (!id || !o || this.outside().running) return false;
    this.outsideFailed.set(false);
    this.outsideUnchecked.set(false);
    if (!this.canBringIn(o)) {
      this.outside.set({ running: false, fraction: null, line: "The app cannot bring these files in", detail: this.outsideLead(o), tone: "error" });
      return false;
    }
    const missing = o.rows.filter((r) => !r.file).map((r) => outsideRowLabel(r));
    if (strict && missing.length && this.chainStandIns) standIns = true; // "Make them with  AI" was  pressed in the box, after a reload
    this.chainStandIns = false;
    if (!missing.length) standIns = false;
    const wait = strict && missing.length > 0 && !standIns;
    if (!o.matched && !standIns && !wait) {
      this.outside.set({ running: false, fraction: null, line: "No input file matched yet", detail: outsideAddLine(o), tone: "error" });
      return false;
    }
    this.inputsError.set("");
    this.outside.set({
      running: true,
      fraction: null,
      line: wait ? `Waiting for ${missing.length} input file${missing.length === 1 ? "" : "s"}: ${nameList(missing)}` : standIns ? "Making the missing input files with the assistant…" : "Starting…",
      detail: "",
      tone: "idle",
    });
    try {
      await this.api.startOutsideInputs(id, standIns, wait);
    } catch (e) {
      // refused: it may run already (started before a reload, or in another tab): then follow that run
      const st = await this.api.outsideInputs(id).catch(() => null);
      if (!st?.running) {
        this.outside.set({ running: false, fraction: null, line: "Not started", detail: errorText(e), tone: "error" });
        this.outsideFailed.set(true);
        return false;
      }
      return await this.followOutside(id, true);
    }
    return await this.followOutside(id);
  }

  /** Follows the job to its end, a look every  second. True when the files came in and the check passed.
   *  `seenRunning`:  the caller saw the job running  (else a status without it waits a  few looks). */
  private async followOutside(id: string, seenRunning = false): Promise<boolean> {
    if (this.outsideFollowing) return false;
    this.outsideFollowing = true;
    const t0 = Date.now();
    let looks = 0;
    let failures = 0;
    try {
      for (;;) {
        if (this.destroyed) return false; // the page was left: the next one picks the job  up
        let st: OutsideRunStatus;
        try {
          st = await this.api.outsideInputs(id);
          failures = 0;
        } catch (e) {
          if (++failures < 5) {
            await this.sleep(1000);
            continue;
          }
          this.outside.set({ running: false, fraction: null, line: "Lost the job", detail: errorText(e), tone: "error" });
          this.outsideFailed.set(true);
          return false;
        }
        looks++;
        if (st.outsideInputs) this.store.outsideInputs.set(st.outsideInputs);
        if (st.waiting) {
          // Do everything waits for the input files:  the box asks, this look-every-second holds the
          // chain here until Go on / Make them with AI (the job runs) or Stop (the job ends, an error)
          seenRunning = true;
          this.showOutsideWaiting(st);
          await this.sleep(1000);
          continue;
        }
        this.endWait(); // the wait is over: the job runs, or ended
        if (st.running) {
          seenRunning = true;
          this.showOutsideRunning(st, t0);
          await this.sleep(1000);
          continue;
        }
        // a server that shows the job a moment after it answered: a few more looks first
        if (!seenRunning && looks < 3) {
          await this.sleep(1000);
          continue;
        }
        await this.reloadSession(id); // the job's new versions, also when it stopped half-way
        if (st.state === "cancelled" || st.state === "stopped") {
          // 17 Nov 2026: stopped with  Cancel -- a running bring-in ("cancelled": its versions are gone) or
          // the wait for the input files ("stopped"). Not an error: one calm sentence, no "Try again".
          await this.refreshOutside(id);
          this.showOutsideCancelled();
          return false;
        }
        if (st.error) {
          this.showOutsideError(st.error);
          return false;
        }
        if (!st.result) {
          this.outside.set({ running: false, fraction: null, line: "The job ended without a result", detail: "", tone: "error" });
          this.outsideFailed.set(true);
          return false;
        }
        return this.showOutsideResult(st.result);
      }
    } finally {
      this.outsideFollowing = false;
    }
  }

  /** The running job on the card: its stage (or its latest line, when that says the stage in full),
   *  then the latest line and the time it has taken. */
  private showOutsideRunning(st: OutsideRunStatus, t0: number): void {
    const secs = Math.round(st.elapsed_s ?? (Date.now() - t0) / 1000);
    const took = secs < 60 ? `${secs} s so far` : `${Math.floor(secs / 60)} min ${secs % 60} s so far`;
    const stage = st.stage || "Working…";
    const said = st.message && st.message !== stage ? st.message : "";
    const full = !!said && said.startsWith(stage);
    this.outside.set({ running: true, fraction: null, line: full ? said : stage, detail: [full ? "" : said, took].filter(Boolean).join(" · "), tone: "idle" });
  }

  /** The wait on the card and in the chain's step: the server's line ("Waiting for 2 input files: ..."),
   *  and the box under Do everything (outsideWaiting). The step counts as  running: the chain is held  here. */
  private showOutsideWaiting(st: OutsideRunStatus): void {
    if (!this.waitT0) this.waitT0 = Date.now();
    this.outsideWaiting.set(true);
    this.outside.set({ running: true, fraction: null, line: st.message || st.stage || "Waiting for the input files", detail: "See the box under Do everything.", tone: "idle" });
  }

  /** The box's "Go on" (every file picked)  or "Make them  with AI"  (`makeStandIns`). While the chain is
   *  held at the wait, the server's job is told to go on and the chain's look-every-second picks the run
   *  up; after a reload (no chain in this page)  Do everything starts again: its first step finds every
   *  file there (or  takes the stand-in path) and the  rest follows. */
  async goOn(makeStandIns = false): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.outsideWaiting()) return;
    if (!makeStandIns && this.waitingMissing().length) return;
    if (this.goingOn()) return; // a second click while the first  request is on its way
    this.inputsError.set("");
    if (this.chainRunning()) {
      this.goingOn.set(true);
      try {
        await this.api.continueOutsideInputs(id, makeStandIns);
        this.endWait(); // the job runs: the box goes now,  not at the next look
      } catch (e) {
        this.inputsError.set(errorText(e)); // still waiting: the  server says what is missing
      } finally {
        this.goingOn.set(false);
      }
      return;
    }
    this.endWait();
    this.chainStandIns = makeStandIns;
    await this.runAll();
  }

  /** The wait for the input files is over (Go on, Make them with AI, Stop, or the server's job moved
   *  on): the box goes, and the time waited is taken out of the running step's own time. */
  private endWait(): void {
    if (this.outsideWaiting()) this.outsideWaiting.set(false);
    if (this.waitT0) {
      if (this.chainRunning()) {
        this.stepT0 += Date.now() - this.waitT0;
        // the countdown and Basic's bar are worked out again from now: not with the wait counted as the step's time
        this.clock.set(Date.now());
      }
      this.waitT0 = 0;
    }
  }

  /** The box's Cancel  (it said "Stop" until 17 Nov 2026): the server's wait ends (nothing was changed) and
   *  the chain stops at this step with the sentence of a Cancel -- nothing had finished, so "Your workbook
   *  is as it was before you pressed Do everything." */
  async stopWaiting(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.outsideWaiting()) return;
    try {
      await this.api.stopOutsideInputs(id);
    } catch (e) {
      this.inputsError.set(errorText(e));
      return;
    }
    if (!this.chainRunning()) {
      // no chain holds here (after a reload): the card and the line say it stopped, as the chain would
      this.endWait();
      this.showOutsideCancelled();
      this.chainStoppedByCancel.set(true);
      this.say(chainStoppedLine(null));
      this.endPlace.set({ id: PLACE_OF_STEP.outside });
    }
  }

  /** 17 Nov 2026: the bring-in (or its wait for the files) was stopped with Cancel: the card says so in one
   *  calm sentence --  not an error, no "Try  again" -- and a chain that runs  this step ends  as cancelled. */
  private showOutsideCancelled(): void {
    this.outside.set(stoppedState());
    this.outsideFailed.set(false);
    this.outsideUnchecked.set(false);
    this.stepEnd = "cancelled";
  }

  /** The  result on the card (and in the chain): the summary, then the details and the new versions. */
  private showOutsideResult(r: OutsideRunResult): boolean {
    this.outsideResult.set(r);
    const out = outsideOutcome(r);
    const labels = (r.versions ?? []).map((v) => this.store.versions().find((x) => x.id === v)?.label ?? "");
    this.outside.set({ running: false, fraction: null, line: out.line, detail: [out.detail, ...labels].filter(Boolean).join(" · "), tone: out.tone });
    this.outsideFailed.set(out.tone === "error"); // the macro stopped early, or a value changed: "Try again"
    this.outsideUnchecked.set(out.unchecked);
    return out.ok;
  }

  /** A job that stopped (the server's  words: why, and which version holds the outside data): its first
   *  clause  as the card's line, all of  it in the details (open), and "Try  again" on the card. */
  private showOutsideError(error: string): void {
    this.outside.set({ running: false, fraction: null, line: firstClause(error) || "Could not bring in the outside files", detail: error, tone: "error", open: true });
    this.outsideFailed.set(true);
    this.outsideUnchecked.set(false);
  }

  /** The job may still run from before a  reload or a visit to Advanced (follow  it), or its result is there to show. */
  private async resumeOutside(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.store.outsideInputs() || this.outsideFollowing) return;
    let st: OutsideRunStatus;
    try {
      st = await this.api.outsideInputs(id);
    } catch {
      return; // nothing to pick up
    }
    if (st.outsideInputs) this.store.outsideInputs.set(st.outsideInputs);
    // the card then follows the server: a result that no longer describes the current version (files
    // replaced since, another version chosen) gives way to why (outsideCard)
    if (st.waiting) {
      // Do everything was waiting when the page was left or reloaded: the box asks again (no chain
      // runs in this page: Go on starts Do everything again, which  goes straight through this step)
      this.showOutsideWaiting(st);
      this.outside.update((s) => ({ ...s, running: false }));
    } else if (st.running) {
      this.showOutsideRunning(st, Date.now());
      await this.followOutside(id, true);
    } else if (st.state === "cancelled" || st.state === "stopped") {
      /* 17 Nov 2026: the last job was stopped with Cancel: nothing was changed, the card is as before it */
    } else if (st.error) this.showOutsideError(st.error);
    else if (st.result) this.showOutsideResult(st.result);
  }

  /** The match and `brought_in` as the server has them now (a version restored under Advanced changes them). */
  private async refreshOutside(id: string): Promise<void> {
    try {
      const st = await this.api.outsideInputs(id);
      if (st.outsideInputs) this.store.outsideInputs.set(st.outsideInputs);
    } catch {
      /* keep what the scan said */
    }
  }

  /** After a background job made new versions: the session as the server holds it now, and the loop
   *  plan of the new current version. */
  private async reloadSession(id: string): Promise<void> {
    try {
      const s = await this.api.resume(id);
      this.store.applyAnalysis({ ...s, version: s.version });
    } catch {
      /* the next step's answer brings the new version */
    }
    await this.loadLoopPlan();
  }

  onAddInputs(event: Event): void {
    const input = event.target as HTMLInputElement;
    const files = Array.from(input.files ?? []);
    input.value = "";
    if (files.length) void this.addInputs(files);
  }

  onCardDragOver(event: DragEvent): void {
    event.preventDefault(); // always: a file dropped on the card must not open  in the browser
    const o = this.store.outsideInputs();
    this.cardDragOver.set(!!o && this.canBringIn(o));
  }

  onCardDrop(event: DragEvent): void {
    event.preventDefault();
    this.cardDragOver.set(false);
    const files = Array.from(event.dataTransfer?.files ?? []);
    if (files.length) void this.addInputs(files);
  }

  /** Add files: more  input files (one of the same name  replaces the old  one); the  server matches them
   *  again. After a bring-in, a file that changes what the model reads (a corrected file, a missing one)
   *  makes the server say "not brought in" (with why): the card offers  "Bring in again", Do everything
   *  runs the step again. */
  async addInputs(files: File[]): Promise<void> {
    const id = this.store.sessionId();
    const o = this.store.outsideInputs();
    if (!id || !o || !this.canBringIn(o)) return;
    if (this.busy() && !this.outsideWaiting()) {
      // (while Do everything waits  for the input files, picking them  is the point)
      this.inputsError.set(`Wait: ${this.busyHint()}.`);
      return;
    }
    const problem = inputFilesProblem(files.map((f) => f.name));
    if (problem) {
      this.inputsError.set(problem);
      return;
    }
    this.inputsError.set("");
    this.addingInputs.set(true);
    try {
      const out = await this.api.addInputs(id, files);
      this.store.applyAnalysis({ outsideInputs: out.outsideInputs, readiness: out.readiness ?? undefined });
      if (this.outside().tone === "error") {
        // the last stop was about  the files before
        this.outside.set(idle());
        this.outsideFailed.set(false);
      }
    } catch (e) {
      this.inputsError.set(errorText(e));
    } finally {
      this.addingInputs.set(false);
    }
  }

  // ---- fix all issues -------------------------------------------------------------------------
  /** A fixer may still be  running from before a reload or a  visit to Advanced: pick it up. */
  private async resumeFixer(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.fixAllFollowing) return;
    try {
      const s = await this.api.autoFixStatus(id, false);
      if (s.state === "running") {
        this.showAutoFix(s);
        await this.followAutoFix(id);
      } else if (s.state === "error" && s.error && /the app was closed/i.test(s.error)) {
        // 1.14.0: a saved workbook reopened after a restart that cut Fix all issues: say so (what it had
        // fixed before is in its version, when the run got that far)
        this.fixAll.set({ running: false, fraction: null, line: "Stopped: the app was closed while it ran", detail: [s.error, s.version?.label ? `What it fixed before is in ${s.version.label}` : "Press Fix all issues to run it again"].join(" · "), tone: "error" });
      }
    } catch {
      /* no fixer, nothing to resume */
    }
  }

  async runFixAll(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.fixAll().running) return false;
    this.fixAll.set({ running: true, fraction: 0, line: "Starting…", detail: "", tone: "idle" });
    try {
      const status = await this.api.startAutoFix(id, { use_assistant: true, keep_good_values: true });
      this.showAutoFix(status);
      return await this.followAutoFix(id);
    } catch (e) {
      // The backend says why in one sentence (already running, a recalculation runs, being analysed).
      const why = errorText(e);
      if (/already running/i.test(why)) {
        this.fixAll.set({ running: true, fraction: null, line: "Still running from before", detail: why, tone: "idle" });
        return await this.followAutoFix(id);
      }
      this.fixAll.set({ running: false, fraction: null, line: "Not started", detail: why, tone: "error" });
      return false;
    }
  }

  async stopFixAll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    try {
      await this.api.stopAutoFix(id);
      this.fixAll.update((s) => ({ ...s, stopping: true, line: "Stopping after this round…" }));
    } catch (e) {
      this.fixAll.update((s) => ({ ...s, detail: errorText(e) }));
    }
  }

  private showAutoFix(s: { stage?: string; title?: string; message?: string; fraction?: number; pass_no?: number; errors?: number; errors_before?: number }): void {
    const stopping = this.fixAll().stopping;
    const line = stopping ? "Stopping after this round…" : s.title || s.stage || "Working…";
    const detail = [s.pass_no ? `pass ${s.pass_no}` : "", s.errors_before != null && s.errors != null ? `${s.errors_before} → ${s.errors} error cells` : "", s.message || ""]
      .filter(Boolean)
      .join(" · ");
    this.fixAll.set({ running: true, fraction: s.fraction ?? null, line, detail, tone: "idle", stopping });
  }

  /** Follows the fixer to its end. Returns true when it ended  clean. */
  private async followAutoFix(id: string): Promise<boolean> {
    if (this.fixAllFollowing) return false;
    this.fixAllFollowing = true;
    try {
      for (;;) {
        await this.sleep(1500);
        let s;
        try {
          s = await this.api.autoFixStatus(id, false);
        } catch (e) {
          this.fixAll.set({ running: false, fraction: null, line: "Lost the fixer", detail: errorText(e), tone: "error" });
          return false;
        }
        if (s.state === "running") {
          this.showAutoFix(s);
          continue;
        }
        if (s.state === "cancelled") {
          // 17 Nov 2026: stopped with Cancel: no version, nothing kept (not "Stop, keep what is fixed",
          // which ends "done"). Nothing of this answer goes into the store: the workbook as it is now
          // comes from the Cancel's own answer.
          this.fixAll.set(stoppedState());
          this.stepEnd = "cancelled";
          return false;
        }
        const full = await this.api.autoFixStatus(id, true).catch(() => s);
        if (full.analysis) this.store.applyAnalysis({ ...full.analysis, version: full.version ?? undefined });
        else if (full.version) this.store.addVersion(full.version);
        const r = full.result;
        if (s.state === "error" || r?.status === "error") {
          this.fixAll.set({ running: false, fraction: null, line: "Stopped with an error", detail: full.error || r?.summary || "", tone: "error" });
          return false;
        }
        const before = r?.errors_before ?? full.errors_before;
        const after = r?.errors_after ?? full.errors;
        const left = r?.left ?? [];
        // "unchanged" with no error left: the workbook already recalculated clean, nothing to fix
        const clean = r?.status === "clean" || ((r?.status === "partial" || r?.status === "unchanged") && left.length === 0 && (after ?? 0) === 0);
        // 1.11.6: the errors left could not be fixed without changing a good value, and the
        // preparation caused none of them (no regression): the workbook was already living with
        // them, so the chain goes on (Recalculate  then compares with the original).  A Stop or a
        // used-up time budget still  stops it.
        const acceptable = !clean && !!r && ["partial", "stuck", "unchanged"].includes(r.status) && !r.stopped && (r.regression_cells ?? 0) === 0;
        // On a  prepared version  "no regression" means something only when the comparison with the
        // original ran (none is made on the original itself): when Excel failed it, nothing was checked.
        const compared = r?.numbers_check == null || r.numbers_check.ran !== false;
        const accepted = acceptable && compared;
        const kept = after ?? left.length;
        const line = r?.status === "clean" ? "Every formula error is fixed" : accepted ? `Fixed what it could; ${kept} error${kept === 1 ? "" : "s"} the workbook already had stay, moving on` : acceptable ? "The errors left could not be checked against the original (the comparison did not run)" : r?.status === "unchanged" ? (clean ? "No formula error to fix" : "Nothing could be changed") : `Fixed what it could (${r?.status ?? "partial"})`;
        const shown = left.slice(0, 3);
        const leftText = left.length ? `${left.length} left for a person: ` + shown.map((l) => `${l.sheet}!${l.cell} ${l.error}`).join(", ") + (left.length > 3 ? ", …" : "") + " (Advanced › Recalculate)" : "";
        // what each error value shown means, once  ("#REF! means the formula points at a cell that no longer  exists")
        const meanings = meaningsText(this.codes.legend(shown.map((l) => l.error)));
        const detail = [before != null && after != null ? `${before} → ${after} error cells` : "", r?.cells_fixed != null ? `${r.cells_fixed} cells rewritten` : "", full.version ? full.version.label : "", leftText, meanings].filter(Boolean).join(" · ");
        this.fixAll.set({ running: false, fraction: null, line, detail, tone: clean ? "ok" : "warn", unchecked: acceptable && !compared });
        return clean || accepted;
      }
    } finally {
      this.fixAllFollowing = false;
    }
  }

  // ---- fix the formulas ----------------------------------------------------------------------
  async runFixFormulas(): Promise<boolean> {
    const id = this.store.sessionId();
    const ops = this.formulaOps();
    if (!id || this.fixFormulas().running) return false;
    if (!ops.length) {
      this.fixFormulas.set({ running: false, fraction: null, line: "No formula to change", detail: "The current version has no formula the Mind rules would rewrite.", tone: "ok" });
      return true;
    }
    this.fixFormulas.set({ running: true, fraction: null, line: `Writing ${ops.length} changes through Excel…`, detail: "", tone: "idle" });
    const stop = this.pollApply(id, this.fixFormulas);
    const errorsBefore = this.errorCount();
    try {
      const out = await this.api.apply(id, ops, "major");
      stop();
      this.store.applyAnalysis({ ...out, version: out.version });
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      const newErrors = Math.max(0, this.errorCount() - errorsBefore);
      const ok = refused === 0 && newErrors === 0;
      this.fixFormulas.set({
        running: false,
        fraction: null,
        line: refused ? `${done} changed, ${refused} refused` : `${done} formulas changed`,
        detail: [out.version?.label ?? "", r.verified_opens_in_excel ? "opens in Excel" : "", failedText(r), newErrors ? `${newErrors} new error findings` : ""].filter(Boolean).join(" · "),
        tone: ok ? "ok" : "warn",
      });
      return ok;
    } catch (e) {
      stop();
      if (wasCancelled(e)) {
        // 17 Nov  2026: stopped with Cancel: no error, and  nothing of this call reaches the store
        this.fixFormulas.set(stoppedState());
        return false;
      }
      this.fixFormulas.set({ running: false, fraction: null, line: "Could not write the changes", detail: errorText(e), tone: "error" });
      return false;
    }
  }

  // ---- fix the formulas and  the structure (the chain's second step) --------------------------------
  /** 17  Nov 2026 (never stop for a person when the app can fix it): the step is one job on the server
   *  (POST .../prep-run), followed every second. The job writes every blocking Prep action that is on by
   *  default, in rounds while a round still changes something; then it hands what is left to the
   *  assistant, and keeps a fix of the assistant only when  no number changed.
   *  True when nothing is left ("done"), and for a question that does not stop Start (the page keeps it
   *  for the chain's end). False for a question that stops Start, no AI key, a Cancel and an error.
   *  Until tonight the page wrote the rounds itself, three at most, and stopped with a count ("954
   *  changes applied") when a fourth was needed. */
  async runPrepare(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.prepare().running) return false;
    this.stepAsk = null;
    this.prepDid = { applied: 0, fixed: 0 };
    this.prepare.set({ running: true, fraction: null, line: "Looking at what blocks the file…", detail: "", tone: "idle" });
    try {
      // the job writes the button's name into its sentences ("... then press Start again.")
      await this.api.startPrepRun(id, { button: this.basic ? "Start" : "Do everything" });
    } catch (e) {
      // refused: it may run already (started before a  reload, or in another tab):  then follow that run
      let st: PrepRunStatus | null = null;
      try {
        st = await this.api.prepRun(id);
      } catch {
        st = null;
      }
      if (st?.state !== "running") {
        this.prepare.set({ running: false, fraction: null, line: "Not started", detail: errorText(e), tone: "error" });
        return false;
      }
    }
    return await this.followPrepRun(id);
  }

  /** Follows the job to  its end, a  look every  second. The bar of the round being written is the Apply's
   *  own progress, as before (pollApply). No look ends at a click on Cancel: the job says "cancelled"
   *  itself, seconds later. */
  private async followPrepRun(id: string): Promise<boolean> {
    if (this.prepFollowing) return false;
    this.prepFollowing = true;
    const stopBar = this.pollApply(id, this.prepare);
    let failures = 0;
    try {
      for (;;) {
        if (this.destroyed) return false; // the page was  left: the next one picks the job up
        let st: PrepRunStatus;
        try {
          st = await this.api.prepRun(id);
          failures = 0;
        } catch (e) {
          if (++failures < 5) {
            await this.sleep(1000);
            continue;
          }
          stopBar();
          this.prepare.set({ running: false, fraction: null, line: "Lost the job", detail: errorText(e), tone: "error" });
          return false;
        }
        if (st.state === "running") {
          this.showPrepRun(st);
          await this.sleep(1000);
          continue;
        }
        stopBar();
        return await this.endPrepRun(id, st);
      }
    } finally {
      stopBar();
      this.prepFollowing = false;
    }
  }

  /** The running job on  the step: its one plain sentence  ("Round 2: writing 966 changes through Excel…",
   *  "The  assistant is working on what is left (1 of 2)."). While  a round or the assistant's fix is  being
   *  written, the bar and the words under it are the Apply's own; in the other stages (the assistant, the
   *  numbers check, the check of the workbook) the bar only moves. */
  private showPrepRun(st: PrepRunStatus): void {
    const writing = !st.stage || st.stage === "pass" || st.stage === "write";
    this.prepare.update((s) => ({ ...s, running: true, tone: "idle", line: st.message || st.title || s.line, ...(writing ? {} : { fraction: null, detail: "" }) }));
  }

  /** How the job ended, on the step; the workbook as the server holds it now goes into the store first
   *  (the job made versions of its own). */
  private async endPrepRun(id: string, st: PrepRunStatus): Promise<boolean> {
    if (st.state === "cancelled") {
      // stopped with Cancel: every round and every fix  of this run was thrown away on the server. Nothing
      // of this answer goes into the store: the workbook as it is now comes from the Cancel's own answer.
      this.prepare.set(stoppedState());
      this.stepEnd = "cancelled";
      return false;
    }
    try {
      const s = await this.api.resume(id);
      if (this.store.sessionId() === id) this.store.applyAnalysis({ ...s, version: s.version });
    } catch {
      /* the next step's answer brings the new version */
    }
    const applied = st.applied ?? 0;
    const refused = st.refused ?? 0;
    const fixed = st.assistant_fixed ?? 0;
    this.prepDid = { applied, fixed };
    this.prepTotal = { applied: this.prepTotal.applied + applied, fixed: this.prepTotal.fixed + fixed };
    const did = prepDoneWords(applied, fixed);
    const version = this.store.currentVersion()?.label ?? "";
    const refusals = refused ? `Excel refused ${refused} change${refused === 1 ? "" : "s"}` : "";
    if (st.state === "done") {
      this.pendingAsk = null; // nothing is left: a question from before (the Prep step's, asked before the loop) is answered
      this.prepare.set({ running: false, fraction: null, line: did || "Nothing blocked the file", detail: [version, refusals].filter(Boolean).join(" · "), tone: "ok" });
      return true;
    }
    const sent = st.ask?.lines?.length ? st.ask : null;
    const ask = sent ?? (st.state === "no_key" ? prepNoKeyAsk(this.basic ? "Start" : "Do everything") : null);
    if ((st.state === "ask" || st.state === "no_key") && ask) {
      // Start stops here only where it stopped before tonight: a Prep fix is still left (the server says
      // so; when it does not, the plan does). With no AI key it stops and  says so.
      const stops = st.state === "no_key" || (ask.stops_start ?? prepWorkLeft(this.store.plan()));
      this.prepare.set({ running: false, fraction: null, line: did || ask.lines[0], detail: [version, refusals, ask.why ?? ""].filter(Boolean).join(" · "), tone: "warn" });
      if (stops) {
        this.stepAsk = ask;
        return false;
      }
      this.pendingAsk = ask;
      return true;
    }
    if (st.state === "ask") {
      // a question with no  words: something is left, and the screen  cannot say what
      this.prepare.set({ running: false, fraction: null, line: did || "Something is left that the app could not fix", detail: [version, refusals].filter(Boolean).join(" · "), tone: "warn" });
      return false;
    }
    if (st.state === "error") {
      // the  job itself  failed (the rounds and fixes that had finished stay). The server's sentence says
      // what: a round that could not be written, or one that was written and whose check did not run  --
      // so the line does not guess ("Could not write the changes"  would be untrue for the second).
      this.prepare.set({ running: false, fraction: null, line: "Stopped with an error", detail: st.error || st.message || "", tone: "error" });
      return false;
    }
    this.prepare.set({ running: false, fraction: null, line: "The job ended without a result", detail: "", tone: "error" });
    return false;
  }

  /** A  Prep job may still be running from  before a reload or a visit to Advanced: pick it up,  as the
   *  fixer is. No chain runs in  this page then (nothing follows by itself): the step stands alone in
   *  Custom's list, with its own Cancel under the bar, and at its end the line says how it ended -- the
   *  question with "Show me" and "Send feedback", as at the end of Start. */
  private async resumePrepRun(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.prepFollowing || this.chainRunning()) return;
    let st: PrepRunStatus;
    try {
      st = await this.api.prepRun(id);
    } catch {
      return; // nothing to pick up
    }
    if (st.state !== "running" || this.destroyed || this.prepFollowing || this.chainRunning() || this.store.sessionId() !== id) return;
    const step = STEPS.find((s) => s.id === "prepare")!;
    this.steps.set([{ ...step, state: "running", line: "", detail: "" }]);
    this.chainStoppedByCancel.set(false);
    this.say("");
    this.stepAsk = null;
    this.pendingAsk = null;
    this.takeStepEnd();
    this.prepare.set({ running: true, fraction: null, line: st.message || "Still running from before", detail: "", tone: "idle" });
    const ok = await this.followPrepRun(id);
    if (this.destroyed) return;
    const r = this.prepare();
    const cancelled = this.takeStepEnd() === "cancelled";
    const ask = this.takeStepAsk() ?? this.pendingAsk;
    this.pendingAsk = null;
    this.setStep("prepare", { state: cancelled ? "stopped" : ok ? "done" : "failed", line: r.line, detail: r.detail });
    if (cancelled) {
      // Basic draws no list: its one line says it (no "Show me": the person stopped  it)
      if (this.basic) {
        this.chainStoppedByCancel.set(true);
        this.say(STOPPED_NOTHING_CHANGED);
        this.endPlace.set({ id: PLACE_OF_STEP.prepare });
      }
    } else if (ask) this.say("", ask);
    else if (!ok) this.say(`Stopped at "${step.label}": ${r.line.replace(/\.$/, "")}. ${stoppedWhere("prepare", r, this.basic)}`);
  }

  private pollApply(id: string, target: typeof this.fixFormulas): () => void {
    let live = true;
    const tick = async () => {
      // (17 Nov 2026: nor on a page that was left: the Prep job is followed for minutes, and a look that
      // was on its way when the page went would otherwise go on asking for ever)
      if (!live || this.destroyed) return;
      try {
        const s = await this.api.scanStatus(id);
        const a = s.apply;
        if (live && a && a.state === "running") target.update((x) => ({ ...x, fraction: a.fraction ?? null, detail: [a.title, a.message].filter(Boolean).join(" · ") }));
      } catch {
        /* ignore */
      }
      if (live && !this.destroyed) this.timers.push(setTimeout(tick, 600));
    };
    void tick();
    return () => (live = false);
  }

  // ---- the Mind loop -------------------------------------------------------------------------
  private async loadLoopPlan(): Promise<void> {
    const id = this.store.sessionId();
    if (!id) return;
    this.loopPlanReads++;
    try {
      const { plan } = await this.api.mindLoopPlan(id);
      this.loopPlan.set(plan);
    } catch {
      this.loopPlan.set(null);
    } finally {
      this.loopPlanReads--;
    }
  }

  loopSummary(): string {
    const p = this.loopPlan();
    if (!p) return "Reading the model map…";
    if (!p.found) return "No run loop in this workbook: nothing to build.";
    const ready = p.loops.filter((l) => l.status === "ready");
    const parts = ready.map((l) => `${l.name} (${l.first} to ${l.last})`);
    // only a loop that waits for an answer counts: another macro's loops and the ones left as is need none
    const open = p.loops.filter((l) => l.status === "needs_decision").length;
    const asked = open === 1 ? "1 loop needs one answer" : `${open} loops need one answer each`;
    return [parts.length ? `Loop ${parts.join(", ")}: ${p.operation_count} formulas to write` : "", open ? `${asked} (Advanced › Mind loop)` : "", this.headerRow() ? "row 1 is a header row: skipped" : ""].filter(Boolean).join(" · ");
  }

  /** The plan says the first row of the table is a header row: the loop then starts at 2. */
  private headerRow(): boolean {
    const p = this.loopPlan();
    return !!p && p.loops.some((l) => l.first === 1 && (l.notes ?? []).some((n) => /header/i.test(n)));
  }

  /** Basic: the loop step is wanted  only while its box is  ticked; Custom always wants it. */
  loopWanted(): boolean {
    return !this.basic || this.buildLoop();
  }

  /** Basic's one short  line: the step  going on  now, in plain words, with how long is left (an estimate,
   *  said so), else how the last Start ended (in Basic's words). 16 Nov 2026: more words on what is going
   *  on, and  an ETA with a disclaimer (Joseph). */
  readonly basicLine = computed(() => this.basicLines().join(" "));
  /** The same words, one short sentence  each: Basic draws them one under  the other (16 Nov 2026, night). */
  readonly basicLines = computed<string[]>(() => {
    if (this.chainRunning()) {
      if (this.chainStopping()) return ["Stopping…"]; // 17 Nov  2026: Cancel was clicked: no step, no time left
      if (this.outsideWaiting()) return ["Waiting for the input files."];
      const list = this.countedSteps();
      const now = this.chainStep();
      const step = now ? `Step ${list.findIndex((s) => s.id === now.id) + 1} of ${list.length}: ${BASIC_WORDS[now.id]}.` : "Getting ready…";
      return [step, ...this.etaLines()];
    }
    // 17 Nov 2026:  the question the Prep step ended with:  its two sentences, one per line
    const ask = this.chainAsk();
    if (ask) return ask.lines.map(basicWords);
    // (17 Nov 2026: or the line after a Start the page lost, L2b-4; and the assistant's fixes, L2b-9)
    const ended = basicWords(this.endLine());
    const note = this.chainNote();
    // (17 Nov 2026: the Mind loop's first-value line is on the finish banner right below, not here as well)
    return ended ? (note ? [ended, note] : [ended]) : [];
  });
  /** On time: "About 6 min left (around 15:40). An estimate only: big models can take longer."
   *  Late (16 Nov 2026, night; Joseph:  once it says "longer than usual", "an estimate only" makes no sense): one
   *  thing is said and the disclaimer goes: "This step is taking longer than usual; about 6 min left (around
   *  15:40), maybe more." */
  readonly basicEta = computed(() => this.etaLines().join(" "));
  private readonly etaLines = computed<string[]>(() => {
    const t = this.chainLeft();
    // no estimate: still being worked out (the first seconds), or the app has none for this run (then nothing is said)
    if (!t) return this.estimateSettled() ? [] : ["Working out how long this will take."];
    const at = new Date(this.clock() + t.left * 1000);
    const left = `${aboutLeft(t.left)} left (around ${String(at.getHours()).padStart(2, "0")}:${String(at.getMinutes()).padStart(2, "0")})`;
    if (t.slow) return [`This step is taking longer than usual; ${left}, maybe more.`];
    return [`${left[0].toUpperCase()}${left.slice(1)}.`, "An estimate only: big models can take longer."];
  });
  /** Basic's bar while Start runs (null: nothing runs, so no bar). `percent` null: no estimate yet, the bar
   *  only moves. With an estimate it fills by time (worked / (worked  + left)); when the app  has no estimate for
   *  this run it fills  by step (each  finished step  whole, the running one half). It never moves backwards
   *  (an estimate that grows holds the bar where it is) and never reaches 100 before the run has ended.
   *  16 Nov 2026 (night). */
  readonly basicBar = computed<{ percent: number | null } | null>(() => {
    // (17 Nov  2026: nor while a Cancel is stopping the run: a bar  that still fills would promise an  end)
    if (!this.chainRunning() || this.outsideWaiting() || this.chainStopping()) return null;
    const t = this.chainLeft();
    if (!t && !this.estimateSettled()) return { percent: null };
    const now = this.clock();
    const list = this.countedSteps();
    // the time worked: every step that ran,  and the running step so far (the wait for input files  is not in it)
    let worked = 0;
    let finished = 0;
    let running = 0;
    for (const s of list) {
      if (s.state === "running") {
        worked += Math.max(0, (now - this.stepT0) / 1000);
        running = 1;
      } else if (s.state !== "waiting") {
        worked += this.stepSeconds[s.id] ?? 0;
        finished++;
      }
    }
    const byStep = list.length ? (finished + running / 2) / list.length : 0;
    const share = t && t.left > 0 ? worked / (worked + t.left) : byStep;
    this.barHigh = Math.max(this.barHigh, Math.min(99, Math.floor(share * 100)));
    return { percent: this.barHigh };
  });

  /** The line says the  assistant is at work (making stand-ins): the thinking dots go beside it.
   *  (17 Nov 2026, review R2-4: the pattern held two backspace characters where the word boundaries were
   *  meant, so it never matched and the dots were never drawn.) */
  aiWords(text: string | null | undefined): boolean {
    return /\b(stand-in|assistant|with AI)\b/i.test(text ?? "");
  }

  canBuildLoop(): boolean {
    const p = this.loopPlan();
    return !!p && p.found && p.operation_count > 0 && p.loops.some((l) => l.status === "ready");
  }

  async runLoop(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.loop().running || !this.canBuildLoop()) return false;
    this.loop.set({ running: true, fraction: null, line: "Writing the loop through Excel…", detail: "", tone: "idle" });
    const stop = this.pollApply(id, this.loop);
    try {
      const out = await this.api.applyMindLoopPlan(id, this.headerRow() ? { first: "2" } : {});
      stop();
      const { plan: loopPlan, ...analysis } = out; // `plan` here is the LOOP plan, not the Prep actions
      this.store.applyAnalysis({ ...analysis, version: out.version });
      this.loopPlan.set(loopPlan);
      // the Prep plan and the readiness of the new version come with the workbook read again  (not waited
      // for here; `runLoopThenFix` waits for it before it looks at what the loop uncovered)
      this.afterLoop = this.api
        .resume(id)
        .then((s) => this.store.applyAnalysis({ summary: s.summary, report: s.report, plan: s.plan, readiness: s.readiness, prep_progress: s.prep_progress, versions: s.versions }))
        .catch(() => {});
      const r = out.result;
      const done = appliedCount(r);
      const refused = failedCount(r);
      this.loop.set({
        running: false,
        fraction: null,
        line: refused ? `${done} written, ${refused} refused` : `${done} loop formulas written`,
        detail: [out.version?.label ?? "", loopPlan.how_it_runs ?? "", failedText(r)].filter(Boolean).join(" · "),
        tone: refused ? "warn" : "ok",
      });
      return refused === 0;
    } catch (e) {
      stop();
      if (wasCancelled(e)) {
        // 17 Nov 2026: stopped with Cancel: the loop was not written
        this.loop.set(stoppedState());
        this.stepEnd = "cancelled";
        return false;
      }
      this.loop.set({ running: false, fraction: null, line: "Could not write the loop", detail: errorText(e), tone: "error" });
      return false;
    }
  }

  /** After the loop, the Prep job runs again when a blocking Prep action still has changes to write, and
   *  (17 Nov 2026, review R2-3) when the readiness lists any other blocking problem but the outside files:
   *  one the loop made that no Prep action repairs goes to the assistant too, as the Prep step's own do.
   *  Before, such a problem  was never given to the  assistant, and Start ended with "Everything ran  and
   *  every check passed"  above a  banner that said the workbook was blocked. A problem the assistant
   *  already tried in the Prep step costs no new call (the server remembers it). */
  private fixesFollowLoop(): boolean {
    if (prepWorkLeft(this.store.plan())) return true;
    // repair round: only a problem the Prep step  did not have already. One its assistant already tried
    // and could not  fix (Start asks about it at  the end) started a whole Prep job  again, Excel and all,
    // for no new answer.
    const saw = this.prepSaw;
    return (this.store.readiness()?.blocking ?? []).some((b) => b.fix !== "outside_inputs" && !saw?.has(problemKey(b)));
  }

  /** Building the loop changes what  the rules see (which policy row the  model computes, the cells it clears): a
   *  blocking action can show up that was not there before (Palermo, 14 Nov 2026: "an empty cell is compared with
   *  a number" went from PASS to ERROR once the loop was in). So after the loop step: one more Prep pass, the same
   *  rounds as the Prep step, only when something blocking is left. */
  async runLoopThenFix(): Promise<boolean> {
    this.afterLoop = null;
    if (!(await this.runLoop())) return false;
    // 17 Nov 2026: the  Prep plan as it is AFTER the loop. The loop's own answer carries only the loop plan;
    // the Prep plan comes with the workbook read again, which `runLoop` asks for and does not wait for.
    // Without this wait the line below read the plan from before the loop, where nothing was left to fix.
    if (this.afterLoop) await this.afterLoop;
    this.afterLoop = null;
    // 17 Nov 2026: after a Cancel nothing new starts: the loop is built (its own answer came normally), its fixes are not begun
    if (this.destroyed || this.chainCancelled || !this.fixesFollowLoop()) return true;
    const loopLine = this.loop().line;
    this.loop.update((s) => ({ ...s, running: true, fraction: null, line: `${loopLine}; fixing what the loop uncovered…` }));
    this.takeStepEnd();
    // the same job as the Prep step (rounds,  then the assistant); its question  replaces the Prep step's
    const ok = await this.runPrepare();
    const p = this.prepare();
    if (this.takeStepEnd() === "cancelled") {
      // 17 Nov 2026: Cancel  during the fixes  after the  loop: they are thrown away, the loop itself stands
      // (it counts as finished: "Your workbook is as it was after the Mind loop was built.")
      this.stepEnd = "after_loop";
      this.loop.update((s) => ({ ...s, running: false, line: `${loopLine}; the fixes after the loop were stopped` }));
      return false;
    }
    // what the fixes wrote, after the loop's own  line: "...; then 5 changes applied for what the loop
    // uncovered", and  the assistant's sentence when it fixed  something (with its "Show me" to Versions)
    const { applied, fixed } = this.prepDid;
    const then = applied ? `${loopLine}; then ${applied} changes applied for what the loop uncovered` : loopLine;
    const assistant = fixed ? `. The assistant fixed ${fixed}${applied ? " more" : ""}: see Advanced › Versions.` : "";
    this.loop.update((s) => ({
      ...s,
      running: false,
      line: ok ? `${then}${assistant}` : `${loopLine}; the fixes after the loop stopped: ${p.line}`,
      detail: [s.detail, p.detail].filter(Boolean).join(" · "),
      tone: ok ? s.tone : "warn",
    }));
    return ok;
  }

  // ----  recalculate ------------------------------------------------------------------------------
  async runRecalc(): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.recalc().running) return false;
    this.recalc.set({ running: true, fraction: null, line: "Recalculating in Excel…", detail: "Tens of seconds to a few minutes.", tone: "idle" });
    try {
      const out = await this.api.recalculate(id);
      this.store.recalc.set(out);
      this.store.applyAnalysis({ report: out.report, plan: out.plan, readiness: out.readiness, version: out.version ?? undefined });
      // 17 Nov 2026 (Palermo): the cells that show an error  only for the Mind loop's first value  never stop
      // Start (the verdict leaves them out): when they are all the new errors, the step's line is the
      // verdict's one plain line on them; else they are named after the count (review LBr: the count is then
      // the verdict's -- only the new errors that also show without the loop, not the original's own)
      const loopOnly = out.readiness?.recalc?.loop_errors ?? 0;
      const loopNote = out.readiness?.recalc?.loop_note ?? "";
      const newErrors = (out.new_errors ?? out.formula_errors.length) - loopOnly;
      const ok = out.status === "PASS" || newErrors <= 0;
      const line =
        out.status === "PASS" ? "Recalculated, no formula errors"
        : ok && loopOnly && loopNote ? loopNote
        : loopOnly ? `Recalculated, ${newErrors} formula errors (the other ${loopOnly} only for the Mind loop's first value)`
        : `Recalculated, ${out.formula_errors.length} formula errors`;
      // the list is named only when there is one:  after a clean run its "Show me" would lead to the
      // Recalculate button, for another recalculation with nothing to list
      const list = out.formula_errors.length ? "The list is under Advanced › Recalculate." : "";
      const detail = [out.message, out.new_errors != null ? `${out.new_errors} new since the original` : "", out.preexisting_errors != null ? `${out.preexisting_errors} already in the original` : "", out.addin_gap_errors.length ? `${out.addin_gap_errors.length} add-in gap` : "", line === loopNote ? "" : loopNote, list].filter(Boolean).join(" · ");
      this.recalc.set({ running: false, fraction: null, line, detail, tone: ok ? "ok" : "warn" });
      return ok;
    } catch (e) {
      if (wasCancelled(e)) {
        // 17 Nov 2026: stopped with Cancel: the last good recalculation stays as it was
        this.recalc.set(stoppedState());
        this.stepEnd = "cancelled";
        return false;
      }
      this.recalc.set({ running: false, fraction: null, line: "Could not recalculate", detail: errorText(e), tone: "error" });
      return false;
    }
  }

  // ---- do everything --------------------------------------------------------------------------
  /** The whole chain, each step followed by its check, stopping at the first failed check. */
  async runAll(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || this.busy()) return;
    // Basic's bar starts from nothing, and this run's estimate is asked for below
    this.barHigh = 0;
    this.estimateSettled.set(false);
    // 17 Nov 2026: a new run: the last run's Cancel is over
    this.chainCancelled = false;
    this.chainStopping.set(false);
    this.chainStoppedByCancel.set(false);
    this.cancelHeard = null;
    this.cancelHeardNow = null;
    this.chainRunning.set(true);
    this.interrupted.set(false);
    this.interruptedMark = null;
    this.say("");
    // 17 Nov 2026: nor is the last run's question this  run's
    this.pendingAsk = null;
    this.stepAsk = null;
    this.prepTotal = { applied: 0, fixed: 0 };
    this.prepSaw = null;
    // the outside step only for  a model whose macros open outside files
    const plan = STEPS.filter((s) => s.id !== "outside" || this.store.outsideInputs() !== null);
    // 17  Nov 2026 (live  test L2b-6):  a loop step that cannot run is left out from the start, so "Step 2 of
    // 3" is said from the first step (it said "Step 2 of 4", then "Step 3  of 3" when the loop's turn came).
    // (Repair round: only when that  is sure. While the loop plan  is still being read, the step is  counted
    // and asked at its turn: left out then, it keeps its  number, "Step 4 of 4"; before, it  was left out
    // and then ran, and "Step 2 of 3" became "Step 3 of 4".)
    const loopSkip = !this.loopWanted() || this.loopPlanReads === 0 ? this.loopSkip() : "";
    this.steps.set(plan.map((s) => ({ ...s, state: s.id === "loop" && loopSkip ? "skipped" : "waiting", line: s.id === "loop" ? loopSkip : "", detail: "" })));
    // 17 Nov 2026 (live test L2b-4): this tab remembers the run until it ends here, for a reload in its middle
    this.markStep(id, null);
    // the clock starts before the estimate is fetched: until then the timer would show the last run's time
    this.stepSeconds = {};
    this.waitT0 = 0;
    this.chainT0 = this.stepT0 = Date.now();
    this.clock.set(Date.now());
    await this.loadEstimate(id);
    this.estimateSettled.set(true);
    if (this.ticker) clearInterval(this.ticker);
    this.ticker = setInterval(() => this.clock.set(Date.now()), 10_000);
    const runners: Record<StepId, () => Promise<boolean>> = {
      outside: () => this.runOutside(true),
      fixAll: () => this.runFixAll(),
      prepare: () => this.runPrepare(),
      loop: () => this.runLoopThenFix(),
      recalc: () => this.runRecalc(),
    };
    const states: Record<StepId, typeof this.fixAll> = { outside: this.outside, fixAll: this.fixAll, prepare: this.prepare, loop: this.loop, recalc: this.recalc };
    // 17 Nov 2026: the last step that really finished (a  skipped step does not count): the sentence after a
    // Cancel names it ("... as it was after the formula errors were fixed.")
    let finished: StepId | null = null;
    try {
      for (const step of plan) {
        if (this.destroyed) return; // the page was left: the next step would run unseen
        // Cancel before the first step, or between two steps (nothing runs on the server): nothing new starts
        if (this.chainCancelled) return await this.endCancelled(finished, step.id);
        if (step.id === "outside") {
          await this.refreshOutside(id);
          if (this.destroyed) return;
          if (this.chainCancelled) return await this.endCancelled(finished, step.id);
          const why = this.outsideSkip();
          if (why) {
            this.setStep(step.id, { state: "skipped", line: why, counted: this.stepsBegun() });
            continue;
          }
        }
        if (step.id === "loop") {
          // asked again at its turn: the outside step may have changed the loop plan. A step left out
          // only now keeps its number (L2b-6); one left out from the start that can run now runs.
          const skip = this.loopSkip();
          if (skip) {
            const counted = this.steps().find((s) => s.id === "loop")?.state !== "skipped" && this.stepsBegun();
            this.setStep(step.id, { state: "skipped", line: skip, counted });
            continue;
          }
        }
        this.stepT0 = Date.now();
        this.clock.set(this.stepT0);
        this.setStep(step.id, { state: "running" });
        this.markStep(id, step.id);
        this.takeStepEnd();
        let ok = false;
        try {
          ok = await runners[step.id]();
        } catch (e) {
          states[step.id].set({ running: false, fraction: null, line: "Failed", detail: errorText(e), tone: "error" });
        }
        const r = states[step.id]();
        this.stepSeconds[step.id] = (Date.now() - this.stepT0) / 1000;
        // 17 Nov 2026:  the step was stopped with  Cancel (its own request answered "cancelled", or  its job
        // said so)  -- by  the Cancel beside Start, the waiting box's, a card's or the strip's. Only this step
        // is thrown away: the steps before it keep their state, and the sentence names the last of  them.
        // When the Mind loop was built and only its fixes afterwards  were stopped, the loop is that  step.
        const end = this.takeStepEnd();
        // 17 Nov 2026: the question the Prep  job stopped Start with (the Prep step, or the fixes after  the loop)
        const ask = this.takeStepAsk();
        if (end) {
          if (end === "after_loop") finished = step.id;
          this.setStep(step.id, { state: "stopped", line: end === "after_loop" ? r.line : "Stopped.", detail: "" });
          return await this.endCancelled(finished, end === "after_loop" ? "prepare" : step.id);
        }
        // the files are in, but no  computed value could be compared: amber, and the chain stops -- going
        // on is the user's choice (Do everything again skips the step: the server says brought in)
        const unchecked = step.id === "outside" && !ok && this.outsideUnchecked();
        this.setStep(step.id, { state: ok ? "done" : unchecked ? "stopped" : "failed", line: r.line, detail: r.detail });
        // 17 Nov 2026: Cancel was clicked while this step ran, and the step's own answer came normally (the
        // click reached the server too late to stop it): its result stands and is shown, it counts as
        // finished,  and the chain stops after it. Only a step that failed by itself, with nothing to show
        // for it (red), keeps its own stop sentence below.
        if (this.chainCancelled && r.tone !== "error") return await this.endCancelled(step.id, plan[plan.indexOf(step) + 1]?.id ?? null);
        if (unchecked) {
          this.say(
            `Stopped at "${step.label}": the computed values were not checked. Check them under Advanced › Deliver (Check the numbers), or press Do everything again to go on without the check.`,
          );
          return;
        }
        if (!ok) {
          // 17 Nov 2026: the Prep job could not fix everything and Start stops here (a Prep fix is still
          // left after the assistant tried, or no AI key is saved): its own two sentences are the stop
          // line, with "Show me" to the problem's cell and "Send feedback". Every other stop keeps the
          // sentence that names the step and where  to fix it.
          // 17 Nov 2026  (live test L2b-2): the server would not  start the step  because other  work of this
          // workbook still ran: its own sentence says so ("Wait a moment: ..."), not "Not started."
          const busyWords = r.tone === "error" ? waitWords(r.detail) : "";
          if (ask) this.say("", ask);
          else if (busyWords) this.say(`${busyWords} Press Do everything again once that has finished.`);
          else this.say(`Stopped at "${step.label}": ${r.line.replace(/\.$/, "")}. ${stoppedWhere(step.id, r, this.basic)}`);
          return;
        }
        // repair round: what the Prep step's job had; after the loop only a problem it  did not have
        // starts the job again (fixesFollowLoop)
        if (step.id === "prepare") this.prepSaw = new Set((this.store.readiness()?.blocking ?? []).map(problemKey));
        finished = step.id;
      }
      // 17 Nov 2026: every step  ran. A question that did not  stop Start (a blocking problem with no  Prep
      // fix, which the assistant could not fix either) is asked now,  in place of "Everything ran and every
      // check  passed" -- unless the steps after it made its problem go.
      const left = this.pendingAsk;
      this.pendingAsk = null;
      const r = this.store.readiness();
      if (left && this.stillBlocking(left)) this.say("", left);
      else if (r?.state === "blocked") {
        // 17 Nov 2026 (review R2-3): never "every check passed" while the workbook is blocked: how many
        // blocking problems are left, and "Show me" to the first (the banner below says it is blocked)
        const first = r.blocking?.[0] ?? null;
        this.say(blockingLeftLine(r.blocking_count || r.blocking?.length || 1, first));
        this.endPlace.set(first ? placeOfProblem(first) : { id: "assistant-ask", open: true });
      } else {
        this.say(`Everything ran and every check passed (${minutes((Date.now() - this.chainT0) / 1000)}).`);
        // 17 Nov 2026 (live test L2b-9): Basic has no step list where Custom says it: the assistant's own
        // fixes in a  second sentence, with "Show me" to Versions
        if (this.basic) this.chainNote.set(assistantFixedWords(this.prepTotal.applied, this.prepTotal.fixed));
        // (17 Nov 2026, Palermo: the cells that show an error only for the Mind loop's first value are said once,
        // on Basic's finish banner; Custom says it on the Recalculate step's line)
      }
    } finally {
      if (this.ticker) clearInterval(this.ticker);
      this.ticker = null;
      // 17 Nov 2026 (live test L2b-4): ended here, so a reload says nothing about it (a page left in the
      // middle keeps the mark: the next one says that the rest of Start did not run)
      if (!this.destroyed) {
        markStart(null);
        // 17 Nov 2026 (live test L2b-3): Start's steps have ended: the  shell's list of running work  drops
        // them now, before the Cancel beside Start  leaves the screen  (else "Still  working: Recalculating."
        // flashed beside "Everything ran and every check passed"). Not after the page was left (repair
        // round): its step may still run on the server, followed by nothing here, and the strip  and the
        // next page's line must still see it.
        this.work.ended(START_KINDS);
      }
      this.chainStopping.set(false);
      this.chainRunning.set(false);
    }
  }


  /** The problem a kept  question asks about is still blocking  now (the loop and the recalculation ran
   *  after  it was asked): read from the readiness the last step brought.  Not known: it stands. */
  private stillBlocking(ask: PrepAsk): boolean {
    const rule = ask.show_me?.rule_id;
    const r = this.store.readiness();
    if (!rule || !r) return true;
    return (r.blocking ?? []).some((b) => b.rule_id === rule);
  }

  private setStep(id: StepId, patch: Partial<Step>): void {
    this.steps.update((list) => list.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  }

  toggleStep(id: StepId): void {
    this.setStep(id, { open: !this.steps().find((s) => s.id === id)?.open });
  }

  chainStep(): Step | null {
    return this.steps().find((s) => s.state === "running") ?? null;
  }

  /** 17 Nov 2026 (live  test L2b-6): the steps Start counts in "Step 2 of 4": every step but those left out
   *  before the first one began (the outside files already brought in, a loop that cannot be built, the
   *  box not ticked). A step skipped later keeps its number: "Step 2 of 4" was followed by "Step 3 of 3". */
  private countedSteps(): Step[] {
    return this.steps().filter((s) => s.state !== "skipped" || s.counted);
  }

  /** A step of this run has begun (a  step skipped from now on keeps its number). */
  private stepsBegun(): boolean {
    return this.steps().some((s) => s.state !== "waiting" && s.state !== "skipped");
  }

  /** Why the loop step is left out of this run ("" when it runs): the box is not ticked, or the workbook
   *  has no loop the app can build. */
  private loopSkip(): string {
    if (this.canBuildLoop() && !this.loopWanted()) return "Left out: the box was not ticked";
    if (!this.canBuildLoop()) return this.loopPlan()?.found ? "The loop needs one answer first (Advanced › Mind loop)" : "No run loop in this workbook";
    return "";
  }

  stateOf(id: StepId): ActionState {
    return this[id]();
  }

  toggle(which: StepId | "fixFormulas" | "nextQuarter"): void {
    this[which].update((s) => ({ ...s, open: !s.open }));
  }

  // ---- next quarter's Mind input files -------------------------------------------------------------
  onNextQuarterFiles(event: Event): void {
    const input = event.target as HTMLInputElement;
    const files = Array.from(input.files ?? []);
    input.value = "";
    if (files.length) void this.makeNextQuarter(files);
  }

  /** Make next quarter's Mind input files: the picked raw files go to the server, which runs the model's
   *  own import macro on a copy with them and writes one file per Mind input grid (a file already shaped
   *  like a grid goes in  as it is). Followed until  it ends; the zip is then offered  for download. No  new
   *  version: the  model in Mind stays as it is. */
  async makeNextQuarter(files: File[]): Promise<boolean> {
    const id = this.store.sessionId();
    if (!id || this.nextQuarter().running) return false;
    const fail = (line: string, detail = "") => {
      this.nextQuarter.set({ running: false, fraction: null, line, detail, tone: "error", open: !!detail });
      return false;
    };
    if (this.busy()) return fail(`Wait: ${this.busyHint()}.`);
    const problem = inputFilesProblem(files.map((f) => f.name));
    if (problem) return fail(problem);
    const seq = ++this.nextQuarterSeq;
    this.nextQuarterResult.set(null);
    this.nextQuarterSession.set(null);
    this.nextQuarter.set({ running: true, fraction: null, line: "Starting…", detail: "", tone: "idle" });
    try {
      await this.api.startNextQuarter(id, files);
    } catch (e) {
      if (seq !== this.nextQuarterSeq) return false;
      return fail(firstClause(errorText(e)) || "Not started", errorText(e));
    }
    return this.followNextQuarter(id, seq, Date.now());
  }

  /** A next-quarter job may still run from before a reload or a visit to Advanced (follow it), or its
   *  result is there to show (the zip  is offered again). */
  private async resumeNextQuarter(): Promise<void> {
    const id = this.store.sessionId();
    if (!id || !this.store.outsideInputs() || this.nextQuarter().running) return;
    const seq = ++this.nextQuarterSeq;
    let st: NextQuarterStatus;
    try {
      st = await this.api.nextQuarter(id);
    } catch {
      return; // nothing to pick up
    }
    if (this.destroyed || seq !== this.nextQuarterSeq || this.store.sessionId() !== id) return;
    if (st.running) {
      this.showNextQuarterRunning(st, Date.now());
      await this.followNextQuarter(id, seq, Date.now());
    } else if (st.result || st.error) this.showNextQuarterEnd(id, st);
  }

  /** Follow the job until it ends.  Stops quietly when the page is  left, the card is reset or the
   *  session  changes: an old job's result never lands on another workbook's card. */
  private async followNextQuarter(id: string, seq: number, t0: number): Promise<boolean> {
    const stale = () => {
      if (this.destroyed || seq !== this.nextQuarterSeq) return true;
      if (this.store.sessionId() === id) return false;
      this.resetNextQuarter(); // another  session without a reset: the card must  not stay "running"
      return true;
    };
    let failures = 0;
    for (;;) {
      if (stale()) return false;
      let st: NextQuarterStatus;
      try {
        st = await this.api.nextQuarter(id);
        failures = 0;
      } catch (e) {
        if (stale()) return false;
        if (++failures < 5) {
          await this.sleep(1000);
          continue;
        }
        this.nextQuarter.set({ running: false, fraction: null, line: "Lost the job", detail: errorText(e), tone: "error", open: true });
        return false;
      }
      if (stale()) return false;
      if (st.running) {
        this.showNextQuarterRunning(st, t0);
        await this.sleep(1000);
        continue;
      }
      return this.showNextQuarterEnd(id, st);
    }
  }

  private showNextQuarterRunning(st: NextQuarterStatus, t0: number): void {
    const secs = Math.round(st.elapsed_s ?? (Date.now() - t0) / 1000);
    const took = secs < 60 ? `${secs} s so far` : `${Math.floor(secs / 60)} min ${secs % 60} s so far`;
    const said = st.message && st.message !== st.stage ? st.message : "";
    this.nextQuarter.set({ running: true, fraction: null, line: st.stage || "Working…", detail: [said, took].filter(Boolean).join(" · "), tone: "idle" });
  }

  private showNextQuarterEnd(id: string, st: NextQuarterStatus): boolean {
    if (st.state === "cancelled") {
      // 17 Nov 2026: stopped with Cancel: no zip was made; the card is emptied (an older job's files and
      // zip go too) and says so in one calm sentence
      this.resetNextQuarter();
      this.nextQuarter.set(stoppedState());
      return false;
    }
    if (st.result) {
      this.nextQuarterResult.set(st.result);
      this.nextQuarterSession.set(id);
      const out = nextQuarterOutcome(st.result);
      this.nextQuarter.set({ running: false, fraction: null, line: out.line, detail: out.detail, tone: out.tone, open: out.tone !== "ok" && !!out.detail });
      return st.result.ok;
    }
    const error = st.error ?? "";
    this.nextQuarter.set({ running: false, fraction: null, line: firstClause(error) || "The job ended without a result", detail: error, tone: "error", open: !!error });
    return false;
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((res) => this.timers.push(setTimeout(res, ms)));
  }
}
