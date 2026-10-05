import ChatMarkdown from "./ChatMarkdown";
import type { ChatDraft, ChatStreamEvent } from "../types";

/** The machine-readable blocks the backend strips from the final answer; hidden while streaming too. */
const HIDDEN_FENCE = /```(changes|lookup)\b/;

const PHASE_LABEL: Record<ChatDraft["phase"], string | null> = {
  answer: null,
  lookup: "Looking up cells in the workbook…",
  repair: "Checking the proposed change…",
};

/** Folds one stream event into the draft: a new model call starts a fresh text, a delta extends it. */
export function nextDraft(draft: ChatDraft | null, ev: ChatStreamEvent): ChatDraft {
  if (ev.type === "phase") return { phase: ev.phase, text: "" };
  return { phase: draft?.phase ?? "answer", text: (draft?.text ?? "") + ev.text };
}

function Dots() {
  return (
    <div className="flex gap-1 py-1">
      {[0, 1, 2].map((i) => (
        <div key={i} className="w-1.5 h-1.5 bg-[#9CA3AF] rounded-full animate-bounce" style={{ animationDelay: `${i * 150}ms` }} />
      ))}
    </div>
  );
}

/**
 * 1.7.2: the assistant's reply while it is being written. Shows the dots until
 * the first text arrives, a status line during lookups and proposal repairs,
 * and "Preparing the change…" once the model starts its ```changes block.
 */
export default function StreamingDraft({ draft, className }: { draft: ChatDraft | null; className: string }) {
  const raw = draft?.text ?? "";
  const cut = raw.search(HIDDEN_FENCE);
  const visible = (cut >= 0 ? raw.slice(0, cut) : raw).trimEnd();
  const label = draft ? PHASE_LABEL[draft.phase] : null;
  const preparing = cut >= 0 ? (raw.startsWith("```lookup", cut) ? "Looking up cells in the workbook…" : "Preparing the change…") : null;
  return (
    <div className={className}>
      {label && label !== preparing && <div className="text-[11px] text-[#6B7280] italic mb-1">{label}</div>}
      {visible ? <ChatMarkdown content={visible} /> : !preparing && <Dots />}
      {preparing && <div className="text-[11px] text-[#6B7280] italic mt-1 flex items-center gap-2">{preparing}<Dots /></div>}
    </div>
  );
}
