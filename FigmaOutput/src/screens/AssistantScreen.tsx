import { useState, useRef, useEffect } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { ChatDraft, ChatMessage, Proposal } from "../types";
import OperationsTable from "../components/OperationsTable";
import ChatMarkdown from "../components/ChatMarkdown";
import StreamingDraft, { nextDraft } from "../components/StreamingDraft";

function ProvenanceLine({ p }: { p: NonNullable<ChatMessage["provenance"]> }) {
  return (
    <div className="text-[11px] text-[#9CA3AF] font-mono mt-1">
      context {(p.context_chars / 1000).toFixed(1)}k chars · detail {(p.detail_chars / 1000).toFixed(1)}k · {p.lookups} lookup{p.lookups !== 1 ? "s" : ""}
    </div>
  );
}

function ProposalCard({
  proposal,
  onApply,
  onDiscard,
}: {
  proposal: Proposal;
  onApply: () => void;
  onDiscard: () => void;
}) {
  return (
    <div className="border border-[#1F3A5F]/30 rounded-xl bg-[#EEF2FF] p-4 mt-2">
      {proposal.summary && (
        <p className="text-[13px] font-semibold text-[#1F3A5F] mb-3">{proposal.summary}</p>
      )}
      <OperationsTable operations={proposal.operations} compact />
      {proposal.errors.length > 0 && (
        <div className="mt-2 flex flex-col gap-1">
          {proposal.errors.map((e, i) => (
            <div key={i} className="text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded px-2 py-1">
              ⚠ {e}
            </div>
          ))}
        </div>
      )}
      <div className="flex gap-2 mt-3">
        <button
          onClick={onApply}
          className="px-4 py-1.5 bg-[#1F3A5F] text-white rounded-md text-[12px] font-semibold hover:bg-[#162d4a] transition-colors"
        >
          Apply {proposal.operations.length} change{proposal.operations.length !== 1 ? "s" : ""}
        </button>
        <button
          onClick={onDiscard}
          className="px-4 py-1.5 border border-[#E5E7EB] text-[#6B7280] rounded-md text-[12px] hover:bg-white transition-colors"
        >
          Discard
        </button>
      </div>
    </div>
  );
}

function Bubble({ msg }: { msg: ChatMessage & { proposal?: Proposal } }) {
  const isUser = msg.role === "user";
  const isNote = msg.note;

  if (isNote) {
    return (
      <div className="flex justify-center">
        <div className="px-3 py-1.5 bg-[#F0FDFA] border border-[#0F766E]/20 rounded-full text-[12px] text-[#0F766E] font-mono">
          {msg.content}
        </div>
      </div>
    );
  }

  return (
    <div className={`flex ${isUser ? "justify-end" : "justify-start"}`}>
      <div className={`max-w-[75%] ${isUser ? "order-2" : ""}`}>
        <div
          className={`rounded-2xl px-4 py-3 text-[13px] leading-relaxed ${
            isUser
              ? "bg-[#1F3A5F] text-white rounded-tr-sm"
              : "bg-white border border-[#E5E7EB] text-[#374151] rounded-tl-sm"
          }`}
        >
          <ChatMarkdown content={msg.content} />
        </div>
        {msg.provenance && <ProvenanceLine p={msg.provenance} />}
      </div>
    </div>
  );
}

export default function AssistantScreen() {
  const sessionId = useStore((s) => s.sessionId);
  const chatHistory = useStore((s) => s.chatHistory);
  const setChatHistory = useStore((s) => s.setChatHistory);
  const appendChatMessage = useStore((s) => s.appendChatMessage);
  const addVersion = useStore((s) => s.addVersion);
  const setReport = useStore((s) => s.setReport);
  const setSummary = useStore((s) => s.setSummary);
  const [downloadName, setDownloadName] = useState<string | null>(null);

  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  // 1.7.2: a chat turn in flight (as opposed to an apply) and its reply so far
  const [streaming, setStreaming] = useState(false);
  const [draft, setDraft] = useState<ChatDraft | null>(null);
  const [pendingProposal, setPendingProposal] = useState<{ proposal: Proposal; msgIdx: number } | null>(null);
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chatHistory, loading, draft]);

  if (!sessionId) {
    return (
      <div className="p-8 flex flex-col items-center justify-center h-64 gap-3">
        <div className="w-12 h-12 rounded-full border-2 border-dashed border-[#E5E7EB] flex items-center justify-center text-[#D1D5DB] text-xl">💬</div>
        <div className="text-[13px] text-[#6B7280] text-center">
          The assistant needs the shared gateway key.<br />
          Upload a workbook first.
        </div>
      </div>
    );
  }

  async function send() {
    if (!input.trim() || loading) return;
    const question = input.trim();
    setInput("");
    appendChatMessage({ role: "user", content: question });
    setLoading(true);
    setStreaming(true);
    setDraft(null);
    try {
      const reply = await api.chatStream(sessionId!, chatHistory, question, undefined, (ev) => setDraft((d) => nextDraft(d, ev)));
      const msg: ChatMessage & { proposal?: Proposal } = {
        role: "assistant",
        content: reply.text,
        provenance: reply.provenance,
      };
      appendChatMessage(msg);
      if (reply.proposal && (reply.proposal.operations.length > 0 || reply.proposal.errors.length > 0)) {
        setPendingProposal({ proposal: reply.proposal, msgIdx: chatHistory.length + 1 });
      }
    } catch (e) {
      appendChatMessage({ role: "assistant", content: `The assistant could not answer: ${e instanceof Error ? e.message : String(e)}`, note: true });
    } finally {
      setLoading(false);
      setStreaming(false);
      setDraft(null);
    }
  }

  async function applyProposal() {
    if (!pendingProposal || !sessionId) return;
    if (pendingProposal.proposal.operations.length === 0) {
      setPendingProposal(null);
      return;
    }
    setPendingProposal(null);
    setLoading(true);
    try {
      // The backend applies through Excel, verifies the file and re-analyzes it, so the
      // conversation continues on the changed workbook.
      const outcome = await api.applyOperations(sessionId, pendingProposal.proposal.operations, { reanalyze: true });
      addVersion(outcome.version);
      if (outcome.summary) setSummary(outcome.summary);
      if (outcome.report && outcome.plan) setReport(outcome.report, outcome.plan);
      const { result } = outcome;
      const verified = result.verified_opens_in_excel ? "verified" : "not verified";
      const failed = result.failed.length ? ` · ${result.failed.length} failed` : "";
      appendChatMessage({
        role: "assistant",
        content: `Applied ${result.applied.length} change${result.applied.length !== 1 ? "s" : ""} via ${result.method === "excel_com" ? "Excel" : "openpyxl"} · ${verified}${failed} · re-analyzed as ${outcome.version.label}`,
        note: true,
      });
      if (result.failed.length) {
        // 1.7.2: say WHY, grouped by reason (17 identical refusals read as one line, not 17)
        const byReason = new Map<string, string[]>();
        for (const f of result.failed) {
          const reason = f.error ?? "unknown reason";
          byReason.set(reason, [...(byReason.get(reason) ?? []), `${f.sheet}!${f.cell ?? f.row ?? f.column ?? ""}`]);
        }
        const lines = [...byReason.entries()].map(
          ([reason, cells]) => `- **${cells.length} cell${cells.length !== 1 ? "s" : ""}** (${cells.slice(0, 6).join(", ")}${cells.length > 6 ? ", …" : ""}): ${reason}`
        );
        appendChatMessage({
          role: "assistant",
          content: `${result.failed.length} change${result.failed.length !== 1 ? "s were" : " was"} not applied — the workbook keeps the old content there:\n\n${lines.join("\n")}`,
        });
      }
      setDownloadName(result.output_name);
    } catch (e) {
      appendChatMessage({ role: "assistant", content: `The change could not be applied: ${e instanceof Error ? e.message : String(e)}`, note: true });
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="flex flex-col flex-1 min-h-0">
      <div className="px-6 pt-5 pb-3 border-b border-[#E5E7EB] bg-white flex-shrink-0 flex items-center justify-between">
        <h1 className="text-lg font-semibold text-[#111827]">Assistant</h1>
        {chatHistory.length > 0 && (
          <button
            onClick={() => { setChatHistory([]); setPendingProposal(null); }}
            className="text-[12px] text-[#9CA3AF] hover:text-[#374151] transition-colors"
          >
            Clear conversation
          </button>
        )}
      </div>

      <div className="flex-1 overflow-auto px-6 py-5 flex flex-col gap-4">
        {chatHistory.length === 0 && (
          <div className="flex flex-col items-center justify-center h-40 text-[#9CA3AF] text-[13px] text-center gap-1">
            <p>Ask anything about the workbook, or request a change.</p>
            <p className="text-[12px] font-mono">e.g. why does RES-002 fail? · rename sheet Notes to Inputs</p>
          </div>
        )}
        {chatHistory.map((msg, i) => (
          <div key={i}>
            <Bubble msg={msg} />
          </div>
        ))}
        {downloadName && !loading && (
          <div className="flex justify-center">
            <a
              href={api.downloadUrl(sessionId, downloadName)}
              className="px-3 py-1.5 border border-[#1F3A5F] text-[#1F3A5F] rounded-md text-[12px] font-medium hover:bg-white transition-colors"
            >
              Download {downloadName}
            </a>
          </div>
        )}
        {pendingProposal && (
          <div className="flex justify-start">
            <div className="max-w-[80%]">
              <ProposalCard
                proposal={pendingProposal.proposal}
                onApply={applyProposal}
                onDiscard={() => setPendingProposal(null)}
              />
            </div>
          </div>
        )}
        {loading && streaming && (
          <div className="flex justify-start">
            <div className="max-w-[75%]">
              <StreamingDraft draft={draft} className="rounded-2xl rounded-tl-sm px-4 py-3 text-[13px] leading-relaxed bg-white border border-[#E5E7EB] text-[#374151]" />
            </div>
          </div>
        )}
        {loading && !streaming && (
          <div className="flex justify-start">
            <div className="bg-white border border-[#E5E7EB] rounded-2xl rounded-tl-sm px-4 py-3 flex items-center gap-2">
              <div className="flex gap-1">
                {[0, 1, 2].map((i) => (
                  <div
                    key={i}
                    className="w-1.5 h-1.5 bg-[#9CA3AF] rounded-full animate-bounce"
                    style={{ animationDelay: `${i * 150}ms` }}
                  />
                ))}
              </div>
            </div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      <div className="px-6 py-4 border-t border-[#E5E7EB] bg-white flex-shrink-0">
        <div className="flex gap-2 items-end">
          <textarea
            ref={textareaRef}
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
            placeholder="Ask anything, or tell me what to change — e.g. rename sheet Notes to Inputs · title the grid at B4 #Premiums /Input · why does RES-002 fail?"
            rows={2}
            className="flex-1 resize-none px-3 py-2.5 border border-[#E5E7EB] rounded-xl text-[13px] text-[#374151] placeholder-[#9CA3AF] focus:outline-none focus:border-[#1F3A5F] leading-relaxed"
          />
          <button
            onClick={send}
            disabled={loading || !input.trim()}
            className="px-4 py-2.5 bg-[#1F3A5F] text-white rounded-xl text-[13px] font-semibold hover:bg-[#162d4a] transition-colors disabled:opacity-40 flex-shrink-0"
          >
            Send
          </button>
        </div>
      </div>
    </div>
  );
}
