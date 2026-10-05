import { useEffect, useState } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { NumbersCheck } from "../services/api";

/**
 * 1.8.0: did the preparation change what the model computes?
 *
 * Prep inserts rows and writes titles. Excel moves ordinary references along,
 * but a 3-D reference or an INDIRECT address does not move -- on one real model
 * an ordinary Prep changed 2,462 computed values and nothing said so. This card
 * asks the backend to recalculate the original and the current version and
 * compare every formula cell (POST /numbers-check). Shown once a version other
 * than the original is current.
 */
export default function NumbersCheckCard() {
  const sessionId = useStore((s) => s.sessionId);
  const currentVersion = useStore((s) => s.currentVersion);
  const versions = useStore((s) => s.versions);
  const [result, setResult] = useState<NumbersCheck | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // another version: the last answer is about a different file
  useEffect(() => {
    setResult(null);
    setError(null);
  }, [currentVersion?.id]);

  if (!sessionId || !currentVersion || versions.length < 2 || currentVersion.source === "upload") return null;

  async function check() {
    setBusy(true);
    setError(null);
    try {
      setResult(await api.numbersCheck(sessionId!));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  const bad = result?.ran && result.clean === false;
  const good = result?.ran && result.clean === true;
  const examples = [...(result?.samples?.good_to_error ?? []), ...(result?.samples?.good_to_other ?? [])].slice(0, 5);

  return (
    <div className={`mb-3 text-[12px] rounded-lg px-3 py-2 border ${bad ? "text-[#9F1D1D] bg-[#FDE2E2] border-[#9F1D1D]/20" : good ? "text-[#0F766E] bg-[#F0FDFA] border-[#0F766E]/20" : "text-[#374151] bg-white border-[#E5E7EB]"}`} aria-live="polite">
      <div className="flex items-center gap-3 flex-wrap">
        <span className="flex-1 min-w-0">
          {result?.ran ? (
            <strong>{good ? "✓ " : "⚠ "}{result.verdict}</strong>
          ) : result ? (
            result.verdict
          ) : (
            <>Did the changes so far alter any number the model computes? Excel recalculates the original and {currentVersion.label.split(" — ")[0]} and compares every formula.</>
          )}
        </span>
        <button onClick={check} disabled={busy} className="px-3 py-1.5 rounded-md text-[12px] font-semibold bg-[#1F3A5F] text-white hover:bg-[#162d4a] disabled:opacity-50 whitespace-nowrap flex items-center gap-2">
          {busy && <span className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin" />}
          {busy ? "Comparing…" : result ? "Check again" : "Check the numbers"}
        </button>
      </div>
      {bad && (
        <div className="mt-2">
          {examples.map((x, i) => (
            <div key={i} className="font-mono text-[11px]">
              {x.now}: {x.was} → {x.is}
            </div>
          ))}
          <div className="mt-1">
            Do not use this version. Go back to the version before the change in History, and apply Prep again without the action that inserts rows (grid titles) on the sheets named above.
          </div>
        </div>
      )}
      {error && <div className="mt-1 text-[#9F1D1D]" role="alert">{error}</div>}
    </div>
  );
}
