import { useEffect, useMemo, useRef, useState } from "react";
import { useStore } from "../store";
import type { Grid } from "../types";
import * as api from "../services/api";

/** 1 -> A, 27 -> AA */
function colName(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseCell(ref: string | null): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec((ref ?? "").trim());
  if (!m) return null;
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: parseInt(m[2], 10), c };
}

interface Rect { r1: number; c1: number; r2: number; c2: number }

function normRect(a: { r: number; c: number }, b: { r: number; c: number }): Rect {
  return { r1: Math.min(a.r, b.r), r2: Math.max(a.r, b.r), c1: Math.min(a.c, b.c), c2: Math.max(a.c, b.c) };
}

function rectRef(s: Rect): string {
  const a = `${colName(s.c1)}${s.r1}`;
  return s.r1 === s.r2 && s.c1 === s.c2 ? a : `${a}:${colName(s.c2)}${s.r2}`;
}

/** Mirror of the backend's resolution: the grids the selection touches (their
 * rectangle, or their title cell), and whether two of them are the documented
 * caption-above-a-table pair that can be named as one. */
function resolveSelection(grids: Grid[], sel: Rect):
  | { kind: "grid"; grid: Grid; hits: Grid[] }
  | { kind: "caption"; grid: Grid; table: Grid; hits: Grid[] }
  | { kind: "none" | "many"; hits: Grid[] } {
  const hits = grids.filter((g) => {
    const t = parseCell(g.title_cell);
    const inRect = !(g.last_row < sel.r1 || g.first_row > sel.r2 || g.last_col < sel.c1 || g.first_col > sel.c2);
    const onTitle = !!t && sel.r1 <= t.r && t.r <= sel.r2 && sel.c1 <= t.c && t.c <= sel.c2;
    return inRect || onTitle;
  });
  const isCap = (g: Grid) => {
    const v = g.header_values[0];
    return !g.name && g.n_cols === 1 && typeof v === "string" && v.trim() !== "" && !v.startsWith("=");
  };
  if (hits.length === 1) {
    // one half of an untitled caption-above-a-table pair is named as the pair
    // (the backend extends the same way)
    const g = hits[0];
    if (!g.name) {
      if (isCap(g)) {
        const t = grids.find((h) => h !== g && !h.name && h.first_row === g.first_row + 1 && h.first_col === g.first_col + 1);
        if (t) return { kind: "caption", grid: g, table: t, hits };
      }
      const cap = grids.find((h) => h !== g && isCap(h) && g.first_row === h.first_row + 1 && g.first_col === h.first_col + 1);
      if (cap) return { kind: "caption", grid: cap, table: g, hits };
    }
    return { kind: "grid", grid: g, hits };
  }
  if (hits.length === 2) {
    for (const [cap, table] of [[hits[0], hits[1]], [hits[1], hits[0]]] as const) {
      const capText = cap.header_values[0];
      if (!cap.name && cap.n_cols === 1 && typeof capText === "string" && capText.trim() !== "" && !capText.startsWith("=")
        && table.first_row === cap.first_row + 1 && table.first_col === cap.first_col + 1) {
        return { kind: "caption", grid: cap, table, hits };
      }
    }
  }
  return { kind: hits.length === 0 ? "none" : "many", hits };
}

/** "/Resize.A /Group.(My group).0.1" or "Resize.A" -> ["Resize.A", "Group.(My group).0.1"] */
function parseExtraFlags(text: string): string[] {
  const t = text.trim();
  if (!t) return [];
  if (t.includes("/")) return t.split("/").map((x) => x.trim()).filter(Boolean);
  return [t];
}

interface PendingArea { sheet: string; ref: string; gridRef: string; name: string; flags: string[]; wasNamed: string | null }

export default function GridNamerScreen() {
  const summary = useStore((s) => s.summary);
  const sessionId = useStore((s) => s.sessionId);
  const addVersion = useStore((s) => s.addVersion);
  const applyAnalysis = useStore((s) => s.applyAnalysis);

  const sheets = summary?.sheets ?? [];
  const [active, setActive] = useState<string>(sheets[0]?.name ?? "");
  const [cells, setCells] = useState<api.SheetCells | null>(null);
  const [loading, setLoading] = useState(false);
  const [big, setBig] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const [sel, setSel] = useState<Rect | null>(null);
  const dragFrom = useRef<{ r: number; c: number } | null>(null);
  const [dragging, setDragging] = useState(false);

  const [flagList, setFlagList] = useState<api.MindFlag[]>([]);
  const [name, setName] = useState("");
  const [chosenFlags, setChosenFlags] = useState<Set<string>>(new Set());
  const [extraFlags, setExtraFlags] = useState("");
  const [pending, setPending] = useState<PendingArea[]>([]);
  const [nameRest, setNameRest] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [outcome, setOutcome] = useState<api.GridNamerOutcome | null>(null);
  const [showSkipped, setShowSkipped] = useState(false);

  useEffect(() => { api.mindFlags().then(setFlagList).catch(() => {}); }, []);
  useEffect(() => {
    if (sheets.length > 0 && !sheets.some((s) => s.name === active)) setActive(sheets[0].name);
  }, [sheets, active]);

  useEffect(() => {
    if (!sessionId || !active) return;
    let gone = false;
    setLoading(true);
    setError(null);
    setSel(null);
    api.sheetCells(sessionId, active, big ? 2000 : 400, big ? 200 : 60)
      .then((c) => { if (!gone) setCells(c); })
      .catch((e) => { if (!gone) { setCells(null); setError(e instanceof Error ? e.message : String(e)); } })
      .finally(() => { if (!gone) setLoading(false); });
    return () => { gone = true; };
  }, [sessionId, active, big, summary]);

  useEffect(() => {
    const up = () => { setDragging(false); dragFrom.current = null; };
    window.addEventListener("mouseup", up);
    return () => window.removeEventListener("mouseup", up);
  }, []);

  const grids = useMemo(() => sheets.find((s) => s.name === active)?.grids ?? [], [sheets, active]);
  const target = useMemo(() => (sel ? resolveSelection(grids, sel) : null), [grids, sel]);
  const pendingByGrid = useMemo(() => new Map(pending.map((p) => [`${p.sheet}!${p.gridRef}`, p])), [pending]);

  // cell -> grid index, and title cells, for painting
  const paint = useMemo(() => {
    const cellOf = new Map<string, number>();
    const titleOf = new Map<string, number>();
    grids.forEach((g, i) => {
      for (let r = g.first_row; r <= g.last_row; r++)
        for (let c = g.first_col; c <= g.last_col; c++) cellOf.set(`${r}:${c}`, i);
      const t = parseCell(g.title_cell);
      if (t) titleOf.set(`${t.r}:${t.c}`, i);
    });
    return { cellOf, titleOf };
  }, [grids]);

  const targetGrid = target && (target.kind === "grid" || target.kind === "caption") ? target.grid : null;
  const documented = useMemo(() => new Set(flagList.map((f) => f.name.toLowerCase())), [flagList]);

  // Prefill name + flags from the resolved grid (its current title, if any).
  useEffect(() => {
    if (!targetGrid) return;
    setName(targetGrid.name ?? "");
    const docFlags = new Set<string>();
    const extras: string[] = [];
    for (const f of targetGrid.flags) {
      const raw = f.raw.replace(/^\//, "");
      const known = flagList.find((d) => d.name.toLowerCase() === f.name.toLowerCase());
      if (known && f.args.length === 0) docFlags.add(known.name);
      else extras.push(raw);
    }
    setChosenFlags(docFlags);
    setExtraFlags(extras.map((e) => "/" + e).join(" "));
  }, [targetGrid, flagList]);

  function addArea() {
    if (!targetGrid || !target || !sel || !name.trim()) return;
    const flags = [...chosenFlags, ...parseExtraFlags(extraFlags)];
    const entry: PendingArea = {
      sheet: active,
      ref: target.kind === "caption" ? rectRef(sel) : targetGrid.ref,
      gridRef: targetGrid.ref,
      name: name.trim(),
      flags,
      wasNamed: targetGrid.name,
    };
    setPending((prev) => [...prev.filter((p) => !(p.sheet === entry.sheet && p.gridRef === entry.gridRef)), entry]);
    setSel(null);
    setName("");
    setChosenFlags(new Set());
    setExtraFlags("");
  }

  async function submit() {
    if (!sessionId || pending.length === 0) return;
    setSubmitting(true);
    setError(null);
    setOutcome(null);
    try {
      const out = await api.applyNamedAreas(
        sessionId,
        pending.map((p) => ({ sheet: p.sheet, ref: p.ref, name: p.name, flags: p.flags })),
        nameRest
      );
      addVersion(out.version);
      applyAnalysis({ summary: out.summary, report: out.report, plan: out.plan, delta: out.delta, readiness: out.readiness, prep_progress: out.prep_progress });
      setOutcome(out);
      setPending([]);
      setSel(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  }

  if (!summary || !sessionId) {
    return (
      <div className="p-8 flex flex-col items-center justify-center h-64 text-[#9CA3AF] text-sm gap-2">
        <span className="text-2xl">🏷️</span>
        Upload and analyze a workbook first — then select areas here and name them.
      </div>
    );
  }

  const nRows = cells?.rows.length ?? 0;
  const nCols = cells?.n_cols ?? 0;

  return (
    <div className="flex flex-1 min-h-0">
      {/* --- sheet view ------------------------------------------------------------ */}
      <div className="flex flex-col flex-1 min-w-0">
        <div className="px-6 pt-5 pb-3 border-b border-[#E5E7EB] bg-white flex-shrink-0">
          <h1 className="text-lg font-semibold text-[#111827] mb-1">Grid Namer</h1>
          <p className="text-[13px] text-[#6B7280]">
            Drag over an area of the sheet and give it a name (and flags). Submit writes the{" "}
            <span className="font-mono">#Name /Flags</span> titles into a new verified copy, so Mind reads those
            areas under your names; every grid you did not touch is named by the usual conventions.
          </p>
          <div className="flex gap-1 mt-3 flex-wrap">
            {sheets.map((s) => (
              <button
                key={s.name}
                onClick={() => setActive(s.name)}
                className={`px-2.5 py-1 rounded-md text-[12px] font-medium transition-colors ${
                  s.name === active ? "bg-[#1F3A5F] text-white" : "bg-[#F3F4F6] text-[#4B5563] hover:bg-[#E5E7EB]"
                }`}
              >
                {s.name}
                {s.state !== "visible" && <span className="opacity-60"> (hidden)</span>}
                <span className={`ml-1.5 font-mono ${s.name === active ? "text-white/60" : "text-[#9CA3AF]"}`}>{s.grids.length}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="flex-1 overflow-auto bg-white select-none" onMouseLeave={() => setDragging(false)}>
          {loading && <div className="p-6 text-[13px] text-[#6B7280]">Loading {active}…</div>}
          {error && !loading && (
            <div className="m-4 text-[12px] text-[#9F1D1D] bg-[#FDE2E2] border border-[#9F1D1D]/20 rounded px-3 py-2" role="alert">{error}</div>
          )}
          {cells && !loading && (
            <table className="border-collapse text-[11px] font-mono">
              <thead>
                <tr>
                  <th className="sticky top-0 left-0 z-30 bg-[#F3F4F6] border border-[#E5E7EB] w-10 min-w-10" />
                  {Array.from({ length: nCols }, (_, i) => (
                    <th key={i} className="sticky top-0 z-20 bg-[#F3F4F6] border border-[#E5E7EB] px-2 py-0.5 text-[#6B7280] font-semibold min-w-16">
                      {colName(i + 1)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {cells.rows.map((row, ri) => {
                  const r = ri + 1;
                  return (
                    <tr key={r}>
                      <td className="sticky left-0 z-10 bg-[#F3F4F6] border border-[#E5E7EB] px-1.5 py-0.5 text-[#6B7280] font-semibold text-right">{r}</td>
                      {Array.from({ length: nCols }, (_, ci) => {
                        const c = ci + 1;
                        const v = row[ci];
                        const gi = paint.cellOf.get(`${r}:${c}`);
                        const ti = paint.titleOf.get(`${r}:${c}`);
                        const g = gi !== undefined ? grids[gi] : undefined;
                        const inSel = sel && sel.r1 <= r && r <= sel.r2 && sel.c1 <= c && c <= sel.c2;
                        const isPending = g && pendingByGrid.has(`${active}!${g.ref}`);
                        let cls = "border border-[#F3F4F6] px-2 py-0.5 whitespace-nowrap max-w-44 overflow-hidden text-ellipsis cursor-crosshair ";
                        if (ti !== undefined) cls += "bg-[#DFF5E6] text-[#0F6E3A] font-semibold ";
                        else if (isPending) cls += "bg-[#FDF4FF] text-[#86198F] ";
                        else if (g && g.name) cls += "bg-[#EEF2FF] text-[#374151] ";
                        else if (g) cls += "bg-[#FFF7E6] text-[#374151] ";
                        else cls += "text-[#6B7280] ";
                        if (g) {
                          if (r === g.first_row) cls += "border-t-[#94A3B8] ";
                          if (r === g.last_row) cls += "border-b-[#94A3B8] ";
                          if (c === g.first_col) cls += "border-l-[#94A3B8] ";
                          if (c === g.last_col) cls += "border-r-[#94A3B8] ";
                        }
                        if (inSel) cls += "!bg-[#1F3A5F]/20 ";
                        const text = v === null || v === undefined ? "" : String(v);
                        return (
                          <td
                            key={c}
                            className={cls}
                            title={`${colName(c)}${r}${g ? ` · ${g.display_name} (${g.ref})` : ""}${text ? ` · ${text.slice(0, 200)}` : ""}`}
                            onMouseDown={(e) => {
                              e.preventDefault();
                              dragFrom.current = { r, c };
                              setDragging(true);
                              setSel(normRect({ r, c }, { r, c }));
                            }}
                            onMouseEnter={() => {
                              if (dragging && dragFrom.current) setSel(normRect(dragFrom.current, { r, c }));
                            }}
                          >
                            {text.slice(0, 60)}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
          {cells?.truncated && !big && (
            <div className="p-3 text-[12px] text-[#8A5A00] bg-[#FFF1CC] border-t border-[#8A5A00]/20 flex items-center gap-3">
              Showing the first {nRows} rows × {nCols} columns of {cells.total_rows} × {cells.total_cols}.
              <button onClick={() => setBig(true)} className="underline underline-offset-2 font-medium">Load more</button>
            </div>
          )}
          {cells && (
            <div className="p-2 text-[11px] text-[#9CA3AF] flex items-center gap-4 border-t border-[#F3F4F6]">
              <span><span className="inline-block w-3 h-3 bg-[#DFF5E6] border border-[#94A3B8] align-middle mr-1" />title cell</span>
              <span><span className="inline-block w-3 h-3 bg-[#EEF2FF] border border-[#94A3B8] align-middle mr-1" />named grid</span>
              <span><span className="inline-block w-3 h-3 bg-[#FFF7E6] border border-[#94A3B8] align-middle mr-1" />untitled grid</span>
              <span><span className="inline-block w-3 h-3 bg-[#FDF4FF] border border-[#94A3B8] align-middle mr-1" />to be named</span>
              <span className="ml-auto">values from the {cells.values_from}</span>
            </div>
          )}
        </div>
      </div>

      {/* --- naming panel ---------------------------------------------------------- */}
      <div className="w-80 flex-shrink-0 border-l border-[#E5E7EB] bg-[#F9FAFB] flex flex-col overflow-auto">
        <div className="p-4 border-b border-[#E5E7EB]">
          <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-2">SELECTION</div>
          {!sel && <div className="text-[12px] text-[#9CA3AF]">Drag over cells on the left.</div>}
          {sel && target?.kind === "none" && (
            <div className="text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded px-2 py-1.5">
              {rectRef(sel)}: Mind detects no grid there — an alone text cell or empty space is ignored.
            </div>
          )}
          {sel && target?.kind === "many" && (
            <div className="text-[12px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/20 rounded px-2 py-1.5">
              {rectRef(sel)} touches {target.hits.length} grids ({target.hits.map((h) => h.ref).join(", ")}). Mind reads
              them separately — select and name one at a time.
            </div>
          )}
          {sel && targetGrid && target && (
            <>
              <div className="text-[12px] text-[#374151] mb-2">
                <span className="font-mono font-semibold">{active}!{targetGrid.ref}</span>
                {target.kind === "caption" && <span> + its table (named as one grid)</span>}
                <div className="text-[11px] text-[#9CA3AF] mt-0.5">
                  {targetGrid.name ? <>currently <span className="font-mono">#{targetGrid.name}</span></> : "currently untitled"}
                  {" · "}{targetGrid.n_rows}×{targetGrid.n_cols}
                </div>
              </div>
              <label className="block text-[11px] font-semibold text-[#6B7280] mb-1">Name</label>
              <input
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="e.g. Mortality Rates"
                className="w-full border border-[#D1D5DB] rounded-md px-2 py-1.5 text-[13px] bg-white focus:outline-none focus:border-[#1F3A5F]"
              />
              <label className="block text-[11px] font-semibold text-[#6B7280] mt-3 mb-1">Flags (documented)</label>
              <div className="max-h-40 overflow-auto border border-[#E5E7EB] rounded-md bg-white p-2 flex flex-col gap-1">
                {flagList.length === 0 && <div className="text-[11px] text-[#9CA3AF]">flag list unavailable</div>}
                {flagList.map((f) => (
                  <label key={f.name} className="flex items-start gap-2 text-[12px] text-[#374151] cursor-pointer" title={f.meaning}>
                    <input
                      type="checkbox"
                      className="mt-0.5 accent-[#1F3A5F]"
                      checked={chosenFlags.has(f.name)}
                      onChange={(e) => {
                        setChosenFlags((prev) => {
                          const next = new Set(prev);
                          e.target.checked ? next.add(f.name) : next.delete(f.name);
                          return next;
                        });
                      }}
                    />
                    <span className="font-mono">/{f.name}</span>
                  </label>
                ))}
              </div>
              <label className="block text-[11px] font-semibold text-[#6B7280] mt-2 mb-1">
                Flags with arguments <span className="font-normal text-[#9CA3AF]">(e.g. /Resize.A /Group.(My group).0.1)</span>
              </label>
              <input
                value={extraFlags}
                onChange={(e) => setExtraFlags(e.target.value)}
                placeholder="/Flag.arg …"
                className="w-full border border-[#D1D5DB] rounded-md px-2 py-1.5 text-[12px] font-mono bg-white focus:outline-none focus:border-[#1F3A5F]"
              />
              <button
                onClick={addArea}
                disabled={!name.trim()}
                className="mt-3 w-full py-1.5 bg-[#1F3A5F] text-white rounded-md text-[13px] font-semibold disabled:opacity-40"
              >
                Add named area
              </button>
            </>
          )}
        </div>

        <div className="p-4 flex-1">
          <div className="text-[11px] font-semibold text-[#9CA3AF] tracking-wider mb-2">
            NAMED AREAS ({pending.length})
          </div>
          {pending.length === 0 && <div className="text-[12px] text-[#9CA3AF]">Nothing added yet.</div>}
          <div className="flex flex-col gap-1.5">
            {pending.map((p) => (
              <div key={`${p.sheet}!${p.gridRef}`} className="bg-white border border-[#E5E7EB] rounded-md px-2 py-1.5 text-[12px]">
                <div className="flex items-center gap-2">
                  <span className="font-mono text-[#86198F] font-semibold truncate" title={`#${p.name}${p.flags.map((f) => " /" + f).join("")}`}>
                    #{p.name}{p.flags.map((f) => ` /${f}`).join("")}
                  </span>
                  <button
                    onClick={() => setPending((prev) => prev.filter((x) => x !== p))}
                    className="ml-auto text-[#9CA3AF] hover:text-[#9F1D1D]"
                    title="Remove"
                  >
                    ✕
                  </button>
                </div>
                <div className="text-[11px] text-[#9CA3AF] font-mono">
                  {p.sheet}!{p.gridRef}
                  {p.wasNamed && <span> (was #{p.wasNamed})</span>}
                </div>
              </div>
            ))}
          </div>

          <label className="flex items-start gap-2 mt-4 text-[12px] text-[#374151] cursor-pointer">
            <input type="checkbox" className="mt-0.5 accent-[#1F3A5F]" checked={nameRest} onChange={(e) => setNameRest(e.target.checked)} />
            <span>Name every other untitled grid automatically (captions, labels, headers — the usual conventions)</span>
          </label>

          <button
            onClick={submit}
            disabled={pending.length === 0 || submitting}
            className="mt-3 w-full py-2 bg-[#0F766E] text-white rounded-md text-[13px] font-semibold disabled:opacity-40"
          >
            {submitting ? "Writing via Excel…" : `Submit ${pending.length > 0 ? `(${pending.length} area${pending.length !== 1 ? "s" : ""})` : ""}`}
          </button>
          <p className="text-[11px] text-[#9CA3AF] mt-1.5">
            Writes a new verified copy — the original upload is never touched. A write site some formula reads is
            refused, with the reader named.
          </p>

          {outcome && (
            <div className="mt-4 border border-[#0F766E]/30 bg-[#F0FDFA] rounded-md p-3 text-[12px]" aria-live="polite">
              <div className="font-semibold text-[#0F766E] mb-1">
                {outcome.result.status}
                {outcome.result.verified_opens_in_excel && " · verified to open in Excel"}
              </div>
              <div className="text-[#374151]">
                {outcome.named.length} area{outcome.named.length !== 1 ? "s" : ""} named by you
                {outcome.auto_ops > 0 && <> · {outcome.auto_ops} automatic title op{outcome.auto_ops !== 1 ? "s" : ""}</>}
                {" → "}<span className="font-mono">{outcome.version.label}</span>
              </div>
              {outcome.result.failed.length > 0 && (
                <div className="mt-1 text-[#9F1D1D]">{outcome.result.failed.length} operation(s) failed — see History.</div>
              )}
              {outcome.skipped.length > 0 && (
                <div className="mt-2">
                  <button onClick={() => setShowSkipped(!showSkipped)} className="text-[#8A5A00] underline underline-offset-2">
                    {showSkipped ? "Hide" : "Show"} {outcome.skipped.length} skipped (with reasons)
                  </button>
                  {showSkipped && (
                    <div className="mt-1.5 flex flex-col gap-1 max-h-48 overflow-auto">
                      {outcome.skipped.map((s, i) => (
                        <div key={i} className="text-[11px] text-[#8A5A00] bg-[#FFF1CC] border border-[#8A5A00]/15 rounded px-2 py-1 font-mono">{s}</div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
