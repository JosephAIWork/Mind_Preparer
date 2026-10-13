import { useEffect, useState } from "react";
import { useStore } from "../store";
import * as api from "../services/api";
import type { CellWindow, ModelMap } from "../services/api";
import MindLoopSection from "./MindLoopSection";

/**
 * Model map (1.9.0): where the model's inputs are, which cell the run loop sets,
 * which settings stay for a run but change between runs (and how), and where
 * the outputs land. Everything comes from the workbook itself -- its macros,
 * names, dropdowns, sheet names and formulas -- and every line says why.
 * "Show" opens the cells around the item.
 */

function fmt(v: unknown): string {
  if (v === null || v === undefined) return "empty";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : v.toPrecision(6).replace(/\.?0+$/, "");
  return String(v);
}

function splitRef(cell: string): { sheet: string; ref: string } | null {
  const i = cell.lastIndexOf("!");
  if (i < 0) return null;
  const sheet = cell.slice(0, i).replace(/^'(.*)'$/, "$1");
  return { sheet, ref: cell.slice(i + 1).split(":")[0] };
}

function CellPeek({ sessionId, cell, onClose }: { sessionId: string; cell: string; onClose: () => void }) {
  const [win, setWin] = useState<CellWindow | null>(null);
  const [err, setErr] = useState<string | null>(null);
  useEffect(() => {
    const sr = splitRef(cell);
    if (!sr) { setErr("no sheet in the address"); return; }
    api.cellWindow(sessionId, sr.sheet, sr.ref, 3, 3).then(setWin).catch((e) => setErr(String(e?.message ?? e)));
  }, [sessionId, cell]);
  return (
    <div className="mt-2 border border-[#E5E7EB] rounded-lg bg-[#F9FAFB] p-3 text-[12px]">
      <div className="flex items-center justify-between mb-2">
        <span className="font-mono text-[#111827]">{cell}</span>
        <button onClick={onClose} className="text-[#6B7280] hover:text-[#111827]">close</button>
      </div>
      {err && <div className="text-[#9F1D1D]">{err}</div>}
      {win && (
        <table className="border-collapse">
          <thead>
            <tr><th className="px-2 text-[#9CA3AF] font-normal"></th>{win.columns.map((c) => <th key={c} className="px-2 text-[#9CA3AF] font-normal">{c}</th>)}</tr>
          </thead>
          <tbody>
            {win.rows.map((r) => (
              <tr key={r.row}>
                <td className="px-2 text-[#9CA3AF]">{r.row}</td>
                {r.cells.map((c) => (
                  <td key={c.ref} title={c.formula ?? undefined} className={`px-2 py-0.5 border border-[#E5E7EB] bg-white max-w-[220px] truncate ${c.focus ? "ring-2 ring-[#2563EB]" : ""}`}>
                    {fmt(c.value)}{c.formula ? <span className="text-[#9CA3AF]"> ƒ</span> : ""}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function Row({ sessionId, cell, title, sub, right, evidence }: { sessionId: string; cell: string | null; title: string; sub?: string | null; right?: string | null; evidence?: string | null }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="py-2 border-b border-[#F3F4F6] last:border-b-0">
      <div className="flex items-start gap-3">
        <div className="flex-1 min-w-0">
          <div className="text-[13px] text-[#111827]">
            {cell && <span className="font-mono font-semibold mr-2">{cell}</span>}
            <span>{title}</span>
          </div>
          {sub && <div className="text-[12px] text-[#6B7280] mt-0.5">{sub}</div>}
          {evidence && <div className="text-[11px] text-[#9CA3AF] mt-0.5 font-mono truncate" title={evidence}>{evidence}</div>}
        </div>
        {right && <span className="text-[12px] text-[#374151] whitespace-nowrap">{right}</span>}
        {cell && cell.includes("!") && !cell.includes("<") && (
          <button onClick={() => setOpen(!open)} className="text-[12px] px-2 py-1 rounded border border-[#E5E7EB] bg-white hover:bg-[#F3F4F6] text-[#374151] whitespace-nowrap">
            {open ? "Hide" : "Show"}
          </button>
        )}
      </div>
      {open && cell && <CellPeek sessionId={sessionId} cell={cell} onClose={() => setOpen(false)} />}
    </div>
  );
}

function Section({ title, hint, children, count }: { title: string; hint: string; children: React.ReactNode; count?: number }) {
  return (
    <section className="border border-[#E5E7EB] rounded-xl bg-white">
      <div className="px-4 py-3 border-b border-[#E5E7EB] flex items-baseline gap-3">
        <h2 className="font-semibold text-[14px] text-[#111827]">{title}</h2>
        {typeof count === "number" && <span className="text-[11px] text-[#6B7280]">{count}</span>}
        <span className="text-[12px] text-[#6B7280]">{hint}</span>
      </div>
      <div className="px-4 py-1">{children}</div>
    </section>
  );
}

const Empty = ({ text }: { text: string }) => <div className="py-3 text-[12px] text-[#9CA3AF]">{text}</div>;

export default function ModelMapScreen() {
  const sessionId = useStore((s) => s.sessionId);
  const summary = useStore((s) => s.summary);
  const [map, setMap] = useState<ModelMap | null>(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);
  const [showAllSettings, setShowAllSettings] = useState(false);

  const load = (refresh = false) => {
    if (!sessionId) return;
    setBusy(true);
    setErr(null);
    api.modelMap(sessionId, refresh).then((r) => setMap(r.map)).catch((e) => setErr(String(e?.message ?? e))).finally(() => setBusy(false));
  };
  useEffect(() => { load(false); }, [sessionId]);

  if (!sessionId) {
    return <div className="p-8 text-[13px] text-[#6B7280]">Upload and analyse a workbook first (Workbook screen). The map is read off the analysed copy.</div>;
  }
  const lp = map?.loop;
  const rs = map?.run_settings;
  const settingsShown = showAllSettings ? rs?.named_settings ?? [] : (rs?.named_settings ?? []).slice(0, 12);
  const labelledShown = showAllSettings ? rs?.labelled ?? [] : (rs?.labelled ?? []).slice(0, 12);

  return (
    <div className="p-6 max-w-[1100px] space-y-4">
      <div className="flex items-center gap-3 flex-wrap">
        <h1 className="text-[18px] font-semibold text-[#111827]">Model map</h1>
        {summary?.file_name && <span className="text-[12px] text-[#6B7280]">{summary.file_name}</span>}
        <button onClick={() => load(true)} disabled={busy} className="ml-auto text-[12px] px-3 py-1.5 rounded border border-[#E5E7EB] bg-white hover:bg-[#F3F4F6] text-[#374151] disabled:opacity-50">
          {busy ? "Reading…" : "Read again"}
        </button>
      </div>
      <p className="text-[12px] text-[#6B7280]">
        Where the inputs are, which cell the run loop sets, which settings stay for a run but change between runs, and where the outputs land.
        Read off the workbook itself: its macros, defined names, dropdowns, sheet names and formulas. Each line says why. Nothing is guessed.
      </p>
      {err && <div className="text-[12px] text-[#9F1D1D] border border-[#9F1D1D]/20 bg-[#FDE2E2] rounded-lg px-3 py-2">{err}</div>}
      {busy && !map && <div className="text-[12px] text-[#6B7280]">Reading the macros, names and dropdowns…</div>}
      {map && lp && rs && (
        <>
          <Section title="The loop" hint={map.has_macros ? "the cell the macro sets on every pass, and the cells that bound the run" : "no macros in this workbook"} count={lp.drivers.length || undefined}>
            {lp.found ? (
              <>
                {lp.drivers.map((d) => (
                  <Row key={d.cell ?? d.evidence} sessionId={sessionId} cell={d.cell} title={`set to ${d.how}`} sub={[d.name ? `name ${d.name}` : null, d.label ? `label “${d.label}”` : null, `macro ${d.macro}, For ${d.loop_var}`].filter(Boolean).join(" · ")} evidence={d.evidence} />
                ))}
                {lp.bounds.map((b) => (
                  <Row key={b.role + (b.cell ?? "")} sessionId={sessionId} cell={b.cell} title={`${b.role} value of the loop`} sub={[b.name ? `name ${b.name}` : null, b.label ? `label “${b.label}”` : null].filter(Boolean).join(" · ")} right={b.formula ? `holds ${b.formula}` : `= ${fmt(b.value)}`} />
                ))}
                {lp.how_to_run && <div className="py-2 text-[12px] text-[#374151] bg-[#EFF6FF] border border-[#BFDBFE] rounded-lg px-3 my-2">{lp.how_to_run}</div>}
              </>
            ) : (
              <>
                {map.notes.map((n) => <div key={n} className="py-2 text-[12px] text-[#374151]">{n}</div>)}
                {lp.item_cells.map((b) => (
                  <Row key={b.cell ?? b.macro} sessionId={sessionId} cell={b.cell} title="item cell: the sheet is copied once per item and this cell names the item" sub={b.how} evidence={b.evidence} />
                ))}
                {lp.key_cells.map((k) => (
                  <Row key={k.cell ?? ""} sessionId={sessionId} cell={k.cell} title={`read by ${k.readers} formulas`} sub={[k.name ? `name ${k.name}` : null, k.label ? `“${k.label}”` : null].filter(Boolean).join(" · ")} right={`= ${fmt(k.value)}`} />
                ))}
              </>
            )}
          </Section>

          {lp.found && <MindLoopSection sessionId={sessionId} />}

          <Section title="Inputs" hint="data the loop walks through, and the sheets that hold inputs" count={map.inputs.sheets.length}>
            {map.inputs.per_item.map((p) => <Row key={p.cell} sessionId={sessionId} cell={p.cell} title={p.what} sub={`macro ${p.macro}`} />)}
            {map.inputs.sheets.map((s) => <Row key={s.sheet} sessionId={sessionId} cell={null} title={s.sheet} sub={`${s.formulas} formulas · ${s.why}${s.state && s.state !== "visible" ? ` · ${s.state}` : ""}`} />)}
            {map.inputs.sheets.length === 0 && map.inputs.per_item.length === 0 && <Empty text="No sheet whose name says input, data or assumptions, and no macro that reads one row per loop value." />}
          </Section>

          <Section title="Run settings — switches" hint="cells that choose how a run goes: dropdowns and the cells the macros test" count={rs.switches.length}>
            {rs.switches.map((sw) => (
              <Row key={sw.cell ?? ""} sessionId={sessionId} cell={sw.cell} title={sw.label ?? sw.name ?? "switch"} sub={[sw.name && sw.label ? `name ${sw.name}` : null, `options: ${sw.options.join(" / ") || "—"}`, sw.how].filter(Boolean).join(" · ")} right={`= ${fmt(sw.value)}`} evidence={sw.evidence} />
            ))}
            {rs.switches.length === 0 && <Empty text="No dropdown and no named cell the macros test." />}
            {rs.status.length > 0 && (
              <div className="mt-2 pt-2 border-t border-dashed border-[#E5E7EB]">
                <div className="text-[11px] text-[#9CA3AF] mb-1">Status cells the macro writes and reads back (not settings)</div>
                {rs.status.map((st) => <Row key={st.cell ?? ""} sessionId={sessionId} cell={st.cell} title={st.label ?? st.name ?? ""} sub={[st.name ? `name ${st.name}` : null, `written in ${st.written_in.join(", ")}`].filter(Boolean).join(" · ")} right={`= ${fmt(st.value)}`} />)}
              </div>
            )}
          </Section>

          <Section title="Run settings — named values and labelled cells" hint="single constants with a defined name that formulas read, and label | value pairs on small control sheets" count={rs.named_settings.length + rs.labelled.length}>
            {settingsShown.map((n) => <Row key={n.cell ?? ""} sessionId={sessionId} cell={n.cell} title={n.name ?? ""} sub={[n.label ? `“${n.label}”` : null, `read by ${n.readers} formulas by name`].filter(Boolean).join(" · ")} right={`= ${fmt(n.value)}`} />)}
            {labelledShown.map((n) => <Row key={n.cell} sessionId={sessionId} cell={n.cell} title={n.label} right={`= ${fmt(n.value)}`} />)}
            {rs.named_settings.length + rs.labelled.length > 24 && (
              <button onClick={() => setShowAllSettings(!showAllSettings)} className="my-2 text-[12px] text-[#2563EB]">{showAllSettings ? "Show fewer" : `Show all ${rs.named_settings.length + rs.labelled.length}`}</button>
            )}
            {rs.named_settings.length + rs.labelled.length === 0 && <Empty text="None found." />}
          </Section>

          <Section title="Run settings — assumption tables" hint="named ranges on input sheets that formulas read" count={rs.tables.length}>
            {rs.tables.map((t) => <Row key={t.name} sessionId={sessionId} cell={t.cell} title={t.name} right={`read by ${t.readers} formulas`} />)}
            {rs.tables.length === 0 && <Empty text="No named table on an input sheet is read by name." />}
          </Section>

          <Section title="Outputs" hint="where the macros put results, and the sheets whose name says output" count={map.outputs.cells.length + map.outputs.sheets.length}>
            {map.outputs.cells.map((o) => (
              <Row key={(o.cell ?? "") + (o.from ?? "")} sessionId={sessionId} cell={o.cell} title={o.from ? `← ${o.from}${o.from_name ? ` (${o.from_name})` : ""}` : o.when} sub={[o.name ? `name ${o.name}` : null, `${o.written_by}: ${o.when}`].filter(Boolean).join(" · ")} evidence={o.evidence} />
            ))}
            {map.outputs.sheets.map((s) => <Row key={s.sheet} sessionId={sessionId} cell={null} title={s.sheet} sub={`${s.formulas} formulas · ${s.why}`} />)}
            {map.outputs.saves.map((sv, i) => <Row key={i} sessionId={sessionId} cell={sv.path_from} title={`saved to the path held here${sv.name ? ` (${sv.name})` : ""}`} sub={`macro ${sv.in}`} />)}
            {map.outputs.cells.length + map.outputs.sheets.length === 0 && <Empty text="No macro writes results and no sheet name says output." />}
          </Section>

          {map.macros.procedures.length > 0 && (
            <Section title="Macros" hint="what the workbook can run" count={map.macros.procedures.length}>
              {map.macros.procedures.map((p) => (
                <div key={p.name} className="py-1.5 text-[12px] text-[#374151] border-b border-[#F3F4F6] last:border-b-0">
                  <span className="font-mono">{p.name}</span>
                  {map.macros.entry_points.includes(p.name) && <span className="ml-2 text-[11px] px-1.5 py-0.5 rounded bg-[#DFF5E6] text-[#0F6E3A]">runs the loop</span>}
                  <span className="text-[#9CA3AF]"> · {p.module} · {p.statements} statements{p.calls.length ? ` · calls ${p.calls.join(", ")}` : ""}</span>
                </div>
              ))}
            </Section>
          )}
        </>
      )}
    </div>
  );
}
