import { useEffect, useState } from "react";
import { NavLink } from "react-router-dom";
import { useStore } from "../store";
import { downloadUrl, health, type AppHealth } from "../services/api";

// 1.7.2: two groups. "Workflow" is the path in order -- upload, read the
// findings, fix, verify, deliver. "Tools" are aids you reach for when needed.
const navGroups: { title: string; items: { to: string; label: string; icon: string; step?: number }[] }[] = [
  {
    title: "WORKFLOW",
    items: [
      { to: "/", label: "Workbook", icon: "📄", step: 1 },
      { to: "/findings", label: "Findings", icon: "🔍", step: 2 },
      { to: "/prep", label: "Prep", icon: "⚙️", step: 3 },
      { to: "/recalculate", label: "Recalculate", icon: "▶", step: 4 },
      { to: "/reports", label: "Deliver", icon: "📋", step: 5 },
    ],
  },
  {
    title: "TOOLS",
    items: [
      { to: "/assistant", label: "Assistant", icon: "💬" },
      { to: "/grid-namer", label: "Grid Namer", icon: "🏷️" },
      { to: "/history", label: "History", icon: "🕐" },
      { to: "/mind", label: "Mind", icon: "🧠" },
    ],
  },
];

interface Props {
  collapsed: boolean;
  onToggle: () => void;
}

export default function Sidebar({ collapsed, onToggle }: Props) {
  const report = useStore((s) => s.report);
  const sessionId = useStore((s) => s.sessionId);
  const currentVersion = useStore((s) => s.currentVersion);
  const [hp, setHp] = useState<AppHealth | null>(null);
  useEffect(() => { health().then(setHp).catch(() => {}); }, []);
  const latest = sessionId && currentVersion ? { href: downloadUrl(sessionId, currentVersion.file_name), short: currentVersion.label.split(" — ")[0], ...currentVersion } : null;
  return (
    <aside
      className={`flex flex-col bg-[#1F3A5F] text-white transition-all duration-200 ${
        collapsed ? "w-14" : "w-52"
      } flex-shrink-0 h-screen sticky top-0`}
    >
      <div className="flex items-center gap-2 px-3 py-4 border-b border-white/10">
        {!collapsed && (
          <span className="font-semibold text-sm tracking-tight text-white/90 truncate">
            Mind Ready
          </span>
        )}
        <button
          onClick={onToggle}
          className="ml-auto p-1 rounded hover:bg-white/10 transition-colors text-white/60 hover:text-white"
          title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
        >
          <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
            {collapsed ? (
              <path d="M6 3l5 5-5 5V3z" />
            ) : (
              <path d="M10 3L5 8l5 5V3z" />
            )}
          </svg>
        </button>
      </div>
      <nav className="flex flex-col gap-0.5 p-2 flex-1 min-h-0 overflow-y-auto">
        {navGroups.map((group) => (
          <div key={group.title} className="flex flex-col gap-0.5 mb-2">
            {!collapsed && <div className="px-3 pt-2 pb-1 text-[10px] font-semibold tracking-wider text-white/40">{group.title}</div>}
            {collapsed && <div className="border-t border-white/10 my-1" />}
            {group.items.map(({ to, label, icon, step }) => (
              <NavLink
                key={to}
                to={to}
                end={to === "/"}
                className={({ isActive }) =>
                  `flex items-center gap-3 px-3 py-2.5 rounded-md text-sm transition-colors ${
                    isActive
                      ? "bg-white/15 text-white font-medium"
                      : "text-white/60 hover:bg-white/10 hover:text-white"
                  }`
                }
                title={step ? `Step ${step}: ${label}` : label}
              >
                {step ? (
                  <span className="w-5 h-5 rounded-full border border-white/40 text-[10px] font-mono flex items-center justify-center flex-shrink-0">{step}</span>
                ) : (
                  <span className="text-base flex-shrink-0">{icon}</span>
                )}
                {!collapsed && <span className="truncate">{label}</span>}
              </NavLink>
            ))}
          </div>
        ))}
      </nav>
      {latest && (
        <div className={`border-t border-white/10 ${collapsed ? "p-2" : "px-3 py-3"}`} aria-label="Download">
          {!collapsed && <div className="text-[10px] font-semibold tracking-wider text-white/40 mb-1.5">DOWNLOAD</div>}
          {!collapsed && (
            <div className="text-[11px] text-white/60 truncate" title={latest.label}>
              Latest: <span className="font-mono">{latest.short}</span>
              {latest.verified_opens_in_excel ? <span className="text-[#9BE7B8]"> · verified</span> : null}
            </div>
          )}
          <a
            href={latest.href}
            download={latest.file_name}
            className={`mt-1.5 flex items-center justify-center gap-1.5 rounded-md bg-white text-[#1F3A5F] font-semibold hover:bg-white/90 ${collapsed ? "text-sm py-1.5" : "text-xs py-1.5"}`}
            title={`Download ${latest.file_name}`}
          >
            <span aria-hidden>⬇</span>
            {!collapsed && <span className="truncate">Download {latest.short}</span>}
          </a>
          {!collapsed && <div className="text-[10px] text-white/30 mt-1 truncate font-mono" title={latest.file_name}>{latest.file_name}</div>}
        </div>
      )}
      {/* App release (from /api/health) -- always visible, bottom-left, grey */}
      <div
        className={`mt-auto shrink-0 border-t border-white/10 font-mono text-[#9CA3AF] ${collapsed ? "px-1 py-2 text-[10px] text-center" : "px-3 py-3 text-[11px]"}`}
        title={hp ? `Excel Upload Preparation v${hp.version} · ${hp.rules} rules` : undefined}
      >
        {collapsed ? (hp ? `v${hp.version}` : "…") : hp ? `v${hp.version} · ${hp.rules} rules` : "Excel-verified outputs"}
      </div>
      {!collapsed && report && (
        <div className="px-3 pb-3">
          <div className={`rounded-md px-2 py-1.5 text-xs font-mono font-semibold text-center ${
            report.status === "PASS" ? "bg-[#DFF5E6] text-[#0F6E3A]"
            : report.status === "ERROR" ? "bg-[#FDE2E2] text-[#9F1D1D]"
            : report.status === "WARNING" ? "bg-[#FFF1CC] text-[#8A5A00]"
            : report.status === "NOT_SUPPORTED" ? "bg-[#E9ECEF] text-[#4B5563]"
            : "bg-[#FFE4C7] text-[#7A3E00]"
          }`}>
            {report.status}
          </div>
        </div>
      )}
    </aside>
  );
}
