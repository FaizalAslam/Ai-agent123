"use client";

import { Bot, FileText, History, Mic2, Settings, Wrench } from "lucide-react";

export type SectionId = "assistant" | "documents" | "pdf-tools" | "history" | "settings";

interface SidebarProps {
  active: SectionId;
  onChange: (section: SectionId) => void;
}

const items = [
  { id: "assistant", label: "Assistant", icon: Bot },
  { id: "documents", label: "Documents", icon: FileText },
  { id: "pdf-tools", label: "PDF Tools", icon: Wrench },
  { id: "history", label: "Outputs & History", icon: History },
  { id: "settings", label: "Settings", icon: Settings }
] satisfies Array<{ id: SectionId; label: string; icon: typeof Bot }>;

export function Sidebar({ active, onChange }: SidebarProps) {
  return (
    <aside className="flex w-full shrink-0 flex-col border-b border-line bg-white px-4 py-5 lg:sticky lg:top-0 lg:h-screen lg:w-[250px] lg:border-b-0 lg:border-r">
      <div className="flex items-center gap-3 px-2">
        <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-600 to-blue-500 text-white shadow-card">
          <Bot className="h-5 w-5" />
        </div>
        <div>
          <div className="text-sm font-bold text-slate-950">AI Agent</div>
          <div className="text-xs text-slate-500">Local automation</div>
        </div>
      </div>

      <nav className="mt-6 grid gap-1 sm:grid-cols-5 lg:mt-8 lg:block lg:space-y-1">
        {items.map((item) => {
          const Icon = item.icon;
          const selected = active === item.id;
          return (
            <button
              key={item.id}
              onClick={() => onChange(item.id)}
              className={[
                "flex w-full items-center gap-3 rounded-2xl px-3 py-3 text-left text-sm font-semibold transition",
                selected
                  ? "bg-indigo-50 text-indigo-700 shadow-sm"
                  : "text-slate-600 hover:bg-slate-50 hover:text-slate-950"
              ].join(" ")}
            >
              <Icon className="h-5 w-5" />
              <span>{item.label}</span>
            </button>
          );
        })}
      </nav>

      <div className="mt-5 rounded-3xl border border-line bg-slate-50 p-4 lg:mt-auto">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-full bg-white text-indigo-600 shadow-sm">
            <Mic2 className="h-4 w-4" />
          </div>
          <div>
            <div className="text-sm font-semibold text-slate-950">Desktop mode</div>
            <div className="text-xs text-slate-500">Backend stays local</div>
          </div>
        </div>
      </div>
    </aside>
  );
}
