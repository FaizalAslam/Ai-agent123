"use client";

import { useEffect, useMemo, useState } from "react";
import { Copy, History, Search, Trash2 } from "lucide-react";
import { clearHistory, readHistory, removeHistoryRecord } from "@/lib/history";
import type { HistoryRecord, HistoryType } from "@/types/history";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { PageHeader } from "@/components/ui/PageHeader";
import { StatusBadge } from "@/components/ui/StatusBadge";

interface HistoryPageProps {
  onToast: (toast: ToastState) => void;
}

type Filter = "all" | "office" | "pdf" | "ocr" | "app";

const filters: Array<{ id: Filter; label: string }> = [
  { id: "all", label: "All" },
  { id: "office", label: "Office Files" },
  { id: "pdf", label: "PDF Files" },
  { id: "ocr", label: "OCR Text" },
  { id: "app", label: "Apps" }
];

function typeMatches(type: HistoryType, filter: Filter) {
  if (filter === "all") return true;
  if (filter === "pdf") return type === "pdf" || type === "reader" || type === "editor";
  if (filter === "app") return type === "app" || type === "system";
  return type === filter;
}

export function HistoryPage({ onToast }: HistoryPageProps) {
  const [records, setRecords] = useState<HistoryRecord[]>([]);
  const [filter, setFilter] = useState<Filter>("all");
  const [search, setSearch] = useState("");

  function refresh() {
    setRecords(readHistory());
  }

  useEffect(() => {
    refresh();
    window.addEventListener("ai-agent-history-changed", refresh);
    return () => window.removeEventListener("ai-agent-history-changed", refresh);
  }, []);

  const visible = useMemo(() => {
    const needle = search.trim().toLowerCase();
    return records.filter((record) => {
      const matchesFilter = typeMatches(record.type, filter);
      const matchesSearch = !needle || [record.name, record.type, record.message, record.filePath, record.route]
        .filter(Boolean)
        .some((value) => String(value).toLowerCase().includes(needle));
      return matchesFilter && matchesSearch;
    });
  }, [records, filter, search]);

  async function copyPath(path?: string) {
    if (!path) return;
    await navigator.clipboard.writeText(path);
    onToast({ tone: "success", message: "Path copied." });
  }

  function remove(id: string) {
    removeHistoryRecord(id);
    refresh();
  }

  function clearAll() {
    clearHistory();
    refresh();
    onToast({ tone: "info", message: "History cleared." });
  }

  return (
    <div>
      <PageHeader
        title="Outputs & History"
        subtitle="Local browser history for actions taken through the new UI. No server persistence is required."
        action={<Button onClick={clearAll} variant="danger" icon={<Trash2 className="h-4 w-4" />}>Clear history</Button>}
      />

      <Card className="mb-5">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex flex-wrap gap-2">
            {filters.map((item) => (
              <button
                key={item.id}
                onClick={() => setFilter(item.id)}
                className={[
                  "rounded-full border px-4 py-2 text-sm font-semibold transition",
                  filter === item.id
                    ? "border-indigo-600 bg-indigo-600 text-white"
                    : "border-line bg-white text-slate-600 hover:bg-slate-50"
                ].join(" ")}
              >
                {item.label}
              </button>
            ))}
          </div>
          <div className="relative w-full sm:w-80">
            <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search history"
              className="w-full rounded-2xl border border-line bg-slate-50 py-2 pl-9 pr-3 text-sm outline-none focus:border-indigo-300 focus:bg-white"
            />
          </div>
        </div>
      </Card>

      <Card className="overflow-hidden p-0">
        <div className="grid grid-cols-[1.4fr_0.7fr_0.9fr_0.7fr_0.8fr] border-b border-line bg-slate-50 px-5 py-3 text-xs font-bold uppercase tracking-wide text-slate-500">
          <div>Name</div>
          <div>Type</div>
          <div>Created At</div>
          <div>Status</div>
          <div className="text-right">Action</div>
        </div>

        {visible.length === 0 ? (
          <div className="flex min-h-72 items-center justify-center text-center">
            <div>
              <History className="mx-auto h-10 w-10 text-slate-300" />
              <h2 className="mt-3 font-bold text-slate-950">No history yet</h2>
              <p className="mt-1 text-sm text-slate-500">Run an Assistant, OCR, PDF, or reader action to populate this table.</p>
            </div>
          </div>
        ) : (
          <div className="divide-y divide-line">
            {visible.map((record) => (
              <div key={record.id} className="grid grid-cols-[1.4fr_0.7fr_0.9fr_0.7fr_0.8fr] items-center gap-3 px-5 py-4 text-sm">
                <div className="min-w-0">
                  <div className="truncate font-semibold text-slate-950">{record.name}</div>
                  <div className="truncate text-xs text-slate-500">{record.message}</div>
                  {record.filePath && <div className="truncate text-xs text-indigo-600">{record.filePath}</div>}
                </div>
                <div className="capitalize text-slate-600">{record.type}</div>
                <div className="text-slate-500">{new Date(record.createdAt).toLocaleString()}</div>
                <div>
                  <StatusBadge tone={record.status === "success" ? "success" : "error"} icon={record.status === "success" ? "check" : "alert"}>
                    {record.status}
                  </StatusBadge>
                </div>
                <div className="flex justify-end gap-2">
                  <button
                    disabled={!record.filePath}
                    onClick={() => copyPath(record.filePath)}
                    className="rounded-full border border-line p-2 text-slate-500 transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40"
                    aria-label="Copy file path"
                  >
                    <Copy className="h-4 w-4" />
                  </button>
                  <button
                    onClick={() => remove(record.id)}
                    className="rounded-full border border-line p-2 text-slate-500 transition hover:bg-red-50 hover:text-red-600"
                    aria-label="Remove history item"
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}
