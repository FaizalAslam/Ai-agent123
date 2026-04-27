"use client";

import { useEffect, useState } from "react";
import { BookOpen, ChevronLeft, ChevronRight, Loader2, Pause, Play, Square } from "lucide-react";
import { apiGet, apiPost } from "@/lib/api";
import { addHistoryRecord } from "@/lib/history";
import type { BackendResponse } from "@/types/api";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { StatusBadge } from "@/components/ui/StatusBadge";

interface PdfReaderProps {
  onToast: (toast: ToastState) => void;
}

export function PdfReader({ onToast }: PdfReaderProps) {
  const [status, setStatus] = useState<BackendResponse | null>(null);
  const [loading, setLoading] = useState<string | null>(null);
  const [active, setActive] = useState(false);
  const [speed, setSpeed] = useState(150);

  async function refreshStatus() {
    const result = await apiGet("/reader/status", 15000);
    if (result.reachable && result.data) {
      setStatus(result.data);
      setActive(Boolean(result.data.is_reading));
      if (typeof result.data.speed === "number") setSpeed(result.data.speed);
    }
  }

  async function readerAction(label: string, route: string, body?: unknown) {
    setLoading(route);
    try {
      const result = await apiPost(route, body || {}, 90000);
      setStatus(result.data || null);
      if (route === "/reader/open") setActive(result.ok);
      if (route === "/reader/stop") setActive(false);
      addHistoryRecord({
        name: label,
        type: "reader",
        status: result.ok ? "success" : "error",
        route,
        message: result.message
      });
      onToast({ tone: result.ok ? "success" : "error", message: result.message });
      await refreshStatus();
    } finally {
      setLoading(null);
    }
  }

  useEffect(() => {
    refreshStatus();
  }, []);

  useEffect(() => {
    if (!active) return;
    const timer = window.setInterval(refreshStatus, 1500);
    return () => window.clearInterval(timer);
  }, [active]);

  const currentPage = typeof status?.current_page === "number" ? status.current_page + 1 : 0;
  const totalPages = typeof status?.total_pages === "number" ? status.total_pages : 0;
  const reading = Boolean(status?.is_reading);
  const paused = Boolean(status?.is_paused);

  return (
    <div className="grid gap-6 lg:grid-cols-[420px_1fr]">
      <Card title="PDF reader controls" description="Select a PDF from the backend native file picker.">
        <Button
          variant="primary"
          disabled={Boolean(loading)}
          onClick={() => readerAction("Open PDF reader", "/reader/open")}
          icon={loading === "/reader/open" ? <Loader2 className="h-4 w-4 animate-spin" /> : <BookOpen className="h-4 w-4" />}
          className="w-full"
        >
          Select PDF and Read
        </Button>

        <div className="mt-5 grid grid-cols-2 gap-3">
          <Button disabled={Boolean(loading) || !reading} onClick={() => readerAction("Pause PDF reader", "/reader/pause")} icon={<Pause className="h-4 w-4" />}>
            Pause
          </Button>
          <Button disabled={Boolean(loading) || !reading} onClick={() => readerAction("Resume PDF reader", "/reader/resume")} icon={<Play className="h-4 w-4" />}>
            Resume
          </Button>
          <Button disabled={Boolean(loading) || !reading} onClick={() => readerAction("Previous PDF page", "/reader/prev")} icon={<ChevronLeft className="h-4 w-4" />}>
            Previous
          </Button>
          <Button disabled={Boolean(loading) || !reading} onClick={() => readerAction("Next PDF page", "/reader/next")} icon={<ChevronRight className="h-4 w-4" />}>
            Next
          </Button>
          <Button className="col-span-2" variant="danger" disabled={Boolean(loading) || !reading} onClick={() => readerAction("Stop PDF reader", "/reader/stop")} icon={<Square className="h-4 w-4" />}>
            Stop Reading
          </Button>
        </div>

        <div className="mt-5 rounded-3xl border border-line bg-slate-50 p-4">
          <label className="text-sm font-semibold text-slate-700">Reading speed</label>
          <div className="mt-3 flex items-center gap-3">
            <input
              type="range"
              min={80}
              max={320}
              step={10}
              value={speed}
              onChange={(event) => setSpeed(Number(event.target.value))}
              onMouseUp={() => readerAction("Update PDF reader speed", "/reader/speed", { speed })}
              className="w-full accent-indigo-600"
            />
            <input
              type="number"
              min={80}
              max={320}
              value={speed}
              onChange={(event) => setSpeed(Number(event.target.value))}
              onBlur={() => readerAction("Update PDF reader speed", "/reader/speed", { speed })}
              className="w-20 rounded-2xl border border-line bg-white px-3 py-2 text-sm"
            />
          </div>
        </div>
      </Card>

      <Card title="Reader status" description="Live reader state from the Flask backend.">
        <div className="grid gap-4 sm:grid-cols-3">
          <div className="rounded-3xl border border-line bg-slate-50 p-5">
            <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">State</div>
            <div className="mt-3">
              <StatusBadge tone={reading ? (paused ? "info" : "success") : "muted"} icon={reading ? "check" : undefined}>
                {reading ? (paused ? "Paused" : "Reading") : "Idle"}
              </StatusBadge>
            </div>
          </div>
          <div className="rounded-3xl border border-line bg-slate-50 p-5">
            <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Page</div>
            <div className="mt-2 text-3xl font-bold text-slate-950">{currentPage || "-"}</div>
            <div className="text-sm text-slate-500">of {totalPages || "-"}</div>
          </div>
          <div className="rounded-3xl border border-line bg-slate-50 p-5">
            <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Speed</div>
            <div className="mt-2 text-3xl font-bold text-slate-950">{speed}</div>
            <div className="text-sm text-slate-500">WPM</div>
          </div>
        </div>
      </Card>
    </div>
  );
}
