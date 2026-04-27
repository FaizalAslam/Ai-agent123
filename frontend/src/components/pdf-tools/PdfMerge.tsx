"use client";

import { useState } from "react";
import { Combine, Loader2 } from "lucide-react";
import { apiPost } from "@/lib/api";
import { addHistoryRecord } from "@/lib/history";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { StatusBadge } from "@/components/ui/StatusBadge";

interface PdfMergeProps {
  onToast: (toast: ToastState) => void;
}

export function PdfMerge({ onToast }: PdfMergeProps) {
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("No merge has run yet.");

  async function merge() {
    setLoading(true);
    setMessage("Opening backend file picker...");
    try {
      const result = await apiPost("/pdf/merge", {}, 90000);
      setMessage(result.message);
      addHistoryRecord({
        name: "Merge PDFs",
        type: "pdf",
        status: result.ok ? "success" : "error",
        route: "/pdf/merge",
        message: result.message
      });
      onToast({ tone: result.ok ? "success" : "error", message: result.message });
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="grid gap-6 lg:grid-cols-[1fr_360px]">
      <Card className="min-h-[360px]">
        <div className="flex h-full min-h-[300px] flex-col items-center justify-center rounded-3xl border border-dashed border-indigo-200 bg-gradient-to-br from-indigo-50 to-blue-50 p-8 text-center">
          <div className="flex h-16 w-16 items-center justify-center rounded-3xl bg-white text-indigo-600 shadow-card">
            <Combine className="h-7 w-7" />
          </div>
          <h2 className="mt-5 text-xl font-bold text-slate-950">Merge PDF documents</h2>
          <p className="mt-2 max-w-md text-sm text-slate-500">
            The Flask backend will open a native picker so you can select the PDFs and choose where to save the merged file.
          </p>
          <Button className="mt-6" variant="primary" disabled={loading} onClick={merge} icon={loading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Combine className="h-4 w-4" />}>
            Select PDFs to Merge
          </Button>
        </div>
      </Card>
      <Card title="Result" description="Merge status returned by the backend.">
        <StatusBadge tone={loading ? "loading" : "info"} icon={loading ? "loading" : undefined}>
          {message}
        </StatusBadge>
      </Card>
    </div>
  );
}
