"use client";

import { useEffect, useState } from "react";
import { Activity, Mic2 } from "lucide-react";
import { AssistantPage } from "@/components/assistant/AssistantPage";
import { DocumentsPage } from "@/components/documents/DocumentsPage";
import { HistoryPage } from "@/components/history/HistoryPage";
import { PdfToolsPage } from "@/components/pdf-tools/PdfToolsPage";
import { SettingsPage } from "@/components/settings/SettingsPage";
import { Sidebar, type SectionId } from "@/components/layout/Sidebar";
import { apiGet } from "@/lib/api";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { Toast, type ToastState } from "@/components/ui/Toast";

export function AppShell() {
  const [active, setActive] = useState<SectionId>("assistant");
  const [backendReachable, setBackendReachable] = useState<boolean | null>(null);
  const [voiceEnabled, setVoiceEnabled] = useState<boolean | null>(null);
  const [toast, setToast] = useState<ToastState | null>(null);

  async function refreshStatus() {
    const result = await apiGet("/voice/status", 12000);
    setBackendReachable(result.reachable);
    setVoiceEnabled(Boolean(result.data?.enabled));
  }

  useEffect(() => {
    refreshStatus();
    const timer = window.setInterval(refreshStatus, 10000);
    return () => window.clearInterval(timer);
  }, []);

  const page =
    active === "assistant" ? <AssistantPage onToast={setToast} /> :
    active === "documents" ? <DocumentsPage onToast={setToast} /> :
    active === "pdf-tools" ? <PdfToolsPage onToast={setToast} /> :
    active === "history" ? <HistoryPage onToast={setToast} /> :
    <SettingsPage onToast={setToast} onStatusChange={refreshStatus} />;

  return (
    <div className="min-h-screen bg-shell text-slate-950">
      <div className="flex min-h-screen flex-col lg:flex-row">
        <Sidebar active={active} onChange={setActive} />
        <main className="min-w-0 flex-1">
          <header className="sticky top-0 z-20 border-b border-line bg-shell/90 px-6 py-4 backdrop-blur">
            <div className="flex items-center justify-between gap-4">
              <div className="text-sm text-slate-500">
                Local Flask backend plus Next.js desktop UI
              </div>
              <div className="flex items-center gap-2">
                <StatusBadge
                  tone={backendReachable ? "success" : backendReachable === false ? "error" : "loading"}
                  icon={backendReachable ? "wifi" : backendReachable === false ? "offline" : "loading"}
                >
                  {backendReachable ? "Backend online" : backendReachable === false ? "Backend offline" : "Checking"}
                </StatusBadge>
                <StatusBadge tone={voiceEnabled ? "info" : "muted"} icon={voiceEnabled ? "check" : undefined}>
                  <Mic2 className="h-3.5 w-3.5" />
                  {voiceEnabled ? "Voice on" : "Voice off"}
                </StatusBadge>
                <button
                  onClick={refreshStatus}
                  className="rounded-full border border-line bg-white p-2 text-slate-500 shadow-sm transition hover:text-indigo-600"
                  aria-label="Refresh backend status"
                >
                  <Activity className="h-4 w-4" />
                </button>
              </div>
            </div>
          </header>
          <div className="mx-auto max-w-7xl px-6 py-6">{page}</div>
        </main>
      </div>
      <Toast toast={toast} onClose={() => setToast(null)} />
    </div>
  );
}
