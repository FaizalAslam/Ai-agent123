"use client";

import { useEffect, useState } from "react";
import { Mic, MicOff, RefreshCw, Server } from "lucide-react";
import { apiGet, apiPost, backendDisplayUrl } from "@/lib/api";
import { addHistoryRecord } from "@/lib/history";
import type { BackendResponse } from "@/types/api";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { PageHeader } from "@/components/ui/PageHeader";
import { StatusBadge } from "@/components/ui/StatusBadge";

interface SettingsPageProps {
  onToast: (toast: ToastState) => void;
  onStatusChange: () => void;
}

export function SettingsPage({ onToast, onStatusChange }: SettingsPageProps) {
  const [voice, setVoice] = useState<BackendResponse | null>(null);
  const [reachable, setReachable] = useState<boolean | null>(null);
  const [loading, setLoading] = useState<string | null>(null);

  async function refreshVoice() {
    setLoading("/voice/status");
    try {
      const result = await apiGet("/voice/status", 15000);
      setReachable(result.reachable);
      setVoice(result.data || null);
      onStatusChange();
      if (!result.reachable) {
        onToast({ tone: "error", message: result.message });
      }
    } finally {
      setLoading(null);
    }
  }

  async function voiceAction(label: string, route: "/voice/start" | "/voice/stop") {
    setLoading(route);
    try {
      const result = await apiPost(route, {}, 30000);
      addHistoryRecord({
        name: label,
        type: "system",
        status: result.ok ? "success" : "error",
        route,
        message: result.message
      });
      onToast({ tone: result.ok ? "success" : "error", message: result.message });
      await refreshVoice();
    } finally {
      setLoading(null);
    }
  }

  useEffect(() => {
    refreshVoice();
  }, []);

  const voiceAvailable = Boolean(voice?.available);
  const voiceEnabled = Boolean(voice?.enabled);

  return (
    <div>
      <PageHeader
        title="Settings"
        subtitle="Local backend connection and voice listener controls. Secrets are never displayed."
        action={<Button onClick={refreshVoice} icon={<RefreshCw className={`h-4 w-4 ${loading === "/voice/status" ? "animate-spin" : ""}`} />}>Refresh</Button>}
      />

      <div className="grid gap-6 lg:grid-cols-2">
        <Card title="Backend connection" description="The Next frontend proxies requests to the Flask service.">
          <div className="rounded-3xl border border-line bg-slate-50 p-5">
            <div className="flex items-center gap-3">
              <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-white text-indigo-600 shadow-sm">
                <Server className="h-5 w-5" />
              </div>
              <div>
                <div className="text-sm font-semibold text-slate-950">API base URL</div>
                <div className="text-sm text-slate-500">{backendDisplayUrl}</div>
              </div>
            </div>
            <div className="mt-4">
              <StatusBadge
                tone={reachable ? "success" : reachable === false ? "error" : "loading"}
                icon={reachable ? "wifi" : reachable === false ? "offline" : "loading"}
              >
                {reachable ? "Backend reachable" : reachable === false ? "Backend unavailable" : "Checking backend"}
              </StatusBadge>
            </div>
          </div>
        </Card>

        <Card title="Voice controls" description="Start or stop the existing backend voice listener.">
          <div className="rounded-3xl border border-line bg-slate-50 p-5">
            <div className="flex items-center justify-between gap-4">
              <div>
                <div className="text-sm font-semibold text-slate-950">Voice listener</div>
                <div className="mt-1 text-sm text-slate-500">
                  {voiceAvailable ? "Speech module available" : voice?.message || "Voice status unknown"}
                </div>
              </div>
              <StatusBadge tone={voiceEnabled ? "success" : "muted"} icon={voiceEnabled ? "check" : undefined}>
                {voiceEnabled ? "Enabled" : "Disabled"}
              </StatusBadge>
            </div>

            <div className="mt-5 flex flex-wrap gap-3">
              <Button
                variant="primary"
                disabled={Boolean(loading) || !voiceAvailable || voiceEnabled}
                onClick={() => voiceAction("Start voice listener", "/voice/start")}
                icon={<Mic className="h-4 w-4" />}
              >
                Start Voice
              </Button>
              <Button
                variant="danger"
                disabled={Boolean(loading) || !voiceAvailable || !voiceEnabled}
                onClick={() => voiceAction("Stop voice listener", "/voice/stop")}
                icon={<MicOff className="h-4 w-4" />}
              >
                Stop Voice
              </Button>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}
