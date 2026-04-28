"use client";

import { FormEvent, useMemo, useRef, useState } from "react";
import { ArrowUp, CheckCircle2, Copy, FileSpreadsheet, Loader2, MessageSquare, Sparkles, Terminal } from "lucide-react";
import { apiPost } from "@/lib/api";
import { classifyAssistantCommand } from "@/lib/assistantRouting";
import { addHistoryRecord } from "@/lib/history";
import type { BackendResponse } from "@/types/api";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { PageHeader } from "@/components/ui/PageHeader";
import { StatusBadge } from "@/components/ui/StatusBadge";

type ChatRole = "user" | "assistant" | "status" | "success" | "error" | "partial";

interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  filePath?: string;
  route?: string;
  parserUsed?: string;
}

interface AssistantPageProps {
  onToast: (toast: ToastState) => void;
}

const examples = [
  "create a new Excel file",
  "create a Word document about meeting notes",
  "create a PowerPoint presentation about sales",
  "open chrome"
];

function messageId() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function resultFile(data?: BackendResponse) {
  return (data?.file_path || data?.output_file || "") as string;
}

export function AssistantPage({ onToast }: AssistantPageProps) {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: "hello",
      role: "assistant",
      content: "Tell me what to automate. I can open apps, create Office files, run OCR tools, and control local document workflows."
    }
  ]);
  const [command, setCommand] = useState("");
  const [loading, setLoading] = useState(false);
  const inputRef = useRef<HTMLTextAreaElement | null>(null);

  const routePreview = useMemo(() => {
    if (!command.trim()) return null;
    return classifyAssistantCommand(command);
  }, [command]);

  async function copyPath(path: string) {
    await navigator.clipboard.writeText(path);
    onToast({ tone: "success", message: "Path copied." });
  }

  async function onSubmit(event?: FormEvent) {
    event?.preventDefault();
    const trimmed = command.trim();
    if (!trimmed || loading) return;

    const route = classifyAssistantCommand(trimmed);
    const userMessage: ChatMessage = { id: messageId(), role: "user", content: trimmed };
    const statusId = messageId();
    setMessages((prev) => [
      ...prev,
      userMessage,
      { id: statusId, role: "status", content: "Understanding request, selecting tool, executing, and saving results..." }
    ]);
    setCommand("");
    setLoading(true);

    try {
      if (route.kind === "unsupported") {
        const message = "This looks like neither an Office automation nor an app-launch command.";
        setMessages((prev) => prev.filter((item) => item.id !== statusId).concat({
          id: messageId(),
          role: "error",
          content: message
        }));
        addHistoryRecord({
          name: trimmed,
          type: "system",
          status: "error",
          route: "frontend-classifier",
          message
        });
        return;
      }

      const endpoint = route.endpoint;
      const payload = route.kind === "office"
        ? { app: route.app, raw: route.raw, command: route.command }
        : { command: route.command };
      const result = await apiPost(endpoint, payload);
      const filePath = resultFile(result.data);
      const isPartial = result.status === "partial_success";
      const ok = result.ok;
      const role: ChatRole = ok ? "success" : isPartial ? "partial" : "error";
      const content = result.message || (ok ? "Command completed." : isPartial ? "Command partially completed." : "Command failed.");
      const parserUsed = (result.data as Record<string, unknown>)?.parser_used as string | undefined;

      setMessages((prev) => prev.filter((item) => item.id !== statusId).concat({
        id: messageId(),
        role,
        content,
        filePath,
        route: endpoint,
        parserUsed
      }));

      addHistoryRecord({
        name: trimmed,
        type: route.kind === "office" ? "office" : "app",
        status: ok ? "success" : isPartial ? "partial" : "error",
        route: endpoint,
        message: content,
        filePath: filePath || undefined
      });

      onToast({ tone: ok ? "success" : isPartial ? "info" : "error", message: content });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unexpected frontend error.";
      setMessages((prev) => prev.filter((item) => item.id !== statusId).concat({
        id: messageId(),
        role: "error",
        content: message
      }));
      addHistoryRecord({
        name: trimmed,
        type: "system",
        status: "error",
        route: "frontend",
        message
      });
      onToast({ tone: "error", message });
    } finally {
      setLoading(false);
      window.setTimeout(() => inputRef.current?.focus(), 0);
    }
  }

  return (
    <div className="flex min-h-[calc(100vh-104px)] flex-col">
      <PageHeader
        title="Assistant"
        subtitle="A single command center for app launching, Office automation, OCR workflows, and local desktop tasks."
        action={<StatusBadge tone="info" icon="check">Local first</StatusBadge>}
      />

      <div className="grid flex-1 gap-6 lg:grid-cols-[1fr_320px]">
        <Card className="flex min-h-[620px] flex-col p-0">
          <div className="border-b border-line px-5 py-4">
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600">
                <MessageSquare className="h-5 w-5" />
              </div>
              <div>
                <h2 className="font-bold text-slate-950">Automation chat</h2>
                <p className="text-sm text-slate-500">Commands are routed to Office or system APIs.</p>
              </div>
            </div>
          </div>

          <div className="scrollbar-soft flex-1 space-y-4 overflow-y-auto p-5">
            {messages.map((message) => (
              <div key={message.id} className={`flex ${message.role === "user" ? "justify-end" : "justify-start"}`}>
                <div
                  className={[
                    "max-w-[78%] rounded-3xl border px-4 py-3 text-sm shadow-sm",
                    message.role === "user" ? "border-indigo-200 bg-indigo-600 text-white" :
                    message.role === "success" ? "border-emerald-200 bg-emerald-50 text-emerald-950" :
                    message.role === "error" ? "border-red-200 bg-red-50 text-red-950" :
                    message.role === "partial" ? "border-amber-200 bg-amber-50 text-amber-950" :
                    message.role === "status" ? "border-blue-200 bg-blue-50 text-blue-950" :
                    "border-line bg-white text-slate-700"
                  ].join(" ")}
                >
                  {message.role === "status" ? (
                    <div className="space-y-3">
                      <div className="flex items-center gap-2 font-semibold">
                        <Loader2 className="h-4 w-4 animate-spin" />
                        Processing request
                      </div>
                      <div className="grid gap-2 text-xs">
                        {["Understanding request", "Selecting tool", "Executing", "Saving/result"].map((step) => (
                          <div key={step} className="flex items-center gap-2">
                            <CheckCircle2 className="h-3.5 w-3.5 text-blue-600" />
                            {step}
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : (
                    <>
                      <p className="whitespace-pre-wrap">{message.content}</p>
                      {message.filePath && (
                        <div className="mt-3 rounded-2xl border border-white/60 bg-white/70 p-3 text-xs text-slate-700">
                          <div className="font-semibold text-slate-950">Output file</div>
                          <div className="mt-1 break-all">{message.filePath}</div>
                          <button
                            onClick={() => copyPath(message.filePath || "")}
                            className="mt-2 inline-flex items-center gap-1 font-semibold text-indigo-700"
                          >
                            <Copy className="h-3.5 w-3.5" />
                            Copy path
                          </button>
                        </div>
                      )}
                      {message.parserUsed && (
                        <p className="mt-2 text-[10px] opacity-50">via {message.parserUsed}</p>
                      )}
                    </>
                  )}
                </div>
              </div>
            ))}
          </div>

          <form onSubmit={onSubmit} className="border-t border-line bg-white p-4">
            <div className="rounded-3xl border border-line bg-slate-50 p-3">
              <textarea
                ref={inputRef}
                value={command}
                onChange={(event) => setCommand(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" && !event.shiftKey) {
                    event.preventDefault();
                    onSubmit();
                  }
                }}
                rows={2}
                placeholder="Type a command, e.g. create a new Excel file"
                className="max-h-36 min-h-12 w-full resize-none bg-transparent px-2 py-1 text-sm text-slate-950 outline-none placeholder:text-slate-400"
                disabled={loading}
              />
              <div className="mt-3 flex flex-wrap items-center justify-between gap-3">
                <div className="text-xs text-slate-500">
                  {routePreview?.kind === "office" && `Routes to Office: ${routePreview.app}`}
                  {routePreview?.kind === "system" && "Routes to app launcher"}
                  {routePreview?.kind === "unsupported" && routePreview.reason}
                </div>
                <Button type="submit" variant="primary" disabled={loading || !command.trim()} icon={loading ? <Loader2 className="h-4 w-4 animate-spin" /> : <ArrowUp className="h-4 w-4" />}>
                  Send
                </Button>
              </div>
            </div>
          </form>
        </Card>

        <div className="space-y-5">
          <Card title="Quick starts" description="Use these to verify routing.">
            <div className="space-y-2">
              {examples.map((example) => (
                <button
                  key={example}
                  onClick={() => {
                    setCommand(example);
                    inputRef.current?.focus();
                  }}
                  className="flex w-full items-center gap-3 rounded-2xl border border-line bg-white px-3 py-3 text-left text-sm text-slate-700 transition hover:border-indigo-200 hover:bg-indigo-50"
                >
                  {example.includes("open") ? <Terminal className="h-4 w-4 text-slate-500" /> : <FileSpreadsheet className="h-4 w-4 text-indigo-600" />}
                  <span>{example}</span>
                </button>
              ))}
            </div>
          </Card>

          <Card title="Routing guard" description="Obvious Office commands never go to executable selection.">
            <div className="rounded-2xl bg-gradient-to-br from-indigo-50 to-blue-50 p-4">
              <Sparkles className="h-5 w-5 text-indigo-600" />
              <p className="mt-3 text-sm text-slate-700">
                The frontend classifier sends Office creation and editing commands to <span className="font-semibold">/office/execute</span>.
              </p>
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}
