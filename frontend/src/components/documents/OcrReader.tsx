"use client";

import { useEffect, useState } from "react";
import { Camera, Clipboard, Copy, FileImage, FileText, Loader2, Save, Scissors, Volume2, VolumeX } from "lucide-react";
import { apiGet, apiPost } from "@/lib/api";
import { addHistoryRecord } from "@/lib/history";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { StatusBadge } from "@/components/ui/StatusBadge";

interface OcrReaderProps {
  onToast: (toast: ToastState) => void;
}

export function OcrReader({ onToast }: OcrReaderProps) {
  const [text, setText] = useState("");
  const [loading, setLoading] = useState<string | null>(null);
  const [message, setMessage] = useState("Ready for OCR capture.");

  async function runAction(label: string, route: string, options?: { expectText?: boolean }) {
    setLoading(route);
    setMessage(`${label} running...`);
    try {
      const result = await apiPost(route, {}, 90000);
      const nextText = typeof result.data?.text === "string" ? result.data.text : "";
      if (options?.expectText && nextText) setText(nextText);
      setMessage(result.message);
      addHistoryRecord({
        name: label,
        type: "ocr",
        status: result.ok ? "success" : "error",
        route,
        message: result.message
      });
      onToast({ tone: result.ok ? "success" : "error", message: result.message });
    } finally {
      setLoading(null);
    }
  }

  async function copyText() {
    if (!text.trim()) {
      onToast({ tone: "error", message: "No OCR text to copy." });
      return;
    }
    await navigator.clipboard.writeText(text);
    await runAction("Copy OCR text", "/ocr/clipboard");
  }

  useEffect(() => {
    let mounted = true;
    const timer = window.setInterval(async () => {
      const result = await apiGet("/ocr/poll", 20000);
      if (!mounted) return;
      if (result.data?.status === "ready" && typeof result.data.text === "string") {
        setText(result.data.text);
        setMessage(result.message);
        addHistoryRecord({
          name: "Hotkey OCR result",
          type: "ocr",
          status: "success",
          route: "/ocr/poll",
          message: result.message
        });
      }
    }, 3000);
    return () => {
      mounted = false;
      window.clearInterval(timer);
    };
  }, []);

  const busy = Boolean(loading);

  return (
    <div className="grid gap-6 xl:grid-cols-[390px_1fr]">
      <Card title="OCR actions" description="Use backend native capture and file dialogs.">
        <div className="grid gap-3">
          <Button disabled={busy} onClick={() => runAction("Snip OCR", "/ocr/snip", { expectText: true })} icon={loading === "/ocr/snip" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Scissors className="h-4 w-4" />}>
            Snip OCR
          </Button>
          <Button disabled={busy} onClick={() => runAction("Screenshot OCR", "/ocr/screenshot", { expectText: true })} icon={loading === "/ocr/screenshot" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Camera className="h-4 w-4" />}>
            Screenshot OCR
          </Button>
          <Button disabled={busy} onClick={() => runAction("Image file OCR", "/ocr/file", { expectText: true })} icon={loading === "/ocr/file" ? <Loader2 className="h-4 w-4 animate-spin" /> : <FileImage className="h-4 w-4" />}>
            Select Image File
          </Button>
        </div>

        <div className="mt-5 grid grid-cols-2 gap-3">
          <Button disabled={busy || !text.trim()} onClick={() => runAction("Read OCR text", "/ocr/read")} icon={<Volume2 className="h-4 w-4" />}>
            Read
          </Button>
          <Button disabled={busy} onClick={() => runAction("Stop OCR reading", "/ocr/stop_read")} icon={<VolumeX className="h-4 w-4" />}>
            Stop
          </Button>
          <Button disabled={busy || !text.trim()} onClick={copyText} icon={<Copy className="h-4 w-4" />}>
            Copy
          </Button>
          <Button disabled={busy || !text.trim()} onClick={() => runAction("Save OCR TXT", "/ocr/save_txt")} icon={<Save className="h-4 w-4" />}>
            Save TXT
          </Button>
          <Button className="col-span-2" disabled={busy || !text.trim()} onClick={() => runAction("Save OCR PDF", "/ocr/save_pdf")} icon={<FileText className="h-4 w-4" />}>
            Save PDF
          </Button>
        </div>

        <div className="mt-5">
          <StatusBadge tone={busy ? "loading" : "info"} icon={busy ? "loading" : undefined}>
            {message}
          </StatusBadge>
        </div>
      </Card>

      <Card title="Extracted text" description="OCR output from screen capture, file selection, or hotkeys.">
        <textarea
          value={text}
          onChange={(event) => setText(event.target.value)}
          placeholder="OCR result will appear here..."
          className="min-h-[480px] w-full resize-none rounded-3xl border border-line bg-slate-50 p-4 text-sm leading-6 text-slate-800 outline-none focus:border-indigo-300 focus:bg-white"
        />
        <div className="mt-3 flex items-center gap-2 text-xs text-slate-500">
          <Clipboard className="h-4 w-4" />
          {text.trim() ? `${text.length} characters extracted` : "No OCR text yet"}
        </div>
      </Card>
    </div>
  );
}
