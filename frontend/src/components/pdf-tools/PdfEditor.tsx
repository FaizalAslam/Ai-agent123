"use client";

import { useState } from "react";
import { ChevronLeft, ChevronRight, Edit3, FilePenLine, Loader2, Save, X } from "lucide-react";
import { apiPost } from "@/lib/api";
import { addHistoryRecord } from "@/lib/history";
import type { BackendResponse, PdfEdit, PdfPage, PdfTextBlock } from "@/types/api";
import type { ToastState } from "@/components/ui/Toast";
import { Button } from "@/components/ui/Button";
import { Card } from "@/components/ui/Card";
import { StatusBadge } from "@/components/ui/StatusBadge";

interface PdfEditorProps {
  onToast: (toast: ToastState) => void;
}

interface RenderState {
  image: string;
  width: number;
  height: number;
  zoom: number;
}

interface SelectedBlock {
  pageNum: number;
  block: PdfTextBlock;
  blockIdx: number;
}

function normalizeHex(value?: string) {
  if (!value) return "#000000";
  const hex = value.startsWith("#") ? value : `#${value}`;
  return /^#[0-9a-fA-F]{6}$/.test(hex) ? hex.toLowerCase() : "#000000";
}

export function PdfEditor({ onToast }: PdfEditorProps) {
  const [filePath, setFilePath] = useState<string | null>(null);
  const [pages, setPages] = useState<PdfPage[]>([]);
  const [currentPage, setCurrentPage] = useState(0);
  const [render, setRender] = useState<RenderState | null>(null);
  const [edits, setEdits] = useState<PdfEdit[]>([]);
  const [selected, setSelected] = useState<SelectedBlock | null>(null);
  const [newText, setNewText] = useState("");
  const [keepStyle, setKeepStyle] = useState(true);
  const [font, setFont] = useState("helv");
  const [fontSize, setFontSize] = useState(12);
  const [fontColor, setFontColor] = useState("#000000");
  const [loading, setLoading] = useState<string | null>(null);
  const [message, setMessage] = useState("Open a PDF to begin editing.");

  async function openPdf() {
    setLoading("/editor/open");
    try {
      const result = await apiPost<BackendResponse>("/editor/open", {}, 90000);
      if (result.ok && result.data) {
        const path = typeof result.data.file_path === "string" ? result.data.file_path : "";
        const nextPages = Array.isArray(result.data.pages) ? result.data.pages : [];
        setFilePath(path);
        setPages(nextPages);
        setCurrentPage(0);
        setEdits([]);
        setSelected(null);
        setMessage(result.message || `Loaded ${nextPages.length} pages.`);
        addHistoryRecord({
          name: "Open PDF editor",
          type: "editor",
          status: "success",
          route: "/editor/open",
          message: result.message,
          filePath: path
        });
        await renderPage(path, 0, nextPages);
      } else {
        setMessage(result.message);
        addHistoryRecord({
          name: "Open PDF editor",
          type: "editor",
          status: "error",
          route: "/editor/open",
          message: result.message
        });
      }
      onToast({ tone: result.ok ? "success" : "error", message: result.message });
    } finally {
      setLoading(null);
    }
  }

  async function renderPage(path = filePath, pageNum = currentPage, knownPages = pages) {
    if (!path) return;
    setLoading("/editor/render-page");
    try {
      const result = await apiPost<BackendResponse>("/editor/render-page", { file_path: path, page_num: pageNum }, 90000);
      if (result.ok && result.data?.image && typeof result.data.image === "string") {
        setRender({
          image: result.data.image,
          width: Number(result.data.width || 0),
          height: Number(result.data.height || 0),
          zoom: Number(result.data.zoom || 1)
        });
        setCurrentPage(pageNum);
        setPages(knownPages);
        setMessage(`Page ${pageNum + 1} ready. Click highlighted text to edit.`);
      } else {
        setMessage(result.message);
        onToast({ tone: "error", message: result.message });
      }
    } finally {
      setLoading(null);
    }
  }

  function openEditModal(block: PdfTextBlock, blockIdx: number) {
    setSelected({ pageNum: currentPage, block, blockIdx });
    setNewText(block.text || "");
    setKeepStyle(true);
    setFont(block.font || "helv");
    setFontSize(Number(block.size || 12));
    setFontColor(normalizeHex(block.color));
  }

  function queueEdit() {
    if (!selected || !newText.trim()) return;
    const block = selected.block;
    const blockId = block.id ?? `${selected.pageNum}-${selected.blockIdx}`;
    const edit: PdfEdit = {
      page: selected.pageNum,
      block_id: blockId,
      original_text: block.text,
      new_text: newText.trim(),
      bbox: { x: block.x, y: block.y, x1: block.x1, y1: block.y1 },
      keep_style: keepStyle,
      style: {
        font,
        size: fontSize,
        color: normalizeHex(fontColor),
        flags: Number(block.flags || 0)
      }
    };
    setEdits((prev) => [
      ...prev.filter((item) => !(item.page === selected.pageNum && item.block_id === blockId)),
      edit
    ]);
    setSelected(null);
    setMessage("Edit queued. Save PDF to apply changes.");
  }

  async function saveEdits() {
    if (!filePath || edits.length === 0) return;
    setLoading("/editor/save");
    try {
      const result = await apiPost<BackendResponse>("/editor/save", { file_path: filePath, edits }, 120000);
      if (result.ok) {
        setEdits([]);
        await renderPage(filePath, currentPage);
      }
      setMessage(result.message);
      addHistoryRecord({
        name: "Save PDF edits",
        type: "editor",
        status: result.ok ? "success" : "error",
        route: "/editor/save",
        message: result.message,
        filePath
      });
      onToast({ tone: result.ok ? "success" : "error", message: result.message });
    } finally {
      setLoading(null);
    }
  }

  const currentBlocks = pages[currentPage]?.text_blocks || [];
  const totalPages = pages.length;
  const busy = Boolean(loading);

  return (
    <div className="grid gap-6 xl:grid-cols-[1fr_360px]">
      <Card className="min-h-[620px] p-0">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-line px-5 py-4">
          <div>
            <h2 className="font-bold text-slate-950">PDF editor</h2>
            <p className="text-sm text-slate-500">{filePath || "No PDF loaded"}</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <Button disabled={busy} onClick={openPdf} icon={loading === "/editor/open" ? <Loader2 className="h-4 w-4 animate-spin" /> : <FilePenLine className="h-4 w-4" />}>
              Open PDF
            </Button>
            <Button disabled={busy || !filePath || currentPage === 0} onClick={() => renderPage(filePath, currentPage - 1)} icon={<ChevronLeft className="h-4 w-4" />}>
              Prev
            </Button>
            <Button disabled={busy || !filePath || currentPage >= totalPages - 1} onClick={() => renderPage(filePath, currentPage + 1)} icon={<ChevronRight className="h-4 w-4" />}>
              Next
            </Button>
            <Button variant="primary" disabled={busy || edits.length === 0} onClick={saveEdits} icon={loading === "/editor/save" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Save className="h-4 w-4" />}>
              Save PDF
            </Button>
          </div>
        </div>

        <div className="scrollbar-soft overflow-auto bg-slate-100 p-6">
          {!render ? (
            <div className="flex min-h-[520px] items-center justify-center rounded-3xl border border-dashed border-line bg-white text-center">
              <div>
                <FilePenLine className="mx-auto h-10 w-10 text-indigo-500" />
                <h3 className="mt-4 text-lg font-bold text-slate-950">Open a PDF to edit text blocks</h3>
                <p className="mt-2 text-sm text-slate-500">The backend renders page images and provides clickable overlays.</p>
              </div>
            </div>
          ) : (
            <div
              className="relative mx-auto inline-block bg-white shadow-soft"
              style={{ width: render.width || "auto", height: render.height || "auto" }}
            >
              <img src={render.image} alt={`PDF page ${currentPage + 1}`} className="block max-w-none" />
              <div className="absolute inset-0">
                {currentBlocks.map((block, index) => {
                  const zoom = render.zoom || 1;
                  return (
                    <button
                      key={`${block.id ?? index}-${currentPage}`}
                      title={block.text || "Text block"}
                      onClick={() => openEditModal(block, index)}
                      className="absolute rounded border border-amber-400/70 bg-amber-300/20 transition hover:bg-indigo-400/25"
                      style={{
                        left: block.x * zoom,
                        top: block.y * zoom,
                        width: Math.max(4, (block.x1 - block.x) * zoom),
                        height: Math.max(4, (block.y1 - block.y) * zoom)
                      }}
                    />
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </Card>

      <div className="space-y-5">
        <Card title="Editor status" description="Page and pending edit state.">
          <div className="space-y-3">
            <StatusBadge tone={busy ? "loading" : filePath ? "success" : "muted"} icon={busy ? "loading" : filePath ? "check" : undefined}>
              {message}
            </StatusBadge>
            <div className="grid grid-cols-2 gap-3">
              <div className="rounded-2xl border border-line bg-slate-50 p-4">
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Page</div>
                <div className="mt-2 text-2xl font-bold text-slate-950">{totalPages ? currentPage + 1 : "-"}</div>
                <div className="text-xs text-slate-500">of {totalPages || "-"}</div>
              </div>
              <div className="rounded-2xl border border-line bg-slate-50 p-4">
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Queued</div>
                <div className="mt-2 text-2xl font-bold text-slate-950">{edits.length}</div>
                <div className="text-xs text-slate-500">edits</div>
              </div>
            </div>
          </div>
        </Card>

        <Card title="Queued edits" description="Changes waiting to be saved.">
          <div className="max-h-80 space-y-2 overflow-auto pr-1">
            {edits.length === 0 ? (
              <p className="text-sm text-slate-500">No edits queued.</p>
            ) : (
              edits.map((edit, index) => (
                <div key={`${edit.page}-${edit.block_id}-${index}`} className="rounded-2xl border border-line bg-slate-50 p-3 text-sm">
                  <div className="font-semibold text-slate-950">Page {edit.page + 1}</div>
                  <div className="mt-1 line-clamp-3 text-slate-600">{edit.new_text}</div>
                </div>
              ))
            )}
          </div>
        </Card>
      </div>

      {selected && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/30 p-4 backdrop-blur-sm">
          <div className="w-full max-w-xl rounded-3xl border border-line bg-white p-5 shadow-soft">
            <div className="flex items-start justify-between gap-4">
              <div>
                <h3 className="text-lg font-bold text-slate-950">Edit Text Block</h3>
                <p className="text-sm text-slate-500">Page {selected.pageNum + 1}</p>
              </div>
              <button onClick={() => setSelected(null)} className="rounded-full p-2 text-slate-500 hover:bg-slate-100" aria-label="Close editor modal">
                <X className="h-5 w-5" />
              </button>
            </div>

            <label className="mt-5 block text-sm font-semibold text-slate-700">Original text</label>
            <input value={selected.block.text || ""} readOnly className="mt-2 w-full rounded-2xl border border-line bg-slate-50 px-4 py-3 text-sm text-slate-500" />

            <label className="mt-4 block text-sm font-semibold text-slate-700">New text</label>
            <textarea value={newText} onChange={(event) => setNewText(event.target.value)} rows={4} className="mt-2 w-full resize-none rounded-2xl border border-line bg-white px-4 py-3 text-sm outline-none focus:border-indigo-300" />

            <label className="mt-4 flex items-center gap-2 text-sm font-semibold text-slate-700">
              <input type="checkbox" checked={keepStyle} onChange={(event) => setKeepStyle(event.target.checked)} className="h-4 w-4 accent-indigo-600" />
              Keep original font style
            </label>

            {!keepStyle && (
              <div className="mt-4 grid gap-3 sm:grid-cols-3">
                <input value={font} onChange={(event) => setFont(event.target.value)} className="rounded-2xl border border-line px-3 py-2 text-sm" placeholder="Font" />
                <input type="number" value={fontSize} onChange={(event) => setFontSize(Number(event.target.value))} className="rounded-2xl border border-line px-3 py-2 text-sm" />
                <input type="color" value={fontColor} onChange={(event) => setFontColor(event.target.value)} className="h-10 rounded-2xl border border-line px-2 py-1" />
              </div>
            )}

            <div className="mt-5 flex justify-end gap-3">
              <Button onClick={() => setSelected(null)}>Cancel</Button>
              <Button variant="primary" onClick={queueEdit} icon={<Edit3 className="h-4 w-4" />}>Queue Edit</Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
