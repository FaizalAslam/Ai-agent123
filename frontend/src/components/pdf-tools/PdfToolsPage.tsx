"use client";

import { useState } from "react";
import { Wrench } from "lucide-react";
import { PdfEditor } from "@/components/pdf-tools/PdfEditor";
import { PdfMerge } from "@/components/pdf-tools/PdfMerge";
import { PdfSplit } from "@/components/pdf-tools/PdfSplit";
import { PageHeader } from "@/components/ui/PageHeader";
import { Tabs } from "@/components/ui/Tabs";
import type { ToastState } from "@/components/ui/Toast";

type PdfTab = "merge" | "split" | "edit";

interface PdfToolsPageProps {
  onToast: (toast: ToastState) => void;
}

export function PdfToolsPage({ onToast }: PdfToolsPageProps) {
  const [active, setActive] = useState<PdfTab>("merge");

  return (
    <div>
      <PageHeader
        title="PDF Tools"
        subtitle="Merge, split, and edit PDFs using the local backend and native desktop file dialogs."
        action={<div className="rounded-2xl border border-line bg-white p-3 text-indigo-600 shadow-sm"><Wrench className="h-5 w-5" /></div>}
      />
      <div className="mb-5">
        <Tabs
          active={active}
          onChange={setActive}
          tabs={[
            { id: "merge", label: "Merge" },
            { id: "split", label: "Split" },
            { id: "edit", label: "Edit" }
          ]}
        />
      </div>
      {active === "merge" && <PdfMerge onToast={onToast} />}
      {active === "split" && <PdfSplit onToast={onToast} />}
      {active === "edit" && <PdfEditor onToast={onToast} />}
    </div>
  );
}
