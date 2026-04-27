"use client";

import { useState } from "react";
import { FileText } from "lucide-react";
import { OcrReader } from "@/components/documents/OcrReader";
import { PdfReader } from "@/components/documents/PdfReader";
import { PageHeader } from "@/components/ui/PageHeader";
import { Tabs } from "@/components/ui/Tabs";
import type { ToastState } from "@/components/ui/Toast";

type DocumentTab = "ocr" | "reader";

interface DocumentsPageProps {
  onToast: (toast: ToastState) => void;
}

export function DocumentsPage({ onToast }: DocumentsPageProps) {
  const [active, setActive] = useState<DocumentTab>("ocr");

  return (
    <div>
      <PageHeader
        title="Documents"
        subtitle="OCR Reader and PDF Reader tools backed by the local Flask automation service."
        action={<div className="rounded-2xl border border-line bg-white p-3 text-indigo-600 shadow-sm"><FileText className="h-5 w-5" /></div>}
      />
      <div className="mb-5">
        <Tabs
          active={active}
          onChange={setActive}
          tabs={[
            { id: "ocr", label: "OCR Reader" },
            { id: "reader", label: "PDF Reader" }
          ]}
        />
      </div>
      {active === "ocr" ? <OcrReader onToast={onToast} /> : <PdfReader onToast={onToast} />}
    </div>
  );
}
