import { X } from "lucide-react";

export interface ToastState {
  tone: "success" | "error" | "info";
  message: string;
}

interface ToastProps {
  toast: ToastState | null;
  onClose: () => void;
}

const toneClasses = {
  success: "border-emerald-200 bg-emerald-50 text-emerald-900",
  error: "border-red-200 bg-red-50 text-red-900",
  info: "border-indigo-200 bg-indigo-50 text-indigo-900"
};

export function Toast({ toast, onClose }: ToastProps) {
  if (!toast) return null;

  return (
    <div className={`fixed bottom-6 right-6 z-50 flex max-w-md items-center gap-3 rounded-2xl border px-4 py-3 text-sm shadow-soft ${toneClasses[toast.tone]}`}>
      <span>{toast.message}</span>
      <button onClick={onClose} className="rounded-full p-1 hover:bg-white/70" aria-label="Dismiss notification">
        <X className="h-4 w-4" />
      </button>
    </div>
  );
}
