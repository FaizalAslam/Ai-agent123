import { CheckCircle2, CircleAlert, Loader2, Wifi, WifiOff } from "lucide-react";

type Tone = "success" | "error" | "info" | "muted" | "loading";

const toneClass: Record<Tone, string> = {
  success: "border-emerald-200 bg-emerald-50 text-emerald-700",
  error: "border-red-200 bg-red-50 text-red-700",
  info: "border-indigo-200 bg-indigo-50 text-indigo-700",
  muted: "border-slate-200 bg-slate-50 text-slate-600",
  loading: "border-blue-200 bg-blue-50 text-blue-700"
};

interface StatusBadgeProps {
  tone?: Tone;
  children: React.ReactNode;
  icon?: "check" | "alert" | "wifi" | "offline" | "loading";
}

export function StatusBadge({ tone = "muted", children, icon }: StatusBadgeProps) {
  const Icon =
    icon === "check" ? CheckCircle2 :
    icon === "alert" ? CircleAlert :
    icon === "wifi" ? Wifi :
    icon === "offline" ? WifiOff :
    icon === "loading" ? Loader2 :
    null;

  return (
    <span className={`inline-flex items-center gap-1.5 rounded-full border px-3 py-1 text-xs font-semibold ${toneClass[tone]}`}>
      {Icon && <Icon className={`h-3.5 w-3.5 ${icon === "loading" ? "animate-spin" : ""}`} />}
      {children}
    </span>
  );
}
