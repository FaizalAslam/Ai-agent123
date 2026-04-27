import type { ButtonHTMLAttributes, ReactNode } from "react";

type Variant = "primary" | "secondary" | "ghost" | "danger" | "success";

const variants: Record<Variant, string> = {
  primary: "bg-gradient-to-r from-indigo-600 to-blue-600 text-white shadow-card hover:from-indigo-700 hover:to-blue-700",
  secondary: "border border-line bg-white text-slate-700 shadow-sm hover:bg-slate-50",
  ghost: "text-slate-600 hover:bg-slate-100",
  danger: "border border-red-200 bg-red-50 text-red-700 hover:bg-red-100",
  success: "border border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100"
};

interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: Variant;
  icon?: ReactNode;
}

export function Button({
  variant = "secondary",
  icon,
  className = "",
  children,
  ...props
}: ButtonProps) {
  return (
    <button
      className={[
        "inline-flex min-h-10 items-center justify-center gap-2 rounded-2xl px-4 py-2 text-sm font-semibold transition disabled:cursor-not-allowed disabled:opacity-50",
        variants[variant],
        className
      ].join(" ")}
      {...props}
    >
      {icon}
      {children}
    </button>
  );
}
