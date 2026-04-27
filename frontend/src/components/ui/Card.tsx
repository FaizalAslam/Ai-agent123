import type { HTMLAttributes, ReactNode } from "react";

interface CardProps extends HTMLAttributes<HTMLDivElement> {
  title?: string;
  description?: string;
  action?: ReactNode;
}

export function Card({
  title,
  description,
  action,
  className = "",
  children,
  ...props
}: CardProps) {
  return (
    <section
      className={["rounded-3xl border border-line bg-white p-5 shadow-card", className].join(" ")}
      {...props}
    >
      {(title || description || action) && (
        <div className="mb-4 flex items-start justify-between gap-4">
          <div>
            {title && <h2 className="text-base font-bold text-slate-950">{title}</h2>}
            {description && <p className="mt-1 text-sm text-slate-500">{description}</p>}
          </div>
          {action}
        </div>
      )}
      {children}
    </section>
  );
}
