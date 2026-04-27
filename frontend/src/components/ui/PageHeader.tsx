interface PageHeaderProps {
  title: string;
  subtitle: string;
  action?: React.ReactNode;
}

export function PageHeader({ title, subtitle, action }: PageHeaderProps) {
  return (
    <div className="mb-6 flex flex-wrap items-start justify-between gap-4">
      <div>
        <p className="text-sm font-semibold uppercase tracking-[0.2em] text-indigo-600">AI Agent</p>
        <h1 className="mt-2 text-3xl font-bold tracking-tight text-slate-950">{title}</h1>
        <p className="mt-2 max-w-2xl text-sm text-slate-500">{subtitle}</p>
      </div>
      {action}
    </div>
  );
}
