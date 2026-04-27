export interface TabItem<T extends string> {
  id: T;
  label: string;
}

interface TabsProps<T extends string> {
  tabs: TabItem<T>[];
  active: T;
  onChange: (tab: T) => void;
}

export function Tabs<T extends string>({ tabs, active, onChange }: TabsProps<T>) {
  return (
    <div className="inline-flex rounded-2xl border border-line bg-white p-1 shadow-sm">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          onClick={() => onChange(tab.id)}
          className={[
            "rounded-xl px-4 py-2 text-sm font-semibold transition",
            active === tab.id
              ? "bg-indigo-600 text-white shadow-sm"
              : "text-slate-600 hover:bg-slate-50 hover:text-slate-950"
          ].join(" ")}
        >
          {tab.label}
        </button>
      ))}
    </div>
  );
}
