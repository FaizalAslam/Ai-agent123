import type { HistoryRecord } from "@/types/history";

const STORAGE_KEY = "ai-agent-history-v1";

export function readHistory(): HistoryRecord[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function writeHistory(records: HistoryRecord[]) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(records.slice(0, 250)));
  window.dispatchEvent(new CustomEvent("ai-agent-history-changed"));
}

export function addHistoryRecord(record: Omit<HistoryRecord, "id" | "createdAt">) {
  const next: HistoryRecord = {
    ...record,
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    createdAt: new Date().toISOString()
  };
  writeHistory([next, ...readHistory()]);
  return next;
}

export function removeHistoryRecord(id: string) {
  writeHistory(readHistory().filter((item) => item.id !== id));
}

export function clearHistory() {
  writeHistory([]);
}
