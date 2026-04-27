export type HistoryType = "office" | "pdf" | "ocr" | "app" | "reader" | "editor" | "system";
export type HistoryStatus = "success" | "error";

export interface HistoryRecord {
  id: string;
  name: string;
  type: HistoryType;
  createdAt: string;
  status: HistoryStatus;
  route: string;
  message: string;
  filePath?: string;
}
