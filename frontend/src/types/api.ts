export type BackendStatus = "success" | "fail" | "waiting" | string;

export interface BackendResponse {
  success?: boolean;
  status?: BackendStatus;
  intent?: string;
  message?: string;
  error?: string;
  error_code?: string;
  details?: string;
  app_type?: string;
  action_type?: string;
  file_path?: string;
  output_file?: string;
  persisted?: boolean;
  opened?: boolean;
  text?: string;
  filePath?: string;
  pages?: PdfPage[];
  total_pages?: number;
  current_page?: number;
  totalPages?: number;
  is_reading?: boolean;
  is_paused?: boolean;
  speed?: number;
  available?: boolean;
  enabled?: boolean;
  armed?: boolean;
  heard?: string;
  image?: string;
  width?: number;
  height?: number;
  zoom?: number;
  [key: string]: unknown;
}

export interface ApiResult<T = BackendResponse> {
  ok: boolean;
  reachable: boolean;
  status?: string;
  message: string;
  errorCode?: string;
  data?: T;
}

export interface PdfTextBlock {
  id?: string | number;
  text?: string;
  x: number;
  y: number;
  x1: number;
  y1: number;
  font?: string;
  size?: number;
  color?: string;
  flags?: number;
}

export interface PdfPage {
  page?: number;
  page_num?: number;
  text_blocks?: PdfTextBlock[];
}

export interface PdfEdit {
  page: number;
  block_id?: string | number;
  original_text?: string;
  new_text: string;
  bbox: {
    x: number;
    y: number;
    x1: number;
    y1: number;
  };
  keep_style: boolean;
  style: {
    font: string;
    size: number;
    color: string;
    flags: number;
  };
}
