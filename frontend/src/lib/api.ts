import type { ApiResult, BackendResponse } from "@/types/api";

const API_PREFIX = "/api/backend";
const DEFAULT_TIMEOUT_MS = 60000;

function normalizeMessage(payload: BackendResponse | undefined, fallback: string) {
  if (!payload) return fallback;
  const text = payload.message || payload.error || payload.details;
  return typeof text === "string" && text.trim() ? text : fallback;
}

function normalizeResponse<T extends BackendResponse>(payload: T, reachable = true): ApiResult<T> {
  const ok = payload.success === true || payload.status === "success";
  const status = typeof payload.status === "string" ? payload.status : undefined;
  return {
    ok,
    reachable,
    status,
    message: normalizeMessage(payload, ok ? "Done." : "Request failed."),
    errorCode: typeof payload.error_code === "string" ? payload.error_code : undefined,
    data: payload
  };
}

async function requestJson<T extends BackendResponse>(
  method: "GET" | "POST",
  path: string,
  body?: unknown,
  timeoutMs = DEFAULT_TIMEOUT_MS
): Promise<ApiResult<T>> {
  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), timeoutMs);
  const route = path.startsWith("/") ? path : `/${path}`;

  try {
    const response = await fetch(`${API_PREFIX}${route}`, {
      method,
      headers: method === "POST" ? { "Content-Type": "application/json" } : undefined,
      body: method === "POST" ? JSON.stringify(body ?? {}) : undefined,
      signal: controller.signal
    });

    const text = await response.text();
    let parsed: T;
    try {
      parsed = text ? (JSON.parse(text) as T) : ({ status: response.ok ? "success" : "fail" } as T);
    } catch {
      return {
        ok: false,
        reachable: true,
        status: "fail",
        message: `Unexpected backend response (${response.status}).`,
        errorCode: "MALFORMED_JSON"
      };
    }

    const result = normalizeResponse(parsed, true);
    if (!response.ok) {
      return {
        ...result,
        ok: false,
        message: result.message || `Backend request failed (${response.status}).`,
        errorCode: result.errorCode || `HTTP_${response.status}`
      };
    }
    return result;
  } catch (error) {
    const aborted = error instanceof DOMException && error.name === "AbortError";
    return {
      ok: false,
      reachable: false,
      status: "fail",
      message: aborted
        ? "Backend request timed out."
        : "Backend is not running at 127.0.0.1:5000.",
      errorCode: aborted ? "REQUEST_TIMEOUT" : "BACKEND_UNAVAILABLE"
    };
  } finally {
    window.clearTimeout(timer);
  }
}

export function apiGet<T extends BackendResponse = BackendResponse>(path: string, timeoutMs?: number) {
  return requestJson<T>("GET", path, undefined, timeoutMs);
}

export function apiPost<T extends BackendResponse = BackendResponse>(path: string, body?: unknown, timeoutMs?: number) {
  return requestJson<T>("POST", path, body, timeoutMs);
}

export const backendDisplayUrl =
  process.env.NEXT_PUBLIC_BACKEND_URL || "http://127.0.0.1:5000";
