import json
import logging
import re
import time
from dataclasses import asdict, dataclass, field

try:
    from openai import (
        APIConnectionError,
        APIStatusError,
        APITimeoutError,
        OpenAI,
        RateLimitError,
    )
except Exception:  # pragma: no cover - handled as OPENAI_DISABLED at runtime
    APIConnectionError = APIStatusError = APITimeoutError = RateLimitError = Exception
    OpenAI = None

from config import OPENAI_API_KEY, OPENAI_MODEL
from utils.office_action_registry import registry_as_prompt_lines
from utils.office_actions import OfficeActionError, normalize_actions, validate_actions


logger = logging.getLogger("OfficeAgent")

MAX_OPENAI_COMMAND_LENGTH = 4000
OPENAI_TIMEOUT_SECONDS = 20
OPENAI_MAX_TOKENS = 900
OPENAI_RETRY_DELAYS = (0.5, 1.5)

SYSTEM_PROMPT = """
You are an Office Automation Assistant for a local desktop automation app.
Convert user commands into executable Office JSON actions.

Rules:
- Return ONLY a valid JSON array.
- Do not wrap the JSON in markdown.
- Do not include explanations.
- For one step, return an array containing one object.
- Use only supported actions for the requested app.
- PowerPoint slide_index values are user-facing 1-based slide numbers.
- Excel range or cell fields must be explicit for formatting/cell actions.
- Word style targets should be conservative: heading, heading_1, body, selection, or all.
"""


@dataclass
class OpenAIParseResult:
    success: bool
    actions: list[dict] = field(default_factory=list)
    error_code: str | None = None
    message: str = ""
    retryable: bool = False
    raw_response_preview: str | None = None
    model: str = ""
    duration_ms: int = 0
    usage: dict | None = None
    warnings: list[str] = field(default_factory=list)

    def to_dict(self):
        return asdict(self)


class OpenAIHandler:
    def __init__(self, api_key=None, model=None, timeout=OPENAI_TIMEOUT_SECONDS):
        self.api_key = api_key if api_key is not None else OPENAI_API_KEY
        self.model = model or OPENAI_MODEL
        self.timeout = timeout
        self.last_error_code = ""
        self.last_error = ""
        self._client = None

    def _get_client(self):
        if OpenAI is None:
            raise RuntimeError("OpenAI SDK is not installed.")
        if not self._client:
            self._client = OpenAI(api_key=self.api_key, timeout=self.timeout)
        return self._client

    def _error_result(self, code, message, retryable=False, start=None, preview=None, usage=None, warnings=None):
        self.last_error_code = code
        self.last_error = message
        duration_ms = int((time.perf_counter() - start) * 1000) if start else 0
        return OpenAIParseResult(
            success=False,
            error_code=code,
            message=message,
            retryable=retryable,
            raw_response_preview=preview,
            model=self.model,
            duration_ms=duration_ms,
            usage=usage,
            warnings=warnings or [],
        )

    def _parse_json(self, text):
        warnings = []
        clean = (text or "").strip()
        clean = re.sub(r"^```(?:json)?\s*", "", clean, flags=re.IGNORECASE)
        clean = re.sub(r"\s*```$", "", clean).strip()

        try:
            parsed = json.loads(clean)
            return parsed, warnings
        except json.JSONDecodeError:
            pass

        # Legacy compatibility only: accept a single complete JSON array/object
        # embedded in surrounding text. Partial object salvage is intentionally
        # avoided because it can silently drop actions.
        match = re.search(r"(\[[\s\S]*\]|\{[\s\S]*\})", clean)
        if match:
            try:
                warnings.append("OpenAI response contained extra text around JSON.")
                return json.loads(match.group(1)), warnings
            except json.JSONDecodeError:
                pass

        raise OfficeActionError(
            "OPENAI_INVALID_JSON",
            "OpenAI returned invalid JSON.",
            clean[:300],
        )

    def _usage_dict(self, usage):
        if usage is None:
            return None
        if hasattr(usage, "model_dump"):
            return usage.model_dump()
        if isinstance(usage, dict):
            return usage
        return {
            key: getattr(usage, key)
            for key in ("prompt_tokens", "completion_tokens", "total_tokens")
            if hasattr(usage, key)
        }

    def _messages(self, app_name, command):
        app = (app_name or "").strip().lower()
        registry_lines = "\n".join(registry_as_prompt_lines(app)) or "- Use supported actions only."
        return [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    f"App: {app}\n"
                    f"Supported actions:\n{registry_lines}\n\n"
                    f"Command:\n{command}"
                ),
            },
        ]

    def interpret_result(self, app_name, command):
        start = time.perf_counter()
        self.last_error_code = ""
        self.last_error = ""

        if OpenAI is None:
            return self._error_result(
                "OPENAI_DISABLED",
                "OpenAI SDK is not installed.",
                start=start,
            )
        if not self.api_key:
            logger.info("OpenAI fallback skipped: API key missing.")
            return self._error_result(
                "OPENAI_API_KEY_MISSING",
                "OpenAI API key is missing; deterministic parser could not handle this command.",
                start=start,
            )
        if len(command or "") > MAX_OPENAI_COMMAND_LENGTH:
            return self._error_result(
                "COMMAND_TOO_LONG",
                f"Command is too long for OpenAI fallback ({MAX_OPENAI_COMMAND_LENGTH} characters max).",
                start=start,
            )

        last_retryable = None
        for attempt in range(len(OPENAI_RETRY_DELAYS) + 1):
            try:
                response = self._get_client().chat.completions.create(
                    model=self.model,
                    messages=self._messages(app_name, command),
                    temperature=0,
                    max_tokens=OPENAI_MAX_TOKENS,
                )
                content = (response.choices[0].message.content or "").strip()
                preview = content[:300]
                parsed, warnings = self._parse_json(content)
                try:
                    actions = normalize_actions(parsed)
                    actions = validate_actions(app_name, actions)
                except OfficeActionError as exc:
                    code = "OPENAI_UNSUPPORTED_ACTION" if exc.error_code == "UNSUPPORTED_ACTION" else "OPENAI_INVALID_ACTION_SCHEMA"
                    logger.warning("OpenAI action validation failed: %s", exc.message)
                    return self._error_result(
                        code,
                        exc.message,
                        start=start,
                        preview=preview,
                        usage=self._usage_dict(getattr(response, "usage", None)),
                        warnings=warnings,
                    )

                duration_ms = int((time.perf_counter() - start) * 1000)
                self.last_error_code = ""
                self.last_error = ""
                logger.info(
                    "OpenAI fallback result: model=%s duration_ms=%s success=True actions=%s warnings=%s",
                    self.model,
                    duration_ms,
                    len(actions),
                    len(warnings),
                )
                return OpenAIParseResult(
                    success=True,
                    actions=actions,
                    message="OpenAI parsed Office actions.",
                    retryable=False,
                    raw_response_preview=preview,
                    model=self.model,
                    duration_ms=duration_ms,
                    usage=self._usage_dict(getattr(response, "usage", None)),
                    warnings=warnings,
                )

            except OfficeActionError as exc:
                logger.warning("OpenAI JSON parse failed: %s", exc.message)
                return self._error_result(
                    exc.error_code,
                    exc.message,
                    start=start,
                    preview=exc.details,
                )
            except APITimeoutError as exc:
                last_retryable = ("OPENAI_TIMEOUT", "OpenAI fallback timed out.", exc)
            except RateLimitError as exc:
                last_retryable = ("OPENAI_RATE_LIMITED", "OpenAI fallback was rate limited.", exc)
            except APIConnectionError as exc:
                last_retryable = ("OPENAI_NETWORK_ERROR", "OpenAI fallback could not reach the API.", exc)
            except APIStatusError as exc:
                status_code = getattr(exc, "status_code", 0)
                if 500 <= int(status_code or 0) < 600:
                    last_retryable = ("OPENAI_SERVER_ERROR", "OpenAI fallback returned a server error.", exc)
                else:
                    logger.warning("OpenAI non-retryable status error: %s", exc)
                    return self._error_result(
                        "OPENAI_REQUEST_FAILED",
                        "OpenAI fallback request failed.",
                        start=start,
                        preview=str(exc)[:300],
                    )
            except Exception as exc:
                logger.warning("OpenAI fallback failed: %s", exc)
                return self._error_result(
                    "OPENAI_NETWORK_ERROR",
                    "OpenAI fallback failed before returning a parseable response.",
                    retryable=True,
                    start=start,
                    preview=str(exc)[:300],
                )

            if attempt < len(OPENAI_RETRY_DELAYS):
                time.sleep(OPENAI_RETRY_DELAYS[attempt])

        code, message, exc = last_retryable or ("OPENAI_NETWORK_ERROR", "OpenAI fallback failed.", None)
        logger.warning("OpenAI fallback retryable failure: code=%s model=%s", code, self.model)
        return self._error_result(
            code,
            message,
            retryable=True,
            start=start,
            preview=str(exc)[:300] if exc else None,
        )

    def interpret(self, app_name, command):
        result = self.interpret_result(app_name, command)
        return result.actions if result.success else None
