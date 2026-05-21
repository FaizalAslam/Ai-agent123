import logging
import os
import re
import traceback

from flask import Blueprint, request
from office.constants import OFFICE_APPS
from office.intent import (
    _canonical_office_app,
    _detect_action_type,
    _detect_office_intent,
    _extract_office_agent_command,
)
from office.file_resolver import (
    _ensure_fresh_file_action,
    _expand_powerpoint_slide_count,
    resolve_office_file_path,
)
from office.resolver import _resolve_actions
from office.runner import _known_office_actions, _run_office_actions
from parser.command_complexity import classify_office_command_complexity
from utils.office_actions import OfficeActionError, validate_actions
from utils.api_response import _json_success, _json_error
from utils import command_map

office_bp = Blueprint("office", __name__)

_FILE_ACCESS_RE = re.compile(
    r'\b(?:open|launch|access|create)\b[^.!]{0,120}\b(?:file|workbook|spreadsheet)\b',
    re.IGNORECASE,
)
_SEGMENT_SPLIT_RE = re.compile(
    r'(?:[.!]\s+(?:after\s+that[,.]?\s+|next[,.]?\s+|finally[,.]?\s+|then\s+)|,\s+then\s+)'
    r'(?=(?:open|launch|access|create)\b)',
    re.IGNORECASE,
)


def _is_multi_file_command(text):
    return len(_FILE_ACCESS_RE.findall(text)) >= 2


def _split_file_segments(command):
    parts = [p.strip() for p in _SEGMENT_SPLIT_RE.split(command) if p.strip()]
    return parts if len(parts) > 1 else []


def _office_execute_multi(app_name, base_data, segments):
    successes, failures = [], []
    for seg in segments:
        seg_data = {**base_data, "app": app_name, "raw": seg}
        seg_data.pop("file_path", None)
        seg_data.pop("file", None)
        resp = _office_execute_impl(seg_data)
        r = resp[0] if isinstance(resp, tuple) else resp
        body = r.get_json(silent=True) or {}
        if body.get("status") == "success":
            successes.append({"command": seg[:80], "file": body.get("output_file", "")})
        else:
            failures.append({"command": seg[:80], "error": body.get("message", "failed")})

    ok, total = len(successes), len(segments)
    if ok == 0:
        return _json_error(
            f"All {total} file operations failed.",
            intent="office_automation",
            error_code="OFFICE_MULTI_ALL_FAILED",
            app_type=app_name,
            action_type="multi_file",
            results=successes,
            failed=failures,
        )
    return _json_success(
        f"Completed {ok}/{total} file operations successfully.",
        intent="office_automation",
        app_type=app_name,
        action_type="multi_file",
        results=successes,
        failed=failures,
        executed_count=ok,
        failed_count=len(failures),
    )


def _office_execute_impl(data):
    app_name = (data.get("app") or "").lower().strip()
    command  = (data.get("raw") or "").strip()
    full     = (data.get("command") or "").strip()

    if not command and full:
        parsed_app, parsed_command = _extract_office_agent_command(full)
        if parsed_app and not app_name:
            app_name = (parsed_app or "").strip()
        if parsed_command:
            command = (parsed_command or "").strip()
        elif app_name:
            command = full

    if not app_name:
        office_route = _detect_office_intent(full or command)
        if office_route.get("is_office"):
            app_name = office_route.get("app_type", "")
            command  = office_route.get("command") or command or full

    app_name = _canonical_office_app(app_name)

    if app_name not in OFFICE_APPS or not command:
        return _json_error(
            "Missing or invalid Office app/command.",
            intent="office_automation",
            error_code="INVALID_OFFICE_REQUEST",
            app_type=app_name or "unknown",
            action_type="unknown",
        )

    if _is_multi_file_command(command):
        segments = _split_file_segments(command)
        if segments:
            logging.info("Multi-file command detected; splitting into %d sub-commands.", len(segments))
            return _office_execute_multi(app_name, data, segments)

    logging.info("Office request: original=%r app=%s", full or command, app_name)
    cmd_complexity = classify_office_command_complexity(command)

    cache_key, actions, source, action_error, plan_info = _resolve_actions(app_name, command)
    diag = (plan_info or {}).get("diag") or {}
    _d = {
        "parser_used":             source,
        "complexity":              cmd_complexity,
        "openai_attempted":        diag.get("openai_attempted", False),
        "openai_success":          diag.get("openai_success", False),
        "openai_error_code":       diag.get("openai_error_code"),
        "fallback_reason":         diag.get("fallback_reason", ""),
        "raw_action_count":        diag.get("raw_action_count", 0),
        "normalized_action_count": diag.get("normalized_action_count", 0),
        "validation_errors":       list(diag.get("validation_errors") or []),
    }
    final_actions:    list = []
    executor_results: list = []
    logging.info(
        "OFFICE_ACTION_PLAN parser_used=%s complexity=%s action_count=%d",
        source, cmd_complexity, len(actions),
    )

    if action_error:
        logging.warning("Office action parse error: %s", action_error.message)
        return _json_error(
            action_error.message,
            intent="office_automation",
            error_code=action_error.error_code,
            app_type=app_name,
            action_type=_detect_action_type(command, actions),
            details=action_error.details,
            source=source,
            plan=plan_info,
            **_d,
            final_actions=[],
            executor_results=[],
        )
    if not actions:
        return _json_error(
            "No matching Office command found. Try a more specific action like 'create a new workbook' or 'add heading Introduction'.",
            intent="office_automation",
            error_code="NO_OFFICE_ACTION_MATCH",
            app_type=app_name,
            action_type="unknown",
            source=source,
            plan=plan_info,
            **_d,
            final_actions=[],
            executor_results=[],
        )

    requested_file_path = (data.get("file_path") or data.get("file") or "").strip()

    if not requested_file_path and plan_info and plan_info.get("output_filename"):
        ai_fname = os.path.basename(str(plan_info["output_filename"])).strip()
        if ai_fname:
            requested_file_path = ai_fname
            logging.info("Using OpenAI-suggested output filename: %r for [%s]", ai_fname, app_name)

    actions = _ensure_fresh_file_action(app_name, command, actions, requested_file_path)
    actions = _expand_powerpoint_slide_count(app_name, command, actions)
    try:
        actions = validate_actions(app_name, actions, known_actions=_known_office_actions(app_name))
    except OfficeActionError as exc:
        _d["validation_errors"].append(exc.message)
        logging.warning("OFFICE_VALIDATION_FAILED errors=%s", _d["validation_errors"])
        return _json_error(
            exc.message,
            intent="office_automation",
            error_code=exc.error_code,
            app_type=app_name,
            action_type=_detect_action_type(command, actions),
            details=exc.details,
            source=source,
            plan=plan_info,
            **_d,
            final_actions=[],
            executor_results=[],
        )
    final_actions = list(actions)

    resolution = resolve_office_file_path(data, actions, app_name, mode=_detect_action_type(command, actions))
    if not resolution.get("success"):
        logging.warning("Office path resolution failed: %s", resolution.get("message"))
        return _json_error(
            resolution.get("message", "Could not resolve Office file path."),
            intent="office_automation",
            error_code=resolution.get("error_code", "INVALID_FILE_PATH"),
            app_type=app_name,
            action_type=resolution.get("action_type", "unknown"),
            details=resolution.get("details", ""),
            source=source,
            plan=plan_info,
            **_d,
            final_actions=final_actions,
            executor_results=[],
        )

    logging.info(
        "Routing decision: original=%r intent=office_automation app=%s action_type=%s handler=office_executor reason=%s output=%s action_count=%s",
        command, app_name, resolution.get("action_type"), resolution.get("reason"),
        resolution.get("file_path"), len(actions),
    )

    summary = _run_office_actions(
        app_name, actions,
        file_path=resolution["file_path"],
        source_path=resolution.get("source_path"),
        command_text=command,
    )
    executor_results = summary.get("results", [])
    logging.info(
        "OFFICE_EXECUTION_RESULTS results=%s",
        [(r.get("action"), r.get("status")) for r in executor_results],
    )

    if summary.get("dependency_error"):
        return _json_error(
            summary["dependency_error"],
            intent="office_automation",
            error_code="OFFICE_DEPENDENCY_MISSING",
            app_type=app_name,
            action_type=resolution.get("action_type", "unknown"),
            source=source,
            plan=plan_info,
            file_path=summary["output_path"],
            output_file=summary["output_path"],
            **_d,
            final_actions=final_actions,
            executor_results=executor_results,
        )

    if summary["failures"] and cache_key:
        command_map.remove_action(app_name, cache_key)
    if summary["failures"]:
        return _json_error(
            "Some Office actions completed, but some failed." if summary.get("ok_count") else f"Could not save {app_name.title()} file.",
            intent="office_automation",
            error_code=summary.get("error_code") or "OFFICE_SAVE_FAILED",
            status="partial_success" if summary.get("ok_count") else "fail",
            app_type=app_name,
            action_type=resolution.get("action_type", "unknown"),
            details=f"{summary['ok_count']}/{summary['total']} done | {' | '.join(summary['failures'])}",
            source=source,
            plan=plan_info,
            file_path=summary["output_path"],
            output_file=summary["output_path"],
            persisted=summary.get("persisted", False),
            results=summary.get("results", []),
            executed_count=summary.get("ok_count", 0),
            failed_count=max(0, summary.get("total", 0) - summary.get("ok_count", 0)),
            **_d,
            final_actions=final_actions,
            executor_results=executor_results,
        )

    if plan_info and plan_info.get("errors"):
        return _json_error(
            "Some Office actions completed, but some command clauses could not be parsed.",
            intent="office_automation",
            error_code="OFFICE_PARTIAL_PARSE",
            status="partial_success",
            app_type=app_name,
            action_type=resolution.get("action_type", "unknown"),
            details=" | ".join(plan_info.get("errors") or []),
            source=source,
            plan=plan_info,
            file_path=summary["output_path"],
            output_file=summary["output_path"],
            persisted=summary.get("persisted", False),
            opened=summary.get("opened", False),
            action_count=summary.get("total", len(actions)),
            executed=summary.get("executed", []),
            results=summary.get("results", []),
            executed_count=summary.get("ok_count", 0),
            failed_count=len(plan_info.get("failed_clauses") or []),
            **_d,
            final_actions=final_actions,
            executor_results=executor_results,
        )

    app_label = {"excel": "Excel", "word": "Word", "powerpoint": "PowerPoint"}.get(app_name, app_name.title())
    return _json_success(
        f"Created {app_label} file successfully." if resolution.get("action_type") == "create" else f"Updated {app_label} file successfully.",
        intent="office_automation",
        app_type=app_name,
        action_type=resolution.get("action_type", "unknown"),
        file_path=summary["output_path"],
        source=source,
        plan=plan_info,
        output_file=summary["output_path"],
        persisted=summary.get("persisted", False),
        opened=summary.get("opened", False),
        action_count=summary.get("total", len(actions)),
        executed=summary.get("executed", []),
        results=summary.get("results", []),
        **_d,
        final_actions=final_actions,
        executor_results=executor_results,
    )


@office_bp.route("/office/execute", methods=["POST"])
def office_execute():
    try:
        return _office_execute_impl(request.get_json(silent=True) or {})
    except Exception as e:
        logging.error("Office route error: %s\n%s", e, traceback.format_exc())
        return _json_error(
            "Office command execution failed.",
            intent="office_automation",
            error_code="OFFICE_ROUTE_ERROR",
            details=str(e),
            http_status=500,
        )


@office_bp.route("/command", methods=["POST"])
def office_command():
    try:
        return _office_execute_impl(request.get_json(silent=True) or {})
    except Exception as e:
        logging.error("Command office route error: %s\n%s", e, traceback.format_exc())
        return _json_error(
            "Office command execution failed.",
            intent="office_automation",
            error_code="OFFICE_ROUTE_ERROR",
            details=str(e),
            http_status=500,
        )
