import logging
import re

from office.intent import _canonical_office_app, _default_create_action
from office.runner import _known_office_actions
from parser.command_parser import parse_command
from parser.command_planner import plan_office_command, split_command_clauses
from parser.command_complexity import classify_office_command_complexity
from utils.office_actions import OfficeActionError, validate_actions
from utils import command_map
import state


def _resolve_actions(app_name, command_text):
    def _estimate_subcommands(text):
        return max(1, len(split_command_clauses(text)))

    def _actions_cover_command_intents(app, text, actions):
        low   = (text or "").lower()
        names = {
            str(a.get("action", "")).strip().lower()
            for a in (actions or [])
            if isinstance(a, dict)
        }
        if not names:
            return False

        checks = []
        if app == "excel":
            if "background" in low:
                checks.append("set_bg_color" in names)
            if "font color" in low:
                checks.append("set_font_color" in names)
            if "font size" in low:
                checks.append("set_font_size" in names)
            if "number format" in low:
                checks.append("set_number_format" in names)
            if "formula" in low or "sum" in low:
                checks.append("write_formula" in names)
            if "rename" in low and "sheet" in low:
                checks.append("rename_sheet" in names)
            if "insert row" in low:
                checks.append("insert_row" in names)
            if "[[" in low and "write" in low and "values" in low:
                checks.append("write_range" in names)
            if re.search(r"\bfill\b", low) and re.search(r"\d", low) and "color" not in low:
                checks.append("write_range" in names or "write_cell" in names)
            if "background color" in low and "cells" in low and re.search(
                r"\b[A-Z]{1,3}\d{1,7}\b.*\band\b.*\b[A-Z]{1,3}\d{1,7}\b", text, re.IGNORECASE
            ):
                bg_count = sum(
                    1 for a in (actions or [])
                    if isinstance(a, dict) and str(a.get("action", "")).lower() == "set_bg_color"
                )
                checks.append(bg_count >= 2)

        return all(checks) if checks else True

    complexity = classify_office_command_complexity(command_text)
    logging.info("Office command complexity [%s]: %s — %r", app_name, complexity, command_text[:120])

    diag = {
        "openai_attempted":        False,
        "openai_success":          False,
        "openai_error_code":       None,
        "fallback_reason":         "",
        "raw_action_count":        0,
        "normalized_action_count": 0,
        "validation_errors":       [],
    }

    def _with_diag(pi, extra_diag=None):
        if extra_diag:
            diag.update(extra_diag)
        base = dict(pi) if isinstance(pi, dict) else {}
        base["diag"] = dict(diag)
        return base

    def _ai_plan_info(result):
        info = {}
        if result.output_filename:
            info["output_filename"] = result.output_filename
        if result.ai_context:
            info["ai_context"] = result.ai_context
        return info

    def _call_openai(reason_label):
        logging.info("OPENAI_PLANNER_CALLED app=%s command_len=%d", app_name, len(command_text))
        result = state.openai_handler.interpret_result(app_name, command_text)
        logging.info(
            "OPENAI_PLANNER_RESULT success=%s error_code=%s action_count=%d",
            result.success, result.error_code or "", len(result.actions),
        )
        diag["openai_attempted"] = True
        diag["openai_success"]   = result.success
        if not result.success:
            diag["openai_error_code"] = result.error_code or ""
        diag["fallback_reason"] = reason_label
        return result

    if complexity == "semantic_complex":
        logging.info("Routing semantic_complex command directly to OpenAI for [%s]", app_name)
        ai_result = _call_openai("semantic_complex_routed_to_openai")
        if ai_result.success:
            diag["raw_action_count"] = len(ai_result.actions)
            try:
                normalized = validate_actions(app_name, ai_result.actions, known_actions=_known_office_actions(app_name))
            except OfficeActionError as exc:
                diag["openai_error_code"]  = exc.error_code
                diag["validation_errors"].append(exc.message)
                diag["fallback_reason"]    = "openai_validation_failed"
                return command_text, [], "openai-fallback", exc, _with_diag(_ai_plan_info(ai_result))
            diag["normalized_action_count"] = len(normalized)
            command_map.save_actions(app_name, command_text, normalized)
            return command_text, normalized, "openai-fallback", None, _with_diag(_ai_plan_info(ai_result))
        diag["fallback_reason"] = "semantic_openai_failed_fell_through"
        logging.info("OpenAI unavailable for semantic command; falling through to deterministic. error=%s", ai_result.error_code)

    cache_key, cached_actions, cache_score = command_map.get_cached_actions(app_name, command_text)
    if cached_actions and cache_score == 100:
        cached_count  = len([a for a in cached_actions if isinstance(a, dict) and a.get("action")])
        clause_count  = _estimate_subcommands(command_text)
        try:
            validated_cached = validate_actions(app_name, cached_actions, known_actions=_known_office_actions(app_name))
        except OfficeActionError as exc:
            logging.info("Ignoring invalid command cache for [%s]: %s", app_name, exc.message)
            diag["validation_errors"].append(f"cache: {exc.message}")
            validated_cached = []
        if (
            validated_cached
            and cached_count >= clause_count
            and _actions_cover_command_intents(app_name, command_text, validated_cached)
        ):
            logging.info("Office cache hit [%s] score=%s: %s", app_name, cache_score, command_text)
            diag["fallback_reason"]         = "cache_hit"
            diag["raw_action_count"]        = len(cached_actions)
            diag["normalized_action_count"] = len(validated_cached)
            return cache_key or command_text, validated_cached, "command-cache", None, _with_diag(None)
        logging.info(
            "Ignoring stale cache for [%s] command (cached=%s, clauses=%s): %s",
            app_name, cached_count, clause_count, command_text,
        )

    plan      = plan_office_command(app_name, command_text)
    plan_info = plan.to_dict()

    _partial_plan_actions = plan.actions or []
    _partial_plan_info    = plan_info
    _partial_validation = {"checked": False, "actions": []}

    def _validated_partial_plan():
        if not _partial_plan_actions:
            return []
        if _partial_validation["checked"]:
            return _partial_validation["actions"]
        _partial_validation["checked"] = True
        try:
            validated_partial = validate_actions(
                app_name, _partial_plan_actions, known_actions=_known_office_actions(app_name)
            )
        except OfficeActionError as exc:
            diag["validation_errors"].append(f"partial-planner: {exc.message}")
            return []
        _partial_validation["actions"] = validated_partial
        return validated_partial

    if plan.actions and not plan.requires_api:
        diag["raw_action_count"] = len(plan.actions)
        try:
            actions = validate_actions(app_name, plan.actions, known_actions=_known_office_actions(app_name))
        except OfficeActionError as exc:
            diag["validation_errors"].append(f"planner: {exc.message}")
            diag["fallback_reason"] = "planner_validation_failed_fell_through"
            logging.info(
                "Office planner validation failed for [%s]; falling through to parser. Command: %s",
                app_name, command_text,
            )
            logging.warning("OFFICE_VALIDATION_FAILED errors=%s", diag["validation_errors"])
        else:
            diag["fallback_reason"]         = "planner_success"
            diag["normalized_action_count"] = len(actions)
            if plan.success and not plan.errors:
                command_map.save_actions(app_name, command_text, actions)
            return command_text, actions, "planner", None, _with_diag(plan_info)
    elif plan.actions and plan.requires_api:
        diag["raw_action_count"] = len(plan.actions)
        diag["fallback_reason"]  = "planner_partial_fell_through"
        logging.info(
            "Office planner partial parse (requires_api=True): app=%s clauses=%s actions=%s errors=%s — falling through",
            app_name, len(plan.clauses), len(plan.actions), len(plan.errors),
        )

    raw_parser_actions = parse_command(app_name, command_text)
    if raw_parser_actions:
        diag["raw_action_count"] = max(diag["raw_action_count"], len(raw_parser_actions))
        try:
            actions = validate_actions(app_name, raw_parser_actions, known_actions=_known_office_actions(app_name))
        except OfficeActionError as exc:
            diag["validation_errors"].append(f"parser: {exc.message}")
            diag["fallback_reason"] = "parser_validation_failed_fell_through"
            logging.info(
                "Office parser validation failed for [%s]; falling through to OpenAI. Command: %s",
                app_name, command_text,
            )
            logging.warning("OFFICE_VALIDATION_FAILED errors=%s", diag["validation_errors"])
            actions = []
        else:
            diag["fallback_reason"]         = "parser_success"
            diag["normalized_action_count"] = len(actions)
            command_map.save_actions(app_name, command_text, actions)
            return command_text, actions, "json-parser", None, _with_diag(None)

    ai_result = _call_openai("openai_fallback")
    if ai_result.success:
        diag["raw_action_count"] = max(diag["raw_action_count"], len(ai_result.actions))
        try:
            normalized = validate_actions(app_name, ai_result.actions, known_actions=_known_office_actions(app_name))
        except OfficeActionError as exc:
            diag["openai_error_code"] = exc.error_code
            diag["validation_errors"].append(f"openai: {exc.message}")
            diag["fallback_reason"]   = "openai_validation_failed"
            logging.warning("OFFICE_VALIDATION_FAILED errors=%s", diag["validation_errors"])
            validated_partial = _validated_partial_plan()
            if validated_partial:
                diag["fallback_reason"]         = "partial_planner_fallback"
                diag["normalized_action_count"] = len(validated_partial)
                return command_text, validated_partial, "planner-partial", None, _with_diag(_partial_plan_info)
            return command_text, [], "openai-fallback", exc, _with_diag(_ai_plan_info(ai_result))
        validated_partial = _validated_partial_plan()
        if validated_partial:
            diag["fallback_reason"]         = "partial_planner_fallback"
            diag["normalized_action_count"] = len(validated_partial)
            return command_text, validated_partial, "planner-partial", None, _with_diag(_partial_plan_info)
        diag["normalized_action_count"] = len(normalized)
        diag["fallback_reason"]         = "openai_fallback_success"
        command_map.save_actions(app_name, command_text, normalized)
        return command_text, normalized, "openai-fallback", None, _with_diag(_ai_plan_info(ai_result))

    if ai_result.error_code in {
        "OPENAI_INVALID_JSON", "OPENAI_INVALID_ACTION_SCHEMA",
        "OPENAI_UNSUPPORTED_ACTION", "COMMAND_TOO_LONG",
    }:
        diag["fallback_reason"] = "openai_hard_failure"
        return command_text, [], "openai-fallback", OfficeActionError(
            ai_result.error_code or "OPENAI_REQUEST_FAILED",
            ai_result.message or "OpenAI fallback could not parse this Office command.",
            ai_result.raw_response_preview or "",
        ), _with_diag(None)

    if _partial_plan_actions:
        logging.info(
            "Using partial planner result as last resort for [%s]: %d actions",
            app_name, len(_partial_plan_actions),
        )
        validated_partial = _validated_partial_plan()
        if validated_partial:
            diag["fallback_reason"]         = "partial_planner_fallback"
            diag["normalized_action_count"] = len(validated_partial)
            return command_text, validated_partial, "planner-partial", None, _with_diag(_partial_plan_info)

    fallback = _default_create_action(app_name, command_text)
    if fallback:
        diag["fallback_reason"]         = "intent_fallback_create"
        diag["normalized_action_count"] = 1
        return command_text, [fallback], "office-intent-fallback", None, _with_diag(None)

    if ai_result.error_code:
        diag["fallback_reason"] = "openai_unavailable_no_match"
        return command_text, [], "openai-fallback", OfficeActionError(
            ai_result.error_code,
            ai_result.message or "OpenAI fallback was unavailable.",
            ai_result.raw_response_preview or "",
        ), _with_diag(None)

    diag["fallback_reason"] = "no_match"
    return command_text, [], "no-match", None, _with_diag(None)
