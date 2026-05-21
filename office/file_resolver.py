import logging
import os
import re
from pathlib import Path

from office.constants import BASE_DIR, OFFICE_OUTPUT_DIR
from office.intent import _action_names, _canonical_office_app, _contains_term, _detect_action_type
from utils.file_paths import (
    OFFICE_EXTENSIONS,
    FilePathError,
    common_user_locations,
    extract_office_filename_hint,
    generate_office_output_path,
    named_output_path,
    next_available_path,
    resolve_existing_office_path,
    resolve_path_value,
    sanitize_filename,
)


def _next_available_path(path):
    return str(next_available_path(path).resolve())


def _generate_new_output_path(app_name):
    return str(generate_office_output_path(app_name))


def _resolve_path_value(value, app_name, for_output=False):
    try:
        resolved = resolve_path_value(value, app_name, for_output=for_output, base_dir=BASE_DIR)
    except FilePathError as exc:
        logging.warning("Office path resolution rejected: %s", exc.message)
        raise
    return str(resolved) if resolved else ""


def _first_action_path(actions, action_names, path_keys=("path", "file_path", "filename", "output_path")):
    for action in actions or []:
        if not isinstance(action, dict):
            continue
        if str(action.get("action", "")).strip().lower() not in action_names:
            continue
        for key in path_keys:
            value = action.get(key)
            if str(value or "").strip():
                return str(value).strip(), action
    return "", None


def _save_as_action_names(app_name):
    return {
        "excel":      {"save_workbook_as"},
        "word":       {"save_document_as"},
        "powerpoint": {"save_presentation_as"},
        "ppt":        {"save_presentation_as"},
    }.get(app_name, set())


def _open_action_names(app_name):
    return {
        "excel":      {"open_workbook"},
        "word":       {"open_document"},
        "powerpoint": {"open_presentation"},
        "ppt":        {"open_presentation"},
    }.get(app_name, set())


def _extract_named_file_path(command_text, app_name):
    text = (command_text or "").strip()
    ext = OFFICE_EXTENSIONS.get(app_name)
    if not text or not ext:
        return ""

    def _output_path_for_name(name):
        candidate = Path(str(name or "").strip())
        has_location_hint = re.search(r"\b(?:desktop|documents?|downloads?)\b", text, re.IGNORECASE)
        if has_location_hint and not candidate.is_absolute() and len(candidate.parts) == 1:
            location = common_user_locations(text)[0]
            return str(resolve_path_value(location / sanitize_filename(candidate.name), app_name, for_output=True, base_dir=BASE_DIR))
        return str(resolve_path_value(str(candidate), app_name, for_output=True, base_dir=BASE_DIR))

    def _clean_base(name):
        cleaned = sanitize_filename(name)
        cleaned = re.split(r"\s+(?:and|then|with|in|on)\b", cleaned, maxsplit=1, flags=re.IGNORECASE)[0]
        return re.sub(r"\s+", " ", cleaned).strip(" .")

    quoted = re.search(r'["\']([^"\']+\.' + re.escape(ext) + r')["\']', text, re.IGNORECASE)
    if quoted:
        return _output_path_for_name(quoted.group(1).strip())

    plain = re.search(r'\b([A-Za-z0-9_\- .]+\.' + re.escape(ext) + r')\b', text, re.IGNORECASE)
    if plain:
        return _output_path_for_name(plain.group(1).strip())

    named = re.search(
        r'\b(?:named|called|name)\s*[:=]?\s*(?:["\']([A-Za-z0-9_\- .]{1,100})["\']|([A-Za-z0-9_\- .]{1,100})\b)',
        text, re.IGNORECASE,
    )
    if named:
        base = _clean_base(named.group(1) or named.group(2))
        if base:
            if re.search(r"\b(?:desktop|documents?|downloads?)\b", text, re.IGNORECASE):
                return _output_path_for_name(base)
            return str(named_output_path(base, app_name))

    return ""


def _is_fresh_file_intent(app_name, command_text, actions):
    names = _action_names(actions)
    create_actions = {
        "excel":      {"create_workbook"},
        "word":       {"create_document"},
        "powerpoint": {"create_presentation"},
        "ppt":        {"create_presentation"},
    }
    if names & create_actions.get(app_name, set()):
        return True
    text = (command_text or "").lower()
    creation_words = ("create", "new", "start", "make")
    target_words   = ("file", "workbook", "document", "presentation", "ppt")
    return any(w in text for w in creation_words) and any(w in text for w in target_words)


def _should_start_fresh(app_name, command_text, actions, file_path):
    if file_path:
        return False
    if _extract_named_file_path(command_text, app_name):
        return False

    names = _action_names(actions)
    lifecycle_only = {
        "save_workbook",    "save_workbook_as",    "close_workbook",
        "save_document",    "save_document_as",    "close_document",
        "save_presentation","save_presentation_as","close_presentation",
    }
    if names and names <= lifecycle_only:
        return False

    open_actions = {
        "excel":      {"open_workbook"},
        "word":       {"open_document"},
        "powerpoint": {"open_presentation"},
        "ppt":        {"open_presentation"},
    }
    return not bool(_action_names(actions) & open_actions.get(app_name, set()))


def _ensure_fresh_file_action(app_name, command_text, actions, file_path):
    actions = list(actions or [])
    if not actions or not _should_start_fresh(app_name, command_text, actions, file_path):
        return actions

    create_action = {
        "excel":      "create_workbook",
        "word":       "create_document",
        "powerpoint": "create_presentation",
        "ppt":        "create_presentation",
    }.get(app_name)
    if not create_action:
        return actions
    if str(actions[0].get("action", "")).strip().lower() == create_action:
        return actions

    logging.info("Prepending %s for fresh %s file: %s", create_action, app_name, command_text)
    return [{"action": create_action}, *actions]


def _expand_powerpoint_slide_count(app_name, command_text, actions):
    if _canonical_office_app(app_name) != "powerpoint":
        return actions
    text = (command_text or "").lower()
    match = re.search(r"\b(?:create|make|generate|build|add)\b.*?\b(\d{1,2})\s+slides?\b", text)
    if not match:
        return actions
    target_count = max(1, min(int(match.group(1)), 50))
    existing = sum(
        1 for a in actions
        if isinstance(a, dict) and str(a.get("action", "")).lower() == "add_slide"
    )
    if existing >= target_count:
        return actions
    return list(actions) + [
        {"action": "add_slide", "layout": "title_content"}
        for _ in range(target_count - existing)
    ]


def _resolve_output_file_path(app_name, command_text, actions, file_path):
    explicit = (file_path or "").strip()
    if explicit:
        return os.path.abspath(explicit)
    named = _extract_named_file_path(command_text, app_name)
    if named:
        if _is_fresh_file_intent(app_name, command_text, actions):
            return _next_available_path(named)
        return named
    if _should_start_fresh(app_name, command_text, actions, ""):
        return _generate_new_output_path(app_name)
    return ""


def resolve_office_file_path(request_payload, actions, app_type, mode=None):
    app_name    = _canonical_office_app(app_type)
    actions     = actions or []
    command_text = (
        (request_payload or {}).get("raw")
        or (request_payload or {}).get("command")
        or ""
    )
    action_type = mode or _detect_action_type(command_text, actions)
    OFFICE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    explicit = (
        (request_payload or {}).get("file_path")
        or (request_payload or {}).get("file")
        or ""
    )
    try:
        explicit_path = _resolve_path_value(explicit, app_name, for_output=False) if explicit else ""
    except FilePathError as exc:
        return {"success": False, "error_code": exc.error_code, "message": exc.message,
                "details": exc.details, "app_type": app_name, "action_type": action_type}

    open_value, open_action = _first_action_path(actions, _open_action_names(app_name))
    if not open_value and open_action is not None:
        open_value = extract_office_filename_hint(command_text, app_name) or ""
    try:
        open_path = str(resolve_existing_office_path(
            open_value, app_name, base_dir=BASE_DIR, command_text=command_text,
        )) if open_value else ""
    except FilePathError as exc:
        return {"success": False, "error_code": exc.error_code, "message": exc.message,
                "details": exc.details, "app_type": app_name, "action_type": action_type}

    save_as_value, save_action = _first_action_path(
        actions, _save_as_action_names(app_name),
        path_keys=("filename", "path", "file_path", "output_path"),
    )
    try:
        save_as_path = _resolve_path_value(save_as_value, app_name, for_output=True) if save_as_value else ""
    except FilePathError as exc:
        return {"success": False, "error_code": exc.error_code, "message": exc.message,
                "details": exc.details, "app_type": app_name, "action_type": action_type}

    try:
        named       = _extract_named_file_path(command_text, app_name)
        named_path  = _resolve_path_value(named, app_name, for_output=True) if named else ""
    except FilePathError as exc:
        return {"success": False, "error_code": exc.error_code, "message": exc.message,
                "details": exc.details, "app_type": app_name, "action_type": action_type}

    fresh = _is_fresh_file_intent(app_name, command_text, actions)
    source_path = (explicit_path if explicit_path and not fresh else "") or open_path

    if source_path and not Path(source_path).exists():
        return {
            "success":     False,
            "error_code":  "FILE_NOT_FOUND",
            "message":     f"Office input file was not found: {source_path}",
            "details":     source_path,
            "app_type":    app_name,
            "action_type": action_type,
        }

    if explicit_path:
        output_path = explicit_path
    elif save_as_path:
        output_path = save_as_path
    elif named_path and (fresh or not source_path):
        output_path = _next_available_path(named_path) if fresh else named_path
    elif source_path:
        output_path = source_path
    else:
        output_path = _generate_new_output_path(app_name)

    output = Path(output_path)
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return {
            "success":    False,
            "error_code": "INVALID_FILE_PATH",
            "message":    f"Could not create output directory: {output.parent}",
            "details":    str(exc),
            "app_type":   app_name,
            "action_type": action_type,
        }

    if save_action is not None:
        save_action["filename"] = str(output.resolve())
    if open_action is not None and open_path:
        open_action["path"] = open_path

    return {
        "success":     True,
        "app_type":    app_name,
        "action_type": action_type,
        "source_path": source_path,
        "file_path":   str(output.resolve()),
        "output_path": str(output.resolve()),
        "reason": (
            "frontend file path"           if explicit_path else
            "save-as action path"          if save_as_path  else
            "open action path"             if open_path     else
            "named path from command"      if named_path    else
            "generated default output path"
        ),
    }
