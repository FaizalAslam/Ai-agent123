import logging
import re

from utils.office_action_registry import (
    MAX_ACTIONS,
    MAX_GENERATED_COLUMNS,
    MAX_GENERATED_ROWS,
    MAX_GENERATED_SLIDES,
    MAX_RANGE_CELLS,
    MAX_TEXT_LENGTH,
    canonical_app,
    get_action_spec,
    get_known_actions,
)


logger = logging.getLogger("OfficeAgent")

COLOR_NAMES = {
    "red", "green", "blue", "yellow", "orange", "purple", "pink", "black",
    "white", "gray", "grey", "dark red", "dark blue", "dark green",
    "light blue", "light gray", "teal", "cyan", "magenta", "gold", "brown",
    "navy",
}
COLOR_FIELDS = {"color", "font_color", "background_color", "fill_color"}


class OfficeActionError(ValueError):
    def __init__(self, error_code, message, details="", action_index=None, action_name=None):
        super().__init__(message)
        self.error_code = error_code
        self.message = message
        self.details = details
        self.action_index = action_index
        self.action_name = action_name


def normalize_actions(raw_actions):
    if isinstance(raw_actions, dict):
        logger.warning("Office action parser returned a single object; wrapping it in a list.")
        actions = [raw_actions]
    elif isinstance(raw_actions, list):
        actions = raw_actions
    else:
        raise OfficeActionError(
            "INVALID_OFFICE_ACTION",
            "Office actions must be a JSON array of action objects.",
            f"Got {type(raw_actions).__name__}.",
        )

    if len(actions) > MAX_ACTIONS:
        raise OfficeActionError(
            "ACTION_LIMIT_EXCEEDED",
            f"Office action plan contains {len(actions)} actions; maximum is {MAX_ACTIONS}.",
        )

    normalized = []
    for idx, item in enumerate(actions):
        if not isinstance(item, dict):
            raise OfficeActionError(
                "INVALID_OFFICE_ACTION",
                f"Office action at index {idx} must be an object.",
                f"Got {type(item).__name__}.",
                action_index=idx,
            )
        action_name = str(item.get("action", "")).strip()
        if not action_name:
            raise OfficeActionError(
                "INVALID_OFFICE_ACTION",
                f"Office action at index {idx} is missing required field: action.",
                action_index=idx,
            )
        cleaned = dict(item)
        cleaned["action"] = action_name
        normalized.append(cleaned)

    return normalized


def _has_any(action, fields):
    return any(str(action.get(field, "")).strip() for field in fields)


def _is_positive_int(value):
    try:
        return int(value) > 0
    except (TypeError, ValueError):
        return False


def _column_index(col):
    n = 0
    for ch in (col or "").upper():
        if not ("A" <= ch <= "Z"):
            return 0
        n = (n * 26) + (ord(ch) - ord("A") + 1)
    return n


def _range_cell_count(range_ref):
    ref = str(range_ref or "").strip().upper().replace(" ", "")
    if not ref:
        return 0
    if re.fullmatch(r"[A-Z]{1,3}\d{1,7}", ref):
        return 1
    m = re.fullmatch(r"([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})", ref)
    if m:
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        cols = abs(_column_index(c2) - _column_index(c1)) + 1
        rows = abs(r2 - r1) + 1
        return max(1, rows * cols)
    m = re.fullmatch(r"([A-Z]{1,3}):([A-Z]{1,3})", ref)
    if m:
        cols = abs(_column_index(m.group(2)) - _column_index(m.group(1))) + 1
        return cols * 10000
    m = re.fullmatch(r"(\d{1,7}):(\d{1,7})", ref)
    if m:
        rows = abs(int(m.group(2)) - int(m.group(1))) + 1
        return rows * 200
    return 1


def _looks_like_range_or_cell(value):
    ref = str(value or "").strip().upper().replace(" ", "")
    return bool(re.fullmatch(r"[A-Z]{1,3}\d{1,7}(:[A-Z]{1,3}\d{1,7})?|[A-Z]{1,3}:[A-Z]{1,3}|\d{1,7}:\d{1,7}", ref))


def _validate_text_lengths(action, idx, name):
    for key, value in action.items():
        if isinstance(value, str) and re.search(r"\{[A-Za-z_][A-Za-z0-9_]*\}", value):
            raise OfficeActionError(
                "INVALID_ACTION",
                f"Action '{name}' has unresolved placeholder in field '{key}'.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )
        if key in {"text", "value", "formula", "find_text", "replace_text"} and isinstance(value, str):
            if len(value) > MAX_TEXT_LENGTH:
                raise OfficeActionError(
                    "INVALID_OFFICE_ACTION",
                    f"Action '{name}' field '{key}' is too long.",
                    f"Action index {idx}.",
                    action_index=idx,
                    action_name=name,
                )


def _validate_path_fields(action, idx, name):
    for key in ("path", "file_path", "filename", "output_path", "compare_path"):
        value = action.get(key)
        if not str(value or "").strip():
            continue
        raw = str(value)
        if "\x00" in raw:
            raise OfficeActionError(
                "UNSAFE_PATH",
                f"Action '{name}' contains an unsafe path.",
                f"Action index {idx}, field {key}.",
                action_index=idx,
                action_name=name,
            )


def _validate_color_fields(action, idx, name):
    for key, value in action.items():
        if key not in COLOR_FIELDS or value is None:
            continue
        raw = str(value or "").strip().lstrip("#")
        if not raw:
            raise OfficeActionError(
                "INVALID_COLOR",
                f"Action '{name}' has an empty color value.",
                f"Action index {idx}, field {key}.",
                action_index=idx,
                action_name=name,
            )
        if raw.lower() in COLOR_NAMES:
            continue
        if re.fullmatch(r"[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8}", raw):
            continue
        raise OfficeActionError(
            "INVALID_COLOR",
            f"Action '{name}' has invalid color: {value}.",
            f"Action index {idx}, field {key}.",
            action_index=idx,
            action_name=name,
        )


def _validate_excel_action(action, idx, name):
    for key in ("cell", "start_cell"):
        if key in action and not _looks_like_range_or_cell(action.get(key)):
            raise OfficeActionError(
                "INVALID_EXCEL_RANGE",
                f"Action '{name}' has invalid cell reference: {action.get(key)}.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )

    if "range" in action:
        rng = str(action.get("range") or "").strip().upper().replace(" ", "")
        if not _looks_like_range_or_cell(rng):
            raise OfficeActionError(
                "INVALID_EXCEL_RANGE",
                f"Action '{name}' has invalid range: {action.get('range')}.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )
        if _range_cell_count(rng) > MAX_RANGE_CELLS:
            raise OfficeActionError(
                "RANGE_TOO_LARGE",
                f"Action '{name}' range is too large to run safely.",
                f"Action index {idx}, range {rng}.",
                action_index=idx,
                action_name=name,
            )
        action["range"] = rng

    if name == "create_table":
        rows = int(action.get("rows", 5))
        cols = int(action.get("cols", 3))
        if rows <= 0 or cols <= 0 or rows > MAX_GENERATED_ROWS or cols > MAX_GENERATED_COLUMNS:
            raise OfficeActionError(
                "RANGE_TOO_LARGE",
                f"Table size {rows}x{cols} exceeds safe limits.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )


def _validate_powerpoint_action(action, idx, name):
    for key in ("slide_index", "from_index", "to_index"):
        if key in action and not _is_positive_int(action.get(key)):
            raise OfficeActionError(
                "INVALID_SLIDE_INDEX",
                f"Action '{name}' field '{key}' must be a positive 1-based slide number.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )
        if key in action and int(action.get(key)) > MAX_GENERATED_SLIDES:
            raise OfficeActionError(
                "ACTION_LIMIT_EXCEEDED",
                f"Action '{name}' slide index exceeds the safe limit of {MAX_GENERATED_SLIDES}.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )


def _validate_types(spec, action, idx, name):
    for field, allowed in spec.field_types.items():
        if field not in action or action.get(field) is None:
            continue
        if not isinstance(action.get(field), allowed):
            raise OfficeActionError(
                "INVALID_ACTION",
                f"Action '{name}' field '{field}' has invalid type.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )


def validate_actions(app_type, actions, known_actions=None):
    app = canonical_app(app_type)
    actions = normalize_actions(actions)
    known = set(known_actions or []) or get_known_actions(app)

    for idx, action in enumerate(actions):
        name = str(action.get("action", "")).strip()
        if known and name not in known:
            raise OfficeActionError(
                "UNKNOWN_ACTION",
                f"Unknown Office action: {name}.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )

        spec = get_action_spec(app, name)
        if not spec:
            if known and name in known:
                _validate_text_lengths(action, idx, name)
                _validate_path_fields(action, idx, name)
                if app == "excel":
                    _validate_excel_action(action, idx, name)
                elif app == "powerpoint":
                    _validate_powerpoint_action(action, idx, name)
                continue
            raise OfficeActionError(
                "UNSUPPORTED_ACTION",
                f"Unsupported Office action for {app}: {name}.",
                f"Action index {idx}.",
                action_index=idx,
                action_name=name,
            )

        for field_group in spec.required_groups:
            if not _has_any(action, field_group):
                field_text = " or ".join(field_group)
                raise OfficeActionError(
                    "MISSING_REQUIRED_FIELD",
                    f"Action '{name}' is missing required field: {field_text}.",
                    f"Action index {idx}.",
                    action_index=idx,
                    action_name=name,
                )

        _validate_types(spec, action, idx, name)
        _validate_text_lengths(action, idx, name)
        _validate_path_fields(action, idx, name)
        _validate_color_fields(action, idx, name)
        if app == "excel":
            _validate_excel_action(action, idx, name)
        elif app == "powerpoint":
            _validate_powerpoint_action(action, idx, name)

    return actions
