import logging


logger = logging.getLogger("OfficeAgent")


class OfficeActionError(ValueError):
    def __init__(self, error_code, message, details=""):
        super().__init__(message)
        self.error_code = error_code
        self.message = message
        self.details = details


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

    normalized = []
    for idx, item in enumerate(actions):
        if not isinstance(item, dict):
            raise OfficeActionError(
                "INVALID_OFFICE_ACTION",
                f"Office action at index {idx} must be an object.",
                f"Got {type(item).__name__}.",
            )
        action_name = str(item.get("action", "")).strip()
        if not action_name:
            raise OfficeActionError(
                "INVALID_OFFICE_ACTION",
                f"Office action at index {idx} is missing required field: action.",
            )
        cleaned = dict(item)
        cleaned["action"] = action_name
        normalized.append(cleaned)

    return normalized


def _has_any(action, fields):
    return any(str(action.get(field, "")).strip() for field in fields)


REQUIRED_FIELDS = {
    "excel": {
        "open_workbook": (("path", "file_path", "filename"),),
        "write_cell": (("cell",),),
        "write_formula": (("cell",), ("formula",)),
        "write_range": (("start_cell",), ("values",)),
        "clear_range": (("range",),),
        "clear_format": (("range",),),
        "set_bold": (("range",),),
        "set_italic": (("range",),),
        "set_underline": (("range",),),
        "set_strikethrough": (("range",),),
        "set_font_size": (("range",), ("size",)),
        "set_font_name": (("range",), ("name",)),
        "set_font_color": (("range",), ("color",)),
        "set_bg_color": (("range",), ("color",)),
        "set_border": (("range",),),
        "remove_border": (("range",),),
        "set_alignment": (("range",),),
        "set_vertical_alignment": (("range",),),
        "set_wrap_text": (("range",),),
        "set_number_format": (("range",), ("format",)),
        "merge_cells": (("range",),),
        "unmerge_cells": (("range",),),
        "insert_row": (("row",),),
        "delete_row": (("row",),),
        "insert_column": (("column",),),
        "delete_column": (("column",),),
        "set_row_height": (("row",), ("height",)),
        "set_column_width": (("column",), ("width",)),
        "set_active_sheet": (("name",),),
        "hide_sheet": (("name",),),
        "unhide_sheet": (("name",),),
        "freeze_panes": (("cell",),),
        "filter_range": (("range",),),
        "sort_range": (("range",),),
        "insert_comment": (("cell",),),
        "delete_comment": (("cell",),),
        "insert_hyperlink": (("cell",), ("url",)),
        "insert_image": (("path",),),
    },
    "word": {
        "open_document": (("path", "file_path", "filename"),),
        "set_font_size": (("size",),),
        "set_font_name": (("name",),),
        "set_font_color": (("color",),),
        "set_alignment": (("alignment",),),
        "set_line_spacing": (("spacing",),),
        "find_text": (("text",),),
        "find_replace": (("find_text",), ("replace_text",)),
        "insert_image": (("path",),),
        "insert_hyperlink": (("url",),),
        "compare_documents": (("path", "compare_path"),),
    },
    "powerpoint": {
        "open_presentation": (("path", "file_path", "filename"),),
        "set_slide_text": (("text",),),
        "add_bullet_point": (("text",),),
        "add_numbered_point": (("text",),),
        "set_speaker_notes": (("text",),),
        "set_font_size": (("size",),),
        "set_font_name": (("name",),),
        "set_font_color": (("color",),),
        "set_bg_color": (("color",),),
        "set_bg_image": (("path",),),
        "add_logo": (("path",),),
        "insert_image": (("path",),),
        "insert_video": (("path",),),
        "insert_audio": (("path",),),
        "insert_hyperlink": (("url",),),
        "set_slide_size": (("width",), ("height",)),
    },
}


def validate_actions(app_type, actions, known_actions=None):
    app = "powerpoint" if app_type == "ppt" else (app_type or "").lower().strip()
    actions = normalize_actions(actions)
    known = set(known_actions or [])

    for idx, action in enumerate(actions):
        name = str(action.get("action", "")).strip()
        if known and name not in known:
            raise OfficeActionError(
                "INVALID_OFFICE_ACTION",
                f"Unknown Office action: {name}.",
                f"Action index {idx}.",
            )

        for field_group in REQUIRED_FIELDS.get(app, {}).get(name, ()):
            if not _has_any(action, field_group):
                field_text = " or ".join(field_group)
                raise OfficeActionError(
                    "INVALID_OFFICE_ACTION",
                    f"Action '{name}' is missing required field: {field_text}.",
                    f"Action index {idx}.",
                )

    return actions
