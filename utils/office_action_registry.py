from dataclasses import dataclass, field
from typing import Any


MAX_ACTIONS = 50
MAX_RANGE_CELLS = 50000
MAX_TEXT_LENGTH = 10000
MAX_GENERATED_ROWS = 5000
MAX_GENERATED_COLUMNS = 200
MAX_GENERATED_SLIDES = 50


@dataclass(frozen=True)
class ActionSpec:
    app: str
    name: str
    required_groups: tuple[tuple[str, ...], ...] = ()
    field_types: dict[str, tuple[type, ...]] = field(default_factory=dict)
    destructive: bool = False
    path_required: bool = False


def _spec(app, name, required_groups=(), field_types=None, destructive=False, path_required=False):
    return ActionSpec(
        app=app,
        name=name,
        required_groups=tuple(tuple(group) for group in required_groups),
        field_types=field_types or {},
        destructive=destructive,
        path_required=path_required,
    )


TEXT = (str,)
INTISH = (int, str)
NUMERIC = (int, float, str)
BOOLISH = (bool, str)
LIST = (list,)


EXCEL_ACTIONS = {
    "create_workbook": _spec("excel", "create_workbook"),
    "open_workbook": _spec("excel", "open_workbook", (("path", "file_path", "filename"),), path_required=True),
    "save_workbook": _spec("excel", "save_workbook"),
    "save_workbook_as": _spec("excel", "save_workbook_as", (("filename", "path", "file_path", "output_path"),)),
    "close_workbook": _spec("excel", "close_workbook"),
    "add_sheet": _spec("excel", "add_sheet", (("name",),), {"name": TEXT}),
    "set_active_sheet": _spec("excel", "set_active_sheet", (("name",),), {"name": TEXT}),
    "rename_sheet": _spec("excel", "rename_sheet", (("new_name",),), {"new_name": TEXT}),
    "delete_sheet": _spec("excel", "delete_sheet", (("name",),), {"name": TEXT}, destructive=True),
    "duplicate_sheet": _spec("excel", "duplicate_sheet", (("name",),), {"name": TEXT}),
    "move_sheet": _spec("excel", "move_sheet", (("name",),)),
    "protect_sheet": _spec("excel", "protect_sheet"),
    "unprotect_sheet": _spec("excel", "unprotect_sheet"),
    "protect_workbook": _spec("excel", "protect_workbook"),
    "unprotect_workbook": _spec("excel", "unprotect_workbook"),
    "create_table": _spec("excel", "create_table", (), {"rows": INTISH, "cols": INTISH, "start_cell": TEXT, "headers": LIST}),
    "write_cell": _spec("excel", "write_cell", (("cell",),)),
    "write_formula": _spec("excel", "write_formula", (("cell",), ("formula",)), {"formula": TEXT}),
    "write_range": _spec("excel", "write_range", (("start_cell",), ("values",)), {"values": LIST}),
    "read_cell": _spec("excel", "read_cell", (("cell",),)),
    "clear_range": _spec("excel", "clear_range", (("range",),), destructive=True),
    "clear_all": _spec("excel", "clear_all", destructive=True),
    "clear_format": _spec("excel", "clear_format", (("range",),)),
    "copy_range": _spec("excel", "copy_range", (("range",),)),
    "cut_range": _spec("excel", "cut_range", (("range",),), destructive=True),
    "paste_range": _spec("excel", "paste_range", (("cell",),)),
    "paste_values_only": _spec("excel", "paste_values_only", (("cell",),)),
    "set_bold": _spec("excel", "set_bold", (("range",),), {"bold": BOOLISH}),
    "set_italic": _spec("excel", "set_italic", (("range",),), {"italic": BOOLISH}),
    "set_underline": _spec("excel", "set_underline", (("range",),), {"underline": BOOLISH}),
    "set_strikethrough": _spec("excel", "set_strikethrough", (("range",),)),
    "set_font_size": _spec("excel", "set_font_size", (("range",), ("size",)), {"size": INTISH}),
    "set_font_name": _spec("excel", "set_font_name", (("range",), ("name",)), {"name": TEXT}),
    "set_font_color": _spec("excel", "set_font_color", (("range",), ("color",)), {"color": TEXT}),
    "set_bg_color": _spec("excel", "set_bg_color", (("range",), ("color",)), {"color": TEXT}),
    "set_border": _spec("excel", "set_border", (("range",),)),
    "remove_border": _spec("excel", "remove_border", (("range",),)),
    "set_alignment": _spec("excel", "set_alignment", (("range",),)),
    "set_vertical_alignment": _spec("excel", "set_vertical_alignment", (("range",),)),
    "set_wrap_text": _spec("excel", "set_wrap_text", (("range",),)),
    "set_number_format": _spec("excel", "set_number_format", (("range",), ("format",)), {"format": TEXT}),
    "merge_cells": _spec("excel", "merge_cells", (("range",),)),
    "unmerge_cells": _spec("excel", "unmerge_cells", (("range",),)),
    "insert_row": _spec("excel", "insert_row", (("row",),), {"row": INTISH}),
    "delete_row": _spec("excel", "delete_row", (("row",),), {"row": INTISH}, destructive=True),
    "insert_column": _spec("excel", "insert_column", (("column",),)),
    "delete_column": _spec("excel", "delete_column", (("column",),), destructive=True),
    "set_row_height": _spec("excel", "set_row_height", (("row",), ("height",)), {"row": INTISH, "height": NUMERIC}),
    "set_column_width": _spec("excel", "set_column_width", (("column",), ("width",)), {"width": NUMERIC}),
    "autofit_columns": _spec("excel", "autofit_columns"),
    "autofit_rows": _spec("excel", "autofit_rows"),
    "freeze_panes": _spec("excel", "freeze_panes", (("cell",),)),
    "filter_range": _spec("excel", "filter_range", (("range",),)),
    "sort_range": _spec("excel", "sort_range", (("range",),)),
    "insert_comment": _spec("excel", "insert_comment", (("cell",),)),
    "delete_comment": _spec("excel", "delete_comment", (("cell",),), destructive=True),
    "insert_hyperlink": _spec("excel", "insert_hyperlink", (("cell",), ("url",))),
    "insert_image": _spec("excel", "insert_image", (("path",),), path_required=True),
}


WORD_ACTIONS = {
    "create_document": _spec("word", "create_document"),
    "open_document": _spec("word", "open_document", (("path", "file_path", "filename"),), path_required=True),
    "save_document": _spec("word", "save_document"),
    "save_document_as": _spec("word", "save_document_as", (("filename", "path", "file_path", "output_path"),)),
    "close_document": _spec("word", "close_document"),
    "add_paragraph": _spec("word", "add_paragraph", (("text",),), {"text": TEXT}),
    "add_heading": _spec("word", "add_heading", (("text",),), {"text": TEXT, "level": INTISH}),
    "add_table": _spec("word", "add_table", (), {"rows": INTISH, "cols": INTISH}),
    "delete_table": _spec("word", "delete_table", destructive=True),
    "add_table_row": _spec("word", "add_table_row"),
    "add_table_column": _spec("word", "add_table_column"),
    "set_table_style": _spec("word", "set_table_style"),
    "add_bullet_list": _spec("word", "add_bullet_list"),
    "add_numbered_list": _spec("word", "add_numbered_list"),
    "continue_list": _spec("word", "continue_list", (("text",),), {"text": TEXT}),
    "set_bold": _spec("word", "set_bold"),
    "set_italic": _spec("word", "set_italic"),
    "set_underline": _spec("word", "set_underline"),
    "remove_underline": _spec("word", "remove_underline"),
    "set_strikethrough": _spec("word", "set_strikethrough"),
    "remove_strikethrough": _spec("word", "remove_strikethrough"),
    "set_font_size": _spec("word", "set_font_size", (("size",),), {"size": INTISH}),
    "set_font_name": _spec("word", "set_font_name", (("name",),), {"name": TEXT}),
    "set_font_color": _spec("word", "set_font_color", (("color",),), {"color": TEXT}),
    "set_highlight": _spec("word", "set_highlight"),
    "remove_highlight": _spec("word", "remove_highlight"),
    "set_alignment": _spec("word", "set_alignment", (("alignment",),)),
    "set_line_spacing": _spec("word", "set_line_spacing", (("spacing",),), {"spacing": NUMERIC}),
    "insert_page_break": _spec("word", "insert_page_break"),
    "set_margins": _spec("word", "set_margins"),
    "find_replace": _spec("word", "find_replace", (("find_text",), ("replace_text",))),
    "find_text": _spec("word", "find_text", (("text",),)),
    "insert_image": _spec("word", "insert_image", (("path",),), path_required=True),
    "insert_hyperlink": _spec("word", "insert_hyperlink", (("url",),)),
    "compare_documents": _spec("word", "compare_documents", (("path", "compare_path"),), path_required=True),
}


POWERPOINT_ACTIONS = {
    "create_presentation": _spec("powerpoint", "create_presentation"),
    "open_presentation": _spec("powerpoint", "open_presentation", (("path", "file_path", "filename"),), path_required=True),
    "save_presentation": _spec("powerpoint", "save_presentation"),
    "save_presentation_as": _spec("powerpoint", "save_presentation_as", (("filename", "path", "file_path", "output_path"),)),
    "close_presentation": _spec("powerpoint", "close_presentation"),
    "add_slide": _spec("powerpoint", "add_slide"),
    "delete_slide": _spec("powerpoint", "delete_slide", (("slide_index",),), {"slide_index": INTISH}, destructive=True),
    "duplicate_slide": _spec("powerpoint", "duplicate_slide", (("slide_index",),), {"slide_index": INTISH}),
    "reorder_slide": _spec("powerpoint", "reorder_slide", (("from_index",), ("to_index",)), {"from_index": INTISH, "to_index": INTISH}),
    "hide_slide": _spec("powerpoint", "hide_slide", (("slide_index",),), {"slide_index": INTISH}),
    "show_slide": _spec("powerpoint", "show_slide", (("slide_index",),), {"slide_index": INTISH}),
    "change_layout": _spec("powerpoint", "change_layout", (("slide_index",),), {"slide_index": INTISH}),
    "set_slide_text": _spec("powerpoint", "set_slide_text", (("text",),), {"slide_index": INTISH, "text": TEXT}),
    "clear_slide_text": _spec("powerpoint", "clear_slide_text", (("slide_index",),), {"slide_index": INTISH}, destructive=True),
    "add_bullet_point": _spec("powerpoint", "add_bullet_point", (("text",),), {"slide_index": INTISH, "text": TEXT}),
    "add_numbered_point": _spec("powerpoint", "add_numbered_point", (("text",),), {"slide_index": INTISH, "text": TEXT}),
    "set_speaker_notes": _spec("powerpoint", "set_speaker_notes", (("text",),), {"slide_index": INTISH, "text": TEXT}),
    "set_font_size": _spec("powerpoint", "set_font_size", (("size",),), {"slide_index": INTISH, "size": INTISH}),
    "set_font_name": _spec("powerpoint", "set_font_name", (("name",),), {"slide_index": INTISH, "name": TEXT}),
    "set_font_color": _spec("powerpoint", "set_font_color", (("color",),), {"slide_index": INTISH, "color": TEXT}),
    "set_bold": _spec("powerpoint", "set_bold", (), {"slide_index": INTISH, "bold": BOOLISH}),
    "set_italic": _spec("powerpoint", "set_italic", (), {"slide_index": INTISH, "italic": BOOLISH}),
    "set_bg_color": _spec("powerpoint", "set_bg_color", (("color",),), {"slide_index": INTISH, "color": TEXT}),
    "set_bg_image": _spec("powerpoint", "set_bg_image", (("path",),), path_required=True),
    "insert_image": _spec("powerpoint", "insert_image", (("path",),), path_required=True),
    "insert_text_box": _spec("powerpoint", "insert_text_box", (("text",),), {"slide_index": INTISH, "text": TEXT}),
    "insert_table": _spec("powerpoint", "insert_table", (), {"slide_index": INTISH, "rows": INTISH, "cols": INTISH}),
    "insert_chart": _spec("powerpoint", "insert_chart"),
    "insert_shape": _spec("powerpoint", "insert_shape"),
    "add_logo": _spec("powerpoint", "add_logo", (("path",),), path_required=True),
    "insert_video": _spec("powerpoint", "insert_video", (("path",),), path_required=True),
    "insert_audio": _spec("powerpoint", "insert_audio", (("path",),), path_required=True),
    "insert_hyperlink": _spec("powerpoint", "insert_hyperlink", (("url",),)),
    "set_slide_size": _spec("powerpoint", "set_slide_size", (("width",), ("height",)), {"width": NUMERIC, "height": NUMERIC}),
    "set_transition": _spec("powerpoint", "set_transition"),
    "set_auto_advance": _spec("powerpoint", "set_auto_advance"),
    "set_animation": _spec("powerpoint", "set_animation"),
}


ACTION_REGISTRY = {
    "excel": EXCEL_ACTIONS,
    "word": WORD_ACTIONS,
    "powerpoint": POWERPOINT_ACTIONS,
}


def canonical_app(app_type):
    app = (app_type or "").strip().lower()
    return "powerpoint" if app == "ppt" else app


def get_action_spec(app_type, action_name) -> ActionSpec | None:
    app = canonical_app(app_type)
    return ACTION_REGISTRY.get(app, {}).get((action_name or "").strip())


def get_known_actions(app_type):
    return set(ACTION_REGISTRY.get(canonical_app(app_type), {}).keys())


def is_destructive_action(app_type, action_name):
    spec = get_action_spec(app_type, action_name)
    return bool(spec and spec.destructive)


def registry_as_prompt_lines(app_type):
    lines = []
    for name, spec in sorted(ACTION_REGISTRY.get(canonical_app(app_type), {}).items()):
        required = [" or ".join(group) for group in spec.required_groups]
        lines.append(f"- {name}: required={required or []}")
    return lines
