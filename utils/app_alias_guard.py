import re
from pathlib import Path


OFFICE_TARGET_TERMS = {
    "excel",
    "spreadsheet",
    "workbook",
    "worksheet",
    "sheet",
    "xlsx",
    "xlsm",
    "xls",
    "csv",
    "word",
    "document",
    "docx",
    "powerpoint",
    "power point",
    "ppt",
    "pptx",
    "presentation",
    "slides",
    "slide",
    "deck",
}

DOCUMENT_ACTION_TERMS = {
    "create",
    "make",
    "generate",
    "build",
    "new",
    "add",
    "insert",
    "write",
    "edit",
    "update",
    "modify",
    "format",
    "bold",
    "italic",
    "underline",
    "color",
    "background",
    "border",
    "formula",
    "bullet",
    "table",
    "chart",
    "row",
    "column",
    "cell",
    "paragraph",
    "heading",
    "title",
    "save",
}

DOCUMENT_EXTENSIONS = {".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".csv", ".ppt", ".pptx", ".pdf"}
MAX_ALIAS_LENGTH = 60


def _contains_term(text, term):
    if " " in term:
        return term in text
    return bool(re.search(rf"\b{re.escape(term)}\b", text))


def looks_like_document_command(alias):
    text = re.sub(r"\s+", " ", (alias or "").strip().lower())
    if not text:
        return False
    if len(text) > MAX_ALIAS_LENGTH:
        return True
    if text.startswith("agent:"):
        return True
    if re.search(r"[a-z]:[\\/]", text) or "\\" in text or "/" in text:
        return True

    suffix = Path(text.strip("\"'")).suffix.lower()
    if suffix in DOCUMENT_EXTENSIONS:
        return True
    if any(ext in text for ext in DOCUMENT_EXTENSIONS):
        return True

    has_office_target = any(_contains_term(text, term) for term in OFFICE_TARGET_TERMS)
    has_document_action = any(_contains_term(text, term) for term in DOCUMENT_ACTION_TERMS)
    return has_office_target and has_document_action


def validate_manual_app_alias(alias):
    if looks_like_document_command(alias):
        return (
            False,
            "MANUAL_APP_ALIAS_REJECTED_DOCUMENT_COMMAND",
            "Document or Office automation commands cannot be saved as app aliases.",
        )
    return True, "", ""
