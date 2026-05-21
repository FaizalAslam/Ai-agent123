from pathlib import Path
from utils.file_paths import OFFICE_OUTPUT_ROOT

BASE_DIR = Path(__file__).resolve().parent.parent
OFFICE_OUTPUT_DIR = OFFICE_OUTPUT_ROOT

OFFICE_APPS = {"excel", "word", "powerpoint", "ppt"}

OFFICE_OUTPUTS = {
    "excel":      str(OFFICE_OUTPUT_DIR / "excel"      / "output.xlsx"),
    "word":       str(OFFICE_OUTPUT_DIR / "word"       / "output.docx"),
    "powerpoint": str(OFFICE_OUTPUT_DIR / "powerpoint" / "output.pptx"),
    "ppt":        str(OFFICE_OUTPUT_DIR / "powerpoint" / "output.pptx"),
}

OFFICE_DEPENDENCIES = {
    "excel":      ("openpyxl", "openpyxl"),
    "word":       ("docx",     "python-docx"),
    "powerpoint": ("pptx",     "python-pptx"),
    "ppt":        ("pptx",     "python-pptx"),
}

OFFICE_TARGET_KEYWORDS = {
    "excel": (
        "excel", "spreadsheet", "workbook", "worksheet", "sheet",
        "xlsx", "xlsm", "xls", "csv",
    ),
    "word": (
        "word", "document", "docx", "doc",
    ),
    "powerpoint": (
        "powerpoint", "power point", "ppt", "pptx", "presentation",
        "slide deck", "slides", "slide", "deck",
    ),
}

OFFICE_ACTION_KEYWORDS = (
    "create", "make", "generate", "build", "new", "open", "save as", "add",
    "insert", "edit", "update", "modify", "write", "format", "table", "chart",
    "row", "column", "cell", "slide", "paragraph", "heading", "workbook",
    "worksheet", "spreadsheet", "document", "presentation", "bold", "italic",
    "color", "underline", "background", "border", "formula", "bullet", "title",
    "save", "close", "deck",
)

APP_LAUNCH_PREFIXES = ("open ", "launch ", "start ", "run ", "boot ")
