import ast
import logging
import re
from dataclasses import dataclass, field

from parser.command_parser import parse_command
from utils.office_actions import OfficeActionError, normalize_actions, validate_actions


logger = logging.getLogger("OfficeAgent")

ACTION_STARTERS = (
    "create", "make", "generate", "build", "new", "start", "add", "insert",
    "write", "set", "apply", "make", "bold", "italic", "rename", "save", "open",
    "load", "format", "delete", "duplicate", "move", "protect", "unprotect",
    "unlock", "close", "underline", "border", "bullet", "autofit", "on slide",
)


@dataclass
class PlanClause:
    index: int
    text: str
    intent: str = "unknown"
    actions: list[dict] = field(default_factory=list)
    status: str = "failed"
    warnings: list[str] = field(default_factory=list)
    reason: str = ""

    def to_dict(self):
        return {
            "index": self.index,
            "text": self.text,
            "intent": self.intent,
            "actions": self.actions,
            "status": self.status,
            "warnings": self.warnings,
            "reason": self.reason,
        }


@dataclass
class CommandPlan:
    success: bool
    app: str
    raw_command: str
    clauses: list[PlanClause]
    actions: list[dict]
    context: dict = field(default_factory=dict)
    requires_api: bool = False
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    source: str = "planner"

    @property
    def partial_success(self):
        return bool(self.actions and self.errors)

    def to_dict(self):
        return {
            "success": self.success,
            "partial_success": self.partial_success,
            "app": self.app,
            "raw_command": self.raw_command,
            "clauses": [clause.to_dict() for clause in self.clauses],
            "failed_clauses": [
                clause.to_dict()
                for clause in self.clauses
                if clause.status == "failed"
            ],
            "actions": self.actions,
            "context": self.context,
            "requires_api": self.requires_api,
            "warnings": self.warnings,
            "errors": self.errors,
            "source": self.source,
        }


class PlanningContext:
    def __init__(self, app):
        self.app = "powerpoint" if app == "ppt" else app
        self.current_workbook = ""
        self.current_sheet = ""
        self.current_document = ""
        self.current_presentation = ""
        self.last_table_range = ""
        self.header_range = ""
        self.data_range = ""
        self.columns_range = ""   # e.g. "A:C" — column span of the active table
        self.table_start_cell = ""
        self.table_rows: int = 0
        self.table_cols: int = 0
        self.used_range = ""
        self.last_range = ""
        self.last_cell = ""
        self.last_formula_cell = ""
        self.last_output_filename = ""
        self.last_save_path = ""
        self.last_heading = ""
        self.last_paragraph = ""
        self.last_table_index = 0
        self.current_section = ""
        self.last_slide_index = 0
        self.slide_count = 0
        self.last_title_placeholder = ""
        self.last_body_placeholder = ""
        self.last_created_slide = 0
        self.created = False

    def to_dict(self):
        return {
            "app": self.app,
            "current_workbook": self.current_workbook,
            "current_sheet": self.current_sheet,
            "current_document": self.current_document,
            "current_presentation": self.current_presentation,
            "last_table_range": self.last_table_range,
            "header_range": self.header_range,
            "data_range": self.data_range,
            "columns_range": self.columns_range,
            "table_start_cell": self.table_start_cell,
            "table_rows": self.table_rows,
            "table_cols": self.table_cols,
            "used_range": self.used_range,
            "last_range": self.last_range,
            "last_cell": self.last_cell,
            "last_formula_cell": self.last_formula_cell,
            "last_output_filename": self.last_output_filename,
            "last_save_path": self.last_save_path,
            "last_heading": self.last_heading,
            "last_paragraph": self.last_paragraph,
            "last_table_index": self.last_table_index,
            "current_section": self.current_section,
            "last_slide_index": self.last_slide_index,
            "slide_count": self.slide_count,
            "last_title_placeholder": self.last_title_placeholder,
            "last_body_placeholder": self.last_body_placeholder,
            "last_created_slide": self.last_created_slide,
            "created": self.created,
        }


def _column_to_index(col):
    n = 0
    for ch in (col or "").upper():
        if "A" <= ch <= "Z":
            n = (n * 26) + (ord(ch) - ord("A") + 1)
    return n


def _index_to_column(n):
    n = max(1, int(n or 1))
    out = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def _cell_parts(cell):
    match = re.match(r"^([A-Z]{1,3})(\d{1,7})$", (cell or "").upper())
    if not match:
        return "A", 1
    return match.group(1), int(match.group(2))


def _table_ranges(start_cell, rows, cols):
    col, row = _cell_parts(start_cell or "A1")
    last_col = _index_to_column(_column_to_index(col) + max(1, int(cols or 1)) - 1)
    last_row = row + max(1, int(rows or 1)) - 1
    return f"{col}{row}:{last_col}{last_row}", f"{col}{row}:{last_col}{row}"


def _data_range_from_table(table_range):
    match = re.match(r"^([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})$", str(table_range or "").upper())
    if not match:
        return ""
    start_col, start_row, end_col, end_row = match.group(1), int(match.group(2)), match.group(3), int(match.group(4))
    if end_row <= start_row:
        return table_range
    return f"{start_col}{start_row + 1}:{end_col}{end_row}"


def _columns_from_range(range_ref):
    match = re.match(r"^([A-Z]{1,3})\d{1,7}:([A-Z]{1,3})\d{1,7}$", str(range_ref or "").upper())
    if match:
        return f"{match.group(1)}:{match.group(2)}"
    match = re.match(r"^([A-Z]{1,3}):([A-Z]{1,3})$", str(range_ref or "").upper())
    if match:
        return f"{match.group(1)}:{match.group(2)}"
    return ""


def _contextual_excel_range(clause, ctx, default="A1"):
    low = (clause or "").lower()
    explicit = _extract_range(clause)
    if explicit:
        return explicit
    if "header" in low and ctx.header_range:
        return ctx.header_range
    if "data" in low and ctx.data_range:
        return ctx.data_range
    if any(term in low for term in ("full table", "entire table", "whole table", "table")) and ctx.last_table_range:
        return ctx.last_table_range
    if any(term in low for term in ("column width", "column widths", "autofit", "auto-fit", "auto fit")) and ctx.columns_range:
        return ctx.columns_range
    if "used range" in low and ctx.used_range:
        return ctx.used_range
    return ctx.last_range or ctx.last_cell or default


def _extract_rows_cols(text):
    low = text.lower()
    row_match = re.search(r"(\d+)\s*(?:row|rows)", low)
    col_match = re.search(r"(\d+)\s*(?:col|cols|column|columns)", low)
    rows = int(row_match.group(1)) if row_match else None
    cols = int(col_match.group(1)) if col_match else None
    by_match = re.search(r"(\d+)\s*[xX]\s*(\d+)", text)
    if by_match:
        rows = int(by_match.group(1))
        cols = int(by_match.group(2))
    return rows, cols


def _extract_cell(text):
    match = re.search(r"\b([A-Z]{1,3}\d{1,7})\b", text.upper())
    return match.group(1) if match else ""


def _extract_range(text):
    match = re.search(r"\b([A-Z]{1,3}\d{1,7})\s*(?:\:|to|-)\s*([A-Z]{1,3}\d{1,7})\b", text.upper())
    if match:
        return f"{match.group(1)}:{match.group(2)}"
    col_match = re.search(r"\b([A-Z]{1,3})\s*:\s*([A-Z]{1,3})\b", text.upper())
    if col_match:
        return f"{col_match.group(1)}:{col_match.group(2)}"
    return _extract_cell(text)


def _extract_headers(clause):
    """Extract explicit column headers from phrases like 'use headers Name, Amount, and Status'."""
    patterns = [
        r"(?:use|with|having|set)\s+headers?\s*(?:of\s+)?([^\n.;]+)",
        r"headers?\s+(?:are|is|:)\s*([^\n.;]+)",
        r"columns?\s+(?:named|called)\s+([^\n.;]+)",
    ]
    for pat in patterns:
        m = re.search(pat, clause, re.IGNORECASE)
        if m:
            raw = m.group(1).strip().rstrip(".")
            # Split on comma, "and", semicolon; strip quotes
            parts = re.split(r"\s*(?:,\s*(?:and\s+)?|and\s+|;\s*)", raw, flags=re.IGNORECASE)
            cleaned = [p.strip().strip("\"'") for p in parts if p.strip()]
            if cleaned:
                return cleaned
    return []


def _extract_color(text):
    colors = (
        "dark red", "dark blue", "dark green", "light blue", "light gray",
        "yellow", "green", "blue", "red", "orange", "purple", "pink",
        "black", "white", "gray", "grey", "teal", "cyan", "magenta",
        "gold", "brown", "navy",
    )
    low = text.lower()
    for color in colors:
        if re.search(rf"\b{re.escape(color)}\b", low):
            return color
    hex_match = re.search(r"#?[0-9A-Fa-f]{6,8}", text)
    return hex_match.group(0) if hex_match else ""


def _parse_literal_value(raw):
    value = (raw or "").strip().strip("\"'")
    if not value:
        return ""
    try:
        parsed = ast.literal_eval(value)
        return parsed
    except Exception:
        pass
    if re.fullmatch(r"-?\d+", value):
        return int(value)
    if re.fullmatch(r"-?\d+\.\d+", value):
        return float(value)
    return value


def _excel_formula_literal(value):
    raw = str(value or "").strip().strip("\"'")
    if not raw:
        return '""'
    if re.fullmatch(r"-?\d+(?:\.\d+)?", raw):
        return raw
    if raw.startswith((">", "<", "=")):
        return '"' + raw.replace('"', '""') + '"'
    return '"' + raw.replace('"', '""') + '"'


def _delimiter_literal(clause):
    low = clause.lower()
    if "tab" in low:
        return '"\\t"'
    if "semicolon" in low:
        return '";"'
    if "pipe" in low:
        return '"|"'
    if "space" in low:
        return '" "'
    return '","'


def _extract_sumifs_formula(clause):
    range_pat = r"(?:[A-Z]{1,3}:[A-Z]{1,3}|[A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7})"
    ranges = re.findall(range_pat, clause, flags=re.IGNORECASE)
    sum_range = ranges[0].upper().replace(" ", "") if ranges else "C:C"
    criteria_pairs = []
    for match in re.finditer(
        rf"({range_pat})\s*(?:is|=|equals|equal to|matching|for)\s*['\"]?([^,'\"\n]+?)['\"]?(?=\s+(?:and|,|in|into|to)\b|$)",
        clause,
        flags=re.IGNORECASE,
    ):
        criteria_range = match.group(1).upper().replace(" ", "")
        if criteria_range == sum_range and not criteria_pairs:
            continue
        criteria_pairs.append((criteria_range, match.group(2).strip(" .")))
    if not criteria_pairs and len(ranges) > 1:
        criteria_pairs.append((ranges[1].upper().replace(" ", ""), ">0"))
    if not criteria_pairs:
        criteria_pairs.append(("A:A", ">0"))
    parts = [sum_range]
    for criteria_range, criteria in criteria_pairs[:5]:
        parts.extend([criteria_range, _excel_formula_literal(criteria)])
    return f"=SUMIFS({','.join(parts)})"


def _extract_password(clause):
    match = re.search(r"\bpassword\s+(?:is\s+)?['\"]?([^\s'\"]+)['\"]?", clause, re.IGNORECASE)
    return match.group(1) if match else ""


def _extract_sheet_name(clause):
    if re.search(r"\b(?:active|current)\s+sheet\b", clause, re.IGNORECASE):
        return ""
    match = re.search(
        r"\bsheet\s+(?:named|called|name)\s*['\"]?([A-Za-z0-9 _-]+?)['\"]?(?:\s+with|\s+password|$)",
        clause,
        re.IGNORECASE,
    )
    if not match:
        return ""
    candidate = match.group(1).strip()
    return "" if candidate.lower() in {"active", "current"} else candidate


def _extract_save_as_filename(clause, ext):
    match = re.search(r"\bsave\b.*?\bas\s+['\"]?(.+?)['\"]?$", clause, re.IGNORECASE)
    if not match:
        return ""
    name = match.group(1).strip().strip("\"'")
    if not re.search(rf"\.{re.escape(ext)}$", name, flags=re.IGNORECASE):
        name = f"{name}.{ext}"
    return name


def _protect(text):
    placeholders = {}
    out = []
    i = 0
    token = 0
    while i < len(text):
        ch = text[i]
        if ch in ("\"", "'"):
            quote = ch
            j = i + 1
            while j < len(text):
                if text[j] == quote:
                    j += 1
                    break
                j += 1
            key = f"__P{token}__"
            placeholders[key] = text[i:j]
            out.append(key)
            token += 1
            i = j
            continue
        if ch == "[":
            depth = 1
            j = i + 1
            while j < len(text) and depth:
                if text[j] == "[":
                    depth += 1
                elif text[j] == "]":
                    depth -= 1
                j += 1
            key = f"__P{token}__"
            placeholders[key] = text[i:j]
            out.append(key)
            token += 1
            i = j
            continue
        out.append(ch)
        i += 1
    protected = "".join(out)
    protected = re.sub(
        r"profit\s+and\s+loss",
        lambda m: re.sub(r"\s+and\s+", " __AND__ ", m.group(0), flags=re.IGNORECASE),
        protected,
        flags=re.IGNORECASE,
    )
    protected = re.sub(
        r"sales\s+and\s+marketing",
        lambda m: re.sub(r"\s+and\s+", " __AND__ ", m.group(0), flags=re.IGNORECASE),
        protected,
        flags=re.IGNORECASE,
    )
    safe_patterns = [
        (r"(\d+\s*(?:col|cols|column|columns)\s+)and(\s*\d+\s*(?:row|rows))", r"\1__AND__\2"),
        (r"(\d+\s*(?:row|rows)\s+)and(\s*\d+\s*(?:col|cols|column|columns))", r"\1__AND__\2"),
        (r"([A-Z]{1,3}\d{1,7}\s+)and(\s+[A-Z]{1,3}\d{1,7})", r"\1__AND__\2"),
    ]
    for pattern, repl in safe_patterns:
        protected = re.sub(pattern, repl, protected, flags=re.IGNORECASE)
    return protected, placeholders


def _restore(text, placeholders):
    restored = text.replace("__AND__", "and")
    for key, value in placeholders.items():
        restored = restored.replace(key, value)
    return restored.strip(" .,:;")


def split_command_clauses(raw_command):
    text = (raw_command or "").strip()
    if not text:
        return []
    text = re.sub(r"^\s*(?:[-*]\s+|\d+[\.)]\s*)", "", text, flags=re.MULTILINE)
    protected, placeholders = _protect(text)
    protected = re.sub(r"[\r\n]+", ";", protected)
    protected = re.sub(
        r"\s+\d+[\.)]\s+(?=(?:create|make|add|insert|write|set|apply|bold|italic|underline|border|rename|save|open|protect|unprotect|unlock|close|autofit|on slide)\b)",
        ";",
        protected,
        flags=re.IGNORECASE,
    )
    protected = re.sub(
        r"(?<=[A-Za-z0-9])\.\s+(?=(?:create|make|add|insert|write|set|apply|bold|italic|underline|border|rename|save|open|protect|unprotect|unlock|close|autofit|on slide)\b)",
        ";",
        protected,
        flags=re.IGNORECASE,
    )
    protected = re.sub(r"\s*;\s*", ";", protected)
    pieces = []
    for part in re.split(r";+", protected):
        part = part.strip()
        if not part:
            continue
        subparts = re.split(r"\s+(?:then|next|also|after that|finally)\s+", part, flags=re.IGNORECASE)
        for sub in subparts:
            sub = sub.strip()
            if not sub:
                continue
            # Split commas and "and" only before a new action verb.
            comma_parts = re.split(
                r",\s+(?=(?:create|make|add|insert|write|set|apply|bold|italic|underline|border|rename|save|open|protect|unprotect|unlock|close|autofit|on slide)\b)",
                sub,
                flags=re.IGNORECASE,
            )
            for comma_part in comma_parts:
                and_parts = re.split(
                    r"\s+and\s+(?=(?:create|make|add|insert|write|set|apply|bold|italic|underline|border|rename|save|open|protect|unprotect|unlock|close|autofit|on slide)\b)",
                    comma_part,
                    flags=re.IGNORECASE,
                )
                pieces.extend(p for p in and_parts if p.strip())
    clauses = [_restore(piece, placeholders) for piece in pieces]
    return [clause for clause in clauses if clause]


def _detect_intent(app, clause):
    low = clause.lower()
    if "open" in low or "load" in low:
        return f"open_{app}"
    if "bold" in low or "italic" in low or "underline" in low or "border" in low:
        return "format"
    if "background" in low or "color" in low or "align" in low or "format" in low:
        return "format"
    if any(w in low for w in ("create", "make", "generate", "build", "new", "start")):
        if app == "excel":
            return "create_workbook" if "table" not in low else "create_table"
        if app == "word":
            return "create_document"
        return "create_presentation"
    if "save" in low:
        return f"save_{app}"
    if "write" in low or "add" in low or "set" in low:
        return "content"
    return "unknown"


def _open_action(app, clause):
    low = (clause or "").lower()
    if not any(term in low for term in ("open", "load", "import")):
        return []
    if "save" in low and " as " in low:
        return []
    ext = {"excel": "xlsx", "word": "docx", "powerpoint": "pptx"}.get(app, "")
    if not ext:
        return []
    ext_group = {
        "excel": r"(?:xlsx|xlsm|xls|csv)",
        "word": r"(?:docx|doc)",
        "powerpoint": r"(?:pptx|ppt)",
    }.get(app, re.escape(ext))
    match = re.search(r"['\"]([^'\"]+\." + ext_group + r")['\"]", clause, re.IGNORECASE)
    if not match:
        match = re.search(r"([A-Za-z]:[\\/][^\"'\n]+?\." + ext_group + r")", clause, re.IGNORECASE)
    if not match:
        match = re.search(r"\b([^\s\"']+\." + ext_group + r")\b", clause, re.IGNORECASE)
    if not match:
        match = re.search(
            r"\b(?:named|called|file|document|workbook|presentation)\s+['\"]?([A-Za-z0-9_\- .]{1,120})['\"]?(?:\s+(?:on|from|in|at)\s+(?:desktop|documents|downloads))?$",
            clause,
            re.IGNORECASE,
        )
    if not match:
        return []
    path_value = match.group(1).strip()
    path_value = re.sub(r"^(?:named|called)\s+", "", path_value, flags=re.IGNORECASE).strip()
    path_value = re.sub(r"\s+(?:on|from|in|at)\s+(?:desktop|documents|downloads)$", "", path_value, flags=re.IGNORECASE).strip()
    return [{"action": {
        "excel": "open_workbook",
        "word": "open_document",
        "powerpoint": "open_presentation",
    }[app], "path": path_value}]


def _save_action(app):
    return {
        "excel": {"action": "save_workbook"},
        "word": {"action": "save_document"},
        "powerpoint": {"action": "save_presentation"},
    }.get(app)


def _planner_excel_actions(clause, ctx):
    low = clause.lower()
    actions = []

    actions.extend(_open_action("excel", clause))
    if actions:
        if actions[0].get("path"):
            ctx.current_workbook = actions[0]["path"]
        return actions

    if "save" in low and not any(w in low for w in ("save as", "saved as")):
        ctx.last_save_path = ctx.last_save_path or ctx.current_workbook
        return [_save_action("excel")]

    if "concatenate" in low:
        cells = re.findall(r"\b[A-Z]{1,3}\d{1,7}\b", clause.upper())
        if len(cells) >= 3:
            ctx.last_formula_cell = cells[-1]
            ctx.last_cell = cells[-1]
            ctx.last_range = cells[-1]
            return [{"action": "write_formula", "cell": cells[-1], "formula": f"=CONCATENATE({cells[0]},{cells[1]})"}]

    if "textjoin" in low or "text join" in low:
        rng = _extract_range(clause) or "A1:A5"
        result_cell = re.findall(r"\b[A-Z]{1,3}\d{1,7}\b", clause.upper())
        result = result_cell[-1] if result_cell else "B1"
        ctx.last_formula_cell = result
        ctx.last_cell = result
        ctx.last_range = result
        return [{"action": "write_formula", "cell": result, "formula": f"=TEXTJOIN({_delimiter_literal(clause)},TRUE,{rng})"}]

    if "sumifs" in low or "sum ifs" in low:
        result_cell = re.findall(r"\b[A-Z]{1,3}\d{1,7}\b", clause.upper())
        result = result_cell[-1] if result_cell else "D1"
        ctx.last_formula_cell = result
        ctx.last_cell = result
        ctx.last_range = result
        return [{"action": "write_formula", "cell": result, "formula": _extract_sumifs_formula(clause)}]

    if "today" in low or "today's date" in low or "todays date" in low or "current date" in low:
        cell = _extract_cell(clause) or "A1"
        ctx.last_formula_cell = cell
        ctx.last_cell = cell
        ctx.last_range = cell
        return [{"action": "write_formula", "cell": cell, "formula": "=TODAY()"}]

    if "current time" in low or "current datetime" in low or "timestamp" in low or re.search(r"\bnow\b", low):
        cell = _extract_cell(clause) or "A1"
        ctx.last_formula_cell = cell
        ctx.last_cell = cell
        ctx.last_range = cell
        return [{"action": "write_formula", "cell": cell, "formula": "=NOW()"}]

    if "format" in low and "time" in low:
        fmt = "hh:mm:ss" if "hh:mm:ss" in low else "hh:mm"
        target = _extract_cell(clause) or "A1"
        return [{"action": "set_number_format", "range": target, "format": fmt}]

    if "unprotect workbook" in low or "unlock workbook" in low:
        return [{"action": "unprotect_workbook", "password": _extract_password(clause)}]

    if "protect workbook" in low:
        return [{"action": "protect_workbook", "password": _extract_password(clause)}]

    if "unprotect sheet" in low or "unlock sheet" in low or "remove sheet password" in low:
        return [{"action": "unprotect_sheet", "sheet_name": _extract_sheet_name(clause), "password": _extract_password(clause)}]

    if "protect sheet" in low or "lock sheet" in low:
        return [{"action": "protect_sheet", "sheet_name": _extract_sheet_name(clause), "password": _extract_password(clause)}]

    # "use headers Name, Amount, Status" as a stand-alone clause — stores headers in context
    # so subsequent create_table or an update action can use them.
    standalone_headers = _extract_headers(clause)
    if standalone_headers and "table" not in low and not any(
        w in low for w in ("create", "make", "add", "insert")
    ):
        ctx._pending_headers = standalone_headers
        return []   # no action emitted; headers stored for next create_table clause

    if any(w in low for w in ("create", "make", "generate", "build", "new", "start")) and any(
        term in low for term in ("excel", "spreadsheet", "workbook", "worksheet", "sheet", "file")
    ):
        actions.append({"action": "create_workbook"})
        ctx.created = True
        ctx.current_workbook = ctx.current_workbook or "<new>"

    if "table" in low and any(w in low for w in ("create", "make", "insert", "add", "build", "include")):
        rows, cols = _extract_rows_cols(clause)
        start = _extract_cell(clause) or "A1"
        rows = rows or 5
        cols = cols or 3
        # Accept headers from this clause or from a preceding "use headers …" clause.
        headers = _extract_headers(clause) or getattr(ctx, "_pending_headers", [])
        if hasattr(ctx, "_pending_headers"):
            del ctx._pending_headers
        if headers:
            cols = max(cols, len(headers))
        tbl_action = {"action": "create_table", "rows": rows, "cols": cols, "start_cell": start}
        if headers:
            tbl_action["headers"] = headers
        actions.append(tbl_action)
        ctx.last_table_range, ctx.header_range = _table_ranges(start, rows, cols)
        ctx.data_range = _data_range_from_table(ctx.last_table_range)
        ctx.columns_range = _columns_from_range(ctx.last_table_range)
        ctx.table_start_cell = start
        ctx.table_rows = rows
        ctx.table_cols = cols
        ctx.last_range = ctx.last_table_range
        ctx.used_range = ctx.last_table_range

    rename = re.search(r"rename\s+(?:the\s+)?sheet\s+(?:to|as)\s+(.+)$", clause, re.IGNORECASE)
    if rename:
        name = rename.group(1).strip().strip("\"'")
        actions.append({"action": "rename_sheet", "new_name": name})

    write_match = re.search(
        r"\b(?:write|put|enter|type)\s+(.+?)\s+(?:in|into|at|to)\s+(?:cell\s*)?([A-Z]{1,3}\d{1,7})\b",
        clause,
        re.IGNORECASE,
    )
    if write_match and "[" not in write_match.group(1):
        value = _parse_literal_value(write_match.group(1))
        cell = write_match.group(2).upper()
        actions.append({"action": "write_cell", "cell": cell, "value": value})
        ctx.last_cell = cell
        ctx.last_range = cell
        ctx.used_range = cell

    formula = re.search(
        r"(?:formula\s+in|write\s+formula\s+in)\s+(?:cell\s*)?([A-Z]{1,3}\d{1,7}).*?(?:total|sum).*?(?:from|of)\s+([A-Z]{1,3}\d{1,7})\s*(?:to|-)\s*([A-Z]{1,3}\d{1,7})",
        clause,
        re.IGNORECASE,
    )
    if formula:
        formula_cell = formula.group(1).upper()
        actions.append({
            "action": "write_formula",
            "cell": formula_cell,
            "formula": f"=SUM({formula.group(2).upper()}:{formula.group(3).upper()})",
        })
        ctx.last_formula_cell = formula_cell
        ctx.last_cell = formula_cell
        ctx.last_range = formula_cell

    target_range = _contextual_excel_range(clause, ctx)
    if "bold" in low:
        actions.append({"action": "set_bold", "range": target_range, "bold": True})
    if "italic" in low:
        actions.append({"action": "set_italic", "range": target_range, "italic": True})
    if "underline" in low:
        actions.append({"action": "set_underline", "range": target_range, "underline": True})
    if "background" in low:
        color = _extract_color(clause) or "yellow"
        if "cells" in low and re.search(r"\b[A-Z]{1,3}\d{1,7}\b.*\band\b.*\b[A-Z]{1,3}\d{1,7}\b", clause, re.IGNORECASE):
            for cell in re.findall(r"\b[A-Z]{1,3}\d{1,7}\b", clause.upper()):
                actions.append({"action": "set_bg_color", "range": cell, "color": color})
        else:
            actions.append({"action": "set_bg_color", "range": target_range, "color": color})
    if "font color" in low:
        color = _extract_color(clause) or "black"
        actions.append({"action": "set_font_color", "range": target_range, "color": color})
    if "border" in low:
        actions.append({"action": "set_border", "range": target_range})
    if "wrap text" in low or "wrap" in low and "text" in low:
        actions.append({"action": "set_wrap_text", "range": target_range})
    if any(term in low for term in ("align", "alignment", "center", "centre", "left align", "right align")):
        alignment = "center" if ("center" in low or "centre" in low) else ("right" if "right" in low else "left")
        actions.append({"action": "set_alignment", "range": target_range, "alignment": alignment})
    if "autofit" in low or ("fit" in low and "column" in low):
        col_range = _contextual_excel_range(clause, ctx) if ctx.columns_range else None
        action = {"action": "autofit_columns"}
        if col_range:
            action["range"] = col_range
        actions.append(action)
    width_match = re.search(
        r"(?:column\s+([A-Z]{1,3})\s+width|width\s+of\s+column\s+([A-Z]{1,3}))\s*(?:to|as)?\s*(\d+(?:\.\d+)?)",
        clause,
        re.IGNORECASE,
    )
    if width_match:
        column = (width_match.group(1) or width_match.group(2)).upper()
        actions.append({"action": "set_column_width", "column": column, "width": float(width_match.group(3))})

    return _dedupe(actions)


def _extract_after_keyword(clause, keyword, default):
    pattern = rf"\b{keyword}\b\s+(.+)$"
    match = re.search(pattern, clause, re.IGNORECASE)
    if not match:
        return default
    value = match.group(1).strip().strip("\"'")
    value = re.split(r"\s+(?:and|then|finally)\b", value, maxsplit=1, flags=re.IGNORECASE)[0].strip(" .")
    return value or default


def _planner_word_actions(clause, ctx):
    low = clause.lower()
    actions = []
    save_as = _extract_save_as_filename(clause, "docx") if "save" in low and " as " in low else ""
    if save_as:
        ctx.last_save_path = save_as
        ctx.last_output_filename = save_as
        return [{"action": "save_document_as", "filename": save_as}]
    actions.extend(_open_action("word", clause))
    if actions:
        if actions[0].get("path"):
            ctx.current_document = actions[0]["path"]
        return actions
    if "save" in low and not any(w in low for w in ("save as", "saved as")):
        ctx.last_save_path = ctx.last_save_path or ctx.current_document
        return [_save_action("word")]
    if re.search(r"\bclose\s+(?:the\s+)?(?:current\s+)?(?:document|file)\b", low):
        return [{"action": "close_document"}]
    if any(w in low for w in ("create", "make", "generate", "build", "new", "start", "write")) and any(
        term in low for term in ("word", "document", "docx", "file")
    ):
        actions.append({"action": "create_document"})
        ctx.created = True
        ctx.current_document = ctx.current_document or "<new>"
        heading = re.search(r"\bheading\s+(.+?)(?:\s+and\s+paragraph|\s+paragraph|$)", clause, re.IGNORECASE)
        paragraph = re.search(r"\bparagraph\s+(.+)$", clause, re.IGNORECASE)
        if heading:
            heading_text = heading.group(1).strip(" .")
            actions.append({"action": "add_heading", "text": heading_text, "level": 1})
            ctx.last_heading = heading_text
        if paragraph:
            paragraph_text = paragraph.group(1).strip(" .")
            actions.append({"action": "add_paragraph", "text": paragraph_text})
            ctx.last_paragraph = paragraph_text
        about = re.search(r"\babout\s+(.+?)(?:\s+and\s+save|\s+save|$)", clause, re.IGNORECASE)
        if about:
            paragraph_text = about.group(1).strip(" .")
            actions.append({"action": "add_paragraph", "text": paragraph_text})
            ctx.last_paragraph = paragraph_text
    if "heading" in low and any(w in low for w in ("add", "create", "insert", "write")):
        heading_text = _extract_after_keyword(clause, "heading", "Heading")
        actions.append({"action": "add_heading", "text": heading_text, "level": 1})
        ctx.last_heading = heading_text
    if "paragraph" in low and any(w in low for w in ("add", "create", "insert", "write")):
        paragraph_text = _extract_after_keyword(clause, "paragraph", "New paragraph")
        actions.append({"action": "add_paragraph", "text": paragraph_text})
        ctx.last_paragraph = paragraph_text
    if "table" in low and any(w in low for w in ("add", "create", "insert", "make")):
        rows, cols = _extract_rows_cols(clause)
        actions.append({"action": "add_table", "rows": rows or 3, "cols": cols or 3})
        ctx.last_table_index += 1
    if "bold" in low:
        target = "heading" if "heading" in low else "selection"
        actions.append({"action": "set_bold", "target": target, "bold": True})
    if "italic" in low:
        target = "heading" if "heading" in low else "selection"
        actions.append({"action": "set_italic", "target": target, "italic": True})
    if "font color" in low or ("color" in low and "heading" in low):
        actions.append({"action": "set_font_color", "target": "heading" if "heading" in low else "selection", "color": _extract_color(clause) or "blue"})
    return _dedupe(actions)


def _planner_powerpoint_actions(clause, ctx):
    low = clause.lower()
    actions = []
    actions.extend(_open_action("powerpoint", clause))
    if actions:
        if actions[0].get("path"):
            ctx.current_presentation = actions[0]["path"]
        return actions
    if "save" in low and not any(w in low for w in ("save as", "saved as")):
        ctx.last_save_path = ctx.last_save_path or ctx.current_presentation
        return [_save_action("powerpoint")]
    if any(w in low for w in ("create", "make", "generate", "build", "new", "start")) and any(
        term in low for term in ("powerpoint", "power point", "ppt", "presentation", "slides", "slide deck", "deck", "slides")
    ):
        actions.append({"action": "create_presentation"})
        ctx.created = True
        ctx.current_presentation = ctx.current_presentation or "<new>"
    if "add" in low and "bullet" in low:
        slide = _slide_index(clause, ctx.last_slide_index)
        text = re.sub(r".*?\bbullet(?:\s+point)?\s*", "", clause, flags=re.IGNORECASE).strip(" .")
        actions.append({"action": "add_bullet_point", "slide_index": slide, "target": "body", "text": text or "Bullet"})
        ctx.last_slide_index = slide
        ctx.last_body_placeholder = "body"
    elif "duplicate" in low and "slide" in low:
        slide = _slide_index(clause, ctx.last_slide_index or 1)
        actions.append({"action": "duplicate_slide", "slide_index": slide})
        ctx.slide_count += 1
        ctx.last_slide_index = max(ctx.last_slide_index, ctx.slide_count)
    elif any(w in low for w in ("delete", "remove")) and "slide" in low:
        slide = _slide_index(clause, ctx.last_slide_index or 1)
        actions.append({"action": "delete_slide", "slide_index": slide})
        ctx.slide_count = max(0, ctx.slide_count - 1)
    elif "title slide" in low or ("agenda slide" in low) or ("conclusion slide" in low):
        ctx.slide_count += 1
        ctx.last_slide_index = ctx.slide_count
        ctx.last_created_slide = ctx.last_slide_index
        title = "Title"
        if "agenda" in low:
            title = "Agenda"
        elif "conclusion" in low:
            title = "Conclusion"
        actions.append({"action": "add_slide", "layout": "title_content"})
        actions.append({"action": "set_slide_text", "slide_index": ctx.last_slide_index, "target": "title", "text": title})
        ctx.last_title_placeholder = "title"
    elif "add slide" in low or "new slide" in low:
        ctx.slide_count += 1
        ctx.last_slide_index = ctx.slide_count
        ctx.last_created_slide = ctx.last_slide_index
        actions.append({"action": "add_slide", "layout": "title_content"})
    title_match = None if "title slide" in low else re.search(
        r"(?:on\s+slide\s+(\d+)\s+)?(?:set|write|add)\s+(?:the\s+)?title\s+(?:to|as)?\s*(.+)$",
        clause,
        re.IGNORECASE,
    )
    if title_match:
        slide = int(title_match.group(1) or ctx.last_slide_index or 1)
        actions.append({"action": "set_slide_text", "slide_index": slide, "target": "title", "text": title_match.group(2).strip(" .") or "Title"})
        ctx.last_slide_index = slide
        ctx.last_title_placeholder = "title"
    return _dedupe(actions)


def _slide_index(clause, default=1):
    match = re.search(r"\bslide\s+(\d+)\b", clause, re.IGNORECASE)
    return int(match.group(1)) if match else int(default or 1)


def _dedupe(actions):
    out = []
    seen = set()
    for action in actions:
        key = (
            action.get("action"),
            action.get("cell"),
            action.get("range"),
            action.get("slide_index"),
            action.get("text"),
            action.get("new_name"),
            action.get("path"),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(action)
    return out


def _planner_actions(app, clause, ctx):
    if app == "excel":
        return _planner_excel_actions(clause, ctx)
    if app == "word":
        return _planner_word_actions(clause, ctx)
    if app == "powerpoint":
        return _planner_powerpoint_actions(clause, ctx)
    return []


def plan_office_command(app_type, raw_command):
    app = "powerpoint" if (app_type or "").lower().strip() == "ppt" else (app_type or "").lower().strip()
    clauses_text = split_command_clauses(raw_command)
    ctx = PlanningContext(app)
    clauses = []
    actions = []
    errors = []
    warnings = []

    for index, clause_text in enumerate(clauses_text, start=1):
        intent = _detect_intent(app, clause_text)
        clause_actions = _planner_actions(app, clause_text, ctx)
        if not clause_actions:
            clause_actions = parse_command(app, clause_text) or []
        status = "parsed" if clause_actions else "failed"
        reason = "" if clause_actions else "Could not parse clause deterministically."
        clause = PlanClause(
            index=index,
            text=clause_text,
            intent=intent,
            actions=clause_actions,
            status=status,
            reason=reason,
        )
        clauses.append(clause)
        if clause_actions:
            actions.extend(clause_actions)
        else:
            errors.append(f"Clause {index}: {reason}")

    try:
        actions = normalize_actions(actions) if actions else []
        if actions:
            actions = validate_actions(app, actions)
    except OfficeActionError as exc:
        errors.append(exc.message)
        warnings.append(exc.details)

    success = bool(actions) and not errors
    requires_api = not bool(actions) or bool(errors)
    logger.info(
        "Command planner result: app=%s clauses=%s actions=%s success=%s requires_api=%s",
        app,
        len(clauses),
        len(actions),
        success,
        requires_api,
    )
    return CommandPlan(
        success=success,
        app=app,
        raw_command=raw_command or "",
        clauses=clauses,
        actions=actions,
        context=ctx.to_dict(),
        requires_api=requires_api,
        warnings=[w for w in warnings if w],
        errors=errors,
    )
