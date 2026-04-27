# modules/command_parser.py
import json
import re
import logging
import ast
from pathlib import Path

logger = logging.getLogger("OfficeAgent")

BASE = Path(__file__).parent.parent

_COMMAND_FILES = {
    "excel":      BASE / "excel_commands.json",
    "word":       BASE / "word_commands.json",
    "powerpoint": BASE / "powerpoint_commands.json",
    "ppt":        BASE / "powerpoint_commands.json",
}

_COMMAND_CACHE = {}
_STOP_WORDS = {"a", "an", "the", "please", "kindly"}


def _load_commands(app):
    if app in _COMMAND_CACHE:
        return _COMMAND_CACHE[app]
    path = _COMMAND_FILES.get(app)
    if not path or not path.exists():
        logger.warning(f"No command file for app: {app}")
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    _COMMAND_CACHE[app] = data
    return data


# ════════════════════════════════════════════════════════════════
#  PARAMETER EXTRACTORS
# ════════════════════════════════════════════════════════════════

def _extract_cell(text):
    match = re.search(r'\b([A-Z]{1,3}[0-9]{1,7})\b', text.upper())
    return match.group(1) if match else None


def _extract_range(text):
    match = re.search(r'\b([A-Z]{1,3}[0-9]{1,7})\s*[:\-to]+\s*([A-Z]{1,3}[0-9]{1,7})\b', text.upper())
    if match:
        return f"{match.group(1)}:{match.group(2)}"
    row_match = re.search(r'\brow\s*(\d+)\b', text.lower())
    if row_match:
        return f"{row_match.group(1)}:{row_match.group(1)}"
    col_match = re.search(r'\bcolumn\s*([A-Za-z]+)\b', text.lower())
    if col_match:
        c = col_match.group(1).upper()
        return f"{c}:{c}"
    return _extract_cell(text)


def _extract_number(text):
    match = re.search(r'\b(\d+(?:\.\d+)?)\b', text)
    return match.group(1) if match else None


def _extract_font_size(text):
    match = re.search(r'\b(\d{1,3})\s*(?:pt|px|point|size)?\b', text.lower())
    if match:
        val = int(match.group(1))
        if 6 <= val <= 200:
            return val
    return None


def _extract_color(text):
    color_map = {
        "red":        "FF0000",
        "green":      "00B050",
        "blue":       "0070C0",
        "yellow":     "FFFF00",
        "orange":     "FFA500",
        "purple":     "7030A0",
        "pink":       "FF69B4",
        "black":      "000000",
        "white":      "FFFFFF",
        "gray":       "808080",
        "grey":       "808080",
        "dark red":   "C00000",
        "dark blue":  "00008B",
        "dark green": "006400",
        "light blue": "ADD8E6",
        "light gray": "D3D3D3",
        "teal":       "008080",
        "cyan":       "00FFFF",
        "magenta":    "FF00FF",
        "gold":       "FFD700",
        "brown":      "A52A2A",
        "navy":       "000080",
    }
    text_lower = text.lower()
    for name, hex_val in color_map.items():
        if name in text_lower:
            return hex_val
    hex_match = re.search(r'#?([0-9A-Fa-f]{6})', text)
    if hex_match:
        return hex_match.group(1).upper()
    return None


def _extract_text_value(text):
    quoted = re.search(r'["\u201c\u201d\u2018\u2019](.*?)["\u201c\u201d\u2018\u2019]', text)
    if quoted:
        return quoted.group(1)
    for kw in ["write", "type", "set", "put", "enter", "add", "with text", "text", "saying", "called", "named", "title", "body", "subtitle", "content"]:
        pattern = rf'{kw}\s+(.+?)(?:\s+in\s+|\s+at\s+|\s+to\s+|\s+on\s+|$)'
        match = re.search(pattern, text.lower())
        if match:
            candidate = match.group(1).strip()
            if len(candidate) > 1:
                return candidate
    return None


def _extract_word_target(text):
    if not text:
        return "selection"

    quoted = re.search(r'["\u201c\u201d\u2018\u2019](.*?)["\u201c\u201d\u2018\u2019]', text)
    if quoted:
        candidate = quoted.group(1).strip()
        if candidate:
            return candidate

    patterns = [
        r'(?:for|on|in)\s+(?:the\s+)?text\s+(.+?)(?:\s+to\s+|\s+as\s+|\s+with\s+|$)',
        r'(?:for|on|in)\s+(?:the\s+)?paragraph\s+(.+?)(?:\s+to\s+|\s+as\s+|\s+with\s+|$)',
        r'(?:for|on|in)\s+(?:the\s+)?heading\s+(.+?)(?:\s+to\s+|\s+as\s+|\s+with\s+|$)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            candidate = match.group(1).strip(" .")
            if candidate:
                return candidate

    return "selection"


def _extract_compare_path(text):
    if not text:
        return ""
    match = re.search(r'["\']([^"\']+\.(?:docx|doc|rtf|odt))["\']', text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def _extract_data_source(text):
    if not text:
        return ""
    match = re.search(r'["\']([^"\']+\.(?:csv|xlsx|xls|json|txt))["\']', text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def _extract_filename(text):
    match = re.search(r'(?:as|named?|called?|to)\s+["\']?([a-zA-Z0-9_\-\s\.]+)["\']?', text.lower())
    if match:
        return match.group(1).strip()
    return None


def _extract_rows_cols(text):
    rows = cols = None
    row_match = re.search(r'(\d+)\s*(?:row|rows)', text.lower())
    col_match = re.search(r'(\d+)\s*(?:col|cols|column|columns)', text.lower())
    if row_match:
        rows = int(row_match.group(1))
    if col_match:
        cols = int(col_match.group(1))
    by_match = re.search(r'(\d+)\s*[xX×]\s*(\d+)', text)
    if by_match:
        rows = int(by_match.group(1))
        cols = int(by_match.group(2))
    return rows, cols


def _normalize_for_match(text):
    text = re.sub(r"[^a-z0-9\s]", " ", (text or "").lower())
    tokens = [t for t in text.split() if t and t not in _STOP_WORDS]
    return " ".join(tokens), set(tokens)


def _split_sub_commands(raw_command):
    text = (raw_command or "").strip()

    # Keep "X columns and Y rows" as one command (same for row/column inverse).
    protected = re.sub(
        r"(\d+\s*(?:col|cols|column|columns)\s+)and(\s*\d+\s*(?:row|rows))",
        r"\1__AND__\2",
        text,
        flags=re.IGNORECASE
    )
    protected = re.sub(
        r"(\d+\s*(?:row|rows)\s+)and(\s*\d+\s*(?:col|cols|column|columns))",
        r"\1__AND__\2",
        protected,
        flags=re.IGNORECASE
    )
    protected = re.sub(
        r"([A-Z]{1,3}\d{1,7}\s+)and(\s+[A-Z]{1,3}\d{1,7})",
        r"\1__AND__\2",
        protected,
        flags=re.IGNORECASE
    )
    protected = re.sub(
        r"profit\s+and\s+loss",
        lambda m: re.sub(r"\s+and\s+", " __AND__ ", m.group(0), flags=re.IGNORECASE),
        protected,
        flags=re.IGNORECASE
    )

    # Do not split on commas: commas are commonly used inside value lists
    # (for example: "fill A1:A4 with 2, 3, 4, 5").
    parts = re.split(r"\s+(?:and|then|also|after that|next)\s+", protected, flags=re.IGNORECASE)
    return [p.replace("__AND__", " and ").strip() for p in parts if p.strip()]


def _column_to_index(col):
    n = 0
    for ch in (col or "").upper():
        if "A" <= ch <= "Z":
            n = (n * 26) + (ord(ch) - ord("A") + 1)
    return n


def _index_to_column(n):
    n = int(n or 1)
    n = max(1, n)
    out = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def _parse_range_bounds(range_text):
    m = re.match(r"^\s*([A-Z]{1,3})(\d{1,7})\s*:\s*([A-Z]{1,3})(\d{1,7})\s*$", (range_text or "").upper())
    if not m:
        return None
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    return c1, r1, c2, r2


def _parse_fill_values(values_text):
    if not values_text:
        return []
    text = values_text.strip()
    text = re.sub(r"\band\b", ",", text, flags=re.IGNORECASE)
    parts = [p.strip() for p in re.split(r"[,\n;]+", text) if p.strip()]
    values = []
    for p in parts:
        if re.fullmatch(r"-?\d+", p):
            values.append(int(p))
        elif re.fullmatch(r"-?\d+\.\d+", p):
            values.append(float(p))
        else:
            values.append(p.strip(' "\''))
    return values


def _extract_literal_list(text):
    if not text:
        return None
    match = re.search(r"(\[\s*.*\s*\])", text, re.DOTALL)
    if not match:
        return None
    chunk = match.group(1)
    try:
        obj = ast.literal_eval(chunk)
        return obj
    except Exception:
        return None


def _parse_excel_structured_actions(raw_command):
    text = (raw_command or "").strip()
    low = text.lower()
    actions = []
    anchor_col = "A"

    if any(p in low for p in ("create workbook", "new workbook", "create a workbook", "create new workbook")):
        actions.append({"action": "create_workbook"})

    rename = re.search(r"rename\s+(?:the\s+)?sheet\s+(?:to|as)\s+['\"]?([^'\"\n]+?)['\"]?(?:,|$|\s+then|\s+and)", text, re.IGNORECASE)
    if rename:
        actions.append({"action": "rename_sheet", "new_name": rename.group(1).strip()})

    write_values = re.search(
        r"write\s+(?:the\s+)?values\s*(\[\[.*?\]\])\s*(?:starting\s+at|at|in)\s*([A-Z]{1,3}\d{1,7})",
        text,
        re.IGNORECASE | re.DOTALL
    )
    if write_values:
        try:
            matrix = ast.literal_eval(write_values.group(1))
            if isinstance(matrix, list):
                start_cell = write_values.group(2).upper()
                m = re.match(r"^([A-Z]{1,3})\d{1,7}$", start_cell)
                if m:
                    anchor_col = m.group(1)
                actions.append({
                    "action": "write_range",
                    "start_cell": start_cell,
                    "values": matrix,
                })
        except Exception:
            pass

    insert_row = re.search(r"insert\s+(?:a\s+)?row\s+(?:at\s+)?(?:index|row)?\s*(\d+)", text, re.IGNORECASE)
    if insert_row:
        actions.append({"action": "insert_row", "row": int(insert_row.group(1))})

    write_pair = re.search(
        r"write\s*(\[[^\]]+\])\s*(?:in|at)\s*cell\s*([A-Z]{1,3})(\d{1,7})",
        text,
        re.IGNORECASE
    )
    if write_pair:
        try:
            arr = ast.literal_eval(write_pair.group(1))
            if isinstance(arr, list) and len(arr) >= 2:
                base_col = write_pair.group(2).upper()
                row = int(write_pair.group(3))
                c1_idx = _column_to_index(base_col)
                c2 = _index_to_column(c1_idx + 1)
                actions.append({"action": "write_cell", "cell": f"{base_col}{row}", "value": arr[0]})
                actions.append({"action": "write_cell", "cell": f"{c2}{row}", "value": arr[1]})
        except Exception:
            pass

    # Example: "write ['Design', 'Done', 10] in row 4"
    write_in_row = re.search(
        r"write\s*(\[[^\]]+\])\s*(?:in|at)\s*row\s*(\d+)",
        text,
        re.IGNORECASE
    )
    if write_in_row:
        try:
            arr = ast.literal_eval(write_in_row.group(1))
            row_num = int(write_in_row.group(2))
            if isinstance(arr, list) and arr:
                base_idx = _column_to_index(anchor_col)
                for i, val in enumerate(arr):
                    col = _index_to_column(base_idx + i)
                    actions.append({"action": "write_cell", "cell": f"{col}{row_num}", "value": val})
        except Exception:
            pass

    def _clean_style_value(v):
        val = (v or "").strip().strip("'\"")
        val = re.split(r"\s+(?:and|then|finally)\b", val, maxsplit=1, flags=re.IGNORECASE)[0].strip()
        return val

    def _expand_targets(target_spec):
        spec = (target_spec or "").strip().upper()
        if not spec:
            return []
        if ":" in spec and not re.search(r"\bAND\b|,", spec):
            return [spec.replace(" ", "")]
        refs = re.findall(r"[A-Z]{1,3}\d{1,7}", spec)
        return refs

    bg_action_range = None
    bg = re.search(
        r"background\s+color\s+(?:of\s+)?([A-Z]{1,3}\d{1,7}\s*:\s*[A-Z]{1,3}\d{1,7})\s+(?:to|as)\s+(.+?)(?:,|$)",
        text,
        re.IGNORECASE
    )
    if bg:
        bg_range = bg.group(1).replace(" ", "").upper()
        bg_color_val = _clean_style_value(bg.group(2))
        bg_color = _extract_color(bg_color_val) or bg_color_val
        actions.append({"action": "set_bg_color", "range": bg_range, "color": bg_color})
        bg_action_range = bg_range

    # Example: "set background color of cells C3 and C5 to green"
    bg_cells = re.search(
        r"background\s+color\s+of\s+cells?\s+((?:[A-Z]{1,3}\d{1,7}\s*(?:,|and)\s*)+[A-Z]{1,3}\d{1,7})\s+(?:to|as)\s+(.+?)(?:,|$)",
        text,
        re.IGNORECASE
    )
    if bg_cells:
        color_val = _clean_style_value(bg_cells.group(2))
        color = _extract_color(color_val) or color_val
        refs = _expand_targets(bg_cells.group(1))
        for ref in refs:
            actions.append({"action": "set_bg_color", "range": ref, "color": color})
        if refs:
            bg_action_range = refs[0]

    font = re.search(
        r"font\s+color(?:\s+of\s+(?:cells?\s+)?((?:[A-Z]{1,3}\d{1,7}(?:\s*:\s*[A-Z]{1,3}\d{1,7})?(?:\s*(?:,|and)\s*)?)+))?\s+(?:to|as)\s+(.+?)(?:,|$)",
        text,
        re.IGNORECASE
    )
    if font:
        font_targets = _expand_targets(font.group(1) or bg_action_range or "")
        font_color_val = _clean_style_value(font.group(2))
        font_color = _extract_color(font_color_val) or font_color_val
        for target in font_targets:
            actions.append({"action": "set_font_color", "range": target, "color": font_color})

    font_size = re.search(
        r"font\s+size(?:\s+of\s+(?:cells?\s+)?((?:[A-Z]{1,3}\d{1,7}(?:\s*:\s*[A-Z]{1,3}\d{1,7})?(?:\s*(?:,|and)\s*)?)+))?\s+(?:to|as)\s+['\"]?(\d{1,3})(?:\s*(?:pt|px))?['\"]?",
        text,
        re.IGNORECASE
    )
    if font_size:
        size_targets = _expand_targets(font_size.group(1) or bg_action_range or "")
        size_val = int(font_size.group(2))
        for target in size_targets:
            actions.append({"action": "set_font_size", "range": target, "size": size_val})

    number_format = re.search(
        r"number\s+format\s+of\s+(?:cells?\s+)?([A-Z]{1,3}\d{1,7}\s*:\s*[A-Z]{1,3}\d{1,7})\s+(?:to|as)\s+['\"]?([^'\"]+)['\"]?",
        text,
        re.IGNORECASE
    )
    if number_format:
        actions.append({
            "action": "set_number_format",
            "range": number_format.group(1).replace(" ", "").upper(),
            "format": number_format.group(2).strip(),
        })

    # Example: "write a formula in cell D7 that calculates the total days from D3 to D6"
    sum_formula = re.search(
        r"formula\s+in\s+cell\s*([A-Z]{1,3}\d{1,7}).*?(?:total|sum).*?from\s*([A-Z]{1,3}\d{1,7})\s*(?:to|-)\s*([A-Z]{1,3}\d{1,7})",
        text,
        re.IGNORECASE
    )
    if sum_formula:
        out_cell = sum_formula.group(1).upper()
        start_ref = sum_formula.group(2).upper()
        end_ref = sum_formula.group(3).upper()
        actions.append({
            "action": "write_formula",
            "cell": out_cell,
            "formula": f"=SUM({start_ref}:{end_ref})",
        })

    concat = re.search(
        r"\bconcatenate\s+([A-Z]{1,3}\d{1,7})\s+and\s+([A-Z]{1,3}\d{1,7})\s+(?:in|into|to)\s+([A-Z]{1,3}\d{1,7})\b",
        text,
        re.IGNORECASE,
    )
    if concat:
        actions.append({
            "action": "write_formula",
            "cell": concat.group(3).upper(),
            "formula": f"=CONCATENATE({concat.group(1).upper()},{concat.group(2).upper()})",
        })

    textjoin = re.search(
        r"\btext\s*join\b|\btextjoin\b",
        text,
        re.IGNORECASE,
    )
    if textjoin:
        rng = _extract_range(text) or "A1:A5"
        result = _extract_result_cell(text)
        actions.append({
            "action": "write_formula",
            "cell": result,
            "formula": f"=TEXTJOIN({_excel_delimiter_literal(text)},TRUE,{rng})",
        })

    if re.search(r"\bsumifs\b|\bsum\s+ifs\b", text, re.IGNORECASE):
        actions.append({
            "action": "write_formula",
            "cell": _extract_result_cell(text),
            "formula": _extract_sumifs_formula(text),
        })

    if re.search(r"\b(today|todays date|today's date|current date)\b", low):
        actions.append({"action": "write_formula", "cell": _extract_result_cell(text), "formula": "=TODAY()"})

    if re.search(r"\b(now|current time|current datetime|timestamp)\b", low):
        actions.append({"action": "write_formula", "cell": _extract_result_cell(text), "formula": "=NOW()"})

    if "format" in low and "time" in low:
        actions.append({
            "action": "set_number_format",
            "range": _extract_cell(text) or "A1",
            "format": _extract_number_format(text),
        })

    if "unprotect workbook" in low or "unlock workbook" in low:
        actions.append({"action": "unprotect_workbook", "password": _resolve_inline_placeholder("password", text, "excel") or ""})
    elif "protect workbook" in low:
        actions.append({"action": "protect_workbook", "password": _resolve_inline_placeholder("password", text, "excel") or ""})
    if "unprotect sheet" in low or "unlock sheet" in low or "remove sheet password" in low:
        actions.append({
            "action": "unprotect_sheet",
            "sheet_name": _resolve_inline_placeholder("sheet_name", text, "excel") or "",
            "password": _resolve_inline_placeholder("password", text, "excel") or "",
        })
    elif "protect sheet" in low or "lock sheet" in low:
        actions.append({
            "action": "protect_sheet",
            "sheet_name": _resolve_inline_placeholder("sheet_name", text, "excel") or "",
            "password": _resolve_inline_placeholder("password", text, "excel") or "",
        })

    bold_range = re.search(
        r"(?:set\s+the\s+font\s+of\s+the\s+entire\s+range|set\s+bold\s+on\s+range|bold\s+range)\s*([A-Z]{1,3}\d{1,7}\s*:\s*[A-Z]{1,3}\d{1,7})",
        text,
        re.IGNORECASE
    )
    if not bold_range and "bold" in low:
        r = _extract_range(text)
        if r and ":" in r:
            bold_range = re.match(r"(.+)", r)
    if bold_range:
        rng = bold_range.group(1).replace(" ", "")
        actions.append({"action": "set_bold", "range": rng, "bold": True})

    # Keep only first occurrence for idempotent setup actions.
    dedup = []
    seen = set()
    for act in actions:
        key = (act.get("action"), act.get("cell"), act.get("range"), act.get("start_cell"), act.get("new_name"))
        if key in seen:
            continue
        seen.add(key)
        dedup.append(act)
    return dedup


def _heuristic_action(app, command_text):
    text = (command_text or "").lower()
    create_words = ("create", "make", "generate", "build", "new", "start", "write")

    if app == "excel" and any(w in text for w in create_words) and any(
        term in text for term in ("excel", "spreadsheet", "workbook", "worksheet", "sheet")
    ):
        if any(term in text for term in ("row", "rows", "column", "columns", "table")):
            rows, cols = _extract_rows_cols(text)
            return {
                "action": "create_table",
                "rows": rows or 5,
                "cols": cols or 3,
                "start_cell": _extract_cell(command_text) or "A1",
            }
        return {"action": "create_workbook"}

    if app == "word" and any(w in text for w in create_words) and any(
        term in text for term in ("word", "document", "docx")
    ):
        return {"action": "create_document"}

    if app in ("powerpoint", "ppt") and any(w in text for w in create_words) and any(
        term in text for term in ("powerpoint", "power point", "ppt", "presentation", "slide deck")
    ):
        return {"action": "create_presentation"}

    create_file_phrases = (
        "create a new file",
        "create new file",
        "make a new file",
        "make new file",
        "new blank file",
        "blank new file",
        "start a new file",
    )

    if any(phrase in text for phrase in create_file_phrases):
        if app == "excel":
            return {"action": "create_workbook"}
        if app == "word":
            return {"action": "create_document"}
        if app in ("powerpoint", "ppt"):
            return {"action": "create_presentation"}

    # Example: "fill A1:A4 with 2, 3, 4, 5"
    if app == "excel":
        fill_match = re.search(
            r"\b(?:fill|populate|write|enter)\s+(?:the\s+values?\s+)?(?:from\s+)?([A-Z]{1,3}\d{1,7}\s*:\s*[A-Z]{1,3}\d{1,7})\s+with\s+(.+?)(?:\s+respectively)?$",
            command_text,
            re.IGNORECASE
        )
        if fill_match:
            range_text = fill_match.group(1).replace(" ", "")
            raw_values = fill_match.group(2).strip()
            parsed_vals = _parse_fill_values(raw_values)
            bounds = _parse_range_bounds(range_text)
            if bounds and parsed_vals:
                c1, r1, c2, r2 = bounds
                row_count = abs(r2 - r1) + 1
                col_count = abs(_column_to_index(c2) - _column_to_index(c1)) + 1
                total_cells = max(1, row_count * col_count)

                values = parsed_vals[:total_cells]
                while len(values) < total_cells:
                    values.append("")

                matrix = []
                idx = 0
                for _ in range(row_count):
                    row_vals = []
                    for _ in range(col_count):
                        row_vals.append(values[idx])
                        idx += 1
                    matrix.append(row_vals)

                return {
                    "action": "write_range",
                    "start_cell": f"{c1}{r1}",
                    "values": matrix,
                }

    # Example: "create a table with 4 columns and 5 rows"
    if app == "excel" and "table" in text and any(w in text for w in ("create", "make", "insert", "add", "build", "include")):
        rows, cols = _extract_rows_cols(text)
        start_cell = _extract_cell(text) or "A1"
        return {
            "action": "create_table",
            "rows": rows or 5,
            "cols": cols or 3,
            "start_cell": start_cell,
        }

    if app == "word":
        if "table" in text and any(w in text for w in ("create", "make", "insert", "add", "build")):
            rows, cols = _extract_rows_cols(text)
            return {"action": "add_table", "rows": rows or 5, "cols": cols or 3}
        if any(w in text for w in ("add heading", "create heading", "heading")):
            return {"action": "add_heading", "text": _extract_text_value(command_text) or "Heading", "level": 1}
        if any(w in text for w in ("add paragraph", "write paragraph", "paragraph", "add text", "write text")):
            return {"action": "add_paragraph", "text": _extract_text_value(command_text) or "New paragraph"}

    if app in ("powerpoint", "ppt"):
        if "slide" in text and any(w in text for w in ("add", "create", "insert", "new")):
            return {"action": "add_slide", "layout": "title_content"}
        if "title" in text and any(w in text for w in ("set", "add", "write")):
            return {
                "action": "set_slide_text",
                "target": "title",
                "text": _extract_text_value(command_text) or "Title",
                "slide_index": _extract_slide_index(command_text) or 1,
            }
        if "table" in text and any(w in text for w in ("add", "create", "insert", "make")):
            rows, cols = _extract_rows_cols(text)
            return {
                "action": "insert_table",
                "rows": rows or 3,
                "cols": cols or 3,
                "slide_index": _extract_slide_index(command_text) or 1,
            }

    return None


def _extract_slide_index(text):
    patterns = [
        r'slide\s*(\d+)',
        r'(\d+)(?:st|nd|rd|th)?\s*slide',
        r'page\s*(\d+)',
    ]
    for p in patterns:
        m = re.search(p, text.lower())
        if m:
            return int(m.group(1))
    return None


def _extract_alignment(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["center", "centre", "middle", "centered"]):
        return "center"
    if any(w in text_lower for w in ["right", "flush right"]):
        return "right"
    if any(w in text_lower for w in ["justify", "justified", "both sides"]):
        return "justify"
    return "left"


def _extract_sort_order(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["descend", "z to a", "high to low", "largest first"]):
        return "descending"
    return "ascending"


def _extract_orientation(text):
    text_lower = text.lower()
    if "landscape" in text_lower:
        return "landscape"
    return "portrait"


def _extract_spacing(text):
    if "double" in text.lower():
        return 2.0
    if "1.5" in text:
        return 1.5
    if "single" in text.lower():
        return 1.0
    match = re.search(r'(\d+(?:\.\d+)?)\s*(?:line|spacing)', text.lower())
    if match:
        return float(match.group(1))
    return 1.15


def _extract_formula(text):
    match = re.search(r'=\s*([A-Z]+\s*\(.*?\))', text.upper())
    if match:
        return match.group(0).replace(" ", "")
    return None


def _extract_shape_type(text):
    text_lower = text.lower()
    shape_map = {
        "rectangle": "RECTANGLE",
        "rect":      "RECTANGLE",
        "square":    "RECTANGLE",
        "circle":    "OVAL",
        "oval":      "OVAL",
        "ellipse":   "OVAL",
        "arrow":     "RIGHT_ARROW",
        "triangle":  "TRIANGLE",
        "star":      "STAR_5_POINT",
        "pentagon":  "PENTAGON",
        "diamond":   "DIAMOND",
        "line":      "LINE",
    }
    for keyword, shape in shape_map.items():
        if keyword in text_lower:
            return shape
    return "RECTANGLE"


def _extract_chart_type(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["bar", "column"]):
        return "bar"
    if any(w in text_lower for w in ["line", "trend"]):
        return "line"
    if any(w in text_lower for w in ["pie", "donut", "doughnut"]):
        return "pie"
    if any(w in text_lower for w in ["area"]):
        return "area"
    if any(w in text_lower for w in ["scatter", "plot"]):
        return "scatter"
    return "bar"


def _extract_transition(text):
    text_lower = text.lower()
    transitions = ["fade", "wipe", "dissolve", "push", "cover", "uncover", "zoom", "split", "reveal", "flash"]
    for t in transitions:
        if t in text_lower:
            return t
    return "fade"


def _extract_target(text):
    text_lower = text.lower()
    if "title" in text_lower:
        return "title"
    if "subtitle" in text_lower:
        return "subtitle"
    if any(w in text_lower for w in ["body", "content", "text"]):
        return "body"
    return "body"


def _extract_level(text):
    for i in range(1, 10):
        if str(i) in text:
            return i
    return 1


def _extract_heading_level(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["heading 1", "h1", "primary", "main heading", "chapter"]):
        return 1
    if any(w in text_lower for w in ["heading 2", "h2", "subheading", "section"]):
        return 2
    if any(w in text_lower for w in ["heading 3", "h3"]):
        return 3
    if any(w in text_lower for w in ["heading 4", "h4"]):
        return 4
    return None


def _extract_delimiter(text):
    text_lower = text.lower()
    if "comma" in text_lower:
        return ","
    if "tab" in text_lower:
        return "\t"
    if "semicolon" in text_lower:
        return ";"
    if "pipe" in text_lower:
        return "|"
    if "space" in text_lower:
        return " "
    return ","


def _extract_url(text):
    match = re.search(r'https?://[^\s]+', text)
    if match:
        return match.group(0)
    return None


def _extract_find_replace(text):
    pattern = r'(?:replace|change)\s+["\']?(.+?)["\']?\s+(?:with|to)\s+["\']?(.+?)["\']?$'
    match = re.search(pattern, text.lower())
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return None, None


def _extract_number_format(text):
    text_lower = text.lower()
    if any(w in text_lower for w in ["currency", "dollar", "money", "$"]):
        return "$#,##0.00"
    if any(w in text_lower for w in ["percent", "%", "percentage"]):
        return "0.00%"
    if any(w in text_lower for w in ["date"]):
        if "mm/dd" in text_lower:
            return "MM/DD/YYYY"
        return "DD/MM/YYYY"
    if any(w in text_lower for w in ["time"]):
        if "hh:mm:ss" in text_lower:
            return "hh:mm:ss"
        if "hh:mm" in text_lower:
            return "hh:mm"
        return "hh:mm:ss"
    if "hh:mm:ss" in text_lower:
        return "hh:mm:ss"
    if "hh:mm" in text_lower:
        return "hh:mm"
    if any(w in text_lower for w in ["comma", "thousands"]):
        return "#,##0"
    return "General"


def _excel_formula_literal(value):
    raw_value = str(value if value is not None else "")
    if raw_value == " ":
        return '" "'
    raw = raw_value.strip()
    if not raw:
        return '""'
    if (raw.startswith('"') and raw.endswith('"')) or (raw.startswith("'") and raw.endswith("'")):
        return '"' + raw[1:-1].replace('"', '""') + '"'
    if re.fullmatch(r"-?\d+(?:\.\d+)?", raw):
        return raw
    if raw.startswith((">", "<", "=")):
        return '"' + raw.replace('"', '""') + '"'
    return '"' + raw.replace('"', '""') + '"'


def _excel_delimiter_literal(text):
    delimiter = _extract_delimiter(text)
    if delimiter == "\t":
        delimiter = "\\t"
    return _excel_formula_literal(delimiter)


def _extract_result_cell(text):
    cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', (text or "").upper())
    return cells[-1] if cells else "A1"


def _extract_cell_by_index(text, index, default="A1"):
    cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', (text or "").upper())
    return cells[index] if len(cells) > index else default


def _extract_sumifs_formula(text):
    raw = text or ""
    range_pat = r"(?:[A-Z]{1,3}:[A-Z]{1,3}|[A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7})"
    sum_range = None
    m = re.search(rf"\bsumifs?\s+({range_pat})", raw, re.IGNORECASE)
    if not m:
        m = re.search(rf"\bsum\s+({range_pat})", raw, re.IGNORECASE)
    if m:
        sum_range = m.group(1).upper().replace(" ", "")
    else:
        ranges = re.findall(range_pat, raw, flags=re.IGNORECASE)
        sum_range = ranges[0].upper().replace(" ", "") if ranges else "C:C"

    criteria_pairs = []
    for match in re.finditer(
        rf"({range_pat})\s*(?:is|=|equals|equal to|matching|for)\s*['\"]?([^,'\"\n]+?)['\"]?(?=\s+(?:and|,|in|into|to)\b|$)",
        raw,
        re.IGNORECASE,
    ):
        criteria_range = match.group(1).upper().replace(" ", "")
        if criteria_range == sum_range and not criteria_pairs:
            continue
        criteria = match.group(2).strip(" .")
        if criteria:
            criteria_pairs.append((criteria_range, criteria))

    if not criteria_pairs:
        ranges = [r.upper().replace(" ", "") for r in re.findall(range_pat, raw, flags=re.IGNORECASE)]
        if len(ranges) >= 2:
            criteria_pairs.append((ranges[1], ""))

    if not criteria_pairs:
        criteria_pairs.append(("A:A", ""))

    parts = [sum_range]
    for criteria_range, criteria in criteria_pairs[:5]:
        parts.extend([criteria_range, _excel_formula_literal(criteria or ">0")])
    return f"=SUMIFS({','.join(parts)})"


def _resolve_inline_placeholder(placeholder, command_text, app):
    placeholder = (placeholder or "").strip()
    if placeholder == "range":
        return _extract_range(command_text) or "A1"
    if placeholder in ("cell", "target_cell", "dest_cell"):
        return _extract_cell_by_index(command_text, 0)
    if placeholder == "result_cell":
        return _extract_result_cell(command_text)
    if placeholder == "cell1":
        return _extract_cell_by_index(command_text, 0)
    if placeholder == "cell2":
        return _extract_cell_by_index(command_text, 1, "B1")
    if placeholder == "delimiter":
        return _excel_delimiter_literal(command_text)
    if placeholder == "sumifs_formula":
        return _extract_sumifs_formula(command_text)
    if placeholder in ("sum_range", "range1", "range2"):
        ranges = re.findall(
            r"(?:[A-Z]{1,3}:[A-Z]{1,3}|[A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7})",
            command_text,
            flags=re.IGNORECASE,
        )
        idx = {"sum_range": 0, "range1": 1, "range2": 2}.get(placeholder, 0)
        return ranges[idx].upper().replace(" ", "") if len(ranges) > idx else (_extract_range(command_text) or "A1:A10")
    if placeholder in ("criteria", "criteria1", "criteria2"):
        criteria = re.findall(r"(?:is|=|equals|matching|for)\s*['\"]?([^,'\"\n]+?)['\"]?(?=\s+(?:and|,|in|into|to)\b|$)", command_text, flags=re.IGNORECASE)
        idx = 0 if placeholder in ("criteria", "criteria1") else 1
        return _excel_formula_literal(criteria[idx].strip()) if len(criteria) > idx else '">0"'
    if placeholder == "formula":
        return _extract_formula(command_text) or ""
    if placeholder == "format":
        return _extract_number_format(command_text)
    if placeholder == "filename":
        return _extract_filename(command_text) or "output"
    if placeholder == "file_path":
        path_match = re.search(r'["\']([^"\']+\.[a-z]{2,5})["\']', command_text, re.IGNORECASE)
        if path_match:
            return path_match.group(1)
        path_match = re.search(r'([A-Za-z]:[\\/][^"\']+\.[a-z]{2,5}|[^\s"\']+\.[a-z]{2,5})', command_text, re.IGNORECASE)
        return path_match.group(1) if path_match else ""
    if placeholder == "password":
        m = re.search(r'password\s+(?:is\s+)?["\']?([^\s"\']+)["\']?', command_text, re.IGNORECASE)
        return m.group(1) if m else ""
    if placeholder == "sheet_name":
        if re.search(r"\b(?:active|current)\s+sheet\b", command_text, re.IGNORECASE):
            return ""
        m = re.search(r'(?:sheet|tab)\s+(?:named|called|name)\s*["\']?([A-Za-z0-9 _-]+?)["\']?(?:\s+with|\s+password|$)', command_text, re.IGNORECASE)
        if m:
            candidate = m.group(1).strip()
            if candidate.lower() not in {"active", "current"}:
                return candidate
        return ""
    return None


def _interpolate_param_template(template, command_text, app):
    def repl(match):
        placeholder = match.group(1)
        value = _resolve_inline_placeholder(placeholder, command_text, app)
        if value is None:
            return match.group(0)
        return str(value)
    return re.sub(r"\{([^{}]+)\}", repl, template)


# ════════════════════════════════════════════════════════════════
#  KEYWORD MATCHING
# ════════════════════════════════════════════════════════════════

def _score_match(command_text, keywords):
    text_norm, text_tokens = _normalize_for_match(command_text)
    score = 0
    matched = []
    for kw in keywords:
        kw_norm, kw_tokens = _normalize_for_match(kw)
        if not kw_norm:
            continue

        if kw_norm in text_norm:
            score += max(2, len(kw_tokens) * 2)
            matched.append(kw)
            continue

        if kw_tokens and kw_tokens.issubset(text_tokens):
            score += max(1, len(kw_tokens))
            matched.append(kw)
            continue

        overlap = len(kw_tokens & text_tokens)
        if len(kw_tokens) >= 2 and overlap >= 2:
            score += overlap
            matched.append(kw)
    return score, matched


def _find_matching_commands(app, command_text):
    commands = _load_commands(app)
    scored = []
    for cmd_name, cmd_data in commands.items():
        keywords = cmd_data.get("keywords", [])
        score, matched = _score_match(command_text, keywords)
        if score > 0:
            scored.append((score, cmd_name, cmd_data, matched))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored


# ════════════════════════════════════════════════════════════════
#  PARAMETER RESOLVER
# ════════════════════════════════════════════════════════════════

def _resolve_params(params, command_text, app):
    resolved = {}
    for key, val in params.items():
        if isinstance(val, str) and re.search(r"\{[^{}]+\}", val):
            full = re.fullmatch(r"\{([^{}]+)\}", val.strip())
            if full:
                inline_value = _resolve_inline_placeholder(full.group(1), command_text, app)
                if inline_value is not None:
                    resolved[key] = inline_value
                    continue
            elif "{" in val and "}" in val:
                resolved[key] = _interpolate_param_template(val, command_text, app)
                continue

        if not isinstance(val, str) or not val.startswith("{"):
            resolved[key] = val
            continue

        placeholder = val.strip("{}")

        if placeholder in ("range",):
            resolved[key] = _extract_range(command_text) or "A1"

        elif placeholder in ("cell", "result_cell", "target_cell", "dest_cell"):
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            if placeholder == "result_cell" and len(cells) >= 1:
                resolved[key] = cells[-1]
            elif cells:
                resolved[key] = cells[0]
            else:
                resolved[key] = "A1"

        elif placeholder == "start_cell":
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = cells[0] if cells else "A1"

        elif placeholder == "value":
            resolved[key] = _extract_text_value(command_text) or _extract_number(command_text) or ""

        elif placeholder == "text":
            resolved[key] = _extract_text_value(command_text) or ""

        elif placeholder == "size":
            resolved[key] = _extract_font_size(command_text) or 12

        elif placeholder == "color":
            resolved[key] = _extract_color(command_text) or "000000"

        elif placeholder == "font_name":
            font_match = re.search(
                r'(?:font|typeface|use|change to|apply)\s+([A-Za-z\s]+?)(?:\s+font|\s+size|\s*$)',
                command_text.lower()
            )
            resolved[key] = font_match.group(1).strip().title() if font_match else "Calibri"

        elif placeholder == "alignment":
            resolved[key] = _extract_alignment(command_text)

        elif placeholder == "rows":
            rows, _ = _extract_rows_cols(command_text)
            resolved[key] = rows or 5

        elif placeholder == "cols":
            _, cols = _extract_rows_cols(command_text)
            resolved[key] = cols or 3

        elif placeholder == "row_number":
            m = re.search(r'(?:row|index)\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 1

        elif placeholder == "column":
            m = re.search(r'column\s*([A-Za-z]+)', command_text.lower())
            resolved[key] = m.group(1).upper() if m else "A"

        elif placeholder == "slide_index":
            idx = _extract_slide_index(command_text)
            resolved[key] = idx if idx is not None else 1

        elif placeholder == "target":
            if app == "powerpoint":
                resolved[key] = _extract_target(command_text)
            elif app == "word":
                resolved[key] = _extract_word_target(command_text)
            else:
                resolved[key] = _extract_range(command_text) or "selection"

        elif placeholder == "filename":
            resolved[key] = _extract_filename(command_text) or "output"

        elif placeholder == "file_path":
            path_match = re.search(r'["\']([^"\']+\.[a-z]{3,5})["\']', command_text)
            resolved[key] = path_match.group(1) if path_match else ""

        elif placeholder == "formula":
            resolved[key] = _extract_formula(command_text) or ""

        elif placeholder == "order":
            resolved[key] = _extract_sort_order(command_text)

        elif placeholder == "orientation":
            resolved[key] = _extract_orientation(command_text)

        elif placeholder == "spacing":
            resolved[key] = _extract_spacing(command_text)

        elif placeholder == "shape_type":
            resolved[key] = _extract_shape_type(command_text)

        elif placeholder == "chart_type":
            resolved[key] = _extract_chart_type(command_text)

        elif placeholder == "transition_type":
            resolved[key] = _extract_transition(command_text)

        elif placeholder == "animation_type":
            anim_match = re.search(
                r'(?:animation|animate|effect)\s+([a-z\s]+?)(?:\s+on|\s+to|$)',
                command_text.lower()
            )
            resolved[key] = anim_match.group(1).strip() if anim_match else "appear"

        elif placeholder == "format":
            resolved[key] = _extract_number_format(command_text)

        elif placeholder == "date_format":
            resolved[key] = "DD/MM/YYYY"

        elif placeholder == "delimiter":
            resolved[key] = _extract_delimiter(command_text)

        elif placeholder == "url":
            resolved[key] = _extract_url(command_text) or ""

        elif placeholder == "find_text":
            find, _ = _extract_find_replace(command_text)
            resolved[key] = find or ""

        elif placeholder == "replace_text":
            _, replace = _extract_find_replace(command_text)
            resolved[key] = replace or ""

        elif placeholder in ("level",):
            resolved[key] = _extract_heading_level(command_text) or _extract_level(command_text)

        elif placeholder == "zoom_level":
            m = re.search(r'(\d+)\s*%', command_text)
            resolved[key] = int(m.group(1)) if m else 100

        elif placeholder == "password":
            m = re.search(r'password\s+["\']?([^\s"\']+)["\']?', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "sheet_name":
            m = re.search(r'(?:sheet|tab)\s+["\']?([^\s"\']+)["\']?', command_text.lower())
            resolved[key] = m.group(1) if m else "Sheet1"

        elif placeholder == "old_name":
            m = re.search(r'rename\s+["\']?([^\s"\']+)["\']?\s+to', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "new_name":
            m = re.search(r'to\s+["\']?([^\s"\']+)["\']?$', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder in ("from_index",):
            m = re.search(r'from\s+(?:slide\s*)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 1

        elif placeholder == "to_index":
            m = re.search(r'to\s+(?:slide\s*)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 1

        elif placeholder == "count":
            m = re.search(r'(\d+)\s*(?:column|col)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder == "image_path":
            m = re.search(r'["\']([^"\']+\.(?:png|jpg|jpeg|gif|bmp|svg))["\']', command_text, re.IGNORECASE)
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "video_path":
            m = re.search(r'["\']([^"\']+\.(?:mp4|avi|mov|mkv|wmv))["\']', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder == "audio_path":
            m = re.search(r'["\']([^"\']+\.(?:mp3|wav|ogg|aac|flac))["\']', command_text.lower())
            resolved[key] = m.group(1) if m else ""

        elif placeholder in ("criteria", "condition"):
            m = re.search(r'(?:where|if|when|equals?|greater|less)\s+(.+?)(?:\s+in\s+|\s*$)', command_text.lower())
            resolved[key] = m.group(1).strip() if m else ""

        elif placeholder == "items":
            items_match = re.findall(r'["\u201c\u201d](.*?)["\u201c\u201d]', command_text)
            if items_match:
                resolved[key] = items_match
            else:
                resolved[key] = [s.strip() for s in re.split(r'[,;]', command_text)
                                 if len(s.strip()) > 1][-5:] or ["Item 1", "Item 2"]

        elif placeholder == "output_path":
            m = re.search(r'["\']([^"\']+\.pdf)["\']', command_text.lower())
            resolved[key] = m.group(1) if m else "output.pdf"

        elif placeholder == "width":
            m = re.search(r'width\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 10

        elif placeholder == "height":
            m = re.search(r'height\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 5

        elif placeholder == "angle":
            m = re.search(r'(\d+)\s*(?:degree|deg|°)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 90

        elif placeholder == "direction":
            if "vertical" in command_text.lower():
                resolved[key] = "vertical"
            else:
                resolved[key] = "horizontal"

        elif placeholder == "speed":
            if "slow" in command_text.lower():
                resolved[key] = "slow"
            elif "fast" in command_text.lower():
                resolved[key] = "fast"
            else:
                resolved[key] = "medium"

        elif placeholder == "seconds":
            m = re.search(r'(\d+(?:\.\d+)?)\s*(?:second|sec|s\b)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 3.0

        elif placeholder == "delay":
            m = re.search(r'(\d+(?:\.\d+)?)\s*(?:second|sec|s\b)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 0.0

        elif placeholder in ("before", "after"):
            m = re.search(rf'{placeholder}\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 6

        elif placeholder == "top":
            m = re.search(r'top\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "bottom":
            m = re.search(r'bottom\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "left":
            m = re.search(r'left\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "right":
            m = re.search(r'right\s*[=:]\s*(\d+(?:\.\d+)?)', command_text.lower())
            resolved[key] = float(m.group(1)) if m else 1.0

        elif placeholder == "position":
            m = re.search(r'position\s*[=:]\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 0

        elif placeholder in ("decimals", "num_chars"):
            m = re.search(r'(\d+)\s*(?:decimal|digit|char)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder == "base":
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = cells[0] if cells else "A1"

        elif placeholder == "exponent":
            m = re.search(r'(?:power|exponent)\s+(?:of\s+)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder in ("start_date", "end_date"):
            m = re.search(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = m.group(0) if m else "A1"

        elif placeholder == "unit":
            if "month" in command_text.lower():
                resolved[key] = "M"
            elif "year" in command_text.lower():
                resolved[key] = "Y"
            else:
                resolved[key] = "D"

        elif placeholder in ("year", "month", "day"):
            m = re.search(r'\b(19|20)\d{2}\b', command_text)
            resolved[key] = m.group(0) if m else placeholder.upper()

        elif placeholder == "error_value":
            resolved[key] = '""'

        elif placeholder in ("true_val", "false_val"):
            resolved[key] = '"Yes"' if "true" in placeholder else '"No"'

        elif placeholder in ("cond1", "cond2", "condition1", "condition2"):
            resolved[key] = "TRUE"

        elif placeholder in ("val1", "val2"):
            resolved[key] = '""'

        elif placeholder == "lookup_value":
            cells = re.findall(r'[A-Z]{1,3}[0-9]{1,7}', command_text.upper())
            resolved[key] = cells[0] if cells else "A1"

        elif placeholder == "table_range":
            resolved[key] = _extract_range(command_text) or "A1:Z100"

        elif placeholder == "col_index":
            m = re.search(r'column\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder == "row_index":
            m = re.search(r'row\s*(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 2

        elif placeholder in ("lookup_array", "return_array"):
            resolved[key] = _extract_range(command_text) or "A1:A10"

        elif placeholder in ("lookup_range", "return_range"):
            resolved[key] = _extract_range(command_text) or "A1:A10"

        elif placeholder in ("sum_range", "range1", "range2"):
            resolved[key] = _extract_range(command_text) or "A1:A10"

        elif placeholder in ("criteria1", "criteria2"):
            resolved[key] = '">0"'

        elif placeholder == "data_source":
            resolved[key] = _extract_data_source(command_text)

        elif placeholder == "compare_path":
            resolved[key] = _extract_compare_path(command_text)

        elif placeholder == "bookmark_name":
            m = re.search(r'(?:bookmark|mark)\s+["\']?([^\s"\']+)["\']?', command_text.lower())
            resolved[key] = m.group(1) if m else "bookmark1"

        elif placeholder == "layout":
            text_lower = command_text.lower()
            if "blank" in text_lower:
                return {**resolved, key: "blank"}
            if "title only" in text_lower:
                return {**resolved, key: "title_only"}
            if "two content" in text_lower or "two column" in text_lower:
                return {**resolved, key: "two_content"}
            resolved[key] = "title_content"

        elif placeholder == "style":
            style_match = re.search(
                r'(?:style|design|format)\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)',
                command_text.lower()
            )
            resolved[key] = style_match.group(1).strip().title() if style_match else "Normal"

        elif placeholder == "theme":
            m = re.search(r'theme\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)', command_text.lower())
            resolved[key] = m.group(1).strip().title() if m else "Office Theme"

        elif placeholder == "case":
            text_lower = command_text.lower()
            if "upper" in text_lower or "caps" in text_lower:
                return {**resolved, key: "upper"}
            if "lower" in text_lower:
                return {**resolved, key: "lower"}
            resolved[key] = "title"

        elif placeholder == "character":
            m = re.search(r'(?:symbol|character)\s+["\']?(.)["\']?', command_text)
            resolved[key] = m.group(1) if m else "©"

        elif placeholder == "type":
            resolved[key] = "list"

        elif placeholder == "values":
            lit = _extract_literal_list(command_text)
            resolved[key] = lit if isinstance(lit, list) else []

        elif placeholder == "icon_name":
            m = re.search(r'icon\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)', command_text.lower())
            resolved[key] = m.group(1).strip() if m else "star"

        elif placeholder == "nper":
            m = re.search(r'(\d+)\s*(?:year|month|period)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 12

        elif placeholder == "pv":
            m = re.search(r'(\d+(?:,\d{3})*(?:\.\d+)?)', command_text)
            resolved[key] = m.group(1).replace(",", "") if m else "10000"

        elif placeholder == "rate":
            m = re.search(r'(\d+(?:\.\d+)?)\s*%', command_text)
            resolved[key] = f"{float(m.group(1)) / 100}" if m else "0.05"

        elif placeholder == "start":
            m = re.search(r'start\s*(?:at\s*)?(\d+)', command_text.lower())
            resolved[key] = int(m.group(1)) if m else 1

        elif placeholder in ("start_row", "end_row"):
            rows = re.findall(r'row\s*(\d+)', command_text.lower())
            if placeholder == "start_row":
                resolved[key] = int(rows[0]) if rows else 1
            else:
                resolved[key] = int(rows[1]) if len(rows) > 1 else 5

        elif placeholder == "source_range":
            resolved[key] = _extract_range(command_text) or "A1:Z100"

        elif placeholder in ("color1", "color2"):
            colors = re.findall(
                r'\b(red|blue|green|yellow|orange|purple|white|black|gray|teal)\b',
                command_text.lower()
            )
            color_map_simple = {
                "red": "FF0000", "blue": "0070C0", "green": "00B050",
                "yellow": "FFFF00", "orange": "FFA500", "purple": "7030A0",
                "white": "FFFFFF", "black": "000000", "gray": "808080", "teal": "008080"
            }
            if placeholder == "color1":
                resolved[key] = color_map_simple.get(colors[0], "FFFFFF") if colors else "FFFFFF"
            else:
                resolved[key] = color_map_simple.get(colors[1], "000000") if len(colors) > 1 else "000000"

        elif placeholder == "scheme":
            m = re.search(r'scheme\s+["\']?([A-Za-z\s]+?)["\']?(?:\s|$)', command_text.lower())
            resolved[key] = m.group(1).strip().title() if m else "Default"

        else:
            resolved[key] = val

    return resolved


# ════════════════════════════════════════════════════════════════
#  MAIN PARSE FUNCTION
# ════════════════════════════════════════════════════════════════

def parse_command(app, raw_command):
    """
    Parse a raw office command string for the given app.
    Returns a list of action dicts ready for the executor.

    Example:
        parse_command("excel", "bold A1:E1 and sum B2 to B10 in C10")
        → [
            {"action": "set_bold",      "range": "A1:E1", "bold": True},
            {"action": "write_formula", "cell": "C10", "formula": "=SUM(B2:B10)"}
          ]
    """
    app = app.lower().strip()
    if app not in _COMMAND_FILES:
        logger.warning(f"Unknown app: {app}")
        return []

    if app == "excel":
        structured_actions = _parse_excel_structured_actions(raw_command)
        if structured_actions:
            logger.info(f"Structured excel parse hit ({len(structured_actions)} actions)")
            return structured_actions

    # Split compound commands by conjunctions
    sub_commands = _split_sub_commands(raw_command)

    all_actions = []
    seen_actions = set()

    for sub in sub_commands:
        sub = sub.strip()
        if not sub:
            continue

        # Prefer targeted PowerPoint edit heuristics before fuzzy navigation
        # matches such as "go_to_slide" consume "on slide N" commands.
        pre_heuristic = _heuristic_action(app, sub) if app in ("powerpoint", "ppt") else None
        if pre_heuristic and pre_heuristic.get("action") in {
            "set_slide_text", "add_bullet_point", "insert_table"
        }:
            all_actions.append(pre_heuristic)
            logger.info(f"Heuristic match for '{sub}' → {pre_heuristic}")
            continue

        matches = _find_matching_commands(app, sub)
        if not matches:
            heuristic = _heuristic_action(app, sub)
            if heuristic:
                all_actions.append(heuristic)
                logger.info(f"Heuristic match for '{sub}' → {heuristic}")
                continue
            logger.info(f"No match found for: '{sub}'")
            continue

        # Take top match (highest score), deduplicate by action+key params
        top_score, cmd_name, cmd_data, matched_kws = matches[0]
        action_name = cmd_data.get("action")
        raw_params  = cmd_data.get("params", {})
        resolved    = _resolve_params(raw_params, sub, app)

        # Dedup key to avoid running same action on same target twice
        dedup_key = f"{action_name}_{resolved.get('range') or resolved.get('cell') or resolved.get('slide_index', '')}"
        if dedup_key in seen_actions:
            continue
        seen_actions.add(dedup_key)

        action_dict = {"action": action_name, **resolved}
        all_actions.append(action_dict)
        logger.info(f"Matched '{cmd_name}' (score={top_score}, kws={matched_kws}) → {action_dict}")

    return all_actions
