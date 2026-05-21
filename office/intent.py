import re
from office.constants import OFFICE_TARGET_KEYWORDS, OFFICE_ACTION_KEYWORDS, APP_LAUNCH_PREFIXES


def _contains_term(text, term):
    term = (term or "").lower().strip()
    if not term:
        return False
    if " " in term:
        return term in text
    return bool(re.search(rf"\b{re.escape(term)}\b", text))


def _canonical_office_app(app_name):
    app_name = (app_name or "").lower().strip()
    return "powerpoint" if app_name == "ppt" else app_name


def _extract_office_agent_command(raw_text):
    text = (raw_text or "").strip()
    match = re.match(r"^agent\s*:\s*(excel|word|powerpoint|ppt)\s*:\s*(.+)$", text, re.IGNORECASE)
    if not match:
        return None, None
    return match.group(1).lower().strip(), match.group(2).strip()


def _is_app_launch_command(command_text):
    return (command_text or "").lower().strip().startswith(APP_LAUNCH_PREFIXES)


def _is_known_office_app(app_name):
    return _canonical_office_app(app_name) in {"excel", "word", "powerpoint"}


def _action_names(actions):
    return {
        str(a.get("action", "")).strip().lower()
        for a in (actions or [])
        if isinstance(a, dict)
    }


def _detect_action_type(command_text, actions):
    names = _action_names(actions)
    if names & {"create_workbook", "create_document", "create_presentation"}:
        return "create"
    if names & {"open_workbook", "open_document", "open_presentation"}:
        return "open"
    text = (command_text or "").lower()
    if any(_contains_term(text, t) for t in ("create", "make", "generate", "build", "new")):
        return "create"
    if any(_contains_term(text, t) for t in ("open", "load", "import")):
        return "open"
    if names or any(_contains_term(text, t) for t in OFFICE_ACTION_KEYWORDS):
        return "edit"
    return "unknown"


def _default_create_action(app_name, command_text):
    text = (command_text or "").lower()
    if not any(_contains_term(text, t) for t in OFFICE_ACTION_KEYWORDS):
        return None
    create_action = {
        "excel":      "create_workbook",
        "word":       "create_document",
        "powerpoint": "create_presentation",
        "ppt":        "create_presentation",
    }.get(app_name)
    return {"action": create_action} if create_action else None


def _detect_office_intent(raw_text):
    original = (raw_text or "").strip()
    parsed_app, parsed_command = _extract_office_agent_command(original)
    if parsed_app and parsed_command:
        app = _canonical_office_app(parsed_app)
        return {
            "is_office":   True,
            "app_type":    app,
            "command":     parsed_command,
            "action_type": _detect_action_type(parsed_command, []),
            "reason":      "agent-prefixed office command",
        }

    text = re.sub(r"\s+", " ", original.lower()).strip()
    if not text:
        return {"is_office": False, "reason": "empty command"}

    if re.match(r"^close\s+(?:the\s+)?(?:current\s+)?(?:document|file)\b", text):
        return {
            "is_office":   True,
            "app_type":    "word",
            "command":     original,
            "action_type": "edit",
            "reason":      "document lifecycle close command",
        }

    app_type = matched_term = ""
    for candidate, terms in OFFICE_TARGET_KEYWORDS.items():
        for term in terms:
            if _contains_term(text, term):
                app_type = candidate
                matched_term = term
                break
        if app_type:
            break

    if not app_type:
        return {"is_office": False, "reason": "no office target keyword"}

    if text.startswith("close "):
        document_close_terms = (
            "file", "workbook", "worksheet", "spreadsheet", "document", "docx",
            "doc", "xlsx", "xlsm", "xls", "csv", "pptx", "ppt", "presentation",
            "slide deck", "deck", "slides",
        )
        if not any(_contains_term(text, t) for t in document_close_terms):
            return {
                "is_office": False,
                "app_type":  app_type,
                "reason":    f"office application close request for '{matched_term}'",
            }

    has_action_term = any(_contains_term(text, t) for t in OFFICE_ACTION_KEYWORDS)
    has_open_doc_term = (
        text.startswith(("open ", "load ", "import "))
        and any(_contains_term(text, t) for t in (
            "file", "workbook", "worksheet", "spreadsheet", "document",
            "docx", "xlsx", "pptx", "presentation", "slide deck",
        ))
    )

    if has_action_term or has_open_doc_term:
        return {
            "is_office":   True,
            "app_type":    app_type,
            "command":     original,
            "action_type": _detect_action_type(original, []),
            "reason":      f"office target '{matched_term}' with document action term",
        }

    return {
        "is_office": False,
        "app_type":  app_type,
        "reason":    f"office app launch only for '{matched_term}'",
    }
