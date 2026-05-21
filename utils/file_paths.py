import re
import os
import time
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
OFFICE_OUTPUT_ROOT = PROJECT_ROOT / "outputs" / "office"

OFFICE_EXTENSIONS = {
    "excel": "xlsx",
    "word": "docx",
    "powerpoint": "pptx",
    "ppt": "pptx",
}

OFFICE_EXTENSION_CANDIDATES = {
    "excel": ("xlsx", "xlsm", "xls", "csv"),
    "word": ("docx", "doc"),
    "powerpoint": ("pptx", "ppt"),
    "ppt": ("pptx", "ppt"),
}

OFFICE_OUTPUT_PREFIXES = {
    "excel": "excel_output",
    "word": "word_output",
    "powerpoint": "powerpoint_output",
    "ppt": "powerpoint_output",
}

WINDOWS_INVALID_FILENAME_CHARS = r'<>:"/\|?*'


class FilePathError(ValueError):
    def __init__(self, error_code, message, details=""):
        super().__init__(message)
        self.error_code = error_code
        self.message = message
        self.details = details


def canonical_office_app(app_type):
    app = (app_type or "").strip().lower()
    return "powerpoint" if app == "ppt" else app


def sanitize_filename(name, default="output", max_length=120):
    raw = str(name or "").strip().strip("\"'")
    raw = re.sub(r"[<>:\"/\\|?*\x00-\x1f]+", "", raw)
    raw = re.sub(r"\s+", " ", raw).strip(" .")
    if not raw:
        raw = default
    if len(raw) > max_length:
        raw = raw[:max_length].rstrip(" .")
    return raw or default


def ensure_office_extension(path, app_type):
    ext = OFFICE_EXTENSIONS.get(canonical_office_app(app_type), "")
    path = Path(path)
    if ext and path.suffix.lower() != f".{ext}":
        return path.with_suffix(f".{ext}")
    return path


def next_available_path(path):
    candidate = Path(path)
    base = candidate.with_suffix("")
    suffix = candidate.suffix
    idx = 1
    while candidate.exists():
        candidate = base.with_name(f"{base.name}_{idx}").with_suffix(suffix)
        idx += 1
    return candidate


def output_dir_for_app(app_type):
    app = canonical_office_app(app_type)
    folder = OFFICE_OUTPUT_ROOT / (app or "unknown")
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def common_user_locations(command_text=""):
    """Return local-first locations used when a user says Desktop/Documents/Downloads."""
    home = Path.home()
    text = str(command_text or "").lower()
    candidates = []
    location_map = {
        "desktop": ("Desktop",),
        "documents": ("Documents",),
        "document": ("Documents",),
        "downloads": ("Downloads",),
        "download": ("Downloads",),
    }
    requested = []
    for word, rels in location_map.items():
        if re.search(rf"\b{re.escape(word)}\b", text):
            requested.extend(rels)
    if not requested:
        requested = ["Desktop", "Documents", "Downloads"]

    for rel in requested:
        candidates.append(home / rel)
        candidates.append(home / "OneDrive" / rel)
    candidates.append(PROJECT_ROOT)
    candidates.append(OFFICE_OUTPUT_ROOT)

    seen = set()
    unique = []
    for path in candidates:
        key = str(path).lower()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def generate_office_output_path(app_type):
    app = canonical_office_app(app_type)
    ext = OFFICE_EXTENSIONS.get(app, "xlsx")
    prefix = OFFICE_OUTPUT_PREFIXES.get(app, f"{app or 'office'}_output")
    stamp = time.strftime("%Y%m%d_%H%M%S")
    millis = int((time.time() * 1000) % 1000)
    path = output_dir_for_app(app) / f"{prefix}_{stamp}_{millis:03d}.{ext}"
    return next_available_path(path).resolve()


def _extension_candidates(path, app_type):
    path = Path(path)
    if path.suffix:
        return [path]
    app = canonical_office_app(app_type)
    return [path.with_suffix(f".{ext}") for ext in OFFICE_EXTENSION_CANDIDATES.get(app, (OFFICE_EXTENSIONS.get(app, ""),)) if ext]


def extract_office_filename_hint(command_text, app_type):
    """Extract a raw Office filename/path hint from an open-file command."""
    text = str(command_text or "").strip()
    if not text:
        return ""
    app = canonical_office_app(app_type)
    extensions = OFFICE_EXTENSION_CANDIDATES.get(app, (OFFICE_EXTENSIONS.get(app, ""),))
    ext_group = "|".join(re.escape(ext) for ext in extensions if ext)

    if ext_group:
        quoted = re.search(rf"['\"]([^'\"]+\.({ext_group}))['\"]", text, re.IGNORECASE)
        if quoted:
            return quoted.group(1).strip()
        absolute = re.search(rf"([A-Za-z]:[\\/][^\"'\n]+?\.({ext_group}))", text, re.IGNORECASE)
        if absolute:
            return absolute.group(1).strip()
        plain = re.search(rf"\b([A-Za-z0-9_\- .]+\.({ext_group}))\b", text, re.IGNORECASE)
        if plain:
            return _clean_filename_hint(plain.group(1))

    named_patterns = (
        r"\b(?:named|called)\s+['\"]?([A-Za-z0-9_\- .]{1,160})['\"]?",
        r"\b(?:file|document|workbook|presentation)\s+(?:named|called)?\s*['\"]?([A-Za-z0-9_\- .]{1,160})['\"]?",
        r"\bopen\b.+?\b([A-Za-z0-9_\-]{1,120})\b(?:\s+(?:on|from|in|at)\s+(?:desktop|documents|downloads))",
    )
    for pattern in named_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            hint = _clean_filename_hint(match.group(1))
            if hint:
                return hint
    return ""


def _clean_filename_hint(value):
    hint = str(value or "").strip().strip("\"'")
    hint = re.split(
        r"\s+(?:and|then|with|in|on|from|at|to|into)\b",
        hint,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    return hint.strip(" .")


def _case_insensitive_existing(candidate):
    candidate = Path(candidate)
    if candidate.exists():
        return candidate
    parent = candidate.parent
    try:
        if not parent.exists():
            return None
        target_name = candidate.name.lower()
        for child in parent.iterdir():
            if child.name.lower() == target_name:
                return child
    except OSError:
        return None
    return None


def candidate_input_paths(value, app_type, base_dir=None, command_text=""):
    raw = str(value or "").strip().strip("\"'")
    if not raw:
        return []
    if "\x00" in raw:
        raise FilePathError("UNSAFE_PATH", "Path contains an invalid null byte.", raw[:120])

    path = Path(os.path.expandvars(raw)).expanduser()
    if path.is_absolute():
        bases = [path]
    elif len(path.parts) > 1:
        bases = [Path(base_dir or PROJECT_ROOT) / path]
    else:
        app = canonical_office_app(app_type)
        bases = [
            Path(base_dir or PROJECT_ROOT) / path,
            output_dir_for_app(app) / path.name,
            *(location / path.name for location in common_user_locations(command_text)),
        ]

    candidates = []
    seen = set()
    for base in bases:
        for candidate in _extension_candidates(base, app_type):
            try:
                resolved = candidate.resolve()
            except OSError:
                resolved = candidate
            key = str(resolved).lower()
            if key not in seen:
                seen.add(key)
                candidates.append(resolved)
    return candidates


def resolve_existing_office_path(value, app_type, base_dir=None, command_text=""):
    raw_value = str(value or "").strip()
    if not raw_value:
        raw_value = extract_office_filename_hint(command_text, app_type)
    attempted = candidate_input_paths(raw_value, app_type, base_dir=base_dir, command_text=command_text)
    for candidate in attempted:
        existing = _case_insensitive_existing(candidate)
        if existing:
            return existing.resolve()
    details = "; ".join(str(path) for path in attempted[:12])
    raise FilePathError(
        "FILE_NOT_FOUND",
        f"Office input file was not found: {value}",
        details,
    )


def resolve_path_value(value, app_type, for_output=False, base_dir=None):
    raw = str(value or "").strip().strip("\"'")
    if not raw:
        return Path()
    if "\x00" in raw:
        raise FilePathError("UNSAFE_PATH", "Path contains an invalid null byte.", raw[:120])

    path = Path(os.path.expandvars(raw)).expanduser()
    raw_was_absolute = path.is_absolute()
    if for_output and not raw_was_absolute and any(part == ".." for part in path.parts):
        raise FilePathError("UNSAFE_PATH", "Output path traversal is not allowed.", raw[:120])
    if for_output:
        path = ensure_office_extension(path, app_type)

    if not path.is_absolute():
        if for_output and len(path.parts) == 1:
            path = output_dir_for_app(app_type) / sanitize_filename(path.name)
        else:
            root = Path(base_dir or PROJECT_ROOT)
            path = root / path

    try:
        resolved = path.resolve()
    except OSError as exc:
        raise FilePathError("INVALID_FILE_PATH", "Could not resolve path.", str(exc)) from exc

    if for_output:
        resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def named_output_path(raw_name, app_type):
    app = canonical_office_app(app_type)
    ext = OFFICE_EXTENSIONS.get(app, "xlsx")
    name = sanitize_filename(raw_name, default=OFFICE_OUTPUT_PREFIXES.get(app, "office_output"))
    path = Path(name)
    if path.suffix.lower() != f".{ext}":
        path = path.with_suffix(f".{ext}")
    if path.is_absolute():
        return path.resolve()
    return (output_dir_for_app(app) / path.name).resolve()
