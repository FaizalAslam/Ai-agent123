import logging
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path

from flask import Flask

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# ---- Core modules ---------------------------------------------------------
from modules import system_core, ui

# ---- Shared state ---------------------------------------------------------
import state  # noqa: F401 — imported for side-effects (listener init)

# ---- Blueprints -----------------------------------------------------------
# ---- Backwards-compatible server exports ----------------------------------
BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"
BACKEND_URL = "http://127.0.0.1:5000"
DEFAULT_FRONTEND_URL = "http://127.0.0.1:3000"
_frontend_process = None


def _url_is_reachable(url, timeout=1.5):
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
            return 200 <= getattr(response, "status", 200) < 400
    except (OSError, urllib.error.URLError):
        return False


def _wait_for_url(url, timeout=20):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _url_is_reachable(url, timeout=1):
            return True
        time.sleep(0.5)
    return False


def _is_local_frontend_url(url):
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or "").lower()
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    return host in {"127.0.0.1", "localhost", "::1"} and port == 3000


def _start_frontend_if_needed(frontend_url):
    global _frontend_process
    if os.environ.get("AI_AGENT_SKIP_FRONTEND_START", "").lower() in {"1", "true", "yes"}:
        print("Frontend auto-start skipped by AI_AGENT_SKIP_FRONTEND_START.")
        return False
    if not _is_local_frontend_url(frontend_url):
        print(f"Frontend auto-start skipped for non-local URL: {frontend_url}")
        return False
    if _url_is_reachable(frontend_url):
        print(f"Frontend already running at {frontend_url}")
        return True
    if not _port_available("127.0.0.1", 3000):
        print(
            "Port 3000 is already in use, but the frontend did not return a healthy response. "
            "Stop the old Next.js process and start again."
        )
        return False
    if not (FRONTEND_DIR / "package.json").exists():
        print(f"Frontend folder not found: {FRONTEND_DIR}")
        return False

    npm_cmd = shutil.which("npm.cmd") or shutil.which("npm")
    if not npm_cmd:
        print("Could not find npm on PATH. Start the frontend manually with: cd frontend && npm run dev")
        return False

    log_path = BASE_DIR / "frontend-dev.log"
    creationflags = 0
    startupinfo = None
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0

    with open(log_path, "wb") as log:
        _frontend_process = subprocess.Popen(
            [npm_cmd, "run", "dev"],
            cwd=str(FRONTEND_DIR),
            stdout=log,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            creationflags=creationflags,
            startupinfo=startupinfo,
        )
    print(f"Starting Next.js frontend at {frontend_url} (log: {log_path})")
    return _wait_for_url(frontend_url, timeout=25)


def _port_available(host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(1)
        return sock.connect_ex((host, port)) != 0


def _backend_port_available():
    return _port_available("127.0.0.1", 5000)


def _ensure_backend_can_start():
    if _url_is_reachable(f"{BACKEND_URL}/health", timeout=1):
        print(f"Backend already running at {BACKEND_URL}")
        return False
    if not _backend_port_available():
        print(
            f"Port 5000 is already in use, but {BACKEND_URL}/health did not respond. "
            "Stop the old server.py process and start again."
        )
        sys.exit(1)
    return True


from office.runner import _known_office_actions
from parser.command_parser import parse_command
from parser.command_planner import plan_office_command
from utils import command_map
import office.resolver as _office_resolver

_openai_handler = state.openai_handler


def _resolve_actions(app_name, command_text):
    """Compatibility wrapper for callers that patched server.py pre-refactor."""
    original_parse_command = _office_resolver.parse_command
    original_plan_office_command = _office_resolver.plan_office_command
    original_command_map = _office_resolver.command_map
    original_known_office_actions = _office_resolver._known_office_actions
    original_openai_handler = _office_resolver.state.openai_handler
    _office_resolver.parse_command = parse_command
    _office_resolver.plan_office_command = plan_office_command
    _office_resolver.command_map = command_map
    _office_resolver._known_office_actions = _known_office_actions
    _office_resolver.state.openai_handler = _openai_handler
    try:
        return _office_resolver._resolve_actions(app_name, command_text)
    finally:
        _office_resolver.parse_command = original_parse_command
        _office_resolver.plan_office_command = original_plan_office_command
        _office_resolver.command_map = original_command_map
        _office_resolver._known_office_actions = original_known_office_actions
        _office_resolver.state.openai_handler = original_openai_handler

from blueprints.office import office_bp
from blueprints.system import system_bp
from blueprints.voice  import voice_bp
from blueprints.ocr    import ocr_bp
from blueprints.pdf    import pdf_bp

# ---- Optional modules availability flags ----------------------------------
try:
    from modules import ocr_utils as _ocr_mod
    OCR_AVAILABLE = True
except Exception as _e:
    _ocr_mod = None
    OCR_AVAILABLE = False
    print(f"OCR unavailable: {_e}")

try:
    from modules import pdf_utils as _pdf_mod
    PDF_AVAILABLE = True
except Exception as _e:
    _pdf_mod = None
    PDF_AVAILABLE = False
    print(f"PDF unavailable: {_e}")

try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    keyboard = None
    KEYBOARD_AVAILABLE = False
    print("keyboard not found — pip install keyboard")

# ---- Logging --------------------------------------------------------------
logging.basicConfig(
    filename="agent.log",
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%H:%M:%S",
    filemode="w",
)
logging.getLogger("werkzeug").setLevel(logging.ERROR)

# ---- Flask app ------------------------------------------------------------
app = Flask(__name__)
app.register_blueprint(office_bp)
app.register_blueprint(system_bp)
app.register_blueprint(voice_bp)
app.register_blueprint(ocr_bp)
app.register_blueprint(pdf_bp)


# ---- Global command callback (voice / keyboard / clipboard) ---------------

def _safe_speak(text):
    try:
        ui.speak(text)
    except Exception:
        pass


def _handle_global_command(raw_text):
    """Handles system-wide  agent: <app>: <command>  triggers."""
    from office.intent import _extract_office_agent_command
    from office.file_resolver import _ensure_fresh_file_action, _expand_powerpoint_slide_count, resolve_office_file_path
    from office.resolver import _resolve_actions
    from office.runner import _run_office_actions, _known_office_actions
    from utils.office_actions import OfficeActionError, validate_actions
    from utils import command_map

    try:
        app_name, command = _extract_office_agent_command(raw_text)
        if app_name and command:
            if app_name == "ppt":
                app_name = "powerpoint"
            cache_key, actions, source, action_error, plan_info = _resolve_actions(app_name, command)
            if action_error:
                logging.warning("Global office action parse failed: %s", action_error.message)
                return
            if not actions:
                logging.warning("No office action match for global command: %s: %s", app_name, command)
                return
            actions = _ensure_fresh_file_action(app_name, command, actions, "")
            actions = _expand_powerpoint_slide_count(app_name, command, actions)
            try:
                actions = validate_actions(app_name, actions, known_actions=_known_office_actions(app_name))
            except OfficeActionError as exc:
                logging.warning("Global office validation failed: %s", exc.message)
                return
            resolution = resolve_office_file_path({"raw": command}, actions, app_name)
            if not resolution.get("success"):
                logging.warning("Global office path resolution failed: %s", resolution.get("message"))
                return
            summary = _run_office_actions(
                app_name, actions,
                file_path=resolution["file_path"],
                source_path=resolution.get("source_path"),
                command_text=command,
            )
            if summary["failures"] and cache_key:
                command_map.remove_action(app_name, cache_key)
            logging.info(
                "Global office [%s] %s: %s -> %s/%s | %s",
                source, app_name, command,
                summary["ok_count"], summary["total"], summary["output_path"],
            )
            if summary.get("persisted"):
                _safe_speak(f"Executed {summary['ok_count']} actions in {app_name}")
            else:
                _safe_speak(f"Executed {summary['ok_count']} actions in {app_name}, not saved")
            return

        txt = (raw_text or "").strip()
        low = txt.lower()
        if low.startswith("agent "):
            sys_cmd = txt[len("agent "):].strip().replace("  ", " ").strip(" .,:;!?")
            if sys_cmd.startswith(("open ", "launch ", "start ", "run ", "boot ")):
                success, message = system_core.find_and_launch(sys_cmd)
                _safe_speak(
                    f"Opening {system_core.normalize_app_name(sys_cmd)}"
                    if success else f"Could not open {sys_cmd}"
                )
                logging.info("Voice system open [%s] => %s: %s", sys_cmd, success, message)
            elif sys_cmd.startswith(("close ", "shut ", "exit ")):
                success, message = system_core.close_app(sys_cmd)
                _safe_speak(
                    f"Closing {system_core.normalize_app_name(sys_cmd)}"
                    if success else f"Could not close {sys_cmd}"
                )
                logging.info("Voice system close [%s] => %s: %s", sys_cmd, success, message)
    except Exception as e:
        logging.error("Global command error: %s\n%s", e, traceback.format_exc())


# Patch listener callbacks now that _handle_global_command is defined.
if state._keyboard_listener:
    state._keyboard_listener.on_command = _handle_global_command
if state._voice_listener:
    state._voice_listener.on_command = _handle_global_command


# ===========================================================================
# ENTRY POINT
# ===========================================================================

if __name__ == "__main__":
    _should_start_backend = _ensure_backend_can_start()
    if not _should_start_backend:
        frontend_url = os.environ.get("AI_AGENT_FRONTEND_URL", DEFAULT_FRONTEND_URL)
        frontend_ready = _start_frontend_if_needed(frontend_url)
        browser_url = frontend_url if frontend_ready or not _is_local_frontend_url(frontend_url) else BACKEND_URL
        if os.environ.get("AI_AGENT_SKIP_BROWSER", "").lower() not in {"1", "true", "yes"}:
            webbrowser.open(browser_url)
        print(f"Agent UI running at {browser_url}")
        sys.exit(0)

    # ---- OCR snip overlay (must be on main thread) ------------------------
    if OCR_AVAILABLE:
        threading.Thread(target=_ocr_mod.run_snip_overlay_main_thread, daemon=True).start()

    # ---- OCR keyboard hotkeys ---------------------------------------------
    if KEYBOARD_AVAILABLE and OCR_AVAILABLE:
        keyboard.add_hotkey(
            "ctrl+shift+s",
            lambda: threading.Thread(
                target=_ocr_mod.trigger_snip_and_ocr, args=(state.last_ocr,), daemon=True
            ).start(),
        )
        keyboard.add_hotkey(
            "ctrl+shift+f",
            lambda: threading.Thread(
                target=_ocr_mod.trigger_screenshot_and_ocr, args=(state.last_ocr,), daemon=True
            ).start(),
        )
        print("Ctrl+Shift+S -> Snip OCR  |  Ctrl+Shift+F -> Fullscreen OCR")

    # ---- Global Office Agent listeners ------------------------------------
    if state._clipboard_listener:
        threading.Thread(target=state._clipboard_listener.start, daemon=True, name="ClipboardListener").start()
    if state._keyboard_listener:
        threading.Thread(target=state._keyboard_listener.start, daemon=True, name="KeyboardListener").start()
    if state._clipboard_listener or state._keyboard_listener:
        print("Global agent listener active")
        print("     Type  agent: excel: <command>  anywhere + Enter")
    else:
        print("Global clipboard/keyboard listener unavailable; backend routes still active.")

    if state._voice_listener and state._voice_listener.available:
        if state._voice_listener.start():
            state.voice_state["enabled"] = True
            print("Voice wake listener active (say: agent <app> <command>)")
        else:
            print(f"Voice listener not started: {state._voice_listener.last_error}")

    # ---- Start Flask ------------------------------------------------------
    flask_thread = threading.Thread(
        target=lambda: app.run(host="127.0.0.1", port=5000, debug=False),
        daemon=True,
    )
    flask_thread.start()
    #time.sleep(1)

    # ---- Start/open frontend ---------------------------------------------
    _wait_for_url(f"{BACKEND_URL}/health", timeout=8)
    frontend_url = os.environ.get("AI_AGENT_FRONTEND_URL", DEFAULT_FRONTEND_URL)
    frontend_ready = _start_frontend_if_needed(frontend_url)
    browser_url = frontend_url if frontend_ready or not _is_local_frontend_url(frontend_url) else BACKEND_URL
    if os.environ.get("AI_AGENT_SKIP_BROWSER", "").lower() not in {"1", "true", "yes"}:
        webbrowser.open(browser_url)
    print(f"Agent UI running at {browser_url}")

    # ---- Dialog listener must be on main thread ---------------------------
    if PDF_AVAILABLE:
        _pdf_mod.run_dialog_listener()
    else:
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nAgent stopped.")
