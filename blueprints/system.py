import logging
import traceback

from flask import Blueprint, jsonify, render_template, request
from modules import system_core, ui, config
from utils.app_alias_guard import validate_manual_app_alias
from office.intent import _detect_office_intent, _is_app_launch_command, _is_known_office_app
from utils.api_response import _json_success, _json_error

system_bp = Blueprint("system", __name__)


def _safe_speak(text):
    try:
        ui.speak(text)
    except Exception:
        pass


def _office_execute_impl_ref():
    from blueprints.office import _office_execute_impl
    return _office_execute_impl


@system_bp.route("/")
def index():
    html = render_template("index.html")
    script_tag = '<script src="/static/reliability.js"></script>'
    if script_tag not in html:
        html = html.replace("</body>", f"{script_tag} </body>")
    return html


@system_bp.route("/health", methods=["GET"])
def health():
    return jsonify(success=True, status="success", message="Backend online")


@system_bp.route("/execute", methods=["POST"])
def execute():
    try:
        data    = request.get_json(silent=True) or {}
        raw_cmd = (data.get("command") or "").strip()
        cmd     = raw_cmd.lower()
        logging.info("Received command: %s", raw_cmd)

        office_route = _detect_office_intent(raw_cmd)
        if office_route.get("is_office"):
            logging.info(
                "Routing decision: original=%r intent=office_automation app=%s handler=office_execute reason=%s",
                raw_cmd, office_route.get("app_type"), office_route.get("reason"),
            )
            return _office_execute_impl_ref()({
                **data,
                "command": raw_cmd,
                "raw":     office_route.get("command") or raw_cmd,
                "app":     office_route.get("app_type"),
            })

        if cmd.startswith(("close ", "shut ", "exit ")):
            for _prefix in ("close ", "shut ", "exit "):
                if cmd.startswith(_prefix):
                    app_name = cmd[len(_prefix):].strip()
                    break
            success, message = system_core.close_app(app_name)
            _safe_speak(f"Closing {app_name}" if success else f"Could not close {app_name}")
            if success:
                return _json_success(message, intent="app_close", app_type=app_name)
            return _json_error(message, intent="app_close", error_code="APP_CLOSE_FAILED", app_type=app_name)

        if not _is_app_launch_command(cmd):
            logging.info(
                "Routing decision: original=%r intent=unknown handler=none reason=not an app-launch or office command",
                raw_cmd,
            )
            return _json_error(
                "Command was not recognized as an Office automation or app-launch request.",
                intent="unknown",
                error_code="UNKNOWN_COMMAND",
            )

        app_name = system_core.normalize_app_name(raw_cmd)
        logging.info(
            "Routing decision: original=%r intent=app_launch app=%s handler=system_core.find_and_launch reason=launch verb",
            raw_cmd, app_name,
        )
        success, message = system_core.find_and_launch(app_name)
        if success:
            _safe_speak(f"Opening {app_name}")
            return _json_success(message, intent="app_launch", app_type=app_name)

        if _is_known_office_app(app_name):
            logging.warning("Office app launch failed without manual selector: app=%s message=%s", app_name, message)
            return _json_error(
                f"Could not open configured Office application: {app_name}.",
                intent="app_launch",
                error_code="OFFICE_APP_LAUNCH_FAILED",
                app_type=app_name,
                details=message,
            )

        _safe_speak(f"I couldn't find {app_name}. Please select it manually.")
        path = ui.manual_selector()
        if path:
            norm_app = system_core.normalize_app_name(app_name)
            alias_ok, alias_code, alias_message = validate_manual_app_alias(norm_app)
            if not alias_ok:
                logging.warning(
                    "Manual executable alias rejected: original=%r normalized=%r code=%s",
                    raw_cmd, norm_app, alias_code,
                )
                return _json_error(
                    alias_message, intent="app_launch", error_code=alias_code,
                    app_type=norm_app, requires_manual_selection=False,
                )
            config.save_memory(norm_app, path, is_store_app=False)
            launched = system_core.open_path(path)
            if launched:
                _safe_speak("Path saved. Opening now.")
                return _json_success(
                    "Manual Selection Saved", intent="app_launch",
                    app_type=norm_app, file_path=path, requires_manual_selection=True,
                )
            return _json_error(
                "Saved path, but launch failed", intent="app_launch",
                error_code="APP_LAUNCH_FAILED", app_type=norm_app,
                file_path=path, requires_manual_selection=True,
            )

        return _json_error(
            "Cancelled", intent="app_launch",
            error_code="MANUAL_SELECTION_CANCELLED",
            app_type=app_name, requires_manual_selection=True,
        )

    except Exception as e:
        logging.error("Command route error: %s\n%s", e, traceback.format_exc())
        return _json_error(
            "Command execution failed.", intent="unknown",
            error_code="COMMAND_ROUTE_ERROR", details=str(e), http_status=500,
        )
