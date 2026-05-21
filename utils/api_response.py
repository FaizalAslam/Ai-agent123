from flask import jsonify


def _json_success(message, intent="unknown", **extra):
    payload = {
        "success": True,
        "status": "success",
        "intent": intent,
        "message": message,
    }
    payload.update(extra)
    payload.setdefault("data", {k: v for k, v in extra.items() if k != "data"})
    return jsonify(payload)


def _json_error(message, intent="unknown", error_code="COMMAND_FAILED", http_status=200, **extra):
    payload = {
        "success": False,
        "status": "fail",
        "intent": intent,
        "error_code": error_code,
        "message": message,
    }
    payload.update(extra)
    payload.setdefault("data", {k: v for k, v in extra.items() if k != "data"})
    return jsonify(payload), http_status
