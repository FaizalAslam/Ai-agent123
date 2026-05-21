from flask import Blueprint, jsonify
import state

voice_bp = Blueprint("voice", __name__)


@voice_bp.route("/voice/status", methods=["GET"])
def voice_status():
    import time
    vl = state._voice_listener
    if not vl:
        return jsonify(
            status="fail", available=False, enabled=False,
            message="Voice module unavailable. Install SpeechRecognition + PyAudio.",
        )
    heard = vl.last_heard
    if time.time() - (vl.last_heard_at or 0) > 8:
        heard = ""
    return jsonify(
        status="success",
        available=vl.available,
        enabled=vl.is_running,
        armed=vl.armed,
        armed_seconds=round(vl.armed_seconds_left, 1),
        heard=heard,
        error=vl.last_error,
    )


@voice_bp.route("/voice/start", methods=["POST"])
def voice_start():
    vl = state._voice_listener
    if not vl:
        return jsonify(status="fail", message="Voice module unavailable")
    ok = vl.start()
    state.voice_state["enabled"] = bool(ok)
    return jsonify(
        status="success" if ok else "fail",
        message="Voice listener started" if ok else (vl.last_error or "Could not start voice listener"),
    )


@voice_bp.route("/voice/stop", methods=["POST"])
def voice_stop():
    vl = state._voice_listener
    if not vl:
        return jsonify(status="fail", message="Voice module unavailable")
    vl.stop()
    state.voice_state["enabled"] = False
    return jsonify(status="success", message="Voice listener stopped")
