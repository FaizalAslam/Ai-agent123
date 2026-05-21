import threading
import traceback

from flask import Blueprint, jsonify, request
import state

ocr_bp = Blueprint("ocr", __name__)


def _ocr_utils():
    from modules import ocr_utils as _m
    return _m


def _pdf_utils():
    from modules import pdf_utils as _m
    return _m


def _ui():
    from modules import ui as _m
    return _m


def _ocr_available():
    try:
        _ocr_utils()
        return True
    except Exception:
        return False


def _pdf_available():
    try:
        _pdf_utils()
        return True
    except Exception:
        return False


@ocr_bp.route("/ocr/snip", methods=["POST"])
def ocr_snip():
    try:
        if not _ocr_available():
            return jsonify(status="fail", message="OCR not available")
        ocr = _ocr_utils()
        ocr.snip_queue.put("snip")
        try:
            path = ocr.result_queue.get(timeout=60)
        except Exception:
            return jsonify(status="fail", message="Snip timed out")
        if not path:
            return jsonify(status="fail", message="Snip cancelled")
        text = ocr.image_to_text(path)
        state.last_ocr["text"]    = text
        state.last_ocr["pending"] = False
        return jsonify(status="success", text=text)
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@ocr_bp.route("/ocr/screenshot", methods=["POST"])
def ocr_screenshot():
    try:
        if not _ocr_available():
            return jsonify(status="fail", message="OCR not available")
        ocr = _ocr_utils()
        path = ocr.capture_fullscreen()
        text = ocr.image_to_text(path)
        state.last_ocr["text"]    = text
        state.last_ocr["pending"] = False
        return jsonify(status="success", text=text)
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@ocr_bp.route("/ocr/file", methods=["POST"])
def ocr_file():
    try:
        if not _ocr_available():
            return jsonify(status="fail", message="OCR not available")
        path = _ui().file_selector(
            "Select an Image File",
            [("Images", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("All Files", "*.*")],
        )
        if not path:
            return jsonify(status="fail", message="No file selected")
        ocr  = _ocr_utils()
        text = ocr.image_to_text(path)
        state.last_ocr["text"]    = text
        state.last_ocr["pending"] = False
        return jsonify(status="success", text=text, message=f"OCR complete — {len(text)} chars")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@ocr_bp.route("/ocr/read", methods=["POST"])
def ocr_read():
    try:
        text = state.last_ocr.get("text", "")
        if not text:
            return jsonify(status="fail", message="No OCR text. Run OCR first.")
        threading.Thread(target=_ocr_utils().speak_text, args=(text,), daemon=True).start()
        return jsonify(status="success", message="Speaking...")
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@ocr_bp.route("/ocr/stop_read", methods=["POST"])
def ocr_stop_read():
    try:
        _ocr_utils().stop_speaking()
        return jsonify(status="success", message="Stopped")
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@ocr_bp.route("/ocr/poll", methods=["GET"])
def ocr_poll():
    try:
        if state.last_ocr.get("pending"):
            state.last_ocr["pending"] = False
            return jsonify(
                status="ready",
                text=state.last_ocr["text"],
                message=f"Hotkey OCR complete — {len(state.last_ocr['text'])} chars",
            )
        return jsonify(status="waiting")
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@ocr_bp.route("/ocr/save_txt", methods=["POST"])
def ocr_save_txt():
    try:
        text = state.last_ocr.get("text", "")
        if not text:
            return jsonify(status="fail", message="No OCR text. Run OCR first.")
        path = _ocr_utils().save_as_txt(text)
        if not path:
            return jsonify(status="fail", message="Save cancelled.")
        return jsonify(status="success", message=f"Saved: {path}")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@ocr_bp.route("/ocr/save_pdf", methods=["POST"])
def ocr_save_pdf():
    try:
        text = state.last_ocr.get("text", "")
        if not text:
            return jsonify(status="fail", message="No OCR text. Run OCR first.")
        if not _pdf_available():
            return jsonify(status="fail", message="Install fpdf2: pip install fpdf2")
        path = _pdf_utils().create_report(text, title="OCR Result")
        if not path:
            return jsonify(status="fail", message="Save cancelled.")
        return jsonify(status="success", message=f"Saved: {path}")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@ocr_bp.route("/ocr/clipboard", methods=["POST"])
def ocr_clipboard():
    try:
        text = state.last_ocr.get("text", "")
        if not text:
            return jsonify(status="fail", message="No OCR text. Run OCR first.")
        _ocr_utils().copy_to_clipboard(text)
        return jsonify(status="success", message="Copied to clipboard")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")
