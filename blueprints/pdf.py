import threading
import time
import traceback
from pathlib import Path

from flask import Blueprint, jsonify, request

pdf_bp = Blueprint("pdf", __name__)


def _pdf_reader():
    from modules import pdf_reader as _m
    return _m


def _pdf_utils():
    from modules import pdf_utils as _m
    return _m


def _pdf_editor():
    from modules import pdf_editor as _m
    return _m


def _ui():
    from modules import ui as _m
    return _m


def _reader_available():
    try:
        _pdf_reader()
        return True
    except Exception:
        return False


def _pdf_available():
    try:
        _pdf_utils()
        return True
    except Exception:
        return False


def _editor_available():
    try:
        _pdf_editor()
        return True
    except Exception:
        return False


# ---- PDF Reader -----------------------------------------------------------

@pdf_bp.route("/reader/open", methods=["POST"])
def reader_open():
    try:
        if not _reader_available():
            return jsonify(status="fail", message="PDF reader module not found")
        path = _ui().file_selector("Select PDF to Read", [("PDFs", "*.pdf")])
        if not path:
            return jsonify(status="fail", message="No file selected")
        r = _pdf_reader()
        threading.Thread(target=r.start_reading, args=(path, 0), daemon=True).start()
        time.sleep(0.5)
        return jsonify(status="success", message="Reading started", **r.get_status())
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@pdf_bp.route("/reader/pause", methods=["POST"])
def reader_pause():
    try:
        r = _pdf_reader()
        r.pause_reading()
        return jsonify(status="success", message="Paused", **r.get_status())
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/reader/resume", methods=["POST"])
def reader_resume():
    try:
        r = _pdf_reader()
        r.resume_reading()
        return jsonify(status="success", message="Resumed", **r.get_status())
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/reader/stop", methods=["POST"])
def reader_stop():
    try:
        _pdf_reader().stop_reading()
        return jsonify(status="success", message="Stopped")
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/reader/next", methods=["POST"])
def reader_next():
    try:
        r = _pdf_reader()
        r.next_page()
        return jsonify(status="success", **r.get_status())
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/reader/prev", methods=["POST"])
def reader_prev():
    try:
        r = _pdf_reader()
        r.prev_page()
        return jsonify(status="success", **r.get_status())
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/reader/speed", methods=["POST"])
def reader_speed():
    try:
        data = request.json
        _pdf_reader().set_speed(data.get("speed", 150))
        return jsonify(status="success", message=f"Speed: {data.get('speed')} WPM")
    except Exception as e:
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/reader/status", methods=["GET"])
def reader_status():
    try:
        return jsonify(_pdf_reader().get_status())
    except Exception:
        return jsonify(is_reading=False, is_paused=False, current_page=0, total_pages=0, speed=150)


# ---- PDF Tools ------------------------------------------------------------

@pdf_bp.route("/pdf/merge", methods=["POST"])
def pdf_merge():
    try:
        if not _pdf_available():
            return jsonify(status="fail", message="Install pypdf: pip install pypdf")
        p = _pdf_utils()
        paths = p.ask(
            kind="openmultiple",
            title="Select PDFs to Merge (hold Ctrl for multiple)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if not paths:
            return jsonify(status="fail", message="No files selected.")
        out = p.merge_pdfs(paths)
        if not out:
            return jsonify(status="fail", message="Save cancelled.")
        return jsonify(status="success", message=f"Merged {len(paths)} PDFs → {out}")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@pdf_bp.route("/pdf/split", methods=["POST"])
def pdf_split():
    try:
        if not _pdf_available():
            return jsonify(status="fail", message="Install pypdf: pip install pypdf")
        path = _ui().file_selector("Select PDF to Split", [("PDFs", "*.pdf")])
        if not path:
            return jsonify(status="fail", message="No file selected.")
        pages = _pdf_utils().split_pdf(path)
        if not pages:
            return jsonify(status="fail", message="Save cancelled or no pages.")
        return jsonify(status="success", message=f"Split into {len(pages)} files")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@pdf_bp.route("/pdf/create", methods=["POST"])
def pdf_create():
    try:
        if not _pdf_available():
            return jsonify(status="fail", message="Install fpdf2: pip install fpdf2")
        data  = request.json
        text  = data.get("text", "").strip()
        title = (data.get("title", "Report") or "Report").strip()
        if not text:
            return jsonify(status="fail", message="No text provided")
        path = _pdf_utils().create_report(text, title=title)
        if not path:
            return jsonify(status="fail", message="Save cancelled.")
        return jsonify(status="success", message=f"PDF saved: {path}")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


# ---- PDF Editor -----------------------------------------------------------

@pdf_bp.route("/editor/open", methods=["POST"])
def editor_open():
    try:
        if not _editor_available():
            return jsonify(status="fail", message="PDF Editor not available")
        data = request.get_json(silent=True) or {}
        path = (data.get("file_path") or "").strip()
        if path and not Path(path).exists():
            return jsonify(status="fail", message=f"File not found: {path}")
        if not path:
            path = _ui().file_selector("Select PDF to Edit", [("PDF Files", "*.pdf")])
        if not path:
            return jsonify(status="fail", message="No file selected")
        ed   = _pdf_editor()
        data = ed.extract_pdf_text(path)
        if data.get("status") != "success":
            return jsonify(status="fail", message=data.get("message", "Failed to open PDF"))
        return jsonify(status="success", file_path=path, pages=data["pages"], total_pages=data["total_pages"])
    except Exception as e:
        import logging
        logging.error(traceback.format_exc())
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/editor/render-page", methods=["POST"])
def editor_render_page():
    try:
        if not _editor_available():
            return jsonify(status="fail", message="PDF Editor not available")
        data     = request.json
        pdf_path = data.get("file_path")
        page_num = data.get("page_num", 0)
        if not pdf_path:
            return jsonify(status="fail", message="No file path provided")
        result = _pdf_editor().render_page_as_image(pdf_path, page_num)
        if result.get("status") != "success":
            return jsonify(status="fail", message=result.get("message", "Render failed"))
        return jsonify(status="success", **{k: v for k, v in result.items() if k != "status"})
    except Exception as e:
        import logging
        logging.error(traceback.format_exc())
        return jsonify(status="fail", message=str(e))


@pdf_bp.route("/editor/save", methods=["POST"])
def editor_save():
    try:
        if not _editor_available():
            return jsonify(success=False, status="fail", intent="editor", error_code="PDF_EDITOR_UNAVAILABLE", message="PDF Editor not available"), 503
        data     = request.get_json(silent=True) or {}
        pdf_path = data.get("file_path")
        edits    = data.get("edits", [])
        if not pdf_path:
            return jsonify(success=False, status="fail", intent="editor", error_code="PDF_FILE_PATH_MISSING", message="No file path provided"), 400
        if not isinstance(edits, list) or not edits:
            return jsonify(success=False, status="fail", intent="editor", error_code="PDF_NO_EDITS", message="No PDF edits were queued."), 400
        result = _pdf_editor().save_edited_pdf(pdf_path, edits)
        if result.get("status") != "success":
            return jsonify(
                success=False,
                status="fail",
                intent="editor",
                error_code="PDF_SAVE_FAILED",
                message=result.get("message", "Save failed"),
            ), 500
        output_path = result.get("output_path") or result.get("file_path")
        if not output_path or not Path(output_path).exists():
            return jsonify(
                success=False,
                status="fail",
                intent="editor",
                error_code="PDF_SAVE_VERIFICATION_FAILED",
                message="Edited PDF save did not create a file on disk.",
                details=str(output_path or ""),
            ), 500
        return jsonify(
            success=True,
            status="success",
            intent="editor",
            message=result.get("message", "Saved successfully"),
            file_path=output_path,
            output_path=output_path,
            persisted=True,
        )
    except Exception as e:
        import logging
        logging.error(traceback.format_exc())
        return jsonify(success=False, status="fail", intent="editor", error_code="PDF_SAVE_ROUTE_ERROR", message=str(e)), 500


@pdf_bp.route("/editor/detect-form", methods=["POST"])
def editor_detect_form():
    try:
        if not _editor_available():
            return jsonify(status="fail", message="PDF Editor not available")
        path = _ui().file_selector("Select PDF", [("PDF Files", "*.pdf")])
        if not path:
            return jsonify(status="fail", message="No file selected")
        ed     = _pdf_editor()
        fields = ed.detect_form_fields(path)
        return jsonify(
            status="success",
            is_form=len(fields) > 0,
            field_count=len(fields),
            fields=list(fields.keys()),
            file_path=path,
        )
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@pdf_bp.route("/editor/fill-form", methods=["POST"])
def editor_fill_form():
    try:
        if not _editor_available():
            return jsonify(status="fail", message="PDF Editor not available")
        data      = request.json
        pdf_path  = data.get("file_path")
        form_data = data.get("form_data", {})
        if not pdf_path:
            return jsonify(status="fail", message="No file path provided")
        result = _pdf_editor().fill_form(pdf_path, form_data)
        if result:
            return jsonify(status="success", message="Form saved successfully")
        return jsonify(status="fail", message="Save cancelled")
    except Exception as e:
        print(traceback.format_exc())
        return jsonify(status="fail", message=f"Error: {e}")


@pdf_bp.route("/editor/get-field-options", methods=["POST"])
def editor_get_field_options():
    try:
        if not _editor_available():
            return jsonify(status="fail", message="PDF Editor not available")
        data       = request.json
        pdf_path   = data.get("file_path")
        field_name = data.get("field_name")
        options    = _pdf_editor().get_form_field_options(pdf_path, field_name)
        return jsonify(status="success", field_name=field_name, options=options)
    except Exception as e:
        return jsonify(status="fail", message=f"Error: {e}")
