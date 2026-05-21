# modules/pdf_editor.py
import base64
import io
import logging
import os
from datetime import datetime
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except:
    PYPDF_AVAILABLE = False

try:
    from PIL import Image
    PIL_AVAILABLE = True
except:
    PIL_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except:
    PYMUPDF_AVAILABLE = False

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PDF_EDITOR_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "pdf" / "editor"


def _next_available_path(path):
    candidate = Path(path)
    if not candidate.suffix:
        candidate = candidate.with_suffix(".pdf")
    candidate.parent.mkdir(parents=True, exist_ok=True)
    if not candidate.exists():
        return candidate
    idx = 1
    while True:
        next_path = candidate.with_name(f"{candidate.stem}_{idx}{candidate.suffix}")
        if not next_path.exists():
            return next_path
        idx += 1


def _safe_pdf_font(font_name):
    font = str(font_name or "").strip().lower()
    builtin = {
        "helv", "helvetica", "tiro", "times-roman", "cour", "courier",
        "symbol", "zapfdingbats",
    }
    if font in builtin:
        return "helv" if font == "helvetica" else font
    return "helv"


def _safe_color_tuple(value):
    color_hex = str(value or "#000000").strip().lstrip("#")
    if len(color_hex) != 6:
        color_hex = "000000"
    try:
        return tuple(int(color_hex[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception:
        return (0, 0, 0)


def open_pdf(path):
    if not PYMUPDF_AVAILABLE:
        return None, "PyMuPDF not installed. Run: pip install pymupdf"
    if not os.path.exists(path):
        return None, f"File not found: {path}"
    try:
        doc   = fitz.open(path)
        pages = _extract_all_pages(doc)
        doc.close()
        return {"total_pages": len(pages), "pages": pages, "file_path": path}, None
    except Exception as e:
        return None, str(e)


def _extract_all_pages(doc):
    pages = []
    for page_num in range(len(doc)):
        page   = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        text_blocks = []
        for block in blocks:
            if block.get("type") == 0:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "").strip()
                        if not text:
                            continue
                        bbox = span.get("bbox", [0, 0, 0, 0])
                        text_blocks.append({
                            "id":    f"p{page_num}_b{len(text_blocks)}",
                            "text":  text,
                            "x":     bbox[0],
                            "y":     bbox[1],
                            "x1":    bbox[2],
                            "y1":    bbox[3],
                            "font":  span.get("font", "helv"),
                            "size":  round(span.get("size", 12), 1),
                            "color": _int_to_hex(span.get("color", 0)),
                            "flags": span.get("flags", 0),
                        })
        pages.append({"text_blocks": text_blocks})
    return pages


def _int_to_hex(color_int):
    try:
        r = (color_int >> 16) & 0xFF
        g = (color_int >> 8)  & 0xFF
        b =  color_int        & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except:
        return "#000000"


def render_page(path, page_num, zoom=2.0):
    if not PYMUPDF_AVAILABLE:
        return None, "PyMuPDF not installed"
    try:
        doc  = fitz.open(path)
        page = doc[page_num]
        mat  = fitz.Matrix(zoom, zoom)
        pix  = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        doc.close()
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        return {
            "image":  f"data:image/png;base64,{b64}",
            "width":  pix.width,
            "height": pix.height,
            "zoom":   zoom,
        }, None
    except Exception as e:
        return None, str(e)


def save_with_edits(path, edits):
    if not PYMUPDF_AVAILABLE:
        return None, "PyMuPDF not installed"
    if not os.path.exists(path):
        return None, f"File not found: {path}"
    if not edits:
        return None, "No edits were provided."

    doc = None
    try:
        doc = fitz.open(path)

        # Group edits by page index so redactions are applied once per page.
        from collections import defaultdict
        edits_by_page = defaultdict(list)
        for edit in edits:
            page_index = int(edit.get("page", 0))
            if page_index < 0 or page_index >= len(doc):
                raise ValueError(f"Invalid page index: {page_index}")
            edits_by_page[page_index].append(edit)

        for page_index, page_edits in sorted(edits_by_page.items()):
            page = doc[page_index]
            # Validate bboxes and add all redaction annotations for this page first.
            pending = []
            for edit in page_edits:
                bbox_data = edit.get("bbox") or {}
                bbox = fitz.Rect(
                    float(bbox_data["x"]),  float(bbox_data["y"]),
                    float(bbox_data["x1"]), float(bbox_data["y1"])
                )
                if bbox.is_empty or bbox.is_infinite:
                    raise ValueError("Invalid edit bounding box.")
                page.add_redact_annot(bbox)
                pending.append((bbox, edit))
            # Apply all redactions for this page in one call.
            page.apply_redactions()
            # Now insert replacement text for each edit.
            for bbox, edit in pending:
                style     = edit.get("style", {})
                font_name = _safe_pdf_font(style.get("font", "helv"))
                font_size = max(1.0, float(style.get("size", 12) or 12))
                color     = _safe_color_tuple(style.get("color", "#000000"))
                page.insert_text(
                    fitz.Point(bbox.x0, bbox.y1),
                    str(edit.get("new_text", "")),
                    fontname=font_name,
                    fontsize=font_size,
                    color=color,
                )

        source = Path(path)
        preferred = source.with_name(f"{source.stem}_edited.pdf")
        out_path = _next_available_path(preferred)
        try:
            doc.save(str(out_path), garbage=4, deflate=True)
        except Exception as first_error:
            fallback = _next_available_path(PDF_EDITOR_OUTPUT_DIR / f"{source.stem}_edited.pdf")
            try:
                doc.save(str(fallback), garbage=4, deflate=True)
                out_path = fallback
            except Exception:
                raise first_error
        if not out_path.exists():
            return None, f"Save failed; output file was not created: {out_path}"
        return str(out_path.resolve()), None
    except Exception as e:
        return None, str(e)
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def detect_form_fields(pdf_path):
    if not PYPDF_AVAILABLE:
        return {}
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields() or {}
        return {name: field.get("/FT", "Unknown") for name, field in fields.items()}
    except Exception as e:
        logger.error(f"Form field detection error: {e}")
        return {}


def get_form_field_options(pdf_path, field_name):
    if not PYPDF_AVAILABLE:
        return []
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields() or {}
        field  = fields.get(field_name)
        if field and "/Opt" in field:
            return [str(o) for o in field["/Opt"]]
        return []
    except Exception as e:
        logger.error(f"Get options error: {e}")
        return []


def fill_form(pdf_path, form_data):
    if not PYPDF_AVAILABLE:
        return False
    try:
        from modules import pdf_utils
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        writer.append(reader)
        for page in writer.pages:
            writer.update_page_form_field_values(page, form_data)
        out_path = pdf_utils.ask(
            kind="savefile",
            defaultname="filled_form.pdf",
            title="Save Filled Form As",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if not out_path:
            return False
        with open(out_path, "wb") as f:
            writer.write(f)
        return out_path
    except Exception as e:
        logger.error(f"Fill form error: {e}")
        return False


# Backward-compatible API used by current server routes.
def extract_pdf_text(path):
    data, err = open_pdf(path)
    if err:
        return {"status": "fail", "message": err}
    return {
        "status": "success",
        "total_pages": data.get("total_pages", 0),
        "pages": data.get("pages", []),
        "file_path": data.get("file_path", path),
    }


def render_page_as_image(path, page_num=0):
    try:
        page_num = int(page_num or 0)
    except Exception:
        page_num = 0
    data, err = render_page(path, page_num, zoom=1.0)
    if err:
        return {"status": "fail", "message": err}
    payload = {"status": "success"}
    payload.update(data or {})
    return payload


def save_edited_pdf(path, edits):
    out_path, err = save_with_edits(path, edits or [])
    if err:
        return {"status": "fail", "message": err}
    return {
        "status": "success",
        "message": f"Saved edited PDF: {out_path}",
        "output_path": out_path,
        "file_path": out_path,
    }
