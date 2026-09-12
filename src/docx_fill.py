# -*- coding: utf-8 -*-
"""
Fill the docx template for one applicant using docxtpl.

fill_cv(template_path, values, out_docx, photo_face=None, photo_full=None)
  - values      : {key: str}  (from record.merge)
  - photo_face  : path to the face photo (optional)
  - photo_full  : path to the full-body photo (optional)
"""
import os
import io
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image, ImageOps

# photo boxes (width_in, height_in). Images are fit INSIDE the box, preserving
# aspect ratio, so portrait and landscape both look right and a tall photo can
# never grow the form onto a second page.
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO_PATH = os.path.join(ROOT, "template", "assets", "logo.png")
LOGO_BOX = (2.8, 1.35)
FACE_BOX = (1.5, 1.6)
FULL_BOX = (2.45, 5.8)
PASSPORT_BOX = (6.0, 7.5)


def _select_passport_page(doc):
    if len(doc) == 1:
        return doc[0]
    if "APPLICATION FOR EMPLOYMENT" in doc[0].get_text():
        for idx in range(1, len(doc)):
            if len(doc[idx].get_images()) > 0 or "PASSPORT" in doc[idx].get_text():
                return doc[idx]
        return doc[-1]
    for idx, page in enumerate(doc):
        if "P<" in page.get_text():
            return page
    return doc[0]


def _safe_inline_image(tpl, source, box):
    """
    Safely load any image (JPEG, PNG, WebP, BMP, TIFF, GIF, or PDF page),
    correct EXIF orientation, convert to RGB/RGBA, and return an InlineImage
    backed by a clean PNG BytesIO stream.
    
    If the image is missing, corrupt, or cannot be read, returns "" (empty string)
    so python-docx NEVER encounters an UnrecognizedImageError.
    """
    if not source:
        return ""
    if isinstance(source, str) and not os.path.exists(source):
        return ""

    try:
        if isinstance(source, str) and source.lower().endswith(".pdf"):
            import fitz
            doc = fitz.open(source)
            page = _select_passport_page(doc)
            pm = page.get_pixmap(dpi=200)
            doc.close()
            im = Image.open(io.BytesIO(pm.tobytes("png")))
        elif isinstance(source, str):
            im = Image.open(source)
        else:
            im = Image.open(source)

        # Fix mobile camera orientation
        im = ImageOps.exif_transpose(im)

        # Normalize mode for clean PNG generation
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            im = im.convert("RGBA")
        else:
            im = im.convert("RGB")

        buf = io.BytesIO()
        im.save(buf, format="PNG")
        buf.seek(0)

        w, h = im.size
        box_w, box_h = box
        aspect = h / float(w) if w else 1.0
        if box_w * aspect <= box_h:
            return InlineImage(tpl, buf, width=Inches(box_w))
        return InlineImage(tpl, buf, height=Inches(box_h))
    except Exception:
        return ""


def fill_cv(template_path, values, out_docx, photo_face=None, photo_full=None,
            passport_path=None, logo_path=None):
    tpl = DocxTemplate(template_path)
    ctx = dict(values)
    ctx["photo_logo"] = _safe_inline_image(tpl, logo_path or LOGO_PATH, LOGO_BOX)
    ctx["photo_face"] = _safe_inline_image(tpl, photo_face, FACE_BOX)
    ctx["photo_full"] = _safe_inline_image(tpl, photo_full, FULL_BOX)
    ctx["photo_passport"] = _safe_inline_image(tpl, passport_path, PASSPORT_BOX)

    # any placeholder not supplied -> blank, so render never fails
    for var in tpl.get_undeclared_template_variables():
        ctx.setdefault(var, "")
    out_dir = os.path.dirname(out_docx)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        tpl.render(ctx)
    except Exception:
        # Failsafe: if an image somehow still triggers an error, clear image placeholders
        for k in ("photo_face", "photo_full", "photo_passport", "photo_logo"):
            ctx[k] = ""
        tpl.render(ctx)

    tpl.save(out_docx)
    return out_docx
