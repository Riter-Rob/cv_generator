# -*- coding: utf-8 -*-
"""
Fill the docx template for one applicant using docxtpl.

fill_cv(template_path, values, out_docx, photo_face=None, photo_full=None)
  - values      : {key: str}  (from record.merge)
  - photo_face  : path to the face photo (optional)
  - photo_full  : path to the full-body photo (optional)
"""
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# photo boxes (width_in, height_in). Images are fit INSIDE the box, preserving
# aspect ratio, so portrait and landscape both look right and a tall photo can
# never grow the form onto a second page.
FACE_BOX = (1.5, 1.6)
FULL_BOX = (2.35, 3.3)
PASSPORT_BOX = (6.5, 8.8)
_IMG_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".webp")


def _fit_image(tpl, path, box):
    """InlineImage scaled to fit within box (w_in, h_in), keeping aspect."""
    box_w, box_h = box
    try:
        from PIL import Image
        with Image.open(path) as im:
            w, h = im.size
        aspect = h / float(w) if w else 1.0
        if box_w * aspect <= box_h:      # width is the binding constraint
            return InlineImage(tpl, path, width=Inches(box_w))
        return InlineImage(tpl, path, height=Inches(box_h))  # height binds
    except Exception:
        return InlineImage(tpl, path, width=Inches(box_w))


def _fit_image_stream(tpl, stream, box):
    """InlineImage from a BytesIO stream scaled to fit within box (w_in, h_in)."""
    box_w, box_h = box
    try:
        from PIL import Image
        with Image.open(stream) as im:
            w, h = im.size
        stream.seek(0)
        aspect = h / float(w) if w else 1.0
        if box_w * aspect <= box_h:
            return InlineImage(tpl, stream, width=Inches(box_w))
        return InlineImage(tpl, stream, height=Inches(box_h))
    except Exception:
        stream.seek(0)
        return InlineImage(tpl, stream, width=Inches(box_w))


def _image_or_blank(tpl, path, box):
    if path and os.path.exists(path) and path.lower().endswith(_IMG_EXT):
        try:
            return _fit_image(tpl, path, box)
        except Exception:
            return ""
    return ""


def _passport_or_blank(tpl, path, box=PASSPORT_BOX):
    """Return InlineImage for an image or PDF passport, or blank string."""
    if not path or not os.path.exists(path):
        return ""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        try:
            import fitz
            import io
            doc = fitz.open(path)
            best_page = doc[0]
            if len(doc) > 1:
                for p in doc:
                    if p.get_images():
                        best_page = p
                        break
            pm = best_page.get_pixmap(dpi=200)
            doc.close()
            buf = io.BytesIO(pm.tobytes("png"))
            return _fit_image_stream(tpl, buf, box)
        except Exception:
            return ""
    elif ext in _IMG_EXT:
        return _image_or_blank(tpl, path, box)
    return ""


def fill_cv(template_path, values, out_docx, photo_face=None, photo_full=None, passport_path=None):
    tpl = DocxTemplate(template_path)
    ctx = dict(values)
    ctx["photo_face"] = _image_or_blank(tpl, photo_face, FACE_BOX)
    ctx["photo_full"] = _image_or_blank(tpl, photo_full, FULL_BOX)
    ctx["photo_passport"] = _passport_or_blank(tpl, passport_path, PASSPORT_BOX)
    # any placeholder not supplied -> blank, so render never fails
    for var in tpl.get_undeclared_template_variables():
        ctx.setdefault(var, "")
    out_dir = os.path.dirname(out_docx)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    tpl.render(ctx)
    tpl.save(out_docx)
    return out_docx
