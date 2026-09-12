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

# photo boxes (width_in, height_in).
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO_PATH = os.path.join(ROOT, "template", "assets", "logo.png")
LOGO_BOX = (2.8, 1.35)
FACE_BOX = (1.5, 1.6)
FULL_BOX = (3.50, 5.75)
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


def _crop_to_fill(im, target_w, target_h):
    """Smart center-crop so the image matches target aspect ratio and fills the container."""
    im_w, im_h = im.size
    target_aspect = target_w / float(target_h)
    im_aspect = im_w / float(im_h)
    if im_aspect > target_aspect:
        # Wider: crop extra sides, keep center
        new_w = int(im_h * target_aspect)
        left = (im_w - new_w) // 2
        return im.crop((left, 0, left + new_w, im_h))
    elif im_aspect < target_aspect:
        # Taller: crop extra top/bottom (keep 20% top, 80% bottom for headroom)
        new_h = int(im_w / target_aspect)
        top_offset = max(0, int((im_h - new_h) * 0.20))
        return im.crop((0, top_offset, im_w, top_offset + new_h))
    return im


def _safe_inline_image(tpl, source, box, cover=False):
    """
    Safely load any image (JPEG, PNG, WebP, BMP, TIFF, GIF, or PDF page),
    correct EXIF orientation, convert to RGB/RGBA, and return an InlineImage
    backed by a clean PNG BytesIO stream.
    
    If cover=True, crops the image to match the container's aspect ratio
    so it completely fills the box without distortion.
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

        box_w, box_h = box
        if cover:
            im = _crop_to_fill(im, box_w, box_h)
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            return InlineImage(tpl, buf, width=Inches(box_w), height=Inches(box_h))

        buf = io.BytesIO()
        im.save(buf, format="PNG")
        buf.seek(0)

        w, h = im.size
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
    ctx["photo_full"] = _safe_inline_image(tpl, photo_full, FULL_BOX, cover=True)
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
