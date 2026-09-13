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
FACE_BOX = (1.388, 1.758)
FULL_BOX = (3.523, 5.799)
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


def _autotrim_whitespace(im, thresh=235):
    """
    Trim surrounding whitespace or screenshot borders from uploaded images so
    the actual portrait or photo cleanly fills the container without an off-center shift.
    """
    try:
        im_rgb = im.convert("RGB")
        w, h = im_rgb.size
        xs, ys = [], []
        # Sample every 2 pixels for fast detection
        for y in range(0, h, 2):
            for x in range(0, w, 2):
                r, g, b = im_rgb.getpixel((x, y))
                if r < thresh or g < thresh or b < thresh:
                    xs.append(x)
                    ys.append(y)
        if not xs or not ys:
            return im
        min_x, max_x = max(0, min(xs) - 1), min(w, max(xs) + 2)
        min_y, max_y = max(0, min(ys) - 1), min(h, max(ys) + 2)
        # Only trim if there is significant outer border whitespace (> 10px)
        if min_x > 10 or min_y > 10 or (w - max_x) > 10 or (h - max_y) > 10:
            return im.crop((min_x, min_y, max_x, max_y))
    except Exception:
        pass
    return im


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
            im = _autotrim_whitespace(im)
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


def _zero_cell_margins(tc):
    """Zero all margins and indents on a table cell and its paragraphs in the output doc."""
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.text.paragraph import Paragraph

    tcPr = tc.get_or_add_tcPr()
    for m in tcPr.findall(_qn("w:tcMar")):
        tcPr.remove(m)
    tcMar = _OE("w:tcMar")
    for side in ("top", "left", "bottom", "right"):
        node = _OE(f"w:{side}")
        node.set(_qn("w:w"), "0")
        node.set(_qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)

    for v in tcPr.findall(_qn("w:vAlign")):
        tcPr.remove(v)
    vAlign = _OE("w:vAlign")
    vAlign.set(_qn("w:val"), "center")
    tcPr.append(vAlign)

    for p in tc.findall(_qn("w:p")):
        Paragraph(p, None).alignment = WD_ALIGN_PARAGRAPH.CENTER
        pPr = p.find(_qn("w:pPr"))
        if pPr is None:
            pPr = _OE("w:pPr")
            p.insert(0, pPr)
        for ind in pPr.findall(_qn("w:ind")):
            pPr.remove(ind)
        for sp in pPr.findall(_qn("w:spacing")):
            pPr.remove(sp)
        sp = _OE("w:spacing")
        sp.set(_qn("w:before"), "0")
        sp.set(_qn("w:after"), "0")
        sp.set(_qn("w:line"), "240")
        sp.set(_qn("w:lineRule"), "auto")
        pPr.append(sp)

        # Remove empty runs before the drawing run so nothing shifts the image left
        runs = p.findall(_qn("w:r"))
        has_drawing = any(r.find(_qn("w:drawing")) is not None for r in runs)
        if has_drawing:
            for r in runs:
                if r.find(_qn("w:drawing")) is None:
                    # Empty/text run before the image — remove it
                    t_els = r.findall(_qn("w:t"))
                    if not t_els or all((t.text or "").strip() == "" for t in t_els):
                        p.remove(r)


def _fix_photo_cells_in_output(docx_path):
    """Post-process the rendered output DOCX to ensure photo cells have zero margins."""
    import docx as _docx
    from docx.oxml.ns import qn as _qn

    doc = _docx.Document(docx_path)
    tbl = None
    for t in doc.element.body.xpath(".//w:tbl"):
        for a in t.iterancestors():
            if a.tag.split("}")[-1] == "Choice":
                tbl = t
                break
        if tbl is not None:
            break
    if tbl is None:
        tbls = doc.element.body.xpath(".//w:tbl")
        tbl = tbls[0] if tbls else None
    if tbl is None:
        return

    rows = tbl.findall(_qn("w:tr"))
    try:
        _zero_cell_margins(rows[1].findall(_qn("w:tc"))[2])   # face photo cell
    except (IndexError, Exception):
        pass
    try:
        _zero_cell_margins(rows[4].findall(_qn("w:tc"))[1])   # full-body photo cell
    except (IndexError, Exception):
        pass
    doc.save(docx_path)


def fill_cv(template_path, values, out_docx, photo_face=None, photo_full=None,
            passport_path=None, logo_path=None):
    tpl = DocxTemplate(template_path)
    ctx = dict(values)
    ctx["photo_logo"] = _safe_inline_image(tpl, logo_path or LOGO_PATH, LOGO_BOX)
    ctx["photo_face"] = _safe_inline_image(tpl, photo_face, FACE_BOX, cover=True)
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

    # Post-process: zero cell margins on photo cells in the rendered output
    try:
        _fix_photo_cells_in_output(out_docx)
    except Exception:
        pass

    return out_docx

