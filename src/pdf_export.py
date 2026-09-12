# -*- coding: utf-8 -*-
"""
Export DOCX -> PDF.

Windows: uses Microsoft Word (docx2pdf).
Linux/macOS/cloud: falls back to LibreOffice headless (`soffice --convert-to pdf`).

to_pdf(src, dst)        : convert one file.
folder_to_pdf(in, out)  : convert every .docx in a folder (batch).
append_passport(pdf, passport) : rewrite pdf as [form] + [passport] page.
"""
import os
import glob
import shutil
import platform
import subprocess


def _ensure_com():
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _word_available():
    if platform.system() != "Windows":
        return False
    try:
        import docx2pdf  # noqa: F401
        return True
    except Exception:
        return False


def _soffice():
    """Locate the LibreOffice binary, or None."""
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/usr/bin/libreoffice",
              "/opt/libreoffice/program/soffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def _libre_convert(src, out_dir):
    exe = _soffice()
    if not exe:
        raise RuntimeError("Neither Microsoft Word nor LibreOffice is available "
                           "to export PDF. Install LibreOffice (see README).")
    os.makedirs(out_dir, exist_ok=True)
    import tempfile
    from pathlib import Path
    profile_dir = tempfile.mkdtemp(prefix="lo_prof_")
    profile_url = Path(profile_dir).as_uri()
    cmd = [
        exe,
        "--headless",
        "--invisible",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--nologo",
        f"-env:UserInstallation={profile_url}",
        "--convert-to", "pdf",
        "--outdir", out_dir,
        src,
    ]
    try:
        res = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=180,
        )
        if res.returncode != 0:
            err = (res.stderr or res.stdout or "").strip()
            raise RuntimeError(f"LibreOffice PDF export failed (exit code {res.returncode}): {err}")
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    expected = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if not os.path.exists(expected):
        raise RuntimeError(f"LibreOffice completed but expected PDF was not found: {expected}")
    return expected


def to_pdf(src, dst):
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    if _word_available():
        from docx2pdf import convert
        _ensure_com()
        convert(src, dst)
        return dst
    produced = _libre_convert(src, os.path.dirname(dst) or ".")
    if os.path.abspath(produced) != os.path.abspath(dst):
        shutil.move(produced, dst)
    return dst


def folder_to_pdf(in_dir, out_dir):
    """Convert all docx in in_dir to pdf in out_dir. Returns list of pdf paths.
    Skips Word lock files (~$...) and isolates per-file failures."""
    os.makedirs(out_dir, exist_ok=True)
    for lock in glob.glob(os.path.join(in_dir, "~$*")):
        try:
            os.remove(lock)
        except OSError:
            pass
    docs = [d for d in sorted(glob.glob(os.path.join(in_dir, "*.docx")))
            if not os.path.basename(d).startswith("~$")]
    if not docs:
        return []

    def _out(d):
        return os.path.join(out_dir, os.path.splitext(os.path.basename(d))[0] + ".pdf")

    if _word_available():
        from docx2pdf import convert
        _ensure_com()
        try:
            convert(in_dir, out_dir)           # one Word session (fast)
        except Exception:
            for d in docs:                     # one bad file shouldn't stop the rest
                try:
                    convert(d, _out(d))
                except Exception:
                    pass
    else:
        for d in docs:
            try:
                _libre_convert(d, out_dir)
            except Exception:
                pass

    return [_out(d) for d in docs]


# ---- passport page ----------------------------------------------------------
def _select_passport_page(doc):
    if len(doc) == 1:
        return doc[0]
    # If page 0 is an application form, the passport is on subsequent pages
    if "APPLICATION FOR EMPLOYMENT" in doc[0].get_text():
        for idx in range(1, len(doc)):
            if len(doc[idx].get_images()) > 0 or "PASSPORT" in doc[idx].get_text():
                return doc[idx]
        return doc[-1]
    # Check for MRZ on any page
    for idx, page in enumerate(doc):
        if "P<" in page.get_text():
            return page
    return doc[0]


def _passport_image_bytes(passport_path):
    """Return (png_bytes, width_px, height_px) for an image or pdf passport."""
    import fitz
    from PIL import Image, ImageOps
    import io

    ext = os.path.splitext(passport_path)[1].lower()
    if ext == ".pdf":
        try:
            doc = fitz.open(passport_path)
            page = _select_passport_page(doc)
            pm = page.get_pixmap(dpi=200)
            doc.close()
            return pm.tobytes("png"), pm.width, pm.height
        except Exception:
            pass

    try:
        with Image.open(passport_path) as im:
            im = ImageOps.exif_transpose(im)
            im = im.convert("RGB")
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            return buf.getvalue(), im.width, im.height
    except Exception:
        # Fallback: try fitz if it was a PDF with unusual extension
        try:
            doc = fitz.open(passport_path)
            page = _select_passport_page(doc)
            pm = page.get_pixmap(dpi=200)
            doc.close()
            return pm.tobytes("png"), pm.width, pm.height
        except Exception:
            raise RuntimeError(f"Could not read passport file: {passport_path}")


def _last_content_page(doc):
    last = 0
    for i in range(doc.page_count):
        page = doc[i]
        if page.get_text().strip() or page.get_images():
            last = i
    return last


def append_passport(pdf_path, passport_path):
    """Ensure pdf_path is exactly 2 pages: [page 1: form, page 2: passport].
    Discards any blank intermediary pages (e.g. from LibreOffice section breaks)."""
    import fitz
    if not os.path.exists(pdf_path):
        return pdf_path
    src = fitz.open(pdf_path)
    out = fitz.open()

    # Always keep Page 0 (the form)
    out.insert_pdf(src, from_page=0, to_page=0)

    # Check if any subsequent page in the converted PDF contains the real passport image
    passport_page_index = None
    for i in range(1, len(src)):
        page = src[i]
        placed_imgs = page.get_image_info()
        txt = page.get_text().upper()
        # Skip if this page is a form spillover
        if "APPLICATION FOR EMPLOYMENT" in txt:
            continue
        # A valid passport page must have an actual rendered image on it
        if len(placed_imgs) > 0 and ("PASSPORT" in txt or len(txt.strip()) < 150):
            w = placed_imgs[0].get("width", 0)
            h = placed_imgs[0].get("height", 0)
            if w >= 200 or h >= 200 or "PASSPORT" in txt:
                passport_page_index = i
                break

    if passport_page_index is not None:
        out.insert_pdf(src, from_page=passport_page_index, to_page=passport_page_index)
    elif passport_path and os.path.exists(passport_path):
        try:
            png, w, h = _passport_image_bytes(passport_path)
            page = out.new_page(width=595, height=842)  # A4 portrait (points)
            margin = 40
            title = "PASSPORT"
            page.insert_text((margin, margin), title, fontsize=12, fontname="hebo")
            top = margin + 18
            max_w, max_h = 595 - 2 * margin, 842 - top - margin
            scale = min(max_w / w, max_h / h)
            dw, dh = w * scale, h * scale
            x0 = (595 - dw) / 2
            rect = fitz.Rect(x0, top, x0 + dw, top + dh)
            page.insert_image(rect, stream=png)
        except Exception:
            pass  # keep the form even if passport image fails

    tmp = pdf_path + ".tmp"
    out.save(tmp)
    out.close()
    src.close()
    os.replace(tmp, pdf_path)
    return pdf_path
