# -*- coding: utf-8 -*-
"""
CV Generator - local desktop app (Streamlit).

Launch:
    .\run_app.ps1              (or double-click run_app.bat)
or:
    .venv\Scripts\streamlit run src\app.py

Two tabs:
  - Single applicant : upload a passport -> OCR auto-fills the form ->
    complete the rest -> add photos -> Generate -> preview + download.
  - Batch (Excel)    : use/upload applicants.xlsx -> Generate all -> report + zip.
"""
import os
import sys
import io
import zipfile
import tempfile
import contextlib

import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import fields          # noqa: E402
import ocr             # noqa: E402
import record          # noqa: E402
import docx_fill       # noqa: E402
import pdf_export      # noqa: E402
import excel_io        # noqa: E402
import generate        # noqa: E402

TEMPLATE = os.path.join(ROOT, "template", "cv_template.docx")
LOGO = os.path.join(ROOT, "template", "assets", "logo.png")
IN_PASS = os.path.join(ROOT, "input", "passports")
IN_PHOTO = os.path.join(ROOT, "input", "photos")
EXCEL = os.path.join(ROOT, "input", "applicants.xlsx")
OUT = os.path.join(ROOT, "output")
TMP = os.path.join(tempfile.gettempdir(), "cvgen_app")
for d in (IN_PASS, IN_PHOTO, os.path.join(OUT, "docx"), os.path.join(OUT, "pdf"), TMP):
    os.makedirs(d, exist_ok=True)

# build the template from original.docx if it hasn't been generated yet
# (so a fresh deploy works without running setup first)
if not os.path.exists(TEMPLATE):
    try:
        import build_template
        build_template.build()
    except Exception:
        pass

GROUP_ORDER = ["top", "passport", "lang", "prev", "personal", "physical"]
GROUP_TITLE = {
    "top": "Applicant / Job",
    "passport": "Passport details  (auto from OCR - editable)",
    "lang": "Languages & education",
    "prev": "Previous employment abroad",
    "personal": "Personal data",
    "physical": "Physical & status",
}

# fields rendered as dropdowns in the Single-applicant form
CHOICES = {
    "religion": ["", "MUSLIM", "CHRISTIAN"],
}

st.set_page_config(page_title="CV Generator", page_icon="🪪", layout="wide")


IMG_TYPES = ["jpg", "jpeg", "png", "webp", "jfif", "bmp", "tiff"]
PASS_TYPES = ["jpg", "jpeg", "png", "webp", "jfif", "bmp", "tiff", "pdf"]


# ---------- helpers -----------------------------------------------------------
def save_upload(uploaded, folder, name=None):
    if uploaded is None:
        return None
    fname = name or getattr(uploaded, "name", "upload")
    fname = os.path.basename(fname).replace("/", "_").replace("\\", "_")
    path = os.path.join(folder, fname)
    try:
        uploaded.seek(0)
    except Exception:
        pass
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    try:
        uploaded.seek(0)
    except Exception:
        pass
    return path


def pdf_to_pngs(pdf_path, dpi=110):
    import fitz
    doc = fitz.open(pdf_path)
    pages = [p.get_pixmap(dpi=dpi).tobytes("png") for p in doc]
    doc.close()
    return pages


def zip_folder_bytes(folder):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for r, _, files in os.walk(folder):
            for f in files:
                fp = os.path.join(r, f)
                z.write(fp, os.path.relpath(fp, folder))
    buf.seek(0)
    return buf.getvalue()


def init_state():
    for key in fields.ALL_KEYS:
        st.session_state.setdefault(key, fields.DEFAULTS.get(key, ""))
    st.session_state.setdefault("output_name", "")


def _resolve_passport_path():
    up = st.session_state.get("passport_up")
    pick = st.session_state.get("passport_pick")
    if up is not None:
        p = save_upload(up, TMP)
        if p and os.path.exists(p):
            st.session_state["_cached_passport_path"] = p
            return p
    if pick:
        p = os.path.join(IN_PASS, pick)
        if os.path.exists(p):
            st.session_state["_cached_passport_path"] = p
            return p
    cached = st.session_state.get("_cached_passport_path")
    if cached and os.path.exists(cached):
        return cached
    return None


def read_passport_cb():
    path = _resolve_passport_path()
    if not path or not os.path.exists(path):
        st.session_state["_ocr_msg"] = ("warning", "Upload or pick a passport first.")
        return
    st.session_state["_cached_passport_path"] = path
    try:
        data = ocr.read_passport(path)
        st.session_state["_ocr"] = data
        for k in list(fields.PASSPORT_KEYS) + ["age"]:
            v = data.get(k, "")
            if v:
                st.session_state[k] = str(v)
        if data.get("_mrz_found"):
            ok = "valid" if data.get("_mrz_valid") else "CHECK-DIGIT FAILED"
            st.session_state["_ocr_msg"] = ("ok", f"Passport read - MRZ {ok}.")
        else:
            st.session_state["_ocr_msg"] = ("warning",
                                            "MRZ not detected - type the passport fields manually.")
    except Exception as e:
        st.session_state["_ocr_msg"] = ("warning", f"Could not read passport: {e}")


# ---------- single applicant tab ---------------------------------------------
def single_tab():
    init_state()
    left, right = st.columns([1, 2], gap="large")

    with left:
        st.subheader("1) Passport")
        st.file_uploader("Upload passport (image or PDF)",
                         type=PASS_TYPES, key="passport_up")
        existing = sorted([f for f in os.listdir(IN_PASS)
                           if f.lower().endswith(tuple("." + ext for ext in PASS_TYPES))])
        st.selectbox("...or pick an existing passport", [""] + existing, key="passport_pick")
        st.button("Read passport & auto-fill", type="secondary", on_click=read_passport_cb)

        msg = st.session_state.get("_ocr_msg")
        if msg:
            (st.success if msg[0] == "ok" else st.warning)(msg[1])
        data = st.session_state.get("_ocr")
        if data and data.get("_low_confidence"):
            st.caption("Verify these OCR guesses: " + ", ".join(data["_low_confidence"]))

        st.subheader("2) Photos (optional)")
        st.file_uploader("Face photo", type=IMG_TYPES, key="face_up")
        st.file_uploader("Full-body photo", type=IMG_TYPES, key="full_up")

    with right:
        st.subheader("3) Details")
        st.text_input("Output file name", key="output_name",
                      placeholder="defaults to the full name")

        by_group = {g: [] for g in GROUP_ORDER}
        for k, en, ar, src, d, g in fields.FIELDS:
            if g in by_group:
                by_group[g].append((k, en, ar))

        for g in GROUP_ORDER:
            with st.expander(GROUP_TITLE[g], expanded=g in ("top", "passport", "personal")):
                cols = st.columns(2)
                for i, (k, en, ar) in enumerate(by_group[g]):
                    col = cols[i % 2]
                    if k in CHOICES:
                        opts = CHOICES[k]
                        cur = st.session_state.get(k, "")
                        if cur not in opts:
                            opts = [cur] + opts if cur else opts
                        col.selectbox(en, opts, key=k, help=ar)
                    else:
                        col.text_input(f"{en}", key=k, help=ar)

        with st.expander("Skills & experiences", expanded=True):
            scols = st.columns(3)
            for i, (k, en, ar, d) in enumerate(fields.SKILLS):
                scols[i % 3].radio(en, ["YES", "NO"], key=k, horizontal=True, help=ar)

        with st.expander("Profile summary", expanded=False):
            st.text_area("Profile summary", key="profile_summary", label_visibility="collapsed")

    st.divider()
    gcol1, gcol2 = st.columns([1, 3])
    engine_name = "Microsoft Word" if sys.platform == "win32" else "LibreOffice"
    make_pdf = gcol2.checkbox(f"Also export PDF (uses {engine_name})", value=True, key="single_pdf")
    if gcol1.button("Generate CV", type="primary", use_container_width=True):
        _generate_single(make_pdf)

    _render_single_result()


def _generate_single(make_pdf):
    values = {}
    for k in fields.ALL_KEYS:
        v = st.session_state.get(k, "")
        values[k] = record._norm_skill(v) if k in fields.SKILL_KEYS else str(v).strip()

    face_up = st.session_state.get("face_up")
    full_up = st.session_state.get("full_up")
    face = save_upload(face_up, TMP, "face_" + os.path.basename(face_up.name)) if face_up else None
    full = save_upload(full_up, TMP, "full_" + os.path.basename(full_up.name)) if full_up else None

    out_name = generate.sanitize(st.session_state.get("output_name")
                                 or values.get("full_name") or "applicant")
    docx_path = os.path.join(OUT, "docx", out_name + ".docx")
    passport = _resolve_passport_path()
    engine_name = "Microsoft Word" if sys.platform == "win32" else "LibreOffice"
    try:
        with st.spinner("Building document..."):
            docx_fill.fill_cv(TEMPLATE, values, docx_path, photo_face=face, photo_full=full,
                              passport_path=passport)
        pdf_path = None
        if make_pdf:
            with st.spinner(f"Exporting PDF via {engine_name}..."):
                pdf_path = os.path.join(OUT, "pdf", out_name + ".pdf")
                pdf_export.to_pdf(docx_path, pdf_path)
                if passport:
                    pdf_export.append_passport(pdf_path, passport)
        st.session_state["single_result"] = {"name": out_name, "docx": docx_path, "pdf": pdf_path}
    except PermissionError:
        st.session_state["single_result"] = None
        st.error(f"Could not save '{out_name}'. The file is probably open - "
                 f"close the .docx/.pdf (and any Word window) and click Generate again.")
    except Exception as e:
        st.session_state["single_result"] = None
        import traceback
        err_type = type(e).__name__
        err_msg = str(e).strip()
        display_err = f"{err_type}: {err_msg}" if err_msg else err_type
        st.error(f"Generation failed: {display_err}")
        with st.expander("Error details"):
            st.code(traceback.format_exc())


def _render_single_result():
    res = st.session_state.get("single_result")
    if not res:
        return
    st.success(f"Generated: {res['name']}")
    c1, c2 = st.columns(2)
    if os.path.exists(res["docx"]):
        with open(res["docx"], "rb") as f:
            c1.download_button("Download .docx", f.read(), file_name=res["name"] + ".docx",
                               use_container_width=True)
    if res.get("pdf") and os.path.exists(res["pdf"]):
        with open(res["pdf"], "rb") as f:
            c2.download_button("Download .pdf", f.read(), file_name=res["name"] + ".pdf",
                               type="primary", use_container_width=True)
        st.subheader("Preview")
        try:
            for png in pdf_to_pngs(res["pdf"]):
                st.image(png, use_container_width=True)
        except Exception:
            st.info("PDF saved; preview unavailable.")
    else:
        st.info("DOCX saved in output/docx. Enable PDF export for a preview.")


# ---------- batch tab ---------------------------------------------------------
def batch_tab():
    st.subheader("Batch from Excel")
    st.write("Use the project Excel sheet, or upload one. One row per applicant; "
             "passport images live in `input/passports/`.")

    up = st.file_uploader("Upload applicants.xlsx (optional)", type=["xlsx"], key="excel_up")
    excel_path = EXCEL
    if up is not None:
        excel_path = os.path.join(TMP, "applicants_uploaded.xlsx")
        save_upload(up, TMP, "applicants_uploaded.xlsx")

    cta1, cta2 = st.columns(2)
    if not os.path.exists(EXCEL):
        if cta2.button("Create starter applicants.xlsx"):
            excel_io.build_input_template(EXCEL)
            st.success("Created input/applicants.xlsx")

    if os.path.exists(excel_path):
        try:
            rows = excel_io.read_applicants(excel_path)
            st.caption(f"{len(rows)} applicant row(s) detected")
            import pandas as pd
            if rows:
                df = pd.DataFrame(rows).drop(columns=["_excel_row"], errors="ignore")
                st.dataframe(df, use_container_width=True, height=240)
        except Exception as e:
            st.error(f"Could not read Excel: {e}")
            rows = []
    else:
        st.warning("No Excel found. Upload one or create the starter sheet.")
        rows = []

    engine_name = "Microsoft Word" if sys.platform == "win32" else "LibreOffice"
    make_pdf = st.checkbox(f"Export PDFs (uses {engine_name})", value=True, key="batch_pdf")
    if st.button("Generate all", type="primary", disabled=not rows):
        argv = ["--excel", excel_path, "--out", OUT]
        if not make_pdf:
            argv.append("--no-pdf")
        buf = io.StringIO()
        with st.spinner("Generating... (OCR + fill + PDF)"):
            try:
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                    generate.main(argv)
                st.session_state["batch_log"] = buf.getvalue()
                st.session_state["batch_done"] = True
            except Exception as e:
                st.session_state["batch_log"] = buf.getvalue() + f"\nERROR: {e}"
                st.session_state["batch_done"] = False

    if st.session_state.get("batch_done"):
        st.success("Batch complete.")
        report = os.path.join(OUT, "report.csv")
        if os.path.exists(report):
            import pandas as pd
            st.subheader("Report")
            st.dataframe(pd.read_csv(report), use_container_width=True, height=240)
        st.download_button("Download all results (.zip)", zip_folder_bytes(OUT),
                           file_name="cv_output.zip", type="primary")
    if st.session_state.get("batch_log"):
        with st.expander("Log"):
            st.code(st.session_state["batch_log"])


# ---------- layout ------------------------------------------------------------
def main():
    with st.sidebar:
        if os.path.exists(LOGO):
            st.image(LOGO, use_column_width=True)
        st.markdown("### CV Generator")
        st.caption("Passport OCR + Excel data -> filled Application form (DOCX + PDF). "
                   "Runs entirely on this PC.")
        st.markdown("**Folders**")
        st.caption(f"passports: `input/passports/`\n\nphotos: `input/photos/`\n\noutput: `output/`")
        if not os.path.exists(TEMPLATE):
            st.error("Template missing. Run: python src/build_template.py")

    st.title("Recruitment CV Generator")
    t1, t2 = st.tabs(["Single applicant", "Batch (Excel)"])
    with t1:
        single_tab()
    with t2:
        batch_tab()


main()
