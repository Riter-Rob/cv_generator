# -*- coding: utf-8 -*-
"""
CV Generator - batch orchestrator.

For every row in the Excel sheet:
  1. read the passport image  -> OCR (offline)
  2. merge  Excel > OCR > default
  3. fill the Word template    -> output/docx/<name>.docx
  4. (optional) export PDF      -> output/pdf/<name>.pdf
and writes output/report.csv summarising sources + warnings.

Usage (from the cvgen folder, using the venv python):
    python src/generate.py                      # process input/applicants.xlsx
    python src/generate.py --no-pdf             # skip PDF export
    python src/generate.py --row 3              # only Excel row 3
    python src/generate.py --list               # show rows, generate nothing
"""
import os
import re
import sys
import csv
import shutil
import tempfile
import argparse
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import excel_io          # noqa: E402
import record            # noqa: E402
import docx_fill         # noqa: E402

DEF_EXCEL = os.path.join(ROOT, "input", "applicants.xlsx")
DEF_TEMPLATE = os.path.join(ROOT, "template", "cv_template.docx")
DEF_PASSPORTS = os.path.join(ROOT, "input", "passports")
DEF_PHOTOS = os.path.join(ROOT, "input", "photos")
DEF_OUT = os.path.join(ROOT, "output")


def sanitize(name):
    name = re.sub(r'[\\/:*?"<>|]+', " ", str(name)).strip()
    name = re.sub(r"\s+", " ", name)
    return name or "applicant"


def _rel(p):
    """Path for display; falls back to absolute if on a different drive."""
    try:
        return os.path.relpath(p, ROOT)
    except ValueError:
        return p


def resolve_file(folder, filename):
    """Locate a file given (possibly bare) filename inside folder."""
    if not filename:
        return None
    if os.path.isabs(filename) and os.path.exists(filename):
        return filename
    direct = os.path.join(folder, filename)
    if os.path.exists(direct):
        return direct
    # case-insensitive fallback
    if os.path.isdir(folder):
        low = filename.lower()
        for f in os.listdir(folder):
            if f.lower() == low:
                return os.path.join(folder, f)
        # try matching by stem (any extension)
        stem = os.path.splitext(low)[0]
        for f in os.listdir(folder):
            if os.path.splitext(f.lower())[0] == stem:
                return os.path.join(folder, f)
    return None


def main(argv=None):
    ap = argparse.ArgumentParser(description="Batch CV generator from passport + Excel")
    ap.add_argument("--excel", default=DEF_EXCEL)
    ap.add_argument("--template", default=DEF_TEMPLATE)
    ap.add_argument("--passports", default=DEF_PASSPORTS)
    ap.add_argument("--photos", default=DEF_PHOTOS)
    ap.add_argument("--out", default=DEF_OUT)
    ap.add_argument("--row", type=int, default=None, help="process only this Excel row number")
    ap.add_argument("--no-pdf", action="store_true", help="skip PDF export")
    ap.add_argument("--list", action="store_true", help="list rows only, generate nothing")
    args = ap.parse_args(argv)

    if not os.path.exists(args.excel):
        print("ERROR: Excel not found:", args.excel)
        print("Create it with:  python src/excel_io.py")
        return 2
    if not os.path.exists(args.template):
        print("ERROR: template not found:", args.template)
        print("Create it with:  python src/build_template.py")
        return 2

    rows = excel_io.read_applicants(args.excel)
    if args.row is not None:
        rows = [r for r in rows if r.get("_excel_row") == args.row]
    if not rows:
        print("No applicant rows found.")
        return 1

    print(f"Found {len(rows)} applicant row(s).")
    if args.list:
        for r in rows:
            print(f"  row {r.get('_excel_row')}: {r.get('output_name') or r.get('full_name')} "
                  f"| passport={r.get('passport_file')}")
        return 0

    # lazy import OCR (loads the model once)
    import ocr as ocr_mod

    docx_dir = os.path.join(args.out, "docx")
    pdf_dir = os.path.join(args.out, "pdf")
    os.makedirs(docx_dir, exist_ok=True)

    report_rows = []
    made_docx = []
    used_names = set()
    pdf_jobs = []
    for r in rows:
        rid = r.get("_excel_row")
        label = r.get("output_name") or r.get("full_name") or f"row{rid}"
        print(f"\n[row {rid}] {label}")

        # 1) passport OCR
        passport_path = resolve_file(args.passports, r.get("passport_file"))
        ocr_data = {}
        if passport_path:
            print(f"    passport: {os.path.basename(passport_path)}  -> OCR ...")
            try:
                ocr_data = ocr_mod.read_passport(passport_path)
            except Exception as e:
                print("    OCR ERROR:", e)
                ocr_data = {"_notes": [f"OCR error: {e}"]}
        else:
            if r.get("passport_file"):
                print(f"    WARNING: passport file not found: {r.get('passport_file')}")
            else:
                print("    no passport file given (Excel data only)")

        # 2) merge
        values, meta = record.merge(r, ocr_data)

        # name / output file
        out_name = sanitize(r.get("output_name") or values.get("full_name") or f"row{rid}")
        if out_name in used_names:
            out_name = f"{out_name} (row {rid})"
        used_names.add(out_name)

        # 3) photos
        face = resolve_file(args.photos, r.get("photo_face_file"))
        full = resolve_file(args.photos, r.get("photo_full_file"))

        # 4) fill docx
        out_docx = os.path.join(docx_dir, out_name + ".docx")
        try:
            docx_fill.fill_cv(args.template, values, out_docx, photo_face=face, photo_full=full)
        except Exception as e:
            print("    FILL ERROR:", e)
            traceback.print_exc()
            meta["warnings"].append(f"fill error: {e}")
        else:
            made_docx.append(out_docx)
            pdf_jobs.append((os.path.join(pdf_dir, out_name + ".pdf"), passport_path))
            print(f"    -> {_rel(out_docx)}")

        for w in meta["warnings"]:
            print("    ! " + w)

        report_rows.append({
            "excel_row": rid,
            "output_name": out_name,
            "passport_file": r.get("passport_file", ""),
            "mrz_found": ocr_data.get("_mrz_found", ""),
            "mrz_valid": ocr_data.get("_mrz_valid", ""),
            "full_name": values.get("full_name", ""),
            "passport_number": values.get("passport_number", ""),
            "face_photo": os.path.basename(face) if face else "",
            "full_photo": os.path.basename(full) if full else "",
            "warnings": " | ".join(meta["warnings"]),
        })

    # 5) PDF export (batch, one Word session) + passport as page 2
    pdf_ok = 0
    if not args.no_pdf and made_docx:
        print("\nExporting PDFs (via Microsoft Word) ...")
        try:
            import pdf_export
            # convert ONLY the docs made in this run (avoid touching stale files)
            staging = tempfile.mkdtemp(prefix="cvgen_pdf_")
            try:
                for d in made_docx:
                    shutil.copy(d, staging)
                pdf_export.folder_to_pdf(staging, pdf_dir)
            finally:
                shutil.rmtree(staging, ignore_errors=True)
            # append each applicant's passport as page 2
            for pdfp, passport_path in pdf_jobs:
                try:
                    pdf_export.append_passport(pdfp, passport_path)
                except Exception as e:
                    print(f"    passport-page warning for {os.path.basename(pdfp)}: {e}")
            pdf_ok = sum(1 for pdfp, _ in pdf_jobs if os.path.exists(pdfp))
            print(f"    {pdf_ok} PDF(s) in {_rel(pdf_dir)}")
        except Exception as e:
            print("    PDF export failed:", e)
            print("    (DOCX files are still available.)")

    # report
    report_path = os.path.join(args.out, "report.csv")
    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(report_rows[0].keys()))
        w.writeheader()
        w.writerows(report_rows)

    print("\n==== SUMMARY ====")
    print(f"  applicants   : {len(rows)}")
    print(f"  docx created : {len(made_docx)}")
    if not args.no_pdf:
        print(f"  pdf created  : {pdf_ok}")
    warned = sum(1 for r in report_rows if r["warnings"])
    print(f"  with warnings: {warned}")
    print(f"  report       : {_rel(report_path)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
