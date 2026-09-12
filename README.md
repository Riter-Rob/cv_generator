# CV Generator (Passport + Excel -> filled Application form)

Automated batch generator for the bilingual (English / Arabic) **Application
for Employment** form. For each applicant it:

1. reads the **passport** image/PDF with **offline OCR** (no internet, no API key)
   and extracts the reliable fields from the passport's Machine-Readable Zone;
2. reads the **external data** you type in an **Excel** sheet (skills, salary,
   languages, etc. - the things that are not on the passport);
3. merges them (your Excel values always win over OCR), inserts the photos, and
   produces a filled **`.docx` + `.pdf`** per person that matches the template.

```
passport image/pdf ─► OCR (MRZ + printed) ─┐
Excel row (skills, salary, ...)  ──────────┼─► merge ─► fill template ─► DOCX ─► PDF
face + full-body photos  ──────────────────┘
```

## Requirements

- Windows with **Microsoft Word** installed (used to export PDF).
- **Python 3.10+** on PATH.
- First run downloads Python packages (~a few hundred MB, incl. the OCR engine).

## One-time setup

From this `cvgen` folder in PowerShell:

```powershell
.\setup.ps1
```

This creates `.venv`, installs everything in `requirements.txt`, and builds:
- `template\cv_template.docx`  - the form template (edit in Word if you like)
- `input\applicants.xlsx`      - the data sheet (already has one example row)

## Two ways to use it

- **Desktop app (easiest)** — a point-and-click app in your browser (see below).
- **Command line / batch** — great for processing a whole Excel at once (further down).

## Desktop app

Launch it:

```powershell
.\run_app.ps1
```

(or double-click `run_app.bat`). It opens in your browser at `http://localhost:8531`.

Two tabs:

- **Single applicant** — upload (or pick) a passport, click **Read passport & auto-fill**
  (offline OCR fills the passport fields), complete the remaining fields, optionally add
  face/full-body photos, then **Generate CV** to preview and download the DOCX + PDF.
- **Batch (Excel)** — use the project `applicants.xlsx` (or upload one), review the rows,
  then **Generate all** to produce every CV and download them as a single `.zip`, plus the report.

**Use it from a phone on the same Wi-Fi:** when the app starts it also prints a
*Network URL* like `http://192.168.1.10:8531` — open that on your phone's browser.
(Allow the app through Windows Firewall if prompted; the PC must stay on.)

## Use it on other devices

**Option 1 - one PC, everyone else uses a browser (no install).**
Run `.\run_app.ps1` on one office PC. Other PCs *and phones* on the same network
open the **Network URL** it prints (e.g. `http://192.168.1.10:8531`). Nothing to
install on the other devices. Keep that PC on; allow it through Windows Firewall
once. The port is fixed to `8531` (change it in `config.json`).

**Option 2 - install on another PC.**
1. On this PC run `.\make_portable.ps1` -> creates `cvgen_portable_<date>.zip`
   (everything except the virtual environment and outputs).
2. Copy the zip to the other PC (which needs **Python 3.10+** and **Microsoft Word**).
3. Unzip, then run `.\setup.ps1` (rebuilds the environment) and `.\run_app.ps1`.

The OCR models ship inside the Python packages, so no extra downloads are needed
at runtime - only the one-time `pip install` in `setup.ps1` needs internet.
(For a fully offline target PC, run `pip download -r requirements.txt -d wheels`
here, copy the `wheels` folder too, and install from it.)

For remote/internet access beyond the office network, host the app in the cloud
(see **Deploying** below).

## Deploying (GitHub + cloud)

The app runs on any Linux host. On Windows it exports PDF via Microsoft Word; on
Linux/cloud it automatically uses **LibreOffice** instead (same result). The
passport OCR and form filling are identical on both.

**Privacy:** the repo intentionally excludes personal data - passports, photos,
generated output and the working `applicants.xlsx` are git-ignored. Keep the repo
**private** since it contains your agency form and logo.

### Option 1 - Streamlit Community Cloud (free, easiest)
1. Push this folder to a GitHub repo (already done if you're reading this there).
2. Go to https://share.streamlit.io -> **New app**, pick the repo/branch.
3. Set **Main file path** to `src/app.py` and deploy.
   `packages.txt` installs LibreOffice + fonts automatically; `requirements.txt`
   installs the Python deps. First build takes a few minutes.

### Option 2 - Docker (any server/VPS)
```bash
docker build -t cvgen .
docker run -p 8501:8501 cvgen
```
Open http://SERVER_IP:8501 . Mount a volume to keep output, e.g.
`-v $PWD/output:/app/output`.

### Option 3 - Office PC (no cloud)
Just run `./run_app.ps1` on one Windows PC (uses Word); everyone else opens the
Network URL it prints. See "Use it on other devices" above.

## Daily use (command line)

1. **Add passports** — drop each applicant's passport in `input\passports\`.
   Accepted: `.jpg`, `.png`, `.pdf` (multi-page PDFs are fine - it finds the
   passport page automatically).
2. **Add photos (optional)** — drop applicant photos in `input\photos\`.
3. **Fill the Excel** — open `input\applicants.xlsx`, one row per applicant:
   - `passport_file` — the passport's file name, e.g. `rewda.jpg`
   - `photo_face_file`, `photo_full_file` — photo file names (optional)
   - the coloured columns — the data that is **not** on the passport
   - the green `*` columns mirror the passport fields; leave them **blank** to
     let OCR fill them, or type a value to **override** OCR.
   - See the **Guide** sheet for what every column means (EN + Arabic).
4. **Generate**:
   ```powershell
   .\run.ps1              # DOCX + PDF for every row
   .\run.ps1 --no-pdf     # DOCX only (faster; skips Word)
   .\run.ps1 --row 3      # only Excel row 3
   .\run.ps1 --list       # preview which rows will be processed
   ```
   (Or just double-click `run.bat`.)

## Output

- `output\docx\<name>.docx` — editable Word document (the agency form, filled)
- `output\pdf\<name>.pdf`  — final PDF: **page 1 = the filled form, page 2 = the passport**
- `output\report.csv`      — per-applicant summary: whether the passport MRZ was
  read and validated, and any warnings to review.

## Which fields come from where

| From the passport (OCR, auto) | From Excel (you type) |
|---|---|
| Full name, Passport number, Nationality, Sex, Date of birth, Date of expiry, Age (computed) | Position, Salary, Preferred country, Contact no., English/Arabic level, Education, Previous employment, Religion, Living town, Marital status, No. of children, Weight, Height, Complexion, Profile summary, all Skills (YES/NO), photos |
| Date of issue, Place of birth *(printed zone - lower confidence, verify)* | Place of issue *(usually typed)* |

**Merge rule:** Excel value (if you typed one) > OCR value > blank/default.

## Notes & limitations

- The **MRZ** (the two `<<<<` lines at the bottom of the passport) is read very
  reliably and is check-digit validated. If validation fails, `report.csv` warns
  you to verify that passport's fields.
- **Printed-zone** fields (date of issue, place of birth) are best-effort and
  flagged low-confidence — glance at them and correct in Excel if needed.
- Passport images should be reasonably sharp and upright. Small scans are
  auto-upscaled before OCR.
- If PDF export fails because Word is busy/locked, close all Word windows and
  re-run; the `.docx` files are already produced regardless.

## Project layout

```
cvgen\
  setup.ps1              one-time setup
  run_app.ps1 / run_app.bat   launch the desktop app
  run.ps1 / run.bat      run the batch generator (command line)
  requirements.txt
  template\
    original.docx        the agency's blank form (design source)
    cv_template.docx     original + {{placeholders}} (generated)
    assets\logo.png
  input\
    applicants.xlsx      one row per applicant
    passports\           passport images / PDFs
    photos\              applicant photos
  output\
    docx\  pdf\  report.csv
  src\
    fields.py            single source of truth for every field (EN/AR/source)
    build_template.py    injects placeholders into template/original.docx
    excel_io.py          builds + reads the Excel sheet
    ocr.py               offline passport reader (RapidOCR + MRZ parser)
    record.py            merge (Excel > OCR > default)
    docx_fill.py         fills the template (docxtpl + photos)
    pdf_export.py        DOCX -> PDF via Word, adds passport as page 2
    generate.py          batch orchestrator (command-line entry point)
    app.py               desktop app (Streamlit UI)
```

## Customising the form

The template is generated from the **original agency form** at
`template\original.docx`. `python src\build_template.py` injects the
`{{ placeholders }}` into that exact design (logo, colours, layout, the
face + full-body photo boxes) and writes `template\cv_template.docx`.

- To restyle fonts/colours/logo, edit `template\original.docx` in Word, then
  re-run `python src\build_template.py`. Keep the table layout/rows the same,
  because placeholders are injected by cell position (the builder prints a
  warning if the expected labels no longer line up).
- The final PDF automatically becomes **page 1 = form, page 2 = passport**
  (the template's trailing blank page is dropped and the passport image added).
- To change the field set/labels used by the Excel sheet and app, edit
  `src\fields.py`.

## Configuring defaults (`config.json`)

Agency-specific defaults live in `config.json` (no code changes needed):

- `field_defaults` — values used when a cell is blank, e.g. `place_of_issue`
  (`ADDIS ABABA`), `preferred_country`, `position`.
- `passport.place_of_issue_by_country` — place of issue per passport country
  (Ethiopian passports default to Addis Ababa).
- `passport.validity_years_by_country` + `derive_issue_from_expiry` — if the
  printed **issue date** can't be read by OCR, it is derived from the expiry date
  minus the passport's validity (5 years for Ethiopia).
- `app.port` — the app's fixed port (default `8531`).

## About the passport OCR fields

- **Reliable (from the MRZ):** name, passport number, nationality, sex, date of
  birth, date of expiry — check-digit validated.
- **Issue date:** printed only (not in the MRZ). Read by OCR; if a scan is
  unclear it's derived from expiry via `config.json`. Always editable.
- **Place of issue:** set from `config.json` (Addis Ababa for Ethiopia) rather
  than OCR, since it's effectively constant per country.
- **Place of birth:** best-effort OCR (flagged low-confidence) — verify/override.
