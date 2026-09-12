# -*- coding: utf-8 -*-
"""
Merge the three data sources into the final field values for one applicant.

Precedence (highest wins):
    Excel (operator typed)  >  OCR (passport)  >  default

Returns (values, meta) where:
    values : {key: str} for every template placeholder
    meta   : provenance + warnings for the run report
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fields import ALL_KEYS, DEFAULTS, PASSPORT_KEYS, SKILL_KEYS  # noqa: E402
import config  # noqa: E402

_YES = {"YES", "Y", "TRUE", "1", "T", "X", "✓", "نعم"}
_NO = {"NO", "N", "FALSE", "0", "F", "لا", ""}

# fields OCR is allowed to fill (everything else comes from Excel/default only)
_OCR_KEYS = set(PASSPORT_KEYS) | {"age"}


def _norm_skill(val):
    if val is None:
        return None
    s = str(val).strip().upper()
    if s in _YES:
        return "YES"
    if s in _NO:
        return "NO"
    return "YES" if s else "NO"


def merge(excel_row, ocr):
    excel_row = excel_row or {}
    ocr = ocr or {}
    values = {}
    provenance = {}
    cfg_defaults = config.field_defaults()

    for key in ALL_KEYS:
        default = cfg_defaults.get(key, DEFAULTS.get(key, ""))
        excel_val = excel_row.get(key, "")
        ocr_val = ocr.get(key, "")

        if key in SKILL_KEYS:
            if excel_val != "":
                values[key] = _norm_skill(excel_val)
                provenance[key] = "excel"
            else:
                values[key] = default or "NO"
                provenance[key] = "default"
            continue

        if excel_val not in ("", None):
            values[key] = str(excel_val).strip()
            provenance[key] = "excel"
        elif key in _OCR_KEYS and ocr_val not in ("", None):
            values[key] = str(ocr_val).strip()
            provenance[key] = "ocr"
        else:
            values[key] = default
            provenance[key] = "default" if default else "empty"

    # ---- warnings ----
    warnings = []
    if ocr:
        if ocr.get("_mrz_found") and not ocr.get("_mrz_valid"):
            warnings.append("passport MRZ check digits FAILED - verify passport fields")
        if not ocr.get("_mrz_found"):
            warnings.append("passport MRZ not detected - passport fields rely on Excel")
        for f in ocr.get("_low_confidence", []):
            if provenance.get(f) == "ocr":
                warnings.append(f"low-confidence OCR: {f}='{values.get(f)}' (verify in Excel)")
    # required-but-empty
    for key in ("full_name", "passport_number", "date_of_birth"):
        if not values.get(key):
            warnings.append(f"missing {key}")

    meta = {"provenance": provenance, "warnings": warnings}
    return values, meta
