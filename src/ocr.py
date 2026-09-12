# -*- coding: utf-8 -*-
"""
Offline passport reader.

read_passport(path) -> dict of extracted fields + provenance/confidence.

Pipeline:
  1. Load the passport (image OR pdf -> one image per page, small images upscaled).
  2. RapidOCR over the page(s).
  3. Locate + parse the TD3 Machine-Readable Zone (2 x 44 chars). This yields the
     RELIABLE fields: full_name, passport_number, nationality, sex,
     date_of_birth, date_of_expiry  (validated with MRZ check digits).
  4. Best-effort visual extraction for fields NOT in the MRZ:
     date_of_issue (the "third date"), place_of_birth, place_of_issue.
  5. Derive age from date_of_birth.

Fields marked low-confidence should be verified/overridden in the Excel sheet.
"""
import os
import re
import datetime

import cv2
import numpy as np

import config

_ENGINE = None


def _engine():
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _ENGINE = RapidOCR()
    return _ENGINE


# ---- image loading -----------------------------------------------------------
def load_images(path):
    """Return a list of BGR numpy images (all pages for a PDF)."""
    ext = os.path.splitext(path)[1].lower()
    images = []
    if ext == ".pdf":
        import fitz
        doc = fitz.open(path)
        for page in doc:
            pm = page.get_pixmap(dpi=300)
            arr = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.width, pm.n)
            if pm.n == 4:
                arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
            elif pm.n == 3:
                arr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
            else:
                arr = cv2.cvtColor(arr, cv2.COLOR_GRAY2BGR)
            images.append(arr)
        doc.close()
    else:
        data = np.fromfile(path, dtype=np.uint8)  # handles unicode paths
        img = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if img is None:
            from PIL import Image
            img = cv2.cvtColor(np.array(Image.open(path).convert("RGB")), cv2.COLOR_RGB2BGR)
        images.append(img)
    return [_upscale(im) for im in images]


def _upscale(img, target_long=1800):
    h, w = img.shape[:2]
    long_side = max(h, w)
    if long_side < target_long:
        s = min(3.0, target_long / long_side)
        img = cv2.resize(img, (int(w * s), int(h * s)), interpolation=cv2.INTER_CUBIC)
    return img


def run_ocr(img):
    res, _ = _engine()(img)
    items = []
    if not res:
        return items
    for box, text, score in res:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        items.append({
            "text": text,
            "score": float(score),
            "x": min(xs), "y": min(ys),
            "cx": sum(xs) / 4.0, "cy": sum(ys) / 4.0,
            "h": max(ys) - min(ys),
        })
    return items


# ---- MRZ ---------------------------------------------------------------------
def _char_val(c):
    if c == "<":
        return 0
    if c.isdigit():
        return int(c)
    if "A" <= c <= "Z":
        return ord(c) - 55
    return 0


def _check_digit(s):
    weights = [7, 3, 1]
    return sum(_char_val(c) * weights[i % 3] for i, c in enumerate(s)) % 10


# ISO-3166 alpha-3 -> nationality (common domestic-worker source countries)
NATIONALITY_BY_CODE = {
    "ETH": "ETHIOPIAN", "KEN": "KENYAN", "UGA": "UGANDAN", "TZA": "TANZANIAN",
    "PHL": "FILIPINO", "IND": "INDIAN", "LKA": "SRI LANKAN", "NPL": "NEPALI",
    "BGD": "BANGLADESHI", "IDN": "INDONESIAN", "PAK": "PAKISTANI",
    "GHA": "GHANAIAN", "NGA": "NIGERIAN", "MMR": "MYANMAR", "VNM": "VIETNAMESE",
}


_LETTER_TO_DIGIT = str.maketrans({
    "O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "Z": "2", "S": "5",
    "B": "8", "G": "6", "T": "7", "A": "4", "U": "0",
})


def _digits(s):
    return s.translate(_LETTER_TO_DIGIT)


def _clean_mrz(text):
    t = text.upper().replace(" ", "").replace("«", "<<")
    return re.sub(r"[^A-Z0-9<]", "<", t)


def find_mrz_lines(items):
    """Pick the two bottom-most long [A-Z0-9<] lines that look like a TD3 MRZ."""
    cands = []
    for it in items:
        c = _clean_mrz(it["text"])
        # MRZ lines are long and contain filler '<'
        if len(c) >= 28 and (c.count("<") >= 3 or re.match(r"^[A-Z0-9<]{30,}$", c)):
            cands.append((it["cy"], c))
    if len(cands) < 2:
        return None
    cands.sort()
    l1, l2 = cands[-2][1], cands[-1][1]
    # pad / trim to 44
    l1 = (l1 + "<" * 44)[:44]
    l2 = (l2 + "<" * 44)[:44]
    return l1, l2


def _yymmdd_to_date(yy, mm, dd, future_ok):
    yy, mm, dd = int(yy), int(mm), int(dd)
    year = 2000 + yy
    now = datetime.date.today().year
    if not future_ok and year > now:
        year = 1900 + yy
    try:
        return datetime.date(year, mm, dd)
    except ValueError:
        return None


def parse_mrz(l1, l2):
    out = {"_mrz_l1": l1, "_mrz_l2": l2, "_mrz_checks": {}}

    # --- line 1: type, country, names ---
    country = l1[2:5].replace("<", "")
    names = l1[5:]
    parts = names.split("<<")
    surname = parts[0].replace("<", " ").strip()
    given = ""
    if len(parts) > 1:
        given_tokens = [t for t in parts[1].split("<") if t and t.isalpha()]
        # stop collecting once we hit filler-only region already handled by isalpha
        given = " ".join(given_tokens)
    full = (given + " " + surname).strip()
    out["full_name"] = re.sub(r"\s+", " ", full)
    out["_issuing_country"] = country

    # --- line 2: number, nationality, dob, sex, expiry ---
    number_raw = l2[0:9]
    number_chk = l2[9]
    nationality = l2[10:13].replace("<", "")
    dob_raw = _digits(l2[13:19])
    dob_chk = l2[19]
    sex = l2[20]
    exp_raw = _digits(l2[21:27])
    exp_chk = l2[27]

    out["passport_number"] = number_raw.replace("<", "")
    out["nationality"] = NATIONALITY_BY_CODE.get(nationality, nationality)
    out["sex"] = {"M": "MALE", "F": "FEMALE"}.get(sex, "")

    dob = _yymmdd_to_date(dob_raw[0:2], dob_raw[2:4], dob_raw[4:6], future_ok=False)
    exp = _yymmdd_to_date(exp_raw[0:2], exp_raw[2:4], exp_raw[4:6], future_ok=True)
    out["date_of_birth"] = dob.isoformat() if dob else ""
    out["date_of_expiry"] = exp.isoformat() if exp else ""

    # --- check digits ---
    out["_mrz_checks"] = {
        "number": _check_digit(number_raw) == _char_val(number_chk) if number_chk.isdigit() else None,
        "dob": _check_digit(dob_raw) == _char_val(dob_chk) if dob_chk.isdigit() else None,
        "expiry": _check_digit(exp_raw) == _char_val(exp_chk) if exp_chk.isdigit() else None,
    }
    valid = [v for v in out["_mrz_checks"].values() if v is not None]
    out["_mrz_valid"] = bool(valid) and all(valid)
    return out


# ---- visual (non-MRZ) fields -------------------------------------------------
_DATE_RE = re.compile(r"(\d{1,2})[\s.\-:/]?(\d{1,2})[\s.\-:/]?(\d{4})")
_DATE8_RE = re.compile(r"(\d{2})(\d{2})(\d{4})")

_PLACE_STOP = {
    "FEDERAL", "DEMOCRATIC", "REPUBLIC", "ETHIOPIA", "PASSPORT", "NATIONALITY",
    "ETHIOPIAN", "IMMIGRATION", "CITIZENSHIP", "SERVICE", "AUTHORITY", "TYPE",
    "CODE", "NAME", "SURNAME", "GIVEN", "SEX", "DATE", "PLACE", "BIRTH", "ISSUE",
    "EXPIRY", "HOLDER", "SIGNATURE", "PERSONAL", "NO", "OF", "AND", "THE",
}


def _extract_dates(items, mrz_texts):
    found = []
    for it in items:
        t = it["text"]
        tc = _clean_mrz(t)
        if tc in mrz_texts or len(tc) > 30:
            continue
        digits = re.sub(r"\D", "", t)
        m = _DATE_RE.search(t)
        d = None
        if m:
            dd, mm, yy = m.group(1), m.group(2), m.group(3)
        elif len(digits) == 8:
            m2 = _DATE8_RE.match(digits)
            dd, mm, yy = m2.group(1), m2.group(2), m2.group(3)
        else:
            continue
        try:
            d = datetime.date(int(yy), int(mm), int(dd))
        except ValueError:
            continue
        if 1950 <= d.year <= 2100:
            found.append(d.isoformat())
    return found


def _extract_places(items, mrz_texts, exclude_prefixes=()):
    """Return candidate place tokens (uppercase alpha, filtered)."""
    cands = []
    for it in items:
        t = it["text"].strip()
        tc = _clean_mrz(t)
        if tc in mrz_texts:
            continue
        up = re.sub(r"[^A-Za-z]", "", t).upper()
        if not (3 <= len(up) <= 12):
            continue
        if not t.replace(" ", "").isalpha():
            continue
        if up in _PLACE_STOP:
            continue
        if any(sw in up for sw in _PLACE_STOP):     # reject merged garbage like CITIZENSHIPSERVICE
            continue
        if any(up.startswith(pfx) for pfx in exclude_prefixes if pfx):
            continue
        cands.append({"text": up, "cx": it["cx"], "cy": it["cy"], "score": it["score"]})
    return cands


def _date_items(items, mrz_texts):
    """Dates with positions (for label-proximity matching)."""
    out = []
    for it in items:
        t = it["text"]
        tc = _clean_mrz(t)
        if tc in mrz_texts or len(tc) > 30:
            continue
        digits = re.sub(r"\D", "", t)
        m = _DATE_RE.search(t)
        if m:
            dd, mm, yy = m.group(1), m.group(2), m.group(3)
        elif len(digits) == 8:
            m2 = _DATE8_RE.match(digits)
            dd, mm, yy = m2.group(1), m2.group(2), m2.group(3)
        else:
            continue
        try:
            d = datetime.date(int(yy), int(mm), int(dd))
        except ValueError:
            continue
        if 1950 <= d.year <= 2100:
            out.append({"date": d.isoformat(), "cx": it["cx"], "cy": it["cy"]})
    return out


def _label_positions(items, keyword, exclude=()):
    pos = []
    for it in items:
        up = it["text"].upper()
        if keyword in up and not any(e in up for e in exclude):
            pos.append((it["cx"], it["cy"]))
    return pos


def _pick_issue_date(items, mrz_texts, dob, exp):
    """Issue date is printed-only (not in MRZ). Prefer a date next to an
    'ISSUE' label; else the plausible date that is neither DOB nor expiry."""
    ditems = _date_items(items, mrz_texts)
    cands = [d for d in ditems if d["date"] not in (dob, exp)]
    if not cands:
        return None
    labels = _label_positions(items, "ISSUE", exclude=("EXPIR",))
    if labels:
        def dist(d):
            return min((d["cx"] - lx) ** 2 + (d["cy"] - ly) ** 2 for lx, ly in labels)
        cands.sort(key=dist)
        return cands[0]["date"]
    plausible = [d["date"] for d in cands
                 if (not exp or d["date"] <= exp) and (not dob or d["date"] > dob)]
    return sorted(plausible)[-1] if plausible else sorted(d["date"] for d in cands)[-1]


def _derive_issue_from_expiry(exp_iso, country):
    """Fallback: for known validities, issue = expiry - validity + 1 day."""
    vy = config.validity_years_for(country)
    if not vy or not exp_iso:
        return None
    try:
        e = datetime.date.fromisoformat(exp_iso)
        return (e.replace(year=e.year - vy) + datetime.timedelta(days=1)).isoformat()
    except ValueError:
        return None


def read_passport(path):
    """Main entry: returns extracted dict (empty dict if file missing)."""
    result = {
        "_source_file": os.path.basename(path),
        "_mrz_found": False,
        "_mrz_valid": False,
        "_low_confidence": [],
        "_notes": [],
    }
    if not path or not os.path.exists(path):
        result["_notes"].append("passport file not found")
        return result

    pages = load_images(path)
    best = None
    best_items = None
    for img in pages:
        items = run_ocr(img)
        mrz = find_mrz_lines(items)
        if mrz:
            best = mrz
            best_items = items
            break
        if best_items is None:
            best_items = items

    items = best_items or []
    mrz_texts = set()
    if best:
        mrz_texts = {best[0].rstrip("<"), best[1].rstrip("<"), best[0], best[1]}
        parsed = parse_mrz(*best)
        result.update(parsed)
        result["_mrz_found"] = True
    else:
        result["_notes"].append("MRZ not detected; passport fields must be typed in Excel")

    # visual: issue date (printed only; not in MRZ)
    dob = result.get("date_of_birth", "")
    exp = result.get("date_of_expiry", "")
    country = result.get("_issuing_country", "")
    issue = _pick_issue_date(items, mrz_texts, dob, exp)
    if issue:
        result["date_of_issue"] = issue
        result["_low_confidence"].append("date_of_issue")
    else:
        derived = _derive_issue_from_expiry(exp, country) \
            if config.passport_cfg().get("derive_issue_from_expiry") else None
        if derived:
            result["date_of_issue"] = derived
            result["_notes"].append(f"issue date derived from expiry ({country}); verify")
            result["_low_confidence"].append("date_of_issue")

    # place of issue: for most single-country agencies this is a constant
    poi = config.place_of_issue_for(country)
    if poi:
        result["place_of_issue"] = poi

    # visual: place of birth (best-effort)
    exclude = []
    nat = result.get("nationality", "")
    if nat:
        exclude.append(nat[:5])
    for tok in result.get("full_name", "").split():
        if len(tok) >= 4:
            exclude.append(tok[:5])
    if result.get("_issuing_country"):
        exclude.append(result["_issuing_country"])
    places = _extract_places(items, mrz_texts, exclude_prefixes=tuple(exclude))
    if places:
        # prefer right-column tokens (place of birth sits in the right data block)
        places_sorted = sorted(places, key=lambda p: (-p["cx"], p["cy"]))
        result["place_of_birth"] = places_sorted[0]["text"]
        result["_low_confidence"].append("place_of_birth")

    # age from DOB
    if dob:
        try:
            b = datetime.date.fromisoformat(dob)
            today = datetime.date.today()
            result["age"] = str(today.year - b.year -
                                ((today.month, today.day) < (b.month, b.day)))
        except ValueError:
            pass

    return result


if __name__ == "__main__":
    import sys, json
    p = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "input", "passports", "rewda.jpg")
    data = read_passport(p)
    print(json.dumps(data, ensure_ascii=False, indent=2))
