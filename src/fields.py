# -*- coding: utf-8 -*-
"""
Single source of truth for all CV fields.

Each field has:
  key     : the Jinja placeholder / Excel column / data-dict key
  en      : English label (shown on form + Excel header)
  ar      : Arabic label (shown on form)
  source  : where the value comes from
              'passport' -> filled by OCR (MRZ)
              'passport_visual' -> filled by OCR (printed zone, less reliable)
              'derived' -> computed (e.g. age from DOB)
              'excel'  -> supplied by the operator in the Excel sheet
  default : value used when nothing is provided
  group   : section on the form

Merge rule at runtime: Excel value (if the operator typed one) always wins;
otherwise the OCR/derived value is used; otherwise the default.
"""

# ---- ordered field catalogue -------------------------------------------------
FIELDS = [
    # key, en, ar, source, default, group
    ("full_name",        "FULL NAME",        "الاسم الكامل",      "passport",        "", "top"),
    ("position",         "POSITION",         "المهنة",            "excel",           "HOUSE MAID", "top"),
    ("salary",           "SALARY",           "الراتب",            "excel",           "1000 SR", "top"),
    ("preferred_country","PREFERRED COUNTRY","الدولة المطلوبة",   "excel",           "SAUDI ARABIA", "top"),
    ("contact_no",       "CONTACT N°",       "رقم الاتصال",       "excel",           "", "top"),

    ("passport_number",  "NUMBER",           "رقم الجواز",        "passport",        "", "passport"),
    ("date_of_issue",    "DATE OF ISSUE",    "تاريخ الصدور",      "passport_visual", "", "passport"),
    ("date_of_expiry",   "DATE OF EXPIRY",   "تاريخ الانتهاء",    "passport",        "", "passport"),
    ("place_of_issue",   "PLACE OF ISSUE",   "مكان الصدور",       "passport_visual", "", "passport"),

    ("english_level",    "ENGLISH",          "الانجليزية",        "excel",           "POOR", "lang"),
    ("arabic_level",     "ARABIC",           "العربية",           "excel",           "POOR", "lang"),
    ("education_level",  "EDUCATION LEVEL",  "المستوى التعليمي",  "excel",           "", "lang"),

    ("prev_country",     "COUNTRY",          "الدولة",            "excel",           "FIRST TIME", "prev"),
    ("prev_position",    "POSITION",         "المهنة",            "excel",           "FIRST TIME", "prev"),
    ("prev_period_from", "FROM",             "من",                "excel",           "FIRST TIME", "prev"),
    ("prev_period_to",   "TO",               "إلى",               "excel",           "FIRST TIME", "prev"),

    ("nationality",      "NATIONALITY",      "الجنسية",           "passport",        "", "personal"),
    ("religion",         "RELIGION",         "الديانة",           "excel",           "", "personal"),
    ("date_of_birth",    "DATE OF BIRTH",    "تاريخ الميلاد",     "passport",        "", "personal"),
    ("place_of_birth",   "PLACE OF BIRTH",   "مكان الميلاد",      "passport_visual", "", "personal"),
    ("living_town",      "LIVING TOWN",      "مكان السكن",        "excel",           "", "personal"),

    ("age",              "AGE",              "العمر",             "derived",         "", "physical"),
    ("marital_status",   "MARITAL STATUS",   "الحالة الاجتماعية", "excel",           "", "physical"),
    ("num_children",     "NO. OF CHILDREN",  "عدد الاطفال",       "excel",           "0", "physical"),
    ("weight",           "WEIGHT (KG)",      "الوزن",             "excel",           "", "physical"),
    ("height",           "HEIGHT (CM)",      "الطول",             "excel",           "", "physical"),
    ("complexion",       "COMPLEXION",       "لون البشرة",        "excel",           "", "physical"),

    ("profile_summary",  "PROFILE SUMMARY",  "ملخص البيانات",     "excel",           "", "profile"),
]

# ---- skills matrix (YES / NO) ------------------------------------------------
SKILLS = [
    # key, en, ar, default
    ("skill_babysitting",    "BABY SITTING",   "عناية الرضع",   "YES"),
    ("skill_children_care",  "CHILDREN CARE",  "عناية الاطفال", "YES"),
    ("skill_tutoring",       "TUTORING",       "تعليم الاطفال", "NO"),
    ("skill_disabled_care",  "DISABLED CARE",  "عناية العجزة",  "NO"),
    ("skill_cleaning",       "CLEANING",       "التنظيف",       "YES"),
    ("skill_washing",        "WASHING",        "الغسيل",        "YES"),
    ("skill_ironing",        "IRONING",        "الكوى",         "YES"),
    ("skill_cooking",        "COOKING",        "الطبخ",         "NO"),
    ("skill_arabic_cooking", "ARABIC COOKING", "الطبخ العربي",  "NO"),
]

# convenience lookups ----------------------------------------------------------
FIELD_KEYS = [f[0] for f in FIELDS]
SKILL_KEYS = [s[0] for s in SKILLS]
ALL_KEYS = FIELD_KEYS + SKILL_KEYS

DEFAULTS = {f[0]: f[4] for f in FIELDS}
DEFAULTS.update({s[0]: s[3] for s in SKILLS})

# fields the OCR is expected to provide
PASSPORT_KEYS = [f[0] for f in FIELDS if f[3] in ("passport", "passport_visual")]
# fields the operator provides in Excel
EXCEL_KEYS = [f[0] for f in FIELDS if f[3] == "excel"] + SKILL_KEYS

# English labels used as Excel column headers (human friendly)
def excel_headers():
    """Ordered (key, header_text) pairs for the Excel input sheet."""
    headers = []
    for key, en, ar, source, default, group in FIELDS:
        if source == "excel":
            headers.append((key, en))
    for key, en, ar, default in SKILLS:
        headers.append((key, en + " (YES/NO)"))
    return headers
