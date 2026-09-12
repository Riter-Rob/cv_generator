# -*- coding: utf-8 -*-
"""
Loads cvgen/config.json (agency-specific settings) with safe built-in
fallbacks, so the tool is configurable without touching code.

Edit config.json to change defaults (e.g. place of issue, preferred country,
passport validity years, app port).
"""
import os
import json

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CONFIG_PATH = os.path.join(ROOT, "config.json")

DEFAULTS = {
    "field_defaults": {
        "position": "HOUSE MAID",
        "preferred_country": "SAUDI ARABIA",
        "place_of_issue": "ADDIS ABABA",
    },
    "passport": {
        "default_place_of_issue": "ADDIS ABABA",
        "place_of_issue_by_country": {"ETH": "ADDIS ABABA"},
        "validity_years_by_country": {"ETH": 5},
        "derive_issue_from_expiry": True,
    },
    "app": {"port": 8531},
}

_CFG = None


def _deep_merge(base, override):
    out = dict(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load():
    global _CFG
    if _CFG is None:
        cfg = DEFAULTS
        try:
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    cfg = _deep_merge(DEFAULTS, json.load(f))
        except Exception:
            cfg = DEFAULTS
        _CFG = cfg
    return _CFG


def field_defaults():
    return load().get("field_defaults", {})


def passport_cfg():
    return load().get("passport", {})


def app_port():
    return int(load().get("app", {}).get("port", 8531))


def place_of_issue_for(country_code):
    p = passport_cfg()
    by_country = p.get("place_of_issue_by_country", {})
    return by_country.get((country_code or "").upper(), p.get("default_place_of_issue", ""))


def validity_years_for(country_code):
    p = passport_cfg()
    return p.get("validity_years_by_country", {}).get((country_code or "").upper(), 0)
