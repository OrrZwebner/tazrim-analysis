#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""common.py — shared API for every tazrim-analysis script.

Purpose : schema constants (normalized CSV, database columns, row types, sheet sets per
          analysis level, Hebrew month names), config loading + validation, month/window
          arithmetic, path helpers, stable ids, CSV/JSON helpers and the standard argparse.
Inputs  : `tazrim.config.json` (see references/config-schema.md) — nothing else.
Outputs : none by itself; `emit()` writes exactly one JSON object to stdout, `log()` writes
          diagnostics to stderr, `fail()` writes `{"ok": false, "error", "hint"}` and exits.
Exit    : `fail()` exits with the given code (default 1); ConfigError -> 2 via parse_args.

Python 3.8 compatible; stdlib only (formats.py is imported lazily for id validation).

Usage in a script:
    import os, sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from common import parse_args, months, emit, fail, log
    args, cfg = parse_args("parse bank statements")
"""
import argparse
import copy
from collections import OrderedDict
import csv
import datetime as _dt
import glob
import hashlib
import json
import os
import sys
from typing import Callable, Dict, List, Optional, Tuple

# ----------------------------------------------------------------------------- schema
#: Columns of every normalized CSV written by a parser (work/normalized/*.csv).
#: Sign convention (F20): amount_ils > 0 = money OUT, < 0 = money IN.
NORMALIZED_COLUMNS = [
    "id", "source", "card", "original_name", "txn_date", "charge_date", "amount_ils",
    "orig_currency", "orig_amount", "details", "source_file", "row_ref",
]

#: Columns of work/database.csv (output of classify.py), in output order.
#: `amount` is positive for expenses AND incomes; `type` carries the direction (F20).
DB_COLUMNS = [
    "id", "source", "card", "pay", "person", "original_name", "name_clean", "type",
    "group_tz", "cat_tz", "group_new", "cat_new", "txn_date", "charge_date", "month",
    "month_name", "amount", "orig_currency", "orig_amount", "in_window", "summed",
    "rule_note", "details", "linked_id", "source_file", "row_ref", "trip",
]

#: Hebrew header per database column (used by the workbook / dashboard / לסיווג ידני sheet).
DB_HEADERS_HE = {
    "id": "מזהה", "source": "מקור", "card": "כרטיס", "pay": "אמצעי תשלום", "person": "אדם",
    "original_name": "שם כפי שמופיע בקובץ", "name_clean": "שם מובן", "type": "סוג תנועה",
    "group_tz": "קבוצה (תזרים)", "cat_tz": "קטגוריה (תזרים)", "group_new": "קבוצה (מוצעת)",
    "cat_new": "קטגוריה (מוצעת)", "txn_date": "תאריך עסקה", "charge_date": "תאריך חיוב",
    "month": "חודש ניתוח", "month_name": "שם חודש", "amount": "סכום ₪",
    "orig_currency": "מטבע מקור", "orig_amount": "סכום מקור", "in_window": "בחלון",
    "summed": "נסכם", "rule_note": "הערת סיווג", "details": "פרטים", "linked_id": "מזהה מקושר",
    "source_file": "קובץ מקור", "row_ref": "שורה בקובץ", "trip": "נסיעה",
}

#: Row-type vocabulary (`type` column). Only SUMMED_TYPES enter the expense/income totals.
TYPE_EXPENSE = "הוצאה"
TYPE_INCOME = "הכנסה"
TYPE_INTERNAL = "העברה פנימית"
TYPE_SAVINGS = "חיסכון והשקעות"
TYPE_CARD_DEBIT = "תשלום כרטיס אשראי"
TYPE_DUPLICATE = "כפילות"
TYPE_REIMBURSABLE = "הוצאה בהחזר"
TYPE_REIMBURSEMENT = "החזר הוצאה"
ROW_TYPES = (
    TYPE_EXPENSE, TYPE_INCOME, TYPE_INTERNAL, TYPE_SAVINGS, TYPE_CARD_DEBIT,
    TYPE_DUPLICATE, TYPE_REIMBURSABLE, TYPE_REIMBURSEMENT,
)
#: Types whose in-window rows are summed (F16): summed = in_window AND type in SUMMED_TYPES.
SUMMED_TYPES = (TYPE_EXPENSE, TYPE_INCOME)
#: Types shown in non-summed blocks below the totals (savings/internal/reimbursables).
NON_SUMMED_TYPES = tuple(t for t in ROW_TYPES if t not in SUMMED_TYPES)

#: Boolean flag literals used in the database (`in_window`, `summed`).
YES, NO = "כן", "לא"

#: Category bucket for rows with no rule (P10 step 1) and its rule note.
UNKNOWN_CAT = "אחר / לא מזוהה"
UNKNOWN_GROUP_EXPENSE = "שונות"
UNKNOWN_GROUP_INCOME = "הכנסות"
NOTE_NO_RULE = "לא נמצא כלל סיווג"

#: Categories any script can assign WITHOUT a merchant rule (generated rules, P2P/FX/abroad/
#: benefit-club fallbacks, estimates). classify.py appends any of these that a user scheme
#: lacks to the in-memory scheme; build_excel.py always lists them in the `לסיווג ידני`
#: dropdown; references/category-scheme-default.csv must contain every one of them (tested).
GROUP_CARD_DEBIT = "תשלומי כרטיסי אשראי"
GROUP_INTERNAL = "העברות פנימיות"
GROUP_ABROAD = 'חופשות וחו"ל'
CAT_CARD_DEBIT = "תשלום כרטיס אשראי"
CAT_BENEFIT_LOAD = "טעינת כרטיס הטבות"
CAT_FX_TRANSFER = 'המרת מט"ח'
CAT_P2P_WITHDRAWAL = "משיכת P2P לבנק"
CAT_P2P_INCOME = "החזרים מחברים (P2P)"
CAT_CLUB_DISCOUNT = "הנחת מועדון (זיכוי מחושב)"
CAT_CASH_FUNDING = "מזומן למימון הוצאות מוערכות"
CAT_ABROAD = 'הוצאות בחו"ל'
GENERATED_CATEGORIES = OrderedDict([
    (UNKNOWN_CAT, UNKNOWN_GROUP_EXPENSE),          # income rows keep UNKNOWN_GROUP_INCOME (never overridden)
    (CAT_CARD_DEBIT, GROUP_CARD_DEBIT),
    (CAT_BENEFIT_LOAD, GROUP_INTERNAL),
    (CAT_FX_TRANSFER, GROUP_INTERNAL),
    (CAT_P2P_WITHDRAWAL, GROUP_INTERNAL),
    (CAT_CASH_FUNDING, GROUP_INTERNAL),
    (CAT_P2P_INCOME, UNKNOWN_GROUP_INCOME),
    (CAT_CLUB_DISCOUNT, UNKNOWN_GROUP_EXPENSE),
    (CAT_ABROAD, GROUP_ABROAD),
])

#: Analysis levels and the workbook sheet set of each (spec level table).
LEVELS = ("overview", "standard", "deep")
SHEET_DB, SHEET_HELPER, SHEET_EXP, SHEET_INC, SHEET_ANALYSIS = (
    "database", "עזר_חודשי", "הוצאות", "הכנסות", "ניתוח נתונים")
SHEET_PIVOT, SHEET_BY_CAT, SHEET_TXNS, SHEET_VARIABLE, SHEET_MANUAL, SHEET_TRANSFERS, SHEET_LISTS = (
    "פילוח", "פירוט לפי קטגוריה", "פירוט עסקאות", "הוצאות משתנות לפי בית עסק",
    "לסיווג ידני", "העברות, BIT ו-PAYBOX", "רשימות")
SHEET_PROPOSED = "קטגוריות מוצעות"
SHEETS_BY_LEVEL = {
    "overview": [SHEET_DB, SHEET_HELPER, SHEET_EXP, SHEET_INC, SHEET_ANALYSIS],
    "standard": [SHEET_DB, SHEET_HELPER, SHEET_EXP, SHEET_INC, SHEET_ANALYSIS,
                 SHEET_PIVOT, SHEET_BY_CAT, SHEET_TXNS, SHEET_VARIABLE, SHEET_MANUAL,
                 SHEET_TRANSFERS, SHEET_LISTS],
    "deep": [SHEET_DB, SHEET_HELPER, SHEET_EXP, SHEET_INC, SHEET_ANALYSIS,
             SHEET_PIVOT, SHEET_BY_CAT, SHEET_TXNS, SHEET_VARIABLE, SHEET_MANUAL,
             SHEET_TRANSFERS, SHEET_LISTS, SHEET_PROPOSED],
}
#: Final sheet order in the workbook (P11 step 15); `הוצאות` is always the first tab.
SHEET_ORDER = [SHEET_EXP, SHEET_INC, SHEET_ANALYSIS, SHEET_TRANSFERS, SHEET_PIVOT, SHEET_BY_CAT,
               SHEET_TXNS, SHEET_VARIABLE, SHEET_PROPOSED, SHEET_MANUAL, SHEET_DB, SHEET_HELPER,
               SHEET_LISTS]
#: Hidden sheets.
HIDDEN_SHEETS = (SHEET_LISTS,)

#: User columns of the "לסיווג ידני" sheet (read back by apply_user_labels.py).
USER_TEXT_COL = "למילוי משתמש (טקסט חופשי)"
USER_CAT_COL = "קטגוריה (בחירה מהרשימה)"
#: Hebrew month names, 1..12.
HEBREW_MONTHS = {
    1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל", 5: "מאי", 6: "יוני",
    7: "יולי", 8: "אוגוסט", 9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר",
}

#: Owner id meaning "the joint account" (label = household.shared_label).
SHARED = "shared"

#: Value of categories.scheme that means the shipped default scheme reference.
DEFAULT_SCHEME = "default"

# ----------------------------------------------------------------------------- defaults
#: Defaults for every optional config key (references/config-schema.md). Deep-merged under
#: the user's file by load_config(); list-valued keys are NOT merged, only defaulted to [].
DEFAULTS = {
    "language": "he",
    "analysis_level": "standard",
    "window": {"month_rule": "charge_date"},
    "household": {"label": "משק הבית", "shared_label": "משותף (עו\"ש)"},
    "accounts": [],
    "cards": [],
    "benefit_programs": [],
    "p2p": [],
    "categories": {"scheme": DEFAULT_SCHEME, "from_template": None, "secondary_scheme": None},
    "rules": {
        "merchant_rules": "rules/merchant_rules.csv",
        "user_labels": "rules/user_labels.csv",
        "user_labels_interpreted": "rules/user_labels_interpreted.csv",
        "fx_overrides": "rules/fx_overrides.csv",
    },
    "fixed_categories": [],
    "estimates": [],
    "trips": [],
    "fx": {"rates_dir": "work/boi_rates", "source": "boi_sdmx", "manual_rates": None},
    "work": {
        "dir": "work",
        "normalized": "work/normalized",
        "database": "work/database.csv",
        "summary": "work/summary.json",
        "findings": "work/findings.json",
        "figures": "work/figures",
    },
    "notes": {
        "dir": "notes",
        "decisions": "notes/decisions.md",
        "questions_and_answers": "notes/questions_and_answers.md",
    },
    "outputs": {
        "dir": "outputs",
        "workbook": "outputs/תזרים.xlsx",
        "dashboard": "outputs/dashboard.html",
        "report_html": "work/report.html",
        "report_pdf": "outputs/דוח.pdf",
    },
    "verify": {"excel_recalc": "auto"},
    "dashboard": {"chartjs": "cdn"},
    "thresholds": {
        "highlight_expense_avg": 3000,
        "highlight_income_avg": 12000,
        "dynamic_merchant_rows": 250,
        "dynamic_txn_rows": 500,
        "top_merchants": 25,
        "hyperlink_rows": 45,
    },
}
CARD_DEFAULTS = {"cycle_day": 10, "fx_settlement": None, "bank_debit_pattern": None, "files": None,
                 "reconstruct_missing_cycles": False}
#: cards[].fx_settlement is None or {"currency": "USD", "transfer_pattern": None}
FX_SETTLEMENT_DEFAULTS = {"transfer_pattern": None}
BENEFIT_DEFAULTS = {"default_debit_day": 2, "bank_debit_pattern": None, "files": None}
P2P_DEFAULTS = {"match_days": 5, "match_tolerance": 0.01, "card_marker": "BIT",
                "csv": None, "screenshots": None, "bank_withdrawal_pattern": None}
ACCOUNT_DEFAULTS = {"opening_balance": None, "files": None}
#: estimates[]: name + amount required; one synthetic expense row per `every_n_months` months on `day`
ESTIMATE_DEFAULTS = {"id": None, "every_n_months": 1, "day": 15, "cat": UNKNOWN_CAT, "group": None,
                     "note": "", "funded_by_cat": None}
#: trips[]: label + start + end required; countries = words/country names that attribute rows
TRIP_DEFAULTS = {"countries": []}

_ISO_FMT = "%Y-%m-%d"


class ConfigError(Exception):
    """Raised by load_config for an invalid config. Message is Hebrew + English, actionable."""


class Config(dict):
    """The validated config: a dict (deep-merged with DEFAULTS) plus
    `.path` (config file), `.project_dir`, `.level` (effective analysis level),
    `.warnings` (non-fatal findings such as globs that matched no files)."""
    path = None          # type: Optional[str]
    project_dir = None   # type: Optional[str]
    level = None         # type: Optional[str]
    warnings = None      # type: Optional[List[str]]


# ----------------------------------------------------------------------------- helpers
def _merge(default, given):
    """Deep-merge dicts: keys of `given` win; nested dicts recurse; lists/None pass through."""
    if isinstance(default, dict) and isinstance(given, dict):
        out = copy.deepcopy(default)
        for k, v in given.items():
            out[k] = _merge(default.get(k), v) if k in default else copy.deepcopy(v)
        return out
    return copy.deepcopy(given)


def _err(he, en):
    return ConfigError("%s | %s" % (he, en))


def _parse_iso(value, key):
    try:
        return _dt.datetime.strptime(str(value), _ISO_FMT).date()
    except (TypeError, ValueError):
        raise _err("תאריך לא תקין ב-%s: %r (נדרש YYYY-MM-DD)" % (key, value),
                   "invalid date in %s: %r (expected YYYY-MM-DD)" % (key, value))


def _require(d, key, path):
    if key not in d or d[key] in (None, ""):
        raise _err("חסר מפתח חובה '%s' בקובץ ההגדרות" % path,
                   "missing required config key '%s'" % path)
    return d[key]


# ----------------------------------------------------------------------------- config
def load_config(path, project_dir=None, level=None):
    # type: (str, Optional[str], Optional[str]) -> Config
    """Read `tazrim.config.json`, deep-merge DEFAULTS, validate, return a Config.

    Raises ConfigError (Hebrew | English message) for: missing file / invalid JSON, missing
    required keys (window.start, window.end, household.people[].id/label), bad analysis level,
    window.end < window.start, unsupported month_rule, duplicate ids, unknown owner id,
    card/benefit settles_from not an account id, unknown bank/issuer id (formats.py),
    invalid p2p thresholds. Globs that match zero files are appended to `cfg.warnings`, not
    raised (mode C adds files later).
    `project_dir` defaults to the config's folder; `level` overrides analysis_level.
    """
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise _err("קובץ ההגדרות לא נמצא: %s" % path, "config file not found: %s" % path)
    try:
        with open(path, encoding="utf-8") as fh:
            raw = json.load(fh)
    except ValueError as e:
        raise _err("קובץ ההגדרות אינו JSON תקין: %s" % e, "config is not valid JSON: %s" % e)
    if not isinstance(raw, dict):
        raise _err("קובץ ההגדרות חייב להיות אובייקט JSON", "config must be a JSON object")

    cfg = Config(_merge(DEFAULTS, raw))
    cfg.path = path
    cfg.project_dir = os.path.abspath(project_dir or os.path.dirname(path))
    cfg.warnings = []

    # --- level
    lvl = level or cfg.get("analysis_level") or "standard"
    if lvl not in LEVELS:
        raise _err("רמת ניתוח לא חוקית: %r (אפשרויות: %s)" % (lvl, ", ".join(LEVELS)),
                   "invalid analysis level %r (choose one of %s)" % (lvl, ", ".join(LEVELS)))
    cfg["analysis_level"] = lvl
    cfg.level = lvl

    # --- window
    win = cfg.get("window") or {}
    start = _parse_iso(_require(win, "start", "window.start"), "window.start")
    end = _parse_iso(_require(win, "end", "window.end"), "window.end")
    if end < start:
        raise _err("סוף החלון (%s) לפני תחילתו (%s)" % (end, start),
                   "window.end (%s) is before window.start (%s)" % (end, start))
    if win.get("month_rule", "charge_date") != "charge_date":
        raise _err("month_rule נתמך רק כ-'charge_date' (חודש הניתוח = חודש החיוב בבנק)",
                   "only month_rule='charge_date' is supported (analysis month = bank charge month)")
    win["start"], win["end"] = start.isoformat(), end.isoformat()

    # --- household
    hh = cfg.get("household") or {}
    people = hh.get("people")
    if not isinstance(people, list) or not people:
        raise _err("household.people חייב להכיל לפחות אדם אחד ({id, label})",
                   "household.people must list at least one person ({id, label})")
    ids = set()
    for i, p in enumerate(people):
        if not isinstance(p, dict) or not p.get("id") or not p.get("label"):
            raise _err("household.people[%d] חייב לכלול id ו-label" % i,
                       "household.people[%d] needs both id and label" % i)
        if p["id"] == SHARED or p["id"] in ids:
            raise _err("מזהה אדם כפול או שמור: %r" % p["id"], "duplicate or reserved person id %r" % p["id"])
        ids.add(p["id"])
    owners = ids | {SHARED}

    # --- accounts / cards / benefit programs / p2p
    try:
        import formats  # noqa: F401  (sibling module; lazy to keep common importable alone)
        known_banks, known_issuers = set(formats.BANKS), set(formats.ISSUERS)
    except ImportError:
        known_banks = known_issuers = None
        cfg.warnings.append("formats.py not importable; bank/issuer ids not validated")

    def check_owner(owner, where):
        if owner not in owners:
            raise _err("בעלים לא מוכר ב-%s: %r (אפשרויות: %s)" % (where, owner, ", ".join(sorted(owners))),
                       "unknown owner %r in %s (choose one of %s)" % (owner, where, ", ".join(sorted(owners))))

    def check_glob(pattern, where):
        if not pattern:
            return
        pats = pattern if isinstance(pattern, list) else [pattern]
        hits = []
        for p in pats:
            hits += glob.glob(os.path.join(cfg.project_dir, p) if not os.path.isabs(p) else p)
        if not hits:
            cfg.warnings.append("%s: no files match %r (מצא 0 קבצים)" % (where, pattern))

    all_ids = set()

    def check_id(entry, where):
        eid = _require(entry, "id", where + ".id")
        if eid in all_ids:
            raise _err("מזהה כפול: %r (%s)" % (eid, where), "duplicate id %r (%s)" % (eid, where))
        all_ids.add(eid)
        return eid

    account_ids = []
    for i, a in enumerate(cfg["accounts"]):
        w = "accounts[%d]" % i
        a.update({k: v for k, v in ACCOUNT_DEFAULTS.items() if k not in a})
        account_ids.append(check_id(a, w))
        bank = _require(a, "bank", w + ".bank")
        if known_banks is not None and bank not in known_banks:
            raise _err("בנק לא מוכר ב-%s: %r (אפשרויות: %s)" % (w, bank, ", ".join(sorted(known_banks))),
                       "unknown bank id %r in %s (formats.BANKS: %s)" % (bank, w, ", ".join(sorted(known_banks))))
        check_owner(a.get("owner", SHARED), w + ".owner")
        a.setdefault("owner", SHARED)
        a.setdefault("label", a["id"])
        check_glob(a.get("files"), w + ".files")

    for i, c in enumerate(cfg["cards"]):
        w = "cards[%d]" % i
        c.update({k: v for k, v in CARD_DEFAULTS.items() if k not in c})
        check_id(c, w)
        issuer = _require(c, "issuer", w + ".issuer")
        if known_issuers is not None and issuer not in known_issuers:
            raise _err("מנפיק לא מוכר ב-%s: %r (אפשרויות: %s)" % (w, issuer, ", ".join(sorted(known_issuers))),
                       "unknown issuer id %r in %s (formats.ISSUERS: %s)" % (issuer, w, ", ".join(sorted(known_issuers))))
        check_owner(_require(c, "owner", w + ".owner"), w + ".owner")
        sf = _require(c, "settles_from", w + ".settles_from")
        if sf not in account_ids:
            raise _err("settles_from ב-%s מצביע על חשבון לא קיים: %r (חשבונות: %s)" % (w, sf, ", ".join(account_ids)),
                       "%s.settles_from %r is not an account id (accounts: %s)" % (w, sf, ", ".join(account_ids)))
        c.setdefault("label", "%s %s" % (issuer, c.get("last4", "")))
        c.setdefault("last4", "")
        if not isinstance(c["cycle_day"], int) or not 1 <= c["cycle_day"] <= 28:
            raise _err("cycle_day ב-%s חייב להיות מספר שלם 1–28" % w, "%s.cycle_day must be an int in 1..28" % w)
        if not isinstance(c["reconstruct_missing_cycles"], bool):
            raise _err("reconstruct_missing_cycles ב-%s חייב להיות true/false" % w,
                       "%s.reconstruct_missing_cycles must be a boolean" % w)
        fxs = c.get("fx_settlement")
        if fxs is not None:
            if not isinstance(fxs, dict) or not fxs.get("currency"):
                raise _err("fx_settlement ב-%s חייב להיות אובייקט עם currency (למשל USD)" % w,
                           "%s.fx_settlement must be an object with a currency code (e.g. USD)" % w)
            fxs["currency"] = str(fxs["currency"]).upper()
            fxs.update({k: v for k, v in FX_SETTLEMENT_DEFAULTS.items() if k not in fxs})
        check_glob(c.get("files"), w + ".files")

    for i, b in enumerate(cfg["benefit_programs"]):
        w = "benefit_programs[%d]" % i
        b.update({k: v for k, v in BENEFIT_DEFAULTS.items() if k not in b})
        check_id(b, w)
        issuer = _require(b, "issuer", w + ".issuer")
        if known_issuers is not None and issuer not in known_issuers:
            raise _err("מנפיק לא מוכר ב-%s: %r" % (w, issuer), "unknown issuer id %r in %s" % (issuer, w))
        check_owner(b.get("owner", SHARED), w + ".owner")
        b.setdefault("owner", SHARED)
        sf = _require(b, "settles_from", w + ".settles_from")
        if sf not in account_ids:
            raise _err("settles_from ב-%s מצביע על חשבון לא קיים: %r" % (w, sf),
                       "%s.settles_from %r is not an account id" % (w, sf))
        b.setdefault("label", b["id"])
        check_glob(b.get("files"), w + ".files")

    for i, p in enumerate(cfg["p2p"]):
        w = "p2p[%d]" % i
        p.update({k: v for k, v in P2P_DEFAULTS.items() if k not in p})
        check_id(p, w)
        _require(p, "app", w + ".app")
        check_owner(_require(p, "owner", w + ".owner"), w + ".owner")
        if not (isinstance(p["match_days"], int) and p["match_days"] >= 0):
            raise _err("match_days ב-%s חייב להיות מספר שלם ≥ 0" % w, "%s.match_days must be an int >= 0" % w)
        if not (isinstance(p["match_tolerance"], (int, float)) and p["match_tolerance"] >= 0):
            raise _err("match_tolerance ב-%s חייב להיות מספר ≥ 0" % w, "%s.match_tolerance must be a number >= 0" % w)
        p.setdefault("label", "%s %s" % (p["app"], owner_label(cfg, p["owner"])))
        if p.get("csv"):
            check_glob(p["csv"], w + ".csv")
        if p.get("screenshots"):
            check_glob(p["screenshots"], w + ".screenshots")

    # --- enums
    if cfg["verify"].get("excel_recalc") not in ("auto", "always", "never"):
        raise _err("verify.excel_recalc חייב להיות auto|always|never", "verify.excel_recalc must be auto|always|never")
    if cfg["dashboard"].get("chartjs") not in ("cdn", "inline"):
        raise _err("dashboard.chartjs חייב להיות cdn|inline", "dashboard.chartjs must be cdn|inline")
    if cfg["fx"].get("source") not in ("boi_sdmx", "manual"):
        raise _err("fx.source חייב להיות boi_sdmx|manual", "fx.source must be boi_sdmx|manual")
    for k in ("highlight_expense_avg", "highlight_income_avg"):
        if not isinstance(cfg["thresholds"].get(k), (int, float)):
            raise _err("thresholds.%s חייב להיות מספר" % k, "thresholds.%s must be a number" % k)
    ft = cfg["categories"].get("from_template")
    if ft:
        for k in ("file", "sheet", "groups_col", "cats_col", "first_row", "last_row"):
            _require(ft, k, "categories.from_template." + k)
        # link_avg_col / manual_cells were dropped in v0.1.1 (the template sheet is no longer copied
        # into the workbook); accepted and ignored for backward compatibility.
        for k in ("link_avg_col", "manual_cells"):
            if ft.get(k):
                sys.stderr.write("WARNING: categories.from_template.%s is ignored since v0.1.1\n" % k)
    for i, t in enumerate(cfg["trips"]):
        for k in ("label", "start", "end"):
            _require(t, k, "trips[%d].%s" % (i, k))
        if _parse_iso(t["end"], "trips[%d].end" % i) < _parse_iso(t["start"], "trips[%d].start" % i):
            raise _err("נסיעה %d: end לפני start" % i, "trips[%d]: end before start" % i)
        t.update({k: copy.deepcopy(v) for k, v in TRIP_DEFAULTS.items() if k not in t})
        if not isinstance(t["countries"], list):
            raise _err("trips[%d].countries חייב להיות רשימה" % i, "trips[%d].countries must be a list" % i)
    for i, e in enumerate(cfg["estimates"]):
        w = "estimates[%d]" % i
        _require(e, "name", w + ".name")
        amt = _require(e, "amount", w + ".amount")
        if not isinstance(amt, (int, float)):
            raise _err("amount ב-%s חייב להיות מספר" % w, "%s.amount must be a number" % w)
        e.update({k: v for k, v in ESTIMATE_DEFAULTS.items() if k not in e})
        if not (isinstance(e["every_n_months"], int) and e["every_n_months"] >= 1):
            raise _err("every_n_months ב-%s חייב להיות מספר שלם ≥ 1" % w, "%s.every_n_months must be an int >= 1" % w)
        if not (isinstance(e["day"], int) and 1 <= e["day"] <= 28):
            raise _err("day ב-%s חייב להיות מספר שלם 1–28" % w, "%s.day must be an int in 1..28" % w)
    return cfg


# ----------------------------------------------------------------------------- months
def months(cfg):
    # type: (Config) -> List[str]
    """All "YYYY-MM" months of the window, first to last inclusive (any N >= 1)."""
    start = _parse_iso(cfg["window"]["start"], "window.start")
    end = _parse_iso(cfg["window"]["end"], "window.end")
    out, y, m = [], start.year, start.month
    while (y, m) <= (end.year, end.month):
        out.append("%04d-%02d" % (y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def month_label(ym):
    # type: (str) -> str
    """'2025-01' -> 'ינואר 2025'. Unknown input is returned unchanged."""
    try:
        y, m = ym.split("-")[:2]
        return "%s %s" % (HEBREW_MONTHS[int(m)], y)
    except (AttributeError, KeyError, ValueError):
        return str(ym)


def analysis_month(charge_date, cfg=None):
    # type: (str, Optional[Config]) -> str
    """Analysis month of a row = the bank charge month (F16; month_rule 'charge_date')."""
    return str(charge_date)[:7]


def in_window(date_iso, cfg):
    # type: (str, Config) -> bool
    """True iff window.start <= date_iso <= window.end (ISO strings compare lexically)."""
    d = str(date_iso)[:10]
    return cfg["window"]["start"] <= d <= cfg["window"]["end"]


def flag(value):
    # type: (bool) -> str
    """bool -> 'כן'/'לא' (database flag literal)."""
    return YES if value else NO


# ----------------------------------------------------------------------------- paths
def skill_root():
    # type: () -> str
    """Absolute path of skills/tazrim-analysis (parent of scripts/)."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _lookup(cfg, dotted):
    cur = cfg
    for part in dotted.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return None
        cur = cur[part]
    return cur if isinstance(cur, str) else None


def project_path(cfg, key_or_relpath):
    # type: (Config, str) -> str
    """Absolute path inside the project. Accepts a dotted config key holding a path
    ('outputs.workbook', 'work.database', 'rules.merchant_rules', 'fx.rates_dir') or a plain
    relative path; absolute paths are returned unchanged."""
    val = _lookup(cfg, key_or_relpath) or key_or_relpath
    if os.path.isabs(val):
        return val
    return os.path.normpath(os.path.join(cfg.project_dir, val))


def scheme_path(cfg):
    # type: (Config) -> str
    """Path of the primary category-scheme CSV: the shipped default when
    categories.scheme == 'default', else project_path(categories.scheme)."""
    s = cfg["categories"].get("scheme") or DEFAULT_SCHEME
    if s == DEFAULT_SCHEME:
        return os.path.join(skill_root(), "references", "category-scheme-default.csv")
    return project_path(cfg, s)


def resolve_files(cfg, pattern):
    # type: (Config, object) -> List[str]
    """Sorted absolute paths matching a glob (or list of globs) relative to the project."""
    if not pattern:
        return []
    pats = pattern if isinstance(pattern, list) else [pattern]
    hits = []
    for p in pats:
        hits += glob.glob(p if os.path.isabs(p) else os.path.join(cfg.project_dir, p))
    return sorted(set(os.path.abspath(h) for h in hits))


def ensure_dirs(cfg):
    # type: (Config) -> List[str]
    """Create work/, work/normalized, work/figures, fx.rates_dir, outputs/, notes/, rules/.
    Returns the created/verified absolute paths."""
    dirs = [project_path(cfg, "work.dir"), project_path(cfg, "work.normalized"),
            project_path(cfg, "work.figures"), project_path(cfg, "fx.rates_dir"),
            project_path(cfg, "outputs.dir"), project_path(cfg, "notes.dir"),
            os.path.dirname(project_path(cfg, "rules.merchant_rules"))]
    for d in dirs:
        os.makedirs(d, exist_ok=True)
    return dirs


# ----------------------------------------------------------------------------- labels
def owner_label(cfg, owner_id):
    # type: (Config, str) -> str
    """Label of a person id; 'shared' -> household.shared_label; unknown id -> the id itself."""
    if owner_id == SHARED or owner_id is None:
        return cfg["household"].get("shared_label") or DEFAULTS["household"]["shared_label"]
    for p in cfg["household"].get("people", []):
        if p.get("id") == owner_id:
            return p["label"]
    return str(owner_id)


def find_entry(cfg, entry_id):
    # type: (Config, str) -> Optional[Dict]
    """The account / card / benefit program / p2p entry with this id, or None."""
    for section in ("accounts", "cards", "benefit_programs", "p2p"):
        for e in cfg.get(section, []):
            if e.get("id") == entry_id:
                return e
    return None


def pay_label(cfg, source_or_card):
    # type: (Config, str) -> str
    """Payment-method label with the owner ALWAYS suffixed in parentheses (judgment call
    'Payment-method labels'): '<entry label> (<owner label>)'. Accepts an entry id or an entry
    label; anything else is returned as '<text> (<shared label>)'."""
    e = find_entry(cfg, source_or_card)
    if e is None:
        for section in ("accounts", "cards", "benefit_programs", "p2p"):
            for cand in cfg.get(section, []):
                if cand.get("label") == source_or_card:
                    e = cand
                    break
            if e:
                break
    if e is None:
        return "%s (%s)" % (source_or_card, owner_label(cfg, SHARED))
    return "%s (%s)" % (e.get("label", e["id"]), owner_label(cfg, e.get("owner", SHARED)))


def person_of(cfg, entry_id):
    # type: (Config, str) -> str
    """Owner label of an entry id (joint account -> shared label)."""
    e = find_entry(cfg, entry_id)
    return owner_label(cfg, e.get("owner", SHARED) if e else SHARED)


# ----------------------------------------------------------------------------- ids
def stable_id(prefix, source_file, row_ref):
    # type: (str, str, object) -> str
    """Deterministic id '<PREFIX>-<6 hex>' from (basename of source_file, row_ref), so ids
    survive adding or re-ordering files (FM6). Same inputs -> same id, always."""
    key = "%s|%s" % (os.path.basename(str(source_file)), row_ref)
    return "%s-%s" % (prefix, hashlib.sha1(key.encode("utf-8")).hexdigest()[:6])


# ----------------------------------------------------------------------------- csv / json
def read_csv(path, skip_comments=False):
    # type: (str, bool) -> List[Dict[str, str]]
    """Read a UTF-8(-BOM) CSV into a list of dicts (all values str, missing -> '').
    skip_comments=True drops lines starting with '#' (rules files)."""
    with open(path, encoding="utf-8-sig", newline="") as fh:
        lines = fh.read().splitlines(True)
    if skip_comments:
        lines = [l for l in lines if not l.lstrip().startswith("#")]
    rows = []
    for r in csv.DictReader(lines):
        rows.append({k: ("" if v is None else v) for k, v in r.items() if k is not None})
    return rows


def write_csv(path, rows, columns):
    # type: (str, List[Dict], List[str]) -> int
    """Write rows (dicts) to a UTF-8-BOM CSV with exactly `columns`; returns the row count."""
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        n = 0
        for r in rows:
            w.writerow({c: ("" if r.get(c) is None else r.get(c)) for c in columns})
            n += 1
    return n


def read_json(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def write_json(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2, default=str)
        fh.write("\n")


def emit(obj):
    # type: (object) -> None
    """Write exactly one JSON object to stdout (the script's only stdout output)."""
    sys.stdout.write(json.dumps(obj, ensure_ascii=False, default=str) + "\n")
    sys.stdout.flush()


def log(msg):
    # type: (object) -> None
    """Diagnostics to stderr."""
    sys.stderr.write(str(msg) + "\n")
    sys.stderr.flush()


def fail(msg, hint="", code=1):
    # type: (str, str, int) -> None
    """Emit {"ok": false, "error", "hint"} on stdout and exit with `code`."""
    emit({"ok": False, "error": str(msg), "hint": str(hint)})
    sys.exit(code)


# ----------------------------------------------------------------------------- argparse
def build_parser(description, extra=None):
    # type: (str, Optional[Callable[[argparse.ArgumentParser], None]]) -> argparse.ArgumentParser
    """Standard parser: --config (default ./tazrim.config.json), --project-dir (default =
    config folder), --level overview|standard|deep. `extra(parser)` adds script-specific args."""
    p = argparse.ArgumentParser(description=description,
                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--config", default="./tazrim.config.json", help="path to tazrim.config.json")
    p.add_argument("--project-dir", default=None, help="project folder (default: the config's folder)")
    p.add_argument("--level", choices=LEVELS, default=None, help="analysis level (overrides config)")
    if extra:
        extra(p)
    return p


def parse_args(description, extra=None, argv=None):
    # type: (str, Optional[Callable], Optional[List[str]]) -> Tuple[argparse.Namespace, Config]
    """Parse the standard args, load + validate the config, return (args, cfg).
    On ConfigError: one-line JSON failure on stdout, exit 2. Warnings are logged to stderr."""
    args = build_parser(description, extra).parse_args(argv)
    try:
        cfg = load_config(args.config, project_dir=args.project_dir, level=args.level)
    except ConfigError as e:
        fail(str(e), "בדוק את tazrim.config.json מול references/config-schema.md | "
                     "check tazrim.config.json against references/config-schema.md", code=2)
        raise  # unreachable (fail exits); keeps type checkers happy
    for w in cfg.warnings:
        log("WARNING: " + w)
    return args, cfg


if __name__ == "__main__":
    emit({
        "ok": True,
        "normalized_columns": NORMALIZED_COLUMNS,
        "db_columns": DB_COLUMNS,
        "row_types": list(ROW_TYPES),
        "summed_types": list(SUMMED_TYPES),
        "levels": list(LEVELS),
        "sheets_by_level": SHEETS_BY_LEVEL,
        "sheet_order": SHEET_ORDER,
        "defaults": DEFAULTS,
    })
