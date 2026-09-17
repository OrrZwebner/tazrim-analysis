#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""synth_db.py — a SYNTHETIC work/database.csv (exactly common.DB_COLUMNS) for the workbook tests.

Purpose : write a small classified database over a 2-month window (2025-01 .. 2025-02) with
          expenses in 6 categories across 3 groups, incomes for p1/p2, one internal transfer,
          one savings row, card-payment rows, one duplicate (כפילות), one out-of-window row,
          one unknown row, one row whose rule note asks for a decision, and a reimbursable pair.
          EXPECTED holds hand-computed totals (typed as literals, never computed from the rows).
Inputs  : none (a config dict is provided by `config()`).
Outputs : `write_project(dir)` -> (config_path, database_path); `python3 synth_db.py --out DIR`
          prints one JSON object.
Exit    : 0.

Every person, merchant, amount and date below is invented.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from common import DB_COLUMNS, month_label, write_csv, write_json  # noqa: E402

P1, P2, SH = "בן/בת זוג א'", "בן/בת זוג ב'", 'משותף (עו"ש)'
BANK = "עו\"ש חשבון א'"
CYCLE_SRC = "כרטיס אשראי (קובץ מחזור)"
CARD1, CARD2 = "ויזה 1234", "ויזה 5678"
PAY1, PAY2, PAYB = "%s (%s)" % (CARD1, P1), "%s (%s)" % (CARD2, P2), "%s (%s)" % (BANK, SH)
BIT_SRC = "BIT %s" % P1
PAYBIT = "%s (%s)" % (BIT_SRC, P1)
JAN, FEB = "2025-01", "2025-02"

CONFIG = {
    "language": "he",
    "analysis_level": "standard",
    "window": {"start": "2025-01-01", "end": "2025-02-28", "month_rule": "charge_date"},
    "household": {"label": "משק בית לדוגמה",
                  "people": [{"id": "p1", "label": P1}, {"id": "p2", "label": P2}],
                  "shared_label": SH},
    "accounts": [{"id": "bank_a", "bank": "bank_xlsx_a", "label": BANK, "owner": "shared",
                  "files": "inputs/bank/bank_a/*.xlsx"}],
    "cards": [
        {"id": "card_1234", "issuer": "card_cycle_xlsx", "last4": "1234", "label": CARD1, "owner": "p1",
         "files": "inputs/cards/card_cycle/card_*.xlsx", "settles_from": "bank_a",
         "bank_debit_pattern": "חיוב לכרטיס ויזה 1234", "cycle_day": 10},
        {"id": "card_5678", "issuer": "card_cycle_xlsx", "last4": "5678", "label": CARD2, "owner": "p2",
         "files": "inputs/cards/card_cycle/card_*.xlsx", "settles_from": "bank_a",
         "bank_debit_pattern": "חיוב לכרטיס ויזה 5678", "cycle_day": 10},
    ],
    "p2p": [{"id": "bit_p1", "app": "BIT", "owner": "p1", "csv": "inputs/p2p/bit_p1.csv",
             "card_marker": "BIT"}],
    "categories": {"scheme": "rules/category_scheme.csv", "from_template": None, "secondary_scheme": None},
    "fixed_categories": ["שכר דירה", "ועד בית"],
    "verify": {"excel_recalc": "never"},
    "thresholds": {"highlight_expense_avg": 3000, "highlight_income_avg": 12000},
}

#: Primary scheme CSV (group,cat,fixed). "גז" has no transactions -> must be dropped from הוצאות.
SCHEME = [
    ("מזון", "סופרמרקט", "no"), ("מזון", "מסעדות", "no"),
    ("דיור", "שכר דירה", "yes"), ("דיור", "ועד בית", "yes"), ("דיור", "גז", "yes"),
    ("תחבורה", "דלק", "no"), ("תחבורה", "תחבורה ציבורית", "no"),
    ("שונות", "אחר / לא מזוהה", "no"),
]

_seq = [0]


def _row(source, card, pay, person, orig, clean, typ, group, cat, txn, charge, amount, **kw):
    _seq[0] += 1
    month = charge[:7]
    inw = "כן" if "2025-01-01" <= charge <= "2025-02-28" else "לא"
    summed = "כן" if (inw == "כן" and typ in ("הוצאה", "הכנסה")) else "לא"
    d = {
        "id": kw.get("id") or "SYN-%03d" % _seq[0], "source": source, "card": card, "pay": pay,
        "person": person, "original_name": orig, "name_clean": clean, "type": typ,
        "group_tz": group, "cat_tz": cat, "group_new": kw.get("group_new", group),
        "cat_new": kw.get("cat_new", cat), "txn_date": txn, "charge_date": charge,
        "month": month, "month_name": month_label(month), "amount": "%.2f" % amount,
        "orig_currency": kw.get("cur", "ILS"), "orig_amount": "%.2f" % kw.get("orig_amount", amount),
        "in_window": inw, "summed": summed, "rule_note": kw.get("note", "כלל בית עסק"),
        "details": kw.get("details", ""), "linked_id": kw.get("linked_id", ""),
        "source_file": kw.get("source_file", "synthetic.xlsx"), "row_ref": str(_seq[0]), "trip": "",
    }
    return d


def rows():
    """The synthetic database rows (list of dicts with DB_COLUMNS)."""
    _seq[0] = 0
    R = []
    c1 = lambda *a, **k: _row(CYCLE_SRC, CARD1, PAY1, P1, *a, **k)      # noqa: E731
    c2 = lambda *a, **k: _row(CYCLE_SRC, CARD2, PAY2, P2, *a, **k)      # noqa: E731
    bk = lambda *a, **k: _row(BANK, "", PAYB, SH, *a, **k)          # noqa: E731
    E, I = "הוצאה", "הכנסה"
    # ---- January expenses (charge month 2025-01) — hand total 7,700
    for d in ("2025-01-03", "2025-01-08", "2025-01-15", "2025-01-22"):
        R.append(c1("סופר לדוגמה בע\"מ", "סופר לדוגמה", E, "מזון", "סופרמרקט", d, "2025-01-10", 250))
    for d in ("2025-01-05", "2025-01-19"):
        R.append(c2("מכולת לדוגמה", "מכולת לדוגמה", E, "מזון", "סופרמרקט", d, "2025-01-10", 250))
    R.append(c1("מסעדה לדוגמה", "מסעדה לדוגמה", E, "מזון", "מסעדות", "2025-01-04", "2025-01-10", 350))
    R.append(c1("מסעדה לדוגמה", "מסעדה לדוגמה", E, "מזון", "מסעדות", "2025-01-25", "2025-01-10", 100))
    R.append(c2("בית קפה לדוגמה", "בית קפה לדוגמה", E, "מזון", "מסעדות", "2025-01-11", "2025-01-10", 150))
    R.append(_row(BIT_SRC, "BIT", PAYBIT, P1, "חבר לדוגמה", "ארוחה משותפת (BIT)", E, "מזון", "מסעדות",
                  "2025-01-12", "2025-01-12", 200, id="SYN-BIT", details="ארוחה | Completed | יוצא",
                  note="נספר כאן; מומן בחיוב כרטיס SYN-DUP", linked_id="SYN-DUP", source_file="bit_p1.csv"))
    R.append(bk("הו\"ק שכר דירה לדוגמה", "שכר דירה", E, "דיור", "שכר דירה", "2025-01-01", "2025-01-01", 4000))
    R.append(bk("העברה לועד בית לדוגמה", "ועד בית", E, "דיור", "ועד בית", "2025-01-02", "2025-01-02", 200))
    for d in ("2025-01-06", "2025-01-16", "2025-01-26"):
        R.append(c1("תחנת דלק לדוגמה", "תחנת דלק לדוגמה", E, "תחבורה", "דלק", d, "2025-01-10", 300))
    for d in ("2025-01-07", "2025-01-21"):
        R.append(c2("רב-קו לדוגמה", "רב-קו לדוגמה", E, "תחבורה", "תחבורה ציבורית", d, "2025-01-10", 100))
    R.append(c2("חנות לא מזוהה", "חנות לא מזוהה", E, "שונות", "אחר / לא מזוהה", "2025-01-09", "2025-01-10", 100,
                id="SYN-UNK", note="לא נמצא כלל סיווג"))
    # ---- February expenses (charge month 2025-02) — hand total 6,650
    for d in ("2025-02-03", "2025-02-13"):
        R.append(c1("סופר לדוגמה בע\"מ", "סופר לדוגמה", E, "מזון", "סופרמרקט", d, "2025-02-10", 400))
    R.append(c2("מכולת לדוגמה", "מכולת לדוגמה", E, "מזון", "סופרמרקט", "2025-02-05", "2025-02-10", 400))
    R.append(c2("ירקן לדוגמה", "ירקן לדוגמה", E, "מזון", "סופרמרקט", "2025-02-06", "2025-02-10", 200))
    R.append(c1("מסעדה לדוגמה", "מסעדה לדוגמה", E, "מזון", "מסעדות", "2025-02-04", "2025-02-10", 250))
    R.append(c2("בית קפה לדוגמה", "בית קפה לדוגמה", E, "מזון", "מסעדות", "2025-02-14", "2025-02-10", 300))
    R.append(c1("PAYBOX", "PAYBOX – ארוחה", E, "מזון", "מסעדות", "2025-02-08", "2025-02-10", 150,
                id="SYN-PBX", details="PAYBOX | הערה לדוגמה",
                note="לבדיקה: תשלום PAYBOX — נדרש: אשר את הקטגוריה או בחר אחרת"))
    R.append(bk("הו\"ק שכר דירה לדוגמה", "שכר דירה", E, "דיור", "שכר דירה", "2025-02-01", "2025-02-01", 4000))
    R.append(c1("תחנת דלק לדוגמה", "תחנת דלק לדוגמה", E, "תחבורה", "דלק", "2025-02-02", "2025-02-10", 250))
    R.append(c1("תחנת דלק לדוגמה", "תחנת דלק לדוגמה", E, "תחבורה", "דלק", "2025-02-20", "2025-02-10", 200))
    for d in ("2025-02-07", "2025-02-21"):
        R.append(c2("רב-קו לדוגמה", "רב-קו לדוגמה", E, "תחבורה", "תחבורה ציבורית", d, "2025-02-10", 50))
    # ---- incomes — hand total 30,600 (Jan 15,500; Feb 15,100)
    G1, G2, GO = "הכנסות %s" % P1, "הכנסות %s" % P2, "הכנסות אחרות"
    R.append(bk("משכורת לדוגמה א", "משכורת א'", I, G1, "משכורת א'", "2025-01-01", "2025-01-01", 9000))
    R.append(bk("משכורת לדוגמה א", "משכורת א'", I, G1, "משכורת א'", "2025-02-01", "2025-02-01", 9000))
    R.append(bk("משכורת לדוגמה ב", "משכורת ב'", I, G2, "משכורת ב'", "2025-01-01", "2025-01-01", 6000))
    R.append(bk("משכורת לדוגמה ב", "משכורת ב'", I, G2, "משכורת ב'", "2025-02-01", "2025-02-01", 6000))
    R.append(bk("קצבה לדוגמה", "קצבה", I, GO, "קצבאות", "2025-01-20", "2025-01-20", 500))
    R.append(_row(BIT_SRC, "BIT", PAYBIT, P1, "חברה לדוגמה", "החזר מחבר (BIT)", I, GO, "החזרים מחברים",
                  "2025-02-15", "2025-02-15", 100, details="החזר | Completed | נכנס", source_file="bit_p1.csv"))
    # ---- non-summed rows
    R.append(bk("העברה לחשבון אחר לדוגמה", "העברה בין חשבונות", "העברה פנימית", "העברות פנימיות",
                "העברה בין חשבונות", "2025-01-15", "2025-01-15", 2000, note="העברה פנימית – לא נסכם"))
    R.append(bk("הפקדות קופ\"ג", "הפקדה לקופת גמל", "חיסכון והשקעות", "חיסכון והשקעות", "קופת גמל",
                "2025-02-12", "2025-02-12", 1750, note="חיסכון – לא נסכם"))
    for m, a in ((JAN, 3000), (FEB, 2800)):
        R.append(bk("חיוב לכרטיס ויזה 1234", CARD1, "תשלום כרטיס אשראי", "תשלומי כרטיסי אשראי", CARD1,
                    m + "-10", m + "-10", a, note="חיוב כרטיס – ההוצאות נספרות לפי בתי העסק"))
    for m, a in ((JAN, 1000), (FEB, 900)):
        R.append(bk("חיוב לכרטיס ויזה 5678", CARD2, "תשלום כרטיס אשראי", "תשלומי כרטיסי אשראי", CARD2,
                    m + "-10", m + "-10", a, note="חיוב כרטיס – ההוצאות נספרות לפי בתי העסק"))
    R.append(c1("BIT העברה לחבר לדוגמה", "BIT (מומן בכרטיס)", "כפילות", "מזון", "מסעדות",
                "2025-01-12", "2025-01-10", 200, id="SYN-DUP", details="העברה לחבר לדוגמה",
                note="חיוב כרטיס שמימן תשלום BIT SYN-BIT – ההוצאה נספרת בשורת ה-BIT", linked_id="SYN-BIT"))
    R.append(c1("סופר לדוגמה בע\"מ", "סופר לדוגמה", E, "מזון", "סופרמרקט", "2025-03-02", "2025-03-10", 999,
                id="SYN-OOW"))
    R.append(bk("נסיעה אקדמית לדוגמה", "נסיעה אקדמית", "הוצאה בהחזר", "הוצאות בהחזר", "נסיעה אקדמית",
                "2025-01-18", "2025-01-18", 700, note="הוצאה בהחזר – צפוי החזר"))
    R.append(bk("החזר נסיעה לדוגמה", "נסיעה אקדמית", "החזר הוצאה", "הוצאות בהחזר", "נסיעה אקדמית",
                "2025-02-18", "2025-02-18", 300, note="החזר שהתקבל"))
    for r in R:
        assert set(r) == set(DB_COLUMNS), set(r) ^ set(DB_COLUMNS)
    return R


#: Hand-computed expectations (literals).
EXPECTED = {
    "months": [JAN, FEB],
    "expense_total": 14350.0, "expense_by_month": {JAN: 7700.0, FEB: 6650.0},
    "income_total": 30600.0, "income_by_month": {JAN: 15500.0, FEB: 15100.0},
    "balance_by_month": {JAN: 7800.0, FEB: 8450.0},
    "cat_totals": {"סופרמרקט": 2900.0, "מסעדות": 1500.0, "שכר דירה": 8000.0, "ועד בית": 200.0,
                   "דלק": 1350.0, "תחבורה ציבורית": 300.0, "אחר / לא מזוהה": 100.0},
    "cat_by_month": {"ועד בית": {JAN: 200.0, FEB: 0.0}, "סופרמרקט": {JAN: 1500.0, FEB: 1400.0},
                     "מסעדות": {JAN: 800.0, FEB: 700.0}},
    "nz_avg": {"ועד בית": 200.0}, "avg": {"ועד בית": 100.0},
    "group_totals": {"דיור": 8200.0, "מזון": 4400.0, "תחבורה": 1650.0, "שונות": 100.0},
    "group_order": ["דיור", "מזון", "תחבורה", "שונות"],
    "dropped_categories": ["גז"],
    "manual_rows": ["SYN-UNK", "SYN-PBX"],
    "reimbursable_open_balance": 400.0,
    "savings_total": 1750.0, "internal_total": 2000.0, "card_debits_total": 7700.0,
    "n_rows": 46,
    "top_merchant": "שכר דירה",
}


def config():
    return json.loads(json.dumps(CONFIG))


def write_project(out):
    """Write tazrim.config.json, rules/category_scheme.csv and work/database.csv under `out`."""
    out = os.path.abspath(out)
    cfg_path = os.path.join(out, "tazrim.config.json")
    write_json(cfg_path, config())
    write_csv(os.path.join(out, "rules", "category_scheme.csv"),
              [{"group": g, "cat": c, "fixed": f} for g, c, f in SCHEME], ["group", "cat", "fixed"])
    db_path = os.path.join(out, "work", "database.csv")
    write_csv(db_path, rows(), DB_COLUMNS)
    for d in ("outputs", "notes", "work/normalized"):
        os.makedirs(os.path.join(out, d), exist_ok=True)
    return cfg_path, db_path


if __name__ == "__main__":
    target = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else "synthetic_project"
    c, d = write_project(target)
    print(json.dumps({"ok": True, "config": c, "database": d, "rows": len(rows()), "expected": EXPECTED},
                     ensure_ascii=False))
