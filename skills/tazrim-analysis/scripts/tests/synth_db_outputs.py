#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""synth_db_outputs.py — SYNTHETIC project for testing the output scripts
(make_summary.py, make_figures.py, make_report_html.py, build_dashboard.py).

Writes into a directory:
  tazrim.config.json                 two people p1/p2, two accounts, two cards, one benefit
                                     program, one P2P entry, window 2025-01-01..2025-02-28
  work/database.csv                  ~63 rows with exactly common.DB_COLUMNS
  work/normalized/bank_*.csv         the two bank statements in the normalized schema, plus a
                                     `balance` column (running balance) so the cash reconciliation
                                     can take the bank change from balances
  work/findings.json                 a tiny narrative file (findings / actions / data_gaps)

Every person, merchant, amount and date is invented. EXPECTED holds hand-derived literals
(computed by hand from the row list below, never by code) — see the derivation comments.
Python 3.8 compatible; stdlib only.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from common import DB_COLUMNS, NORMALIZED_COLUMNS, write_csv, write_json  # noqa: E402

P1, P2, SHARED_LBL = "בן/בת זוג א'", "בן/בת זוג ב'", "משותף (עו\"ש)"
BANK1, BANK2 = "עו\"ש חשבון א'", "עו\"ש חשבון ב'"
CARD1, CARD2, CLUB, P2P, EST = "ויזה 1234", "מאסטרקארד 9012", "מועדון הטבות", "BIT " + P1, "ידני – הערכה"
EXP, INC, INT, SAV, CDEB, DUP, RPAID, RRECV = (
    "הוצאה", "הכנסה", "העברה פנימית", "חיסכון והשקעות", "תשלום כרטיס אשראי", "כפילות",
    "הוצאה בהחזר", "החזר הוצאה")
UNKNOWN = "אחר / לא מזוהה"

CONFIG = {
    "language": "he",
    "analysis_level": "deep",
    "window": {"start": "2025-01-01", "end": "2025-02-28", "month_rule": "charge_date"},
    "household": {"label": "משק בית לדוגמה",
                  "people": [{"id": "p1", "label": P1}, {"id": "p2", "label": P2}],
                  "shared_label": SHARED_LBL},
    "accounts": [
        {"id": "bank_a", "bank": "bank_xlsx_a", "label": BANK1, "owner": "shared",
         "files": "inputs/bank/bank_a/*.xlsx"},
        {"id": "bank_b", "bank": "bank_xls_b", "label": BANK2, "owner": "p2",
         "files": "inputs/bank/bank_b/*.xls*"},
    ],
    "cards": [
        {"id": "card_1234", "issuer": "card_cycle_xlsx", "last4": "1234", "label": CARD1, "owner": "p1",
         "files": "inputs/cards/card_cycle/card_*.xlsx", "settles_from": "bank_a",
         "bank_debit_pattern": "חיוב לכרטיס ויזה 1234", "cycle_day": 10,
         "fx_settlement": {"currency": "USD"}},
        {"id": "card_9012", "issuer": "card_blocks_xls", "last4": "9012", "label": CARD2, "owner": "p2",
         "files": "inputs/cards/card_9012/9012_*.xls*", "settles_from": "bank_b",
         "bank_debit_pattern": "9012 - כרטיס אשראי", "cycle_day": 2},
    ],
    "benefit_programs": [
        {"id": "club", "issuer": "club_docx", "label": CLUB, "owner": "p2",
         "files": "inputs/benefits/club/*.docx", "settles_from": "bank_b",
         "bank_debit_pattern": "חיוב מועדון הטבות", "default_debit_day": 2},
    ],
    "p2p": [{"id": "bit_p1", "app": "BIT", "owner": "p1", "csv": "inputs/p2p/bit_p1.csv",
             "card_marker": "BIT", "match_days": 5, "match_tolerance": 0.01}],
    "categories": {"scheme": "default", "from_template": None, "secondary_scheme": None},
    "fixed_categories": ["שכר דירה", "חשמל"],
    "dashboard": {"chartjs": "cdn"},
    "thresholds": {"highlight_expense_avg": 3000, "highlight_income_avg": 12000},
}

# (id, source, card, person, original_name, name_clean, type, group, cat, txn_date, charge_date,
#  amount, orig_currency, orig_amount, rule_note, details, linked_id, source_file, trip)
_R = []


def _row(rid, source, card, person, oname, cname, typ, group, cat, txn, charge, amount,
         ccy="ILS", oamt=None, note="", details="", linked="", sfile="", trip=""):
    _R.append((rid, source, card, person, oname, cname, typ, group, cat, txn, charge, amount,
               ccy, amount if oamt is None else oamt, note, details, linked, sfile, trip))


# ---------------------------------------------------------------- January 2025 (charge month)
_row("T-001", BANK1, "", P1, "משכורת", "משכורת", INC, "הכנסות", "משכורת", "2025-01-10", "2025-01-10", 12000)
_row("T-002", BANK2, "", P2, "משכורת", "משכורת", INC, "הכנסות", "משכורת", "2025-01-10", "2025-01-10", 8000)
_row("T-003", BANK1, "", SHARED_LBL, "הו\"ק שכר דירה", "שכר דירה", EXP, "דיור", "שכר דירה", "2025-01-01", "2025-01-01", 4000)
_row("T-004", BANK1, "", SHARED_LBL, "חברת החשמל", "חשמל", EXP, "דיור", "חשמל", "2025-01-15", "2025-01-15", 300)
_row("T-005", BANK1, "", SHARED_LBL, "חיוב לכרטיס ויזה 1234", "חיוב כרטיס ויזה 1234", CDEB, "תשלומי כרטיסי אשראי", CARD1, "2025-01-10", "2025-01-10", 1200)
_row("T-006", BANK1, "", SHARED_LBL, "הפקדה לקופת גמל", "קופת גמל", SAV, "חיסכון והשקעות", "קופת גמל", "2025-01-12", "2025-01-12", 1000)
_row("T-007", BANK1, "", SHARED_LBL, "המרת מט\"ח", "המרת מט\"ח", INT, "העברות פנימיות", "המרת מט\"ח", "2025-01-14", "2025-01-14", 700)
_row("T-008", BANK1, "", SHARED_LBL, "משיכת מזומן", "משיכת מזומן", INT, "העברות פנימיות", "משיכת מזומן", "2025-01-20", "2025-01-20", 400)
_row("T-009", BANK2, "", P2, "9012 - כרטיס אשראי", "חיוב כרטיס 9012", CDEB, "תשלומי כרטיסי אשראי", CARD2, "2025-01-02", "2025-01-02", 700)
_row("T-010", BANK2, "", P2, "חיוב מועדון הטבות", "חיוב מועדון", CDEB, "תשלומי כרטיסי אשראי", CLUB, "2025-01-02", "2025-01-02", 280)
# card 1 (cycle 10/01): ILS rows sum to 1,200 = T-005
_row("T-011", CARD1, CARD1, P1, "סופר לדוגמה", "סופר לדוגמה", EXP, "מזון", "סופרמרקט", "2024-12-20", "2025-01-10", 500)
_row("T-012", CARD1, CARD1, P1, "תחנת דלק לדוגמה", "דלק לדוגמה", EXP, "תחבורה", "דלק", "2024-12-22", "2025-01-10", 250)
_row("T-013", CARD1, CARD1, P1, "מסעדה לדוגמה", "מסעדה לדוגמה", EXP, "מזון", "מסעדות", "2024-12-28", "2025-01-10", 150)
_row("T-014", CARD1, CARD1, P1, "BIT העברה", "BIT העברה", DUP, "מזון", "מסעדות", "2025-01-05", "2025-01-10", 200, linked="T-030",
     note="מומן תשלום BIT")
_row("T-015", CARD1, CARD1, P1, "תחבורה ציבורית לדוגמה", "תחבורה ציבורית לדוגמה", EXP, "תחבורה", "תחבורה ציבורית", "2025-01-03", "2025-01-10", 100)
_row("T-016", CARD1, CARD1, P1, "RESTAURANT ABROAD", "מסעדה בחו\"ל", EXP, "מזון", "מסעדות", "2025-01-15", "2025-01-15", 360,
     ccy="USD", oamt=100, details="חיוב בחשבון מט\"י | המרה משוערת", trip="נסיעה לדוגמה")
# card 2 (cycle 02/01): 700 = T-009
_row("T-017", CARD2, CARD2, P2, "סופר לדוגמה", "סופר לדוגמה", EXP, "מזון", "סופרמרקט", "2024-12-15", "2025-01-02", 400)
_row("T-018", CARD2, CARD2, P2, "תחנת דלק לדוגמה", "דלק לדוגמה", EXP, "תחבורה", "דלק", "2024-12-18", "2025-01-02", 300)
# benefits club: load 200 (internal) + benefit 100 - bank 280 => discount 20 (F7)
_row("T-019", CLUB, CLUB, P2, "טעינה", "טעינה – לא נסכם", INT, "העברות פנימיות", "טעינת כרטיס מועדון", "2024-12-10", "2025-01-02", 200)
_row("T-020", CLUB, CLUB, P2, "חיוב", "רכישה בכרטיס מועדון", EXP, "מזון", "סופרמרקט", "2024-12-12", "2025-01-02", 150)
_row("T-021", CLUB, CLUB, P2, "הטבה לדוגמה", "הטבה לדוגמה", EXP, "מזון", "מסעדות", "2024-12-14", "2025-01-02", 100)
_row("T-022", CLUB, CLUB, P2, "הנחת מועדון", "הנחת מועדון (זיכוי מחושב)", EXP, "שונות", "הנחת מועדון (זיכוי מחושב)", "2025-01-02", "2025-01-02", -20,
     linked="T-010", sfile="מחושב")
# P2P app (p1)
_row("T-030", P2P, "BIT", P1, "תשלום BIT לדוגמה", "תשלום BIT לדוגמה", EXP, "מזון", "מסעדות", "2025-01-05", "2025-01-05", 200, linked="T-014")
_row("T-031", P2P, "BIT", P1, "החזר מחבר", "החזר מחבר", INC, "הכנסות", "אחר", "2025-01-06", "2025-01-06", 80)
_row("T-032", P2P, "BIT", P1, "תשלום BIT מיתרה", "תשלום BIT מיתרה", EXP, "תחבורה", "תחבורה ציבורית", "2025-01-08", "2025-01-08", 50)
# estimated cash expenses (not on any statement)
_row("T-033", EST, "", SHARED_LBL, "ניקיון – הערכה", "ניקיון – הערכה", EXP, "שונות", "הוצאות במזומן (הערכה)", "2025-01-31", "2025-01-31", 450,
     note="הערכה")

# ---------------------------------------------------------------- February 2025
_row("T-040", BANK1, "", P1, "משכורת", "משכורת", INC, "הכנסות", "משכורת", "2025-02-10", "2025-02-10", 12000)
_row("T-041", BANK1, "", SHARED_LBL, "הו\"ק שכר דירה", "שכר דירה", EXP, "דיור", "שכר דירה", "2025-02-01", "2025-02-01", 4000)
_row("T-043", BANK1, "", SHARED_LBL, "חיוב לכרטיס ויזה 1234", "חיוב כרטיס ויזה 1234", CDEB, "תשלומי כרטיסי אשראי", CARD1, "2025-02-10", "2025-02-10", 1720)
_row("T-044", BANK1, "", SHARED_LBL, "הפקדה לקופת גמל", "קופת גמל", SAV, "חיסכון והשקעות", "קופת גמל", "2025-02-12", "2025-02-12", 1000)
_row("T-045", BANK1, "", SHARED_LBL, "משיכה מיתרת BIT", "משיכת יתרת BIT", INT, "העברות פנימיות", "משיכת יתרת BIT", "2025-02-14", "2025-02-14", -30)
_row("T-046", BANK1, "", SHARED_LBL, "החזר הוצאות", "החזר הוצאות", RRECV, "הוצאות בהחזר", "החזר הוצאה", "2025-02-18", "2025-02-18", 500)
_row("T-047", BANK1, "", SHARED_LBL, "טיסה לכנס", "טיסה לכנס", RPAID, "הוצאות בהחזר", "נסיעה בהחזר", "2025-02-20", "2025-02-20", 800)
_row("T-048", BANK2, "", P2, "משכורת", "משכורת", INC, "הכנסות", "משכורת", "2025-02-10", "2025-02-10", 8000)
_row("T-049", BANK2, "", P2, "9012 - כרטיס אשראי", "חיוב כרטיס 9012", CDEB, "תשלומי כרטיסי אשראי", CARD2, "2025-02-02", "2025-02-02", 675)
_row("T-050", BANK2, "", P2, "העברה מחשבון אחר", "העברה מחשבון אחר", INT, "העברות פנימיות", "העברה מחשבון עצמי אחר", "2025-02-05", "2025-02-05", -2000)
# card 1 (cycle 10/02): 575+250+250-50+120+575 = 1,720 = T-043 (T-051/T-063 = deliberate duplicate charge)
_row("T-051", CARD1, CARD1, P1, "סופר לדוגמה", "סופר לדוגמה", EXP, "מזון", "סופרמרקט", "2025-01-20", "2025-02-10", 575)
_row("T-052", CARD1, CARD1, P1, "תחנת דלק לדוגמה", "דלק לדוגמה", EXP, "תחבורה", "דלק", "2025-01-22", "2025-02-10", 250)
_row("T-053", CARD1, CARD1, P1, "מסעדה לדוגמה", "מסעדה לדוגמה", EXP, "מזון", "מסעדות", "2025-01-25", "2025-02-10", 250)
_row("T-054", CARD1, CARD1, P1, "זיכוי מסעדה לדוגמה", "מסעדה לדוגמה", EXP, "מזון", "מסעדות", "2025-01-27", "2025-02-10", -50)
_row("T-058", CARD1, CARD1, P1, "BIT העברה", "BIT העברה", DUP, "מזון", "מסעדות", "2025-02-03", "2025-02-10", 120, linked="T-057",
     note="מומן תשלום BIT")
_row("T-063", CARD1, CARD1, P1, "סופר לדוגמה", "סופר לדוגמה", EXP, "מזון", "סופרמרקט", "2025-01-20", "2025-02-10", 575)
# card 2 (cycle 02/02): 450+150+75 = 675 = T-049
_row("T-055", CARD2, CARD2, P2, "סופר לדוגמה", "סופר לדוגמה", EXP, "מזון", "סופרמרקט", "2025-01-15", "2025-02-02", 450)
_row("T-056", CARD2, CARD2, P2, "תחבורה ציבורית לדוגמה", "תחבורה ציבורית לדוגמה", EXP, "תחבורה", "תחבורה ציבורית", "2025-01-18", "2025-02-02", 150)
_row("T-062", CARD2, CARD2, P2, "XYZ 123", "XYZ 123", EXP, "שונות", UNKNOWN, "2025-01-19", "2025-02-02", 75, note="לא נמצא כלל סיווג")
# P2P
_row("T-057", P2P, "BIT", P1, "תשלום BIT לדוגמה", "תשלום BIT לדוגמה", EXP, "מזון", "מסעדות", "2025-02-03", "2025-02-03", 120, linked="T-058")
# estimate
_row("T-059", EST, "", SHARED_LBL, "ניקיון – הערכה", "ניקיון – הערכה", EXP, "שונות", "הוצאות במזומן (הערכה)", "2025-02-28", "2025-02-28", 400,
     note="הערכה")
# ---------------------------------------------------------------- outside the window
_row("T-060", BANK1, "", SHARED_LBL, "הו\"ק שכר דירה", "שכר דירה", EXP, "דיור", "שכר דירה", "2025-03-01", "2025-03-01", 4000)
_row("T-061", CARD1, CARD1, P1, "סופר לדוגמה", "סופר לדוגמה", EXP, "מזון", "סופרמרקט", "2025-02-20", "2025-03-10", 300)

MONTH_NAME = {"2025-01": "ינואר 2025", "2025-02": "פברואר 2025", "2025-03": "מרץ 2025"}
PAY = {BANK1: "%s (%s)" % (BANK1, SHARED_LBL), BANK2: "%s (%s)" % (BANK2, P2), CARD1: "%s (%s)" % (CARD1, P1),
       CARD2: "%s (%s)" % (CARD2, P2), CLUB: "%s (%s)" % (CLUB, P2), P2P: "%s (%s)" % (P2P, P1),
       EST: "%s (%s)" % (EST, SHARED_LBL)}


def db_rows():
    out = []
    for (rid, source, card, person, oname, cname, typ, group, cat, txn, charge, amount, ccy, oamt,
         note, details, linked, sfile, trip) in _R:
        month = charge[:7]
        inwin = "2025-01-01" <= charge <= "2025-02-28"
        summed = inwin and typ in (EXP, INC)
        out.append({
            "id": rid, "source": source, "card": card, "pay": PAY[source], "person": person,
            "original_name": oname, "name_clean": cname, "type": typ,
            "group_tz": group, "cat_tz": cat, "group_new": group, "cat_new": cat,
            "txn_date": txn, "charge_date": charge, "month": month, "month_name": MONTH_NAME[month],
            "amount": amount, "orig_currency": ccy, "orig_amount": oamt,
            "in_window": "כן" if inwin else "לא", "summed": "כן" if summed else "לא",
            "rule_note": note, "details": details, "linked_id": linked,
            "source_file": sfile or ("%s.xlsx" % source), "row_ref": rid[-3:], "trip": trip,
        })
    return out


def normalized_bank_rows():
    """Bank rows in the normalized schema (+ `balance`), running balance from a synthetic
    opening balance (bank1 7,500; bank2 3,000). amount_ils > 0 = money out."""
    opening = {BANK1: 7500.0, BANK2: 3000.0}
    rows = {BANK1: [], BANK2: []}
    for r in db_rows():
        if r["source"] not in rows:
            continue
        out = r["amount"] if r["type"] not in (INC, RRECV) else -r["amount"]
        rows[r["source"]].append((r["charge_date"], r["id"], out, r))
    files = {}
    for src, lst in rows.items():
        lst.sort()
        bal = opening[src]
        recs = []
        for date, rid, out, r in lst:
            bal -= out
            recs.append({"id": rid, "source": src, "card": "", "original_name": r["original_name"],
                         "txn_date": date, "charge_date": date, "amount_ils": out,
                         "orig_currency": "ILS", "orig_amount": out, "details": "יתרה: %.2f" % bal,
                         "source_file": r["source_file"], "row_ref": r["row_ref"], "balance": round(bal, 2)})
        files[src] = recs
    return files


FINDINGS = {
    "findings": ["ממצא לדוגמה 1 — נכתב ידנית לאחר קריאת summary.json", {"title": "ממצא 2", "text": "טקסט לדוגמה"}],
    "actions": [{"what": "פעולה לדוגמה", "why": "נימוק לדוגמה", "how": "אופן ביצוע לדוגמה"}],
    "data_gaps": ["פער נתונים לדוגמה"],
}

# ---------------------------------------------------------------- hand-derived expectations
# Expenses (summed): Jan = bank 4,300 + card1 1,000 + FX 360 + card2 700 + club 230 + P2P 250 + est 450 = 7,290
#                    Feb = bank 4,000 + card1 1,600 + card2 675 + P2P 120 + est 400 = 6,795
# Income (summed):   Jan = 12,000 + 8,000 + 80 = 20,080 ; Feb = 12,000 + 8,000 = 20,000
# Groups: מזון 3,780 (סופרמרקט 2,650 + מסעדות 810+320); דיור 8,300 (8,000 + 300); תחבורה 1,100 (דלק 800 + ציבורית 100+50+150);
#         שונות 905 (-20 + 450 + 400 + 75)
# Fixed (שכר דירה, חשמל) = 8,300 -> variable 5,785 ; abroad (USD) = 360
# Bank change (rows): bank1 Jan 4,400 + Feb 5,010 = 9,410 ; bank2 Jan 7,020 + Feb 9,325 = 16,345 ; total 25,755
# Identity: 40,080 - 14,085 - 2,000 + 930 - 800 + 500 + 360 - 80 + 50 - 0 - 0 - 50 + 850 - 0 = 25,755 -> residual 0
EXPECTED = {
    "n_months": 2, "months": ["2025-01", "2025-02"],
    "n_rows_db": 49, "n_expense_rows": 26, "n_income_rows": 5,
    "total_expenses": 14085.0, "total_income": 40080.0,
    "expenses_by_month": [7290.0, 6795.0], "income_by_month": [20080.0, 20000.0],
    "avg_expenses": 7042.5, "avg_income": 20040.0,
    "expenses_by_group": {"דיור": 8300.0, "מזון": 3780.0, "תחבורה": 1100.0, "שונות": 905.0},
    "cat_electricity": {"total": 300.0, "avg": 150.0, "avg_nz": 300.0, "max_month": "2025-01", "min_month": "2025-02",
                        "median": 150.0, "stdev": 212.13, "n_nonzero": 1},
    "fixed_total": 8300.0, "variable_total": 5785.0, "abroad_total": 360.0, "israel_total": 13725.0,
    "savings_total": 2000.0, "internal_bank_net": -930.0,
    "card_debits_total": 4575.0, "card_vs_bank_diff": 0.0,
    "p2p": {"income": 80.0, "expenses": 370.0, "paid_from_balance": 50.0, "matched_pairs": 2},
    "club": {"face_net": 230.0, "bank": 280.0, "diff": -50.0},
    "fx_expenses": 360.0, "reimb_paid_bank": 800.0, "reimb_recv_bank": 500.0, "reimb_open": 300.0,
    "other_expenses": 850.0,
    "bank_change_rows": 25755.0, "bank_change_balances": 25755.0, "residual": 0.0,
    "top_merchant": {"name": "שכר דירה", "total": 8000.0},
    "super_total": 2500.0, "super_count": 5,
    "expenses_by_person": {P1: 3330.0, P2: 1605.0, SHARED_LBL: 9150.0},
    "income_by_person": {P1: 24080.0, P2: 16000.0},
    "expenses_by_card": {CARD1: 2960.0, CARD2: 1375.0, CLUB: 230.0},
    "unknown_total": 75.0, "n_unknown": 1, "n_refunds": 2, "n_double_charges": 1,
    "trip_total": {"נסיעה לדוגמה": 360.0},
}


def make_project(out_dir):
    """Write the synthetic project into out_dir; returns (config_path, EXPECTED)."""
    os.makedirs(os.path.join(out_dir, "work", "normalized"), exist_ok=True)
    cfg_path = os.path.join(out_dir, "tazrim.config.json")
    write_json(cfg_path, CONFIG)
    write_csv(os.path.join(out_dir, "work", "database.csv"), db_rows(), DB_COLUMNS)
    for i, (src, recs) in enumerate(sorted(normalized_bank_rows().items()), 1):
        write_csv(os.path.join(out_dir, "work", "normalized", "bank_%d.csv" % i), recs,
                  NORMALIZED_COLUMNS + ["balance"])
    write_json(os.path.join(out_dir, "work", "findings.json"), FINDINGS)
    return cfg_path, EXPECTED


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    p, _ = make_project(out)
    print(json.dumps({"ok": True, "config": p, "rows": len(db_rows())}, ensure_ascii=False))
