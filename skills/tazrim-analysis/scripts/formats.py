#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""formats.py — export-file formats, keyed by structural id. NOT user config.

Purpose : the single place that knows how each supported bank / card / benefits-club export
          is laid out (sheet names, header rows, exact header lists, block markers,
          date encodings, sign conventions, docx record shapes) plus the Bank of Israel SDMX
          rate endpoint. Parsers assert these and fail loudly on mismatch (P3.1, P4.1, P5.1).
Inputs  : none.
Outputs : `python3 formats.py` prints all specs as one JSON object on stdout; exit 0.
          `describe(format_id)` returns a human-readable summary string.

Formats are keyed by their STRUCTURE (bank_xlsx_a, bank_xls_b, card_cycle_xlsx, card_blocks_xls,
club_docx), never by an institution's name. The Hebrew strings are the files' own headers and
markers — format knowledge, not personal data. To support a new institution's layout, add an
entry HERE and a parser branch; when a layout changes, update it HERE only.
Python 3.8 compatible; stdlib only.
"""
import json
import sys

# --------------------------------------------------------------------------- banks
BANKS = {
    "bank_xlsx_a": {
        "label": "עובר ושב — קובץ A (.xlsx, גיליון 'עובר ושב', כותרת בשורה 8, חדש→ישן)",
        "extensions": [".xlsx"],
        "reader": {".xlsx": "openpyxl"},
        "sheet": "עובר ושב",
        "header_row": 8,                 # 1-based sheet row holding the header list
        "data_from_row": 9,
        "header": ["תאריך", "יום ערך", "תיאור התנועה", "₪ זכות/חובה", "₪ יתרה",
                   "אסמכתה", "עמלה", "ערוץ ביצוע"],
        "columns": {"date": "תאריך", "value_date": "יום ערך", "description": "תיאור התנועה",
                    "amount": "₪ זכות/חובה", "balance": "₪ יתרה", "reference": "אסמכתה",
                    "fee": "עמלה", "channel": "ערוץ ביצוע"},
        "date_encoding": "datetime",     # openpyxl returns datetime cells
        "order": "newest_first",
        "file_sign": "+credit/-debit",   # normalized amount_ils = -amount (F20: + = money out)
        "balance": "every_row",
        "balance_check": "F1",           # balance[i] - amount[i] == balance[i+1], tol 0.005
        "opening_balance": "implied",    # last balance - last amount
        "stop": "first fully blank row or end of sheet",
        "notes": ["value date recorded in details only when it differs from the date",
                  "reference / channel / fee / running balance appended to details",
                  "resolve columns by header name, never by index"],
    },
    "bank_xls_b": {
        "label": "עובר ושב — קובץ B (.xls/.xlsx, גיליון 'Activities', כותרת בשורה 6, ישן→חדש)",
        "extensions": [".xls", ".xlsx"],
        "reader": {".xls": "xlrd", ".xlsx": "openpyxl"},
        "sheet": "Activities",
        "header_row": 6,                 # 1-based (xlrd row index 5)
        "header_row_index0": 5,
        "opening_row": 7,                # 1-based (xlrd index 6): description 'יתרת פתיחה', value in 'יתרה'
        "opening_marker": "יתרת פתיחה",
        "data_from_row": 8,
        "header": ["", "יתרה", "תאריך ערך", "זכות", "חובה", "תאור", "אסמכתא", "סוג פעולה", "תאריך"],
        "columns": {"balance": "יתרה", "value_date": "תאריך ערך", "credit": "זכות", "debit": "חובה",
                    "description": "תאור", "reference": "אסמכתא", "op_type": "סוג פעולה",
                    "date": "תאריך"},
        "date_encoding": {".xls": "excel_serial", ".xlsx": "datetime"},
        "order": "oldest_first",
        "file_sign": "separate credit/debit columns",   # amount_ils = debit - credit
        "balance": "sparse_last_row_of_date_group",
        "balance_check": "F2",           # running += credit - debit from the opening balance; compare at filled cells
        "opening_balance": "row 7 'יתרת פתיחה' (or config accounts[].opening_balance)",
        "blank_cell": " ",               # xlrd blank cells arrive as ' '
        "card_debit_pattern": "<last4> - <issuer>",   # e.g. '1234 - כרטיס אשראי' — configured per card
        "notes": ["skip the 'יתרת פתיחה' row as a transaction but keep its value",
                  "resolve columns by header name"],
    },
}

# --------------------------------------------------------------------------- card issuers
ISSUERS = {
    "card_cycle_xlsx": {
        "label": "פירוט אשראי — קובץ מחזור חודשי (.xlsx, כל הכרטיסים, כותרת בשורה 9, סיכומים בשורות 7–8)",
        "extensions": [".xlsx"],
        "reader": {".xlsx": "openpyxl"},
        "file_pattern": "<prefix>_MMYY.xlsx",
        "file_month_rule": "MMYY is the month AFTER the charge; cycle date = cycle_day of month MM-1",
        "cycle_day_default": 10,
        "sheet": "כרטיסי אשראי",
        "title_row": 6,                  # 'פירוט עסקאות - כל הכרטיסים הבנקאיים'
        "cycle_header_rows": [7, 8],     # row 7: 'N עסקאות לחיוב בחודש <month> בכל הכרטיסים' (or 'N עסקאות לחיוב הקודם')
                                         # row 8: 'עסקאות באשראי - סה"כ חיוב:  + ₪<ils> + $<usd> + €<eur> ...'
        "cycle_count_pattern": r"^(\d+) עסקאות לחיוב",
        "cycle_total_marker": 'סה"כ חיוב:',
        "header_row": 9,
        "data_from_row": 10,
        "stop_marker": "הודעה",          # data ends at the first row whose col A starts with this (or is empty)
        "header": ["כרטיס", "בית עסק", "תאריך עסקה", "סכום העסקה", "מנפיק", "סוג העסקה",
                   "פירוט", "תאריך החיוב", "סכום החיוב", "כרטיס הוצג במעמד העסקה?"],
        "header_optional_tail": ["מטבע העסקה", "שער ההמרה", "תאריך שער", "עמלת ההמרה",
                                 "מדד בסיס", "שם המועדון", "אחוז הנחה", "סכום הנחה"],
        "header_assert": "first 10 cells equal `header`; further cells optional",
        "columns_fixed": {"card": 0, "merchant": 1, "txn_date": 2, "txn_amount": 3,
                          "issuer": 4, "txn_type": 5},
        "columns_shift": "from col 6 on, cells shift left when a field is empty: locate the charge "
                         "date BY PATTERN (dd/mm/yyyy or '--') at col >= 6; charge amount = next "
                         "numeric cell; פירוט = the text cells between col 6 and the charge date",
        "date_encoding": "dd/mm/yyyy string",
        "charge_date_pattern": r"^\d{2}/\d{2}/\d{4}$",
        "uncharged_marker": "--",        # charge date '--' or empty = standing order not yet charged (F14)
        "unposted_amount_marker": "--",  # txn amount '--' = not yet posted (newest file); skip row
        "card_cell": "<card type> <last4>, e.g. 'ויזה 1234'",
        "txn_types": {"ישראל": "domestic", 'חו"ל': "foreign", 'זיכוי-חו"ל': "foreign_credit",
                      "תשלום-ישראל": "installment_domestic", "תשלום": "installment"},
        "installment_pattern": r"^\s*(\d+)\s*מ\s*-\s*(\d+)",   # 'k מ - n' in פירוט
        "standing_order_markers": ["הוראת קבע", 'הו"ק'],
        "fx_conversion_pattern": r"הומר ל-\s*([\d.,]+)\s*\$?\s*(?:בשער\s*([\d.]+))?",
        "fx_rule": "row of type חו\"ל / זיכוי-חו\"ל whose charge date != cycle date is settled from the "
                   "card's FX account in fx_settlement.currency; amount_ils = charge amount x BOI rate on "
                   "the charge date (F4); flag 'המרה משוערת'",
        "sign": "+ = charge; credits negative (זיכוי)",
        "dedupe": "F15: (card, merchant, txn_date, charge_date, charge_amount) seen in two files -> drop the later",
    },
    "card_blocks_xls": {
        "label": "פירוט אשראי — דף חיוב לכרטיס (.xls: 3 בלוקים לכרטיס; .xlsx: גיליון 'פירוט עסקאות')",
        "extensions": [".xls", ".xlsx"],
        "reader": {".xls": "xlrd", ".xlsx": "openpyxl"},
        "file_pattern": "<last4>_MMYY.xls | <last4>_MMYY.xlsx",
        "file_month_rule": "MMYY is the month of the charge; the charge date is read from the file, never from the name",
        "cycle_day_default": 2,
        "blocks_xls": {
            "sheet": "Activities",
            "text_column": 1,            # markers live in column B (index 1)
            "block_card_marker": "כרטיס:",          # 'כרטיס:NNNN - <issuer label> חודש החיוב: dd/mm/yyyy'
            "block_card_pattern": r"כרטיס:(\d{4})",
            "block_ils_marker": "עסקאות בשקלים",    # 'עסקאות בשקלים חיוב בתאריך dd/mm/yyyy'
            "block_fx_marker": 'עסקאות במט"ח',      # 'עסקאות במט"ח חיוב בתאריך dd/mm/yyyy'
            "table_header_marker": "תאריך עסקה",    # first cell of the table header row
            "total_marker": 'סה"כ',                 # block total row; ILS total in col E, FX total in col F
            "header_ils": ["תאריך עסקה", "שם  העסק", "סכום עסקה", "סכום חיוב", "פירוט"],
            "header_fx": ["תאריך עסקה", "שם  העסק", "סכום מקורי", "מטבע מקורי", "סכום חיוב",
                          "מטבע חיוב", "פירוט"],
            "date_encoding": "excel_serial (xlrd.xldate_as_datetime with wb.datemode)",
            "date_pattern_in_markers": r"(\d{2})/(\d{2})/(\d{4})",
            "blank_cell": " ",
            "credit_rule": "F24: charge < 0 with a positive original amount and memo 'זיכוי' -> store both negative",
            "self_check": "sum of parsed rows per block == the block's סה\"כ (tol 0.005)",
        },
        "detail_xlsx": {
            "sheet": "פירוט עסקאות",
            "scan_cells": "rows 1..15 x cols 1..8 for the title cells",
            "charge_cell_prefix": "לחיוב ב-",       # 'לחיוב ב-dd.mm' ; year from a short title cell containing 20yy
            "charge_cell_pattern": r"(\d{2})\.(\d{2})$",
            "year_cell_pattern": r"\b20\d\d\b",
            "card_cell_pattern": r"(\d{4})",       # the cell containing the card's last 4 digits
            "header_marker": "תאריך רכישה",         # header row located dynamically by this first cell
            "header": ["תאריך רכישה", "שם בית עסק", "סכום עסקה", "מטבע עסקה", "סכום חיוב",
                       "מטבע חיוב", "מס' שובר", "פירוט נוסף"],
            "date_encoding": "dd.mm.yy string",
            "total_marker": 'סה"כ',                 # e.g. 'סה"כ לחיוב החודש בכרטיס בש"ח' in col A, value in col B
            "sign": "amounts already signed (credits negative, no זיכוי label)",
        },
        "currency_labels": {'דולר ארה"ב': "USD", "באט תאילנד": "THB", "אירו": "EUR",
                            "לירה שטרלינג": "GBP", 'ש"ח': "ILS", "ש''ח": "ILS", "₪": "ILS"},
        "bank_debit_pattern": "<last4> - <issuer>",
        "reconcile": "F3: Σ block == block סה\"כ == bank debit '<last4> - <issuer>' on the charge date",
        "missing_block_rule": "a card without a block in a cycle is reported, never fabricated "
                              "(reconstruct_missing_cycles option, default false)",
    },
    "club_docx": {
        "label": "מועדון הטבות — פירוט רכישות (.docx, פסקאות בלבד, שני מקטעים)",
        "extensions": [".docx"],
        "reader": {".docx": "python-docx"},
        "structure": "paragraph-only document pasted from the club site; two segments",
        "benefits_segment": {
            "header_sequence": ["תאריך רכישה", "שם ההטבה", "שם המוצר", "כמות", 'סה"כ', "סטטוס"],
            "record_paragraphs": 9,
            "record_start_marker": None,           # paragraph 0 = the club's own label; set
                                                   # benefit_programs[].record_marker to pin it, else any
                                                   # non-empty paragraph followed by a date starts a record
            "record_shape": ["club label", "purchase date dd.mm.yy", "product name (original_name)",
                             "benefit name", "quantity", "total ₪", "status", "extra", "extra"],
            "date_encoding": "dd.mm.yy",
            "sign": "face value; expense at face value",
        },
        "card_segment": {
            "header_sequence": ["תאריך ושעה", "מספר פעולה", "מזהה פיתקית", "רשת", "סניף", "סוג פעולה", "סכום"],
            "record_paragraphs": 7,
            "record_start_pattern": r"^\d{2}/\d{2}/\d{2}$",
            "record_shape": ["date dd/mm/yy", "operation number", "voucher id", "chain",
                             "branch", "operation type", "amount ₪ (signed)"],
            "date_encoding": "dd/mm/yy",
            "op_types": {"טעינה": "load (positive amount) -> internal transfer, details start 'טעינה – לא נסכם'",
                         "חיוב": "purchase (negative amount) -> expense at face value, name '<chain> - <branch>'"},
            "load_flag": "טעינה – לא נסכם",
            "load_name": "טעינת כסף - כרטיס נטען",
        },
        "amount_pattern": r"-?\d[\d,]*(?:\.\d+)?",
        "tail_rule": "ignore trailing stray paragraphs and a truncated last record",
        "charge_date_rule": "activity month m is billed by the bank in m+1: use the actual bank debit date "
                            "(bank_debit_pattern) when present, else default_debit_day of m+1",
        "bank_debit_pattern_default": "חיוב מועדון הטבות",
        "reconcile": "F7: discount_m = (Σ loads_m + Σ benefits_m) - bank_debit_m; add a negative expense row "
                     "'-discount_m' when |discount_m| > 0.005",
    },
}

# `card_detail_xlsx` names the .xlsx sub-layout of the per-card statement; it is the same spec and
# the same parser (parse_card_blocks.py reads both sub-layouts by extension).
ISSUERS["card_detail_xlsx"] = ISSUERS["card_blocks_xls"]

# --------------------------------------------------------------------------- BOI rates
BOI = {
    "label": "בנק ישראל — שערים יציגים (SDMX)",
    "endpoint_template": ("https://edge.boi.org.il/FusionEdgeServer/sdmx/v2/data/dataflow/"
                          "BOI.STATISTICS/EXR/1.0/RER_{ccy}_ILS?startperiod={start}&endperiod={end}&format=csv"),
    "cache_file_template": "RER_{ccy}_ILS.csv",     # under fx.rates_dir (default work/boi_rates)
    "csv_columns": ["SERIES_CODE", "FREQ", "BASE_CURRENCY", "COUNTER_CURRENCY", "UNIT_MEASURE",
                    "DATA_TYPE", "DATA_SOURCE", "TIME_COLLECT", "CONF_STATUS", "PUB_WEBSITE",
                    "UNIT_MULT", "COMMENTS", "TIME_PERIOD", "OBS_VALUE", "RELEASE_STATUS"],
    "date_column": "TIME_PERIOD",        # YYYY-MM-DD
    "rate_column": "OBS_VALUE",          # ILS per 1 unit of ccy
    "lookback_days": 10,                 # fall back to the last published day, at most 10 days back
    "manual_fallback": "fx.manual_rates CSV with columns date,ccy,rate; rows flagged 'שער ידני'",
    "volatile": True,
}

#: Normalized P2P CSV (BIT / PayBox) transcribed by Claude from screenshots — same schema as
#: the parsers' output (common.NORMALIZED_COLUMNS).
P2P_CSV = {
    "source": "<app> <person label>",
    "card": "<app>",
    "original_name": "counterparty",
    "details": "<note> | <status> | יוצא/נכנס",
    "sign": "outgoing positive, incoming negative; 'Withdrawal to bank' rows kept as internal",
    "source_file": "screenshot file name or the csv name",
}

ALL = {"banks": BANKS, "issuers": ISSUERS, "boi": BOI, "p2p_csv": P2P_CSV}


def get(format_id):
    """Spec dict for a bank or issuer id; KeyError if unknown."""
    if format_id in BANKS:
        return BANKS[format_id]
    if format_id in ISSUERS:
        return ISSUERS[format_id]
    if format_id == "boi":
        return BOI
    raise KeyError("unknown format id %r (banks: %s; issuers: %s)" % (
        format_id, ", ".join(sorted(BANKS)), ", ".join(sorted(ISSUERS))))


def reader_for(format_id, path):
    """Library to read `path` with, per the spec's extension map ('openpyxl'|'xlrd'|'python-docx')."""
    import os
    ext = os.path.splitext(path)[1].lower()
    spec = get(format_id)
    try:
        return spec["reader"][ext]
    except KeyError:
        raise ValueError("%s: extension %r not supported (expected %s)" % (
            format_id, ext, ", ".join(spec["extensions"])))


def describe(format_id):
    """Human-readable summary of one format."""
    spec = get(format_id)
    lines = ["%s (%s)" % (spec.get("label", format_id), format_id)]
    for k in ("extensions", "sheet", "header_row", "header", "date_encoding", "order",
              "file_sign", "sign", "balance_check", "file_pattern", "file_month_rule",
              "cycle_day_default", "stop_marker", "fx_rule", "reconcile", "structure",
              "endpoint_template", "date_column", "rate_column"):
        if k in spec:
            lines.append("  %-18s %s" % (k + ":", spec[k]))
    for sub in ("blocks_xls", "detail_xlsx", "benefits_segment", "card_segment"):
        if sub in spec:
            s = spec[sub]
            lines.append("  [%s]" % sub)
            for k in ("sheet", "header", "header_ils", "header_fx", "header_sequence",
                      "record_paragraphs", "record_start_marker", "record_start_pattern",
                      "date_encoding", "total_marker", "charge_cell_prefix"):
                if k in s:
                    lines.append("    %-18s %s" % (k + ":", s[k]))
    return "\n".join(lines)


if __name__ == "__main__":
    if any(a in ("-h", "--help") for a in sys.argv[1:]):
        sys.stderr.write("usage: formats.py [--json] [FORMAT_ID ...]\n"
                         "  prints the institution layout registry as one JSON object on stdout;\n"
                         "  FORMAT_ID (%s) writes a human-readable description to stderr.\n"
                         % ", ".join(sorted(list(BANKS) + list(ISSUERS))))
        sys.exit(0)
    for fid in [a for a in sys.argv[1:] if a != "--json"]:
        try:
            sys.stderr.write(describe(fid) + "\n")
        except KeyError as e:
            sys.stdout.write(json.dumps({"ok": False, "error": str(e), "hint": "run formats.py --help"}, ensure_ascii=False) + "\n")
            sys.exit(1)
    sys.stdout.write(json.dumps({"ok": True, "formats": ALL}, ensure_ascii=False, indent=2) + "\n")
