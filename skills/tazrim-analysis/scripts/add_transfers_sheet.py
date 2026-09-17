#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""add_transfers_sheet.py — post-build step (P12): the sheet 'העברות, BIT ו-PAYBOX'.

Purpose : drop and rebuild the sheet right after 'ניתוח נתונים' with three sections —
          (1) bank transfers, (2) P2P payments (app rows, card rows funding them, withdrawals
          to the bank, duplicates), (3) PayBox-like payments — one row per database row, sorted
          by amount desc, with SUMIFS subtotals per section and direction (F17: נכנס iff
          (type == הכנסה) XOR (amount < 0); amounts shown absolute). Payee / annotation columns
          come from the row itself (counterparty in `original_name`, 'העברה ל…' in `details`),
          from `rule_note`, from the user's label files (rules.user_labels,
          rules.user_labels_interpreted) and from the p2p[] config — no built-in payee dictionary.
Inputs  : tazrim.config.json, work/database.csv, outputs/תזרים.xlsx (built by build_excel.py).
Outputs : the workbook saved in place (other sheets untouched); one JSON object on stdout
          {ok, workbook, sheet, skipped?, sections{bank,p2p,paybox: {rows,out,in}}, flagged}.
Exit    : 0 ok (also when the level has no transfers sheet: skipped=true); 1 on failure; 2 config.

openpyxl only (no Excel automation); re-runnable. Python 3.8 compatible.
"""
import datetime as _dt
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (  # noqa: E402
    DB_COLUMNS, SHEETS_BY_LEVEL, SHEET_ANALYSIS, SHEET_TRANSFERS, TYPE_CARD_DEBIT, TYPE_DUPLICATE,
    TYPE_INCOME, TYPE_INTERNAL, TYPE_REIMBURSABLE, TYPE_REIMBURSEMENT, TYPE_SAVINGS, emit, fail,
    log, parse_args, project_path, read_csv,
)

try:
    import pandas as pd
    import openpyxl
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L
except ImportError as _e:  # pragma: no cover
    fail("missing dependency: %s" % _e, "pip install -r requirements.txt (openpyxl, pandas)")

# ---------------------------------------------------------------- styles
HDR_FILL, HDR_FONT = PatternFill("solid", fgColor="1F4E78"), Font(bold=True, color="FFFFFF")
TOT_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD, TITLE = Font(bold=True), Font(bold=True, size=13)
SECTION, GREY, RED = Font(bold=True, size=12, color="1F4E78"), Font(italic=True, color="808080"), Font(bold=True, color="C00000")
_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(top=_thin, bottom=_thin, left=_thin, right=_thin)
NUM, DATE = "#,##0.00", "DD/MM/YYYY"
CENTER, NOWRAP = Alignment(horizontal="center", vertical="center"), Alignment(wrap_text=False)

COLS = ["מזהה", "סוג תנועה", "כיוון", "תאריך עסקה", "חודש", "סכום ₪", "מוטב / צד שני", "שם מובן",
        "קטגוריה (תזרים)", "קבוצה", "אמצעי תשלום / חשבון", "אדם", "בחלון", "נסכם", "הערת סיווג",
        "הערה מהמשתמש / מהצילום", "פרטים מהמקור", "מזהה מקושר"]
WIDTHS = [11, 17, 8, 12, 12, 12, 28, 36, 24, 16, 24, 14, 7, 7, 40, 36, 50, 12]
(C_ID, C_TYPE, C_DIR, C_DATE, C_MONTH, C_AMT, C_PAYEE, C_NAME, C_CAT, C_GRP, C_PAY, C_PERSON,
 C_WIN, C_SUM, C_RULE, C_NOTE, C_DET, C_LINK) = range(1, 19)
OUT, IN = "יוצא", "נכנס"

#: Generic bank-description patterns of transfers / cheques / cash / own-account movements.
#: (institution vocabulary, not personal data)
BANK_PATS = ["העברה", "הע.", 'הו"ק', "הו״ק", "משיכת שיק", "הפקדת שיק", "ביטול הפקדת", "משיכה", "כספומט",
             "סניפומט", 'הפקדות קופ"ג', "קופת גמל", "נע-קניה", "ניירות ערך", "העברה מהחשבון", "Withdrawal"]
PAYBOX_RE = re.compile(r"paybox|pay box|פייבוקס", re.IGNORECASE)
UNKNOWN_RE = re.compile(r"לא ידוע|לבדיקה|לא מזוהה|לאימות")


def s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def marker_re(marker):
    """The P2P card marker as a whole Latin word (so 'מוביט' or 'HABIT' do not match)."""
    return re.compile(r"(?<![A-Za-z])%s(?![A-Za-z])" % re.escape(marker), re.IGNORECASE)


HEB_P2P_RE = re.compile(r"(^|[\s\-–.,/])ביט($|[\s\-–.,/])")


INBOUND_TYPES = (TYPE_INCOME, TYPE_REIMBURSEMENT)


def direction(typ, amount):
    """F17: נכנס iff (type is inbound: הכנסה / החזר הוצאה) XOR (amount < 0)."""
    return IN if ((typ in INBOUND_TYPES) != (float(amount) < 0)) else OUT


def card_recipient(details):
    m = re.search(r"העברה ל\s*(.+?)(?:\s*\||$)", s(details))
    return m.group(1).strip() if m else ""


def screenshot_note(details):
    """Free text of an app row's details: drop status words and the direction word."""
    out = []
    for seg in s(details).split("|"):
        seg = seg.strip()
        if not seg or seg.lower() in ("completed", "הושלם", OUT, IN):
            continue
        m = re.match(r"^(יוצא|נכנס)\s*(.*)$", seg)
        if m:
            seg = m.group(2).strip()
            if not seg:
                continue
        out.append(seg)
    return " | ".join(out)


def load_user_notes(cfg):
    """{id: text} from rules.user_labels_interpreted (note) then rules.user_labels (user_text)."""
    notes = {}
    for key, col in (("rules.user_labels", "user_text"), ("rules.user_labels_interpreted", "note")):
        path = project_path(cfg, key)
        if os.path.isfile(path):
            try:
                for r in read_csv(path):
                    if r.get("id") and (r.get(col) or "").strip():
                        notes[r["id"]] = r[col].strip()
            except Exception as e:  # noqa: BLE001
                log("WARNING: could not read %s: %s" % (path, e))
    return notes


def build_rows(cfg, db, notes):
    """Three lists of row dicts (bank, p2p, paybox)."""
    acct_names = set()
    for a in cfg.get("accounts") or []:
        acct_names.update([a.get("label") or "", a["id"]])
    p2p_entries = cfg.get("p2p") or []
    pb_names, p2p_names, markers = set(), set(), []
    for p in p2p_entries:
        names = {p.get("label") or "", p["id"], p.get("app") or ""}
        if PAYBOX_RE.search(p.get("app") or ""):
            pb_names |= names
        else:
            p2p_names |= names
            if p.get("card_marker"):
                markers.append(marker_re(p["card_marker"]))
    if not markers:
        markers.append(marker_re("BIT"))

    name, det, src, card = db["original_name"], db["details"], db["source"], db["card"]
    is_bank = src.isin(acct_names) | src.str.startswith('עו"ש') | src.str.startswith("עו״ש")
    hit_marker = (name + " " + card).apply(lambda t: any(m.search(t) for m in markers))
    heb_p2p = name.apply(lambda t: bool(HEB_P2P_RE.search(t)))
    sel_pb = name.apply(lambda t: bool(PAYBOX_RE.search(t))) | det.apply(lambda t: bool(PAYBOX_RE.search(t))) | src.isin(pb_names) | card.isin(pb_names)
    sel_p2p = (src.isin(p2p_names) | card.isin(p2p_names) | hit_marker | heb_p2p | (db["type"] == TYPE_DUPLICATE)) & ~sel_pb
    bank_pat = name.apply(lambda t: any(p in t for p in BANK_PATS))
    bank_type = db["type"].isin([TYPE_INTERNAL, TYPE_SAVINGS, TYPE_REIMBURSABLE, TYPE_REIMBURSEMENT])
    sel_bank = is_bank & (bank_pat | bank_type) & (db["type"] != TYPE_CARD_DEBIT) & ~sel_p2p & ~sel_pb

    linked_note = {}     # card row id -> (counterparty, note) of the app row it funded
    for _, r in db[src.isin(p2p_names)].iterrows():
        if r["linked_id"]:
            linked_note[r["linked_id"]] = (r["original_name"], screenshot_note(r["details"]))

    def base(r):
        try:
            d = _dt.datetime.strptime(r["txn_date"][:10], "%Y-%m-%d").date()
        except ValueError:
            d = r["txn_date"] or None
        return {"id": r["id"], "type": r["type"], "dir": direction(r["type"], r["amount"]), "date": d,
                "month": r["month_name"] or r["month"], "amount": abs(float(r["amount"])), "payee": "",
                "name": r["name_clean"], "cat": r["cat_tz"], "grp": r["group_tz"], "pay": r["pay"],
                "person": r["person"], "win": r["in_window"], "summed": r["summed"], "rule": r["rule_note"],
                "note": notes.get(r["id"], ""), "details": r["details"], "link": r["linked_id"]}

    bank_rows = []
    for _, r in db[sel_bank].iterrows():
        d = base(r)
        d["payee"] = card_recipient(r["details"]) or r["original_name"]
        bank_rows.append(d)

    p2p_rows = []
    for _, r in db[sel_p2p].iterrows():
        d = base(r)
        if r["source"] in p2p_names:                       # app row: counterparty is the original name
            d["payee"] = r["original_name"]
            sn = screenshot_note(r["details"])
            d["note"] = " | ".join(x for x in (d["note"], sn) if x)
        elif is_bank[_]:                                    # bank row (withdrawal to bank etc.)
            d["payee"] = card_recipient(r["details"]) or r["original_name"]
        else:                                               # card row funding an app payment
            rec = card_recipient(r["details"])
            if r["id"] in linked_note:
                cp, note = linked_note[r["id"]]
                d["payee"] = "%s (בצילום: %s)" % (rec, cp) if rec and rec != cp else cp
                d["note"] = " | ".join(x for x in (d["note"], ("%s: %s" % (cp, note)) if note else cp) if x)
            else:
                d["payee"] = rec or "לא ידוע (אין שם מוטב במקור)"
        p2p_rows.append(d)

    pb_rows = []
    for _, r in db[sel_pb].iterrows():
        d = base(r)
        rec = card_recipient(r["details"])
        segs = [x.strip() for x in s(r["details"]).split("|") if x.strip() and not PAYBOX_RE.search(x)]
        d["payee"] = rec or (r["original_name"] if r["source"] in pb_names else "") or r["name_clean"]
        d["note"] = " | ".join(x for x in [d["note"]] + segs if x and "העברה ל" not in x)
        pb_rows.append(d)

    key = lambda d: -d["amount"]  # noqa: E731
    return sorted(bank_rows, key=key), sorted(p2p_rows, key=key), sorted(pb_rows, key=key)


def write_section(ws, r, title, rows):
    ws.cell(r, 1, title).font = SECTION
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(COLS))
    ws.cell(r, 1).alignment = Alignment(horizontal="right", vertical="center")
    r += 1
    for c, h in enumerate(COLS, 1):
        cell = ws.cell(r, c, h)
        cell.fill, cell.font, cell.alignment, cell.border = HDR_FILL, HDR_FONT, CENTER, BORDER
    r += 1
    first = r
    for d in rows:
        vals = [d["id"], d["type"], d["dir"], d["date"], d["month"], d["amount"], d["payee"], d["name"], d["cat"],
                d["grp"], d["pay"], d["person"], d["win"], d["summed"], d["rule"], d["note"], d["details"], d["link"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v if v != "" else None)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"
            cell.border, cell.alignment = BORDER, NOWRAP
        ws.cell(r, C_AMT).number_format = NUM
        ws.cell(r, C_DATE).number_format = DATE
        for c in (C_DIR, C_WIN, C_SUM):
            ws.cell(r, c).alignment = CENTER
        if UNKNOWN_RE.search(d["rule"] or "") or d["payee"].startswith("לא ידוע"):
            ws.cell(r, C_PAYEE).font = RED
        r += 1
    last = r - 1
    if rows:
        amt = "%s%d:%s%d" % (L(C_AMT), first, L(C_AMT), last)
        dr = "%s%d:%s%d" % (L(C_DIR), first, L(C_DIR), last)
        ids = "%s%d:%s%d" % (L(C_ID), first, L(C_ID), last)
        ws.conditional_formatting.add(amt, ColorScaleRule(start_type="min", start_color="FFFFFF", mid_type="percentile",
                                                          mid_value=50, mid_color="FFE699", end_type="max", end_color="F4B183"))
    else:
        amt = dr = ids = None
    ws.cell(r, 1, 'סה"כ').font = BOLD
    ws.cell(r, 2, "מס' שורות:").font = BOLD
    ws.cell(r, 3, "=COUNTA(%s)" % ids if ids else 0).font = BOLD
    ws.cell(r, 5, 'סה"כ יוצא:').font = BOLD
    ws.cell(r, 6, '=SUMIFS(%s,%s,"%s")' % (amt, dr, OUT) if amt else 0).font = BOLD
    ws.cell(r, 6).number_format = NUM
    ws.cell(r, 7, 'סה"כ נכנס:').font = BOLD
    ws.cell(r, 8, '=SUMIFS(%s,%s,"%s")' % (amt, dr, IN) if amt else 0).font = BOLD
    ws.cell(r, 8).number_format = NUM
    for c in range(1, len(COLS) + 1):
        ws.cell(r, c).fill, ws.cell(r, c).border = TOT_FILL, BORDER
    return r + 2, (first, last, r)


def build_sheet(wb, bank_rows, p2p_rows, pb_rows):
    if SHEET_TRANSFERS in wb.sheetnames:
        del wb[SHEET_TRANSFERS]
    idx = wb.sheetnames.index(SHEET_ANALYSIS) + 1 if SHEET_ANALYSIS in wb.sheetnames else len(wb.sheetnames)
    ws = wb.create_sheet(SHEET_TRANSFERS, index=idx)
    ws.sheet_view.rightToLeft = True
    ws["A1"] = "העברות בנקאיות, תשלומי P2P (BIT וכד') ו-PAYBOX — פירוט מלא (הוצאות, הכנסות והעברות פנימיות)"
    ws["A1"].font = TITLE
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    ws["A1"].alignment = Alignment(horizontal="right")
    ws["A2"] = ('כל התנועות מחשבונות העו"ש, מאפליקציות ה-P2P ומכרטיסי האשראי (P2P / PAYBOX), בתוך החלון ומחוצה לו. '
                '"נסכם"=כן: התנועה נספרת בתזרים (הוצאה/הכנסה). העברות פנימיות, חיסכון והשקעות וכפילויות (שורת כרטיס שמימנה תשלום P2P) אינם נסכמים. '
                'כיוון: נכנס=זיכוי, יוצא=חיוב. סכומים בערך מוחלט. סכומי הביניים הם נוסחאות SUMIFS לפי "כיוון". מוטב באדום = לא ידוע / לבדיקה.')
    ws["A2"].font = GREY
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    ws["A2"].alignment = Alignment(horizontal="right", wrap_text=True)
    ws.row_dimensions[2].height = 32
    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[L(c)].width = w
    r = 4
    r, sec1 = write_section(ws, r, '1. העברות בנקאיות (חשבונות עו"ש): העברות, הו"ק, שיקים, מזומן, חיסכון והשקעות', bank_rows)
    r, sec2 = write_section(ws, r, "2. תשלומי P2P (BIT וכד') — שורות האפליקציה, שורות הכרטיס שמימנו אותן ומשיכות לבנק", p2p_rows)
    r, sec3 = write_section(ws, r, "3. PAYBOX ודומיו (שורות כרטיס / אפליקציה)", pb_rows)
    return ws, (sec1, sec2, sec3)


def main(argv=None):
    def extra(p):
        p.add_argument("--database", default=None, help="database.csv (default: config work.database)")
        p.add_argument("--workbook", default=None, help="workbook to update (default: config outputs.workbook)")
    args, cfg = parse_args(__doc__.splitlines()[0], extra, argv)
    xlsx = args.workbook or project_path(cfg, "outputs.workbook")
    db_path = args.database or project_path(cfg, "work.database")
    if SHEET_TRANSFERS not in SHEETS_BY_LEVEL[cfg.level]:
        emit({"ok": True, "workbook": xlsx, "sheet": SHEET_TRANSFERS, "skipped": True,
              "reason": "level %s has no transfers sheet" % cfg.level})
        return
    if not os.path.isfile(xlsx):
        fail("workbook not found: %s" % xlsx, "run build_excel.py first")
    if not os.path.isfile(db_path):
        fail("database not found: %s" % db_path, "run classify.py first")
    db = pd.read_csv(db_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    missing = [c for c in DB_COLUMNS if c not in db.columns]
    if missing:
        fail("database.csv is missing columns: %s" % ", ".join(missing), "work/database.csv must have common.DB_COLUMNS")
    db["amount"] = pd.to_numeric(db["amount"], errors="coerce").fillna(0.0)
    notes = load_user_notes(cfg)
    bank_rows, p2p_rows, pb_rows = build_rows(cfg, db, notes)
    try:
        wb = openpyxl.load_workbook(xlsx)
        ws, secs = build_sheet(wb, bank_rows, p2p_rows, pb_rows)
        wb.save(xlsx)
    except Exception as e:  # noqa: BLE001
        fail("could not update the workbook: %s: %s" % (type(e).__name__, e),
             "close the workbook in Excel (without saving) and rerun")
        return

    def tot(rows):
        return {"rows": len(rows), "out": round(sum(d["amount"] for d in rows if d["dir"] == OUT), 2),
                "in": round(sum(d["amount"] for d in rows if d["dir"] == IN), 2)}
    flagged = [d["id"] for d in bank_rows + p2p_rows + pb_rows if UNKNOWN_RE.search(d["rule"] or "") or d["payee"].startswith("לא ידוע")]
    log("saved %s | sheet index %d | bank %d / p2p %d / paybox %d rows | flagged %d" % (
        xlsx, wb.sheetnames.index(SHEET_TRANSFERS), len(bank_rows), len(p2p_rows), len(pb_rows), len(flagged)))
    emit({"ok": True, "workbook": xlsx, "sheet": SHEET_TRANSFERS, "sheet_index": wb.sheetnames.index(SHEET_TRANSFERS),
          "sections": {"bank": tot(bank_rows), "p2p": tot(p2p_rows), "paybox": tot(pb_rows)},
          "section_rows": {"bank": list(secs[0]), "p2p": list(secs[1]), "paybox": list(secs[2])},
          "flagged": flagged})


if __name__ == "__main__":
    main()
