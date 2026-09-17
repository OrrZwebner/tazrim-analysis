#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify.py — the workbook verification gate (P13).

Purpose : prove that the formula workbook agrees with an independent pandas recomputation of
          work/database.csv.
          (1) formula-error scan of every sheet of the workbook;
          (2) when verify.excel_recalc is `always`, or `auto` on macOS with Microsoft Excel
              installed: AppleScript recalculation of a temp copy (display alerts off), saved as
              a calculated copy, loaded with data_only=True; totals (rows found by their EXACT
              label — subtotal rows share the prefix), every category × month cell, the rows-check
              cell, the transfers-sheet subtotals and a LIVE filter of the `פילוח` sheet (first
              group / month / person found in the data) are compared with pandas (tolerance 0.01);
          (3) otherwise the pandas-only fallback: every SUMIFS criterion of the statistic sheets
              references an existing named range and a category present in the data, expected
              totals are recomputed, no `#REF!`/`#NAME?` appears in any formula or cached value,
              duplicate-row report, card-vs-bank reconciliation — and the Excel checks are
              REPORTED AS `skipped` with the reason (never as passed).
Inputs  : tazrim.config.json, work/database.csv, outputs/תזרים.xlsx;
          --excel-recalc auto|always|never overrides the config.
Outputs : one JSON object on stdout {ok, excel_recalc: {mode, ran, reason}, checks: [{name,
          status: pass|fail|skipped, detail}], totals}; diagnostics on stderr. Temp copies
          the temp copy work/verify_calc.xlsx (saved in place by Excel) is deleted afterwards.
Exit    : 0 when no check failed (skipped checks do not fail); 1 when a check failed or the
          workbook / database is missing; 2 on a config error.

Python 3.8 compatible. Needs openpyxl and pandas; Excel checks need macOS + Microsoft Excel.
"""
import os
import platform
import re
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (  # noqa: E402
    DB_COLUMNS, SHEETS_BY_LEVEL, SHEET_EXP, SHEET_HELPER, SHEET_INC, SHEET_PIVOT, SHEET_PROPOSED,
    SHEET_TRANSFERS, TYPE_CARD_DEBIT, TYPE_EXPENSE, TYPE_INCOME, TYPE_INTERNAL, TYPE_REIMBURSABLE,
    TYPE_REIMBURSEMENT, TYPE_SAVINGS, YES, emit, fail, log, month_label, months, parse_args,
    project_path,
)

try:
    import pandas as pd
    import openpyxl
except ImportError as _e:  # pragma: no cover
    fail("missing dependency: %s" % _e, "pip install -r requirements.txt (openpyxl, pandas)")

TOTAL_EXP_LABEL = 'סה"כ הוצאות שוטפות (נסכם ישירות מ-database)'
TOTAL_INC_LABEL = 'סה"כ הכנסות (נסכם ישירות מ-database)'
ERR_RE = re.compile(r"#(REF!|NAME\?|DIV/0!|VALUE!|N/A|NUM!|NULL!)")
NAME_RE = re.compile(r"(?<![A-Za-z_])(DB[a-z0-9]+|L_[a-z_]+|A_[a-z_]+)(?![A-Za-z0-9_])")
TOL = 0.01
EXCEL_APP = "/Applications/Microsoft Excel.app"
TMP_NAME, CALC_NAME = "verify_tmp.xlsx", "verify_calc.xlsx"


class Checks(object):
    def __init__(self):
        self.items = []

    def add(self, name, status, detail=""):
        self.items.append({"name": name, "status": status, "detail": detail})
        log("[%s] %s — %s" % (status.upper(), name, detail))

    def passed(self, name, ok, detail=""):
        self.add(name, "pass" if ok else "fail", detail)

    @property
    def ok(self):
        return not any(c["status"] == "fail" for c in self.items)


# ----------------------------------------------------------------------------- data
def load_db(path):
    db = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    missing = [c for c in DB_COLUMNS if c not in db.columns]
    if missing:
        fail("database.csv is missing columns: %s" % ", ".join(missing), "work/database.csv must have common.DB_COLUMNS")
    db["amount"] = pd.to_numeric(db["amount"], errors="coerce").fillna(0.0)
    db["_in"] = db["in_window"].str.strip() == YES
    db["_summed"] = db["summed"].str.strip() == YES
    return db


def month_labels(cfg, db):
    out = {}
    for m in months(cfg):
        names = [x for x in db.loc[db["month"] == m, "month_name"] if x]
        out[m] = names[0] if names else month_label(m)
    return out


def expected_totals(cfg, db):
    ms = months(cfg)
    w = db[db["_summed"]]
    E, I = w[w["type"] == TYPE_EXPENSE], w[w["type"] == TYPE_INCOME]
    by = lambda d: {m: round(float(d[d["month"] == m]["amount"].sum()), 2) for m in ms}  # noqa: E731
    inw = db[db["_in"]]
    return {
        "months": ms, "expense": round(float(E["amount"].sum()), 2), "income": round(float(I["amount"].sum()), 2),
        "expense_by_month": by(E), "income_by_month": by(I),
        "expense_by_category": {c: round(float(v), 2) for c, v in E.groupby("cat_tz")["amount"].sum().items()},
        "income_by_category": {c: round(float(v), 2) for c, v in I.groupby("cat_tz")["amount"].sum().items()},
        "savings": round(float(inw[inw["type"] == TYPE_SAVINGS]["amount"].sum()), 2),
        "internal": round(float(inw[inw["type"] == TYPE_INTERNAL]["amount"].sum()), 2),
        "card_debits": round(float(inw[inw["type"] == TYPE_CARD_DEBIT]["amount"].sum()), 2),
        "reimbursable_open": round(float(inw[inw["type"] == TYPE_REIMBURSABLE]["amount"].sum() -
                                         inw[inw["type"] == TYPE_REIMBURSEMENT]["amount"].sum()), 2),
    }


# ----------------------------------------------------------------------------- sheet lookups
def header_map(ws, row=2):
    return {str(ws.cell(row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(row, c).value not in (None, "")}


def find_total_row(ws, label, col=2):
    """Row whose column `col` equals `label` EXACTLY (never startswith: subtotals share the prefix)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == label:
            return r
    return None


def stat_layout(ws, cfg, mlabels):
    """(month_cols {m: col}, total_col, avg_col, nz_col) of a statistic sheet from its header row."""
    hm = header_map(ws)
    mcols = {m: hm[mlabels[m]] for m in months(cfg) if mlabels[m] in hm}
    tot = next((c for h, c in hm.items() if h.startswith('סה"כ') and h.endswith("חודשים")), None)
    return mcols, tot, hm.get("ממוצע חודשי"), hm.get("ממוצע ללא אפס")


def category_rows(ws, first, last, cats):
    """{category: row} for rows whose column 2 is a data category (subtotals / totals excluded)."""
    out = {}
    for r in range(first, last):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v in cats and not str(ws.cell(r, 1).value or "").startswith('סה"כ'):
            out[v] = r
    return out


def scan_errors(wb, skip_sheets):
    errs = []
    for ws in wb.worksheets:
        if ws.title in skip_sheets:
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                text = v.text if hasattr(v, "text") else v
                if isinstance(text, str) and ERR_RE.search(text) and (text.startswith("#") or text.startswith("=")):
                    errs.append("%s!%s: %s" % (ws.title, c.coordinate, text[:60]))
    return errs


# ----------------------------------------------------------------------------- pandas-only checks
def check_formulas(wb, cfg, db, mlabels, chk):
    """Named ranges exist; helper SUMIFS criteria reference categories present in the data;
    every category row of the statistic sheets points at a helper row of the same category/type."""
    names = set(wb.defined_names.keys()) if hasattr(wb.defined_names, "keys") else {d.name for d in wb.defined_names.definedName}
    used, missing = set(), set()
    for title in (SHEET_HELPER, SHEET_EXP, SHEET_INC, SHEET_PIVOT, SHEET_PROPOSED):
        if title not in wb.sheetnames:
            continue
        for row in wb[title].iter_rows():
            for c in row:
                v = c.value
                text = v.text if hasattr(v, "text") else v
                if isinstance(text, str) and text.startswith("="):
                    for nm in NAME_RE.findall(text):
                        used.add(nm)
                        if nm not in names:
                            missing.add(nm)
    chk.passed("named ranges referenced by formulas exist", not missing,
               "%d names used; missing: %s" % (len(used), ", ".join(sorted(missing)) or "none"))
    # helper rows: category present in the data for that type
    pv = wb[SHEET_HELPER]
    cats_by_type = {t: set(db.loc[(db["type"] == t) & (db["_in"]), "cat_tz"]) for t in set(db["type"])}
    cats_new = set(db.loc[db["_summed"], "cat_new"])
    bad, nrows = [], 0
    for r in range(3, pv.max_row + 1):
        scheme, typ, cat = pv.cell(r, 1).value, pv.cell(r, 2).value, pv.cell(r, 4).value
        if pv.cell(r, 3).value == 'סה"כ' or not typ:
            continue
        nrows += 1
        f = pv.cell(r, 5).value
        if not (isinstance(f, str) and f.startswith("=SUMIFS(DBamt,")):
            bad.append("row %d: no SUMIFS" % r)
            continue
        universe = cats_new if scheme == "מוצעת" else cats_by_type.get(typ, set())
        if cat not in universe:
            bad.append("row %d: %s/%s not in data" % (r, typ, cat))
    chk.passed("helper SUMIFS rows reference categories present in the data", not bad,
               "%d helper rows; problems: %s" % (nrows, "; ".join(bad[:5]) or "none"))
    # statistic sheets: category rows -> helper row with the same category and type
    for title, typ, label in ((SHEET_EXP, TYPE_EXPENSE, TOTAL_EXP_LABEL), (SHEET_INC, TYPE_INCOME, TOTAL_INC_LABEL)):
        ws = wb[title]
        tr = find_total_row(ws, label)
        if tr is None:
            chk.add("%s: total row found by exact label" % title, "fail", "label %r not found" % label)
            continue
        mcols, tot_col, _, _ = stat_layout(ws, cfg, mlabels)
        rows = category_rows(ws, 3, tr, cats_by_type.get(typ, set()))
        problems = []
        for cat, r in rows.items():
            f = ws.cell(r, mcols[months(cfg)[0]]).value
            m = re.match(r"^='?%s'?!([A-Z]+)(\d+)$" % re.escape(SHEET_HELPER), str(f))
            if not m:
                problems.append("%s: month cell is not a helper reference (%s)" % (cat, f))
                continue
            hr = int(m.group(2))
            if pv.cell(hr, 4).value != cat or pv.cell(hr, 2).value != typ:
                problems.append("%s: helper row %d holds %s/%s" % (cat, hr, pv.cell(hr, 2).value, pv.cell(hr, 4).value))
        f = ws.cell(tr, mcols[months(cfg)[0]]).value
        total_ok = isinstance(f, str) and SHEET_HELPER in f
        chk.passed("%s: total row found by exact label and references the helper/database" % title, total_ok and tr is not None,
                   "row %d, month formula %s" % (tr, f))
        chk.passed("%s: %d category rows point at helper rows of the same category" % (title, len(rows)), not problems,
                   "; ".join(problems[:5]) or "all consistent; months %s; total col %s" % (sorted(mcols.values()), tot_col))
        expected_cats = cats_by_type.get(typ, set()) & set(db.loc[db["_summed"] & (db["type"] == typ), "cat_tz"])
        chk.passed("%s: every summed category has a row" % title, expected_cats <= set(rows),
                   "missing: %s" % (", ".join(sorted(expected_cats - set(rows))) or "none"))
        subtot = [r for r in range(3, tr) if str(ws.cell(r, 2).value or "").startswith('סה"כ ')]
        chk.passed("%s: group subtotal rows present" % title, len(subtot) >= 1 or not rows, "%d subtotal rows" % len(subtot))


def check_duplicates(db, chk):
    e = db[db["_summed"]]
    d = e[e.duplicated(["source", "txn_date", "amount", "original_name"], keep=False)]
    chk.add("potential duplicate rows (same source/date/amount/name) — report only", "pass",
            "%d rows: %s" % (len(d), ", ".join(d["id"].head(10))))


def card_vs_bank(cfg, db):
    out = {}
    inw = db[db["_in"]]
    for c in cfg.get("cards") or []:
        label, last4, pat = c.get("label") or c["id"], str(c.get("last4") or ""), c.get("bank_debit_pattern") or ""
        stmt = inw[(inw["type"] == TYPE_EXPENSE) & ((inw["source"] == label) | (inw["card"] == label) | (inw["source"] == c["id"]) |
                                                    ((inw["card"] == last4) & (last4 != "")))]
        bank = inw[(inw["type"] == TYPE_CARD_DEBIT) & ((inw["original_name"].str.contains(re.escape(pat), regex=True) if pat else False) |
                                                        (inw["cat_tz"] == label))]
        out[c["id"]] = {"statement_rows": round(float(stmt["amount"].sum()), 2), "bank_debits": round(float(bank["amount"].sum()), 2),
                        "diff": round(float(stmt["amount"].sum() - bank["amount"].sum()), 2)}
    return out


# ----------------------------------------------------------------------------- Excel
def excel_available():
    return platform.system() == "Darwin" and os.path.isdir(EXCEL_APP)


def osa(script, timeout=300):
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)
    return r.returncode, r.stdout.strip(), r.stderr.strip()


def esc(v):
    return str(v).replace("\\", "\\\\").replace('"', '\\"')


def close_temp_workbooks():
    """Close our temp workbooks if Excel holds them (saving no). Returns the still-open names."""
    script = '''
if application "Microsoft Excel" is running then
  tell application "Microsoft Excel"
    set display alerts to false
    repeat with nm in {"%s", "%s"}
      try
        close workbook nm saving no
      end try
    end repeat
    set display alerts to true
    return name of every workbook
  end tell
else
  return ""
end if''' % (TMP_NAME, CALC_NAME)
    code, out, err = osa(script, timeout=60)
    if code != 0:
        return ["<osascript error: %s>" % err]
    return [n.strip() for n in out.split(",") if n.strip() in (TMP_NAME, CALC_NAME)]


def pivot_cells(ws):
    """Input / KPI / table cells of the פילוח sheet found by their labels."""
    inputs = {}
    for r in range(3, 15):
        lab = ws.cell(r, 2).value
        if isinstance(lab, str):
            inputs[lab] = "C%d" % r
    kpi = {}
    for r in range(3, 15):
        lab = ws.cell(r, 5).value
        if isinstance(lab, str):
            kpi[lab] = "F%d" % r
    t0 = next((r for r in range(3, 60) if str(ws.cell(r, 2).value or "").startswith("בית עסק")), None)
    return inputs, kpi, (t0 + 1) if t0 else None


def dismiss_sandbox_dialog():
    """Best effort: cancel Excel's sandbox 'Grant File Access' dialog (needs assistive access; ignored if not)."""
    script = """
with timeout of 10 seconds
tell application "System Events" to tell process "Microsoft Excel"
  if exists window "Grant File Access" then
    click button "Cancel" of window "Grant File Access"
    return "dismissed"
  end if
  return ""
end tell
end timeout"""
    try:
        return osa(script, timeout=20)[1] == "dismissed"
    except Exception:  # noqa: BLE001
        return False


def excel_recalc(cfg, xlsx, workdir, live, chk):
    """Copy the workbook to work/verify_calc.xlsx, open it in Excel, calculate and save IN PLACE
    (a 'save as' into a folder without a sandbox grant pops 'Grant File Access'), then set the
    פילוח live filter and read its cells; close without saving. Returns (calc_path, live_values)."""
    calc = os.path.join(workdir, CALC_NAME)
    still = close_temp_workbooks()
    if still:
        chk.add("excel: temp copies not open in Excel", "fail",
                "%s is still open in Excel — close it (without saving) and rerun" % ", ".join(still))
        return None, None
    for p in (calc, os.path.join(workdir, TMP_NAME)):
        if os.path.exists(p):
            try:
                os.remove(p)
            except OSError as e:
                chk.add("excel: temp copies removable", "fail", "cannot delete %s (%s) — is it open in Excel?" % (p, e))
                return None, None
    shutil.copy(xlsx, calc)
    live_script = ""
    if live:
        sets = "\n".join('    set value of range "%s" of d to "%s"' % (cell, esc(val)) for cell, val in live["set"])
        reads = ' & "|" & '.join('(value of range "%s" of d)' % c for c in live["read"])
        live_script = '''
  set d to sheet "%s" of wb
%s
  calculate
  set out to %s''' % (esc(SHEET_PIVOT), sets, reads)
    script = '''
with timeout of 180 seconds
tell application "Microsoft Excel"
  set display alerts to false
  set screen updating to false
  set out to ""
  open POSIX file "%s"
  set wb to workbook "%s"
  calculate
  save wb
%s
  close wb saving no
  set screen updating to true
  set display alerts to true
  return out
end tell
end timeout''' % (esc(calc), CALC_NAME, live_script)
    try:
        code, out, err = osa(script, timeout=400)
    except subprocess.TimeoutExpired:
        code, out, err = 1, "", "osascript timed out after 400 s"
    if code != 0:
        dismissed = dismiss_sandbox_dialog()
        close_temp_workbooks()
        hint = ""
        if dismissed or "-1712" in err:
            hint = (" | Excel showed its sandbox 'Grant File Access' dialog%s — open the project folder once in Excel "
                    "(File > Open) to grant access, or set verify.excel_recalc=never" % (" (dismissed)" if dismissed else ""))
        elif "-50" in err and not os.path.abspath(workdir).startswith(os.path.expanduser("~")):
            hint = " | hint: sandboxed Excel cannot save outside your home folder — keep the project under ~/"
        chk.add("excel: recalculation", "fail", "osascript rc=%d; stderr: %s%s" % (code, err[:300], hint))
        return None, None
    chk.add("excel: recalculation of a temp copy (display alerts off, saved in place)", "pass", "saved %s" % calc)
    return calc, out


def compare_with_excel(calc_path, cfg, db, mlabels, exp, live, chk):
    wb = openpyxl.load_workbook(calc_path, data_only=True)
    errs = scan_errors(wb, set())
    chk.passed("excel: no formula errors in the calculated sheets", not errs, "; ".join(errs[:10]) or "0 errors")
    ms = months(cfg)
    for title, typ, label, key in ((SHEET_EXP, TYPE_EXPENSE, TOTAL_EXP_LABEL, "expense"), (SHEET_INC, TYPE_INCOME, TOTAL_INC_LABEL, "income")):
        ws = wb[title]
        tr = find_total_row(ws, label)
        if tr is None:
            chk.add("excel: %s total row" % title, "fail", "exact label %r not found" % label)
            continue
        mcols, tot_col, avg_col, nz_col = stat_layout(ws, cfg, mlabels)
        xl = float(ws.cell(tr, tot_col).value or 0)
        chk.passed("excel: %s total equals pandas" % title, abs(xl - exp[key]) < TOL, "excel %.2f | pandas %.2f" % (xl, exp[key]))
        for m in ms:
            xm = float(ws.cell(tr, mcols[m]).value or 0)
            pm = exp[key + "_by_month"][m]
            chk.passed("excel: %s total %s" % (title, m), abs(xm - pm) < TOL, "excel %.2f | pandas %.2f" % (xm, pm))
        if title == SHEET_EXP:
            chkv = ws.cell(tr + 1, tot_col).value
            chk.passed("excel: rows-check cell equals the total", chkv is not None and abs(float(chkv) - xl) < 1,
                       "check %s | total %.2f" % (chkv, xl))
        w = db[db["_summed"] & (db["type"] == typ)]
        rows = category_rows(ws, 3, tr, set(w["cat_tz"]))
        bad = []
        for cat, r in rows.items():
            for m in ms:
                py = float(w[(w["cat_tz"] == cat) & (w["month"] == m)]["amount"].sum())
                xv = float(ws.cell(r, mcols[m]).value or 0)
                if abs(py - xv) > TOL:
                    bad.append("%s %s: excel %.2f pandas %.2f" % (cat, m, xv, py))
            vals = [float(w[(w["cat_tz"] == cat) & (w["month"] == m)]["amount"].sum()) for m in ms]
            nz = [v for v in vals if abs(v) > 0.005]
            py_nz = sum(nz) / len(nz) if nz else 0.0
            xnz = float(ws.cell(r, nz_col).value or 0)
            if abs(py_nz - xnz) > TOL:
                bad.append("%s no-zero avg: excel %.2f pandas %.2f" % (cat, xnz, py_nz))
        chk.passed("excel: %s per-category/month cells equal pandas (%d categories)" % (title, len(rows)), not bad, "; ".join(bad[:8]) or "0 mismatches")
    # transfers sheet subtotals
    if SHEET_TRANSFERS in wb.sheetnames:
        try:
            import add_transfers_sheet as ats
            secs = ats.build_rows(cfg, db, ats.load_user_notes(cfg))
            ws = wb[SHEET_TRANSFERS]
            found = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'סה"כ']
            bad = []
            for (rows, r) in zip(secs, found):
                cnt, out_, in_ = ws.cell(r, 3).value, ws.cell(r, 6).value, ws.cell(r, 8).value
                po = sum(d["amount"] for d in rows if d["dir"] == ats.OUT)
                pi = sum(d["amount"] for d in rows if d["dir"] == ats.IN)
                if int(cnt or 0) != len(rows) or abs(float(out_ or 0) - po) > TOL or abs(float(in_ or 0) - pi) > TOL:
                    bad.append("row %d: excel %s/%s/%s pandas %d/%.2f/%.2f" % (r, cnt, out_, in_, len(rows), po, pi))
            chk.passed("excel: transfers-sheet subtotals equal pandas", len(found) == 3 and not bad, "; ".join(bad) or "3 sections match")
        except Exception as e:  # noqa: BLE001
            chk.add("excel: transfers-sheet subtotals", "skipped", "could not cross-check: %s" % e)
    # פילוח default selection + live filter
    if SHEET_PIVOT in wb.sheetnames and live:
        ws = wb[SHEET_PIVOT]
        sel = ws[live["sel_cell"]].value
        py = float(db[db["_summed"] & (db["type"] == TYPE_EXPENSE) & (db["cat_tz"] == sel)]["amount"].sum())
        xv = ws[live["kpi_total"]].value
        chk.passed("excel: פילוח default selection KPI equals pandas", xv is not None and abs(float(xv) - py) < TOL,
                   "selection %r: excel %s | pandas %.2f" % (sel, xv, py))
        if live.get("result") is not None:
            parts = [p.strip().strip(",").strip() for p in live["result"].split("|")]
            f = live["filter"]
            sub = db[db["_summed"] & (db["type"] == TYPE_EXPENSE) & (db["group_tz"] == f["group"]) & (db["month"] == f["month"]) & (db["person"] == f["person"])]
            top = sub.groupby("name_clean")["amount"].sum().sort_values(ascending=False)
            try:
                ok = (abs(float(parts[0]) - float(sub["amount"].sum())) < TOL and int(float(parts[1])) == len(sub) and
                      (len(top) == 0 or (parts[2] == top.index[0] and abs(float(parts[3]) - float(top.iloc[0])) < TOL)))
                detail = "filter %s: excel %s | pandas %.2f / %d / %s" % (f, parts, float(sub["amount"].sum()), len(sub), top.index[0] if len(top) else None)
            except (ValueError, IndexError) as e:
                ok, detail = False, "unparseable Excel result %r (%s)" % (live["result"], e)
            chk.passed("excel: פילוח live filter (group/month/person) equals pandas", ok, detail)
    wb.close()


# ----------------------------------------------------------------------------- main
def main(argv=None):
    def extra(p):
        p.add_argument("--excel-recalc", choices=["auto", "always", "never"], default=None, help="override verify.excel_recalc")
        p.add_argument("--workbook", default=None, help="workbook to verify (default: config outputs.workbook)")
        p.add_argument("--database", default=None, help="database.csv (default: config work.database)")
        p.add_argument("--keep-temp", action="store_true", help="keep work/verify_calc.xlsx")
    args, cfg = parse_args(__doc__.splitlines()[0], extra, argv)
    xlsx = args.workbook or project_path(cfg, "outputs.workbook")
    db_path = args.database or project_path(cfg, "work.database")
    if not os.path.isfile(xlsx):
        fail("workbook not found: %s" % xlsx, "run build_excel.py first")
    if not os.path.isfile(db_path):
        fail("database not found: %s" % db_path, "run classify.py first")
    mode = args.excel_recalc or cfg["verify"]["excel_recalc"]
    db = load_db(db_path)
    mlabels = month_labels(cfg, db)
    exp = expected_totals(cfg, db)
    chk = Checks()

    # (1) formula scan + pandas-only structural checks on the workbook as written
    wb = openpyxl.load_workbook(xlsx)
    level_sheets = [s for s in SHEETS_BY_LEVEL[cfg.level] if s != SHEET_TRANSFERS or SHEET_TRANSFERS in wb.sheetnames]
    missing_sheets = [s for s in level_sheets if s not in wb.sheetnames]
    chk.passed("sheet set for level %s" % cfg.level, not missing_sheets, "missing: %s" % (", ".join(missing_sheets) or "none"))
    errs = scan_errors(wb, set())
    chk.passed("no #REF!/#NAME? in formulas or cached values", not errs, "; ".join(errs[:10]) or "0 errors")
    check_formulas(wb, cfg, db, mlabels, chk)
    e = db[db["_summed"]]
    chk.passed("every summed row has a primary category", bool((e["cat_tz"] != "").all()),
               "%d summed rows; empty cat_tz: %d" % (len(e), int((e["cat_tz"] == "").sum())))
    check_duplicates(db, chk)
    recon = card_vs_bank(cfg, db)
    chk.add("card statements vs bank debits in the window — report only (FX rows settle outside the debit)", "pass",
            "; ".join("%s: statement %.2f vs bank %.2f (diff %.2f)" % (k, v["statement_rows"], v["bank_debits"], v["diff"]) for k, v in recon.items()) or "no cards configured")

    # live-filter parameters from the data (first summed expense row)
    live = None
    if SHEET_PIVOT in wb.sheetnames:
        inputs, kpi, t0 = pivot_cells(wb[SHEET_PIVOT])
        E = db[db["_summed"] & (db["type"] == TYPE_EXPENSE)]
        if len(E) and t0 and inputs and kpi:
            first = E.iloc[0]
            f = {"group": first["group_tz"], "month": first["month"], "person": first["person"]}
            lab_sel = next((k for k in inputs if k.startswith("בחירה")), None)
            lab_tot = next((k for k in kpi if k.startswith('סה"כ')), None)
            lab_cnt = next((k for k in kpi if k.startswith("מס")), None)
            if lab_sel and lab_tot and lab_cnt:
                live = {"filter": f, "sel_cell": inputs[lab_sel], "kpi_total": kpi[lab_tot],
                        "set": [(inputs["רמה"], "קבוצה"), (inputs[lab_sel], f["group"]), (inputs["חודש"], mlabels[f["month"]]),
                                (inputs["אמצעי תשלום"], "הכל"), (inputs["אדם"], f["person"])],
                        "read": [kpi[lab_tot], kpi[lab_cnt], "B%d" % t0, "C%d" % t0]}
    wb.close()

    # (2) Excel recalculation
    excel = {"mode": mode, "ran": False, "reason": ""}
    run_excel = mode == "always" or (mode == "auto" and excel_available())
    if mode == "never":
        excel["reason"] = "verify.excel_recalc = never"
    elif mode == "auto" and not excel_available():
        excel["reason"] = "not macOS with %s" % EXCEL_APP
    elif mode == "always" and platform.system() != "Darwin":
        run_excel, excel["reason"] = False, "excel_recalc=always but this is not macOS (%s)" % platform.system()
    workdir = project_path(cfg, "work.dir")
    os.makedirs(workdir, exist_ok=True)
    if run_excel:
        calc, out = excel_recalc(cfg, xlsx, workdir, live, chk)
        if calc:
            excel["ran"] = True
            if live is not None:
                live["result"] = out
            compare_with_excel(calc, cfg, db, mlabels, exp, live, chk)
        if not args.keep_temp:
            for p in (os.path.join(workdir, TMP_NAME), os.path.join(workdir, CALC_NAME),
                      os.path.join(workdir, "~$" + TMP_NAME), os.path.join(workdir, "~$" + CALC_NAME)):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
    else:
        for name in ("excel: recalculation of a temp copy", "excel: totals and per-category/month cells vs pandas",
                     "excel: rows-check cell", "excel: transfers-sheet subtotals", "excel: פילוח dynamic arrays (default + live filter)"):
            chk.add(name, "skipped", excel["reason"])

    result = {"ok": chk.ok, "workbook": xlsx, "level": cfg.level, "excel_recalc": excel, "checks": chk.items,
              "totals": exp, "card_vs_bank": recon,
              "summary": {"pass": sum(c["status"] == "pass" for c in chk.items), "fail": sum(c["status"] == "fail" for c in chk.items),
                          "skipped": sum(c["status"] == "skipped" for c in chk.items)}}
    log("RESULT: %s (%d pass / %d fail / %d skipped)" % ("PASS" if chk.ok else "FAIL", result["summary"]["pass"], result["summary"]["fail"], result["summary"]["skipped"]))
    emit(result)
    if not chk.ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
