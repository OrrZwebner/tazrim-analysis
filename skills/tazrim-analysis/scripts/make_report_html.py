#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_report_html.py — Hebrew RTL report (HTML, print-ready A4) from work/summary.json (P16).

Purpose : render the fixed 8-section report. EVERY number comes from summary.json (no numeric
          literal in this generator); every narrative bullet comes from work/findings.json,
          which Claude writes per run after reading summary.json:
              {"findings": [str | {"title","text"}], "actions": [{"what","why","how"}],
               "data_gaps": [str]}
          When findings.json is absent the report renders placeholders saying so.
Inputs  : tazrim.config.json; work/summary.json; work/findings.json (optional);
          work/figures/*.png (embedded as base64 data URIs; a missing figure -> placeholder).
Outputs : work/report.html (single self-contained file; dir="rtl" lang="he"; @page A4);
          ONE JSON object on stdout: {"ok", "html", "bytes", "level", "sections",
          "figures_embedded", "figures_missing", "findings_present", "residual", "residual_ok"}.
Exit    : 0 ok; 1 summary missing; 2 config error; 4 with --strict when |residual| >= 1.

Sections (Output shape §3): cover; 1 תקציר מנהלים; 2 מתודולוגיה ומקורות; 3 הכנסות; 4 הוצאות;
5 מאזן, חיסכון ותזרים; 6 ממצאים מיוחדים; 7 המלצות; 8 נספח.
Level gating: overview = cover + sections 1–4 (short); standard/deep = all 8; deep adds the
per-person income split and the per-trip / per-card detail.
Python 3.8 compatible; stdlib only.
"""
import base64
import datetime as _dt
import html as _html
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (  # noqa: E402
    SHEET_ANALYSIS, SHEET_BY_CAT, SHEET_DB, SHEET_EXP, SHEET_INC, SHEET_MANUAL, SHEET_PIVOT,
    SHEET_TRANSFERS, SHEET_TXNS, SHEET_VARIABLE, USER_CAT_COL, USER_TEXT_COL,
    emit, fail, log, owner_label, parse_args, project_path, read_json)

SECTION_TITLES = ["תקציר מנהלים", "מתודולוגיה ומקורות", "הכנסות", "הוצאות", "מאזן, חיסכון ותזרים",
                  "ממצאים מיוחדים", "המלצות", "נספח"]
OVERVIEW_SECTIONS = 4
FINDINGS_MISSING = "קובץ work/findings.json חסר — יש לכתוב אותו (ממצאים / המלצות / פערי נתונים) לאחר קריאת summary.json"
RECON_LABELS = {
    "income": "הכנסות (נסכמות, כל המקורות)",
    "expenses": "הוצאות (נסכמות, כל המקורות)",
    "savings_bank": "חיסכון והשקעות (הפקדות מהעו\"ש)",
    "internal_bank_net": "העברות פנימיות בעו\"ש, נטו (יוצא חיובי / נכנס שלילי)",
    "reimb_paid_bank": "הוצאות בהחזר ששולמו מהעו\"ש (אינן בסך ההוצאות)",
    "reimb_recv_bank": "החזרי הוצאות שהתקבלו בעו\"ש (אינם בסך ההכנסות)",
    "fx_expenses": "חיובי מט\"ח בפירוטי הכרטיסים — נספרו כהוצאה אך נסלקו מחשבון המט\"ח, לא מהעו\"ש",
    "p2p_income": "הכנסות באפליקציית תשלומים (התקבלו ביתרת האפליקציה, לא בבנק)",
    "p2p_paid_from_balance": "תשלומי אפליקציה שנספרו כהוצאה אך לא מומנו בחיוב כרטיס בחלון",
    "card_duplicates_unlinked": "שורות כפילות בכרטיס שאינן מקושרות לתשלום אפליקציה בחלון",
    "card_nonexpense_rows": "שורות כרטיס שאינן הוצאה (בהחזר / פנימי) הכלולות בחיוב הבנק",
    "club_balance_change": "מועדון הטבות: רכישות נטו פחות חיובי הבנק (שינוי יתרת הכרטיס הנטען)",
    "nonstatement_expenses": "הוצאות שלא מדף חשבון (הערכות / שורות מחושבות)",
    "nonbank_income": "הכנסות ממקור שאינו בנק/אפליקציה",
}
FIG_CAPTIONS = {
    "01": "הוצאות מול הכנסות לפי חודש", "02": "הוצאות לפי קבוצה — ממוצע חודשי",
    "03": "חלק כל קבוצה מסך ההוצאות (במרכז: ממוצע הוצאות חודשי)", "04": "הרכב ההוצאות לפי קבוצה וחודש",
    "05": "הכנסות לפי מקור — ממוצע חודשי", "06": "הכנסות לפי אדם וחודש",
    "07": "הוצאות לפי אמצעי תשלום — ממוצע חודשי", "08": "בתי העסק הגדולים — ממוצע חודשי",
    "09": "בתי העסק הגדולים — ממוצע חודשי ללא אפס", "10": "הוצאות קבועות מול משתנות לפי חודש",
    "11": "הוצאות בישראל מול חו\"ל לפי חודש", "12": "הפקדות לחיסכון והשקעות לפי חודש",
    "13": "פילוח הוצאות לפי כרטיס — ממוצע חודשי", "14": "הוצאות לפי אדם וחודש",
}

CSS = """
@page { size: A4; margin: 18mm; }
* { box-sizing: border-box; }
html, body { direction: rtl; }
body { font-family: "Arial Hebrew", Arial, "Helvetica Neue", "Noto Sans Hebrew", sans-serif; font-size: 11pt; line-height: 1.45; color: #1a1a1a; background: #fff; margin: 0; padding: 0 16px; }
@media screen { body { max-width: 190mm; margin: 0 auto; padding: 24px 16px; } }
h1 { font-size: 24pt; margin: 0 0 8pt; color: #0f2a4a; }
h2 { page-break-after: avoid; break-after: avoid; font-size: 17pt; margin: 0 0 10pt; padding-bottom: 4pt; border-bottom: 2px solid #0f2a4a; color: #0f2a4a; }
h3 { font-size: 13pt; margin: 14pt 0 6pt; color: #1f4e79; page-break-after: avoid; break-after: avoid; }
h4 { font-size: 11.5pt; margin: 10pt 0 4pt; }
p { margin: 0 0 7pt; }
ul, ol { margin: 0 0 8pt; padding-right: 20pt; padding-left: 0; }
li { margin-bottom: 3pt; }
section.page { page-break-before: always; break-before: page; }
section.page:first-of-type { page-break-before: auto; break-before: auto; }
.num { direction: ltr; unicode-bidi: isolate; display: inline-block; font-variant-numeric: tabular-nums; }
.tw { overflow-x: auto; margin: 6pt 0 10pt; }
table { border-collapse: collapse; width: 100%; font-size: 9.5pt; page-break-inside: auto; }
th, td { border: 1px solid #b9c4d0; padding: 3pt 5pt; text-align: right; vertical-align: top; }
th { background: #dbe5f1; color: #0f2a4a; font-weight: bold; }
tbody tr:nth-child(even) td { background: #f3f6fa; }
tfoot td { background: #e6ecf4; font-weight: bold; }
td.n, th.n { text-align: left; white-space: nowrap; }
tr { page-break-inside: avoid; break-inside: avoid; }
thead { display: table-header-group; }
figure { margin: 8pt 0 12pt; text-align: center; page-break-inside: avoid; break-inside: avoid; }
figure img { max-width: 100%; height: auto; border: 1px solid #d5dbe3; }
figcaption { font-size: 9.5pt; color: #4a5568; margin-top: 3pt; }
.kpis { display: flex; flex-wrap: wrap; gap: 8pt; margin: 8pt 0 12pt; }
.kpi { flex: 1 1 28%; min-width: 120px; border: 1px solid #b9c4d0; border-radius: 6px; padding: 7pt 9pt; background: #f3f6fa; page-break-inside: avoid; }
.kpi .v { font-size: 16pt; font-weight: bold; color: #0f2a4a; display: block; }
.kpi .l { font-size: 9pt; color: #4a5568; display: block; }
.cover { min-height: 230mm; display: flex; flex-direction: column; justify-content: center; }
.cover .meta { color: #4a5568; font-size: 11pt; margin-bottom: 18pt; }
.cover .box { border: 1px solid #b9c4d0; background: #f3f6fa; padding: 12pt 14pt; border-radius: 6px; }
.note { border-right: 3px solid #1f4e79; background: #f3f6fa; padding: 5pt 9pt; margin: 6pt 0 10pt; font-size: 10pt; }
.warn { border-right-color: #b7791f; background: #fdf6e3; }
.bad { color: #b42318; font-weight: bold; }
.placeholder { border: 1px dashed #b7791f; background: #fdf6e3; padding: 6pt 9pt; margin: 6pt 0 10pt; font-size: 10pt; color: #7a3a10; }
.action { border: 1px solid #b9c4d0; border-radius: 6px; padding: 6pt 9pt; margin: 0 0 7pt; page-break-inside: avoid; break-inside: avoid; }
.action h4 { margin: 0 0 3pt; color: #0f2a4a; }
.action p { margin: 0 0 2pt; font-size: 10pt; }
.action b { color: #1f4e79; }
.small { font-size: 9.5pt; color: #4a5568; }
code { direction: ltr; unicode-bidi: embed; font-size: 9.5pt; }
@media print { body { padding: 0; } .tw { overflow: visible; } }
"""


# ----------------------------------------------------------------------------- formatting
def esc(x):
    return _html.escape(str(x), quote=False)


def n0(x):
    v = int(round(float(x or 0)))
    s = "{:,}".format(abs(v))
    if v < 0:
        s = "−" + s
    return '<span class="num">%s</span>' % s


def n2(x):
    v = float(x or 0)
    s = "{:,.2f}".format(abs(v))
    if v < 0:
        s = "−" + s
    return '<span class="num">%s</span>' % s


def ils0(x):
    return n0(x) + " ₪"


def ils2(x):
    return n2(x) + " ₪"


def pct(x):
    return '<span class="num">%.1f%%</span>' % (float(x or 0) * 100.0)


def cnt(x):
    return '<span class="num">%s</span>' % "{:,}".format(int(x or 0))


def dmy(iso):
    s = str(iso or "")
    parts = s[:10].split("-")
    if len(parts) == 3:
        return '<span class="num">%s/%s/%s</span>' % (parts[2], parts[1], parts[0])
    return esc(s)


def nz_cell(st):
    k = st.get("n_nonzero", 0)
    return "%s <span class=\"small\">(%s ח')</span>" % (ils0(st.get("avg_nz", 0)), cnt(k)) if k else "—"


def table(headers, rows, foot=None, colcls=None):
    colcls = colcls or []

    def cell(tag, i, c):
        k = colcls[i] if i < len(colcls) else ""
        return "<%s class=\"%s\">%s</%s>" % (tag, k, c, tag)
    h = "".join(cell("th", i, esc(c)) for i, c in enumerate(headers))
    b = "".join("<tr>%s</tr>" % "".join(cell("td", i, c) for i, c in enumerate(r)) for r in rows)
    if not rows:
        b = "<tr><td colspan=\"%d\" class=\"small\">אין נתונים</td></tr>" % len(headers)
    f = ""
    if foot:
        f = "<tfoot><tr>%s</tr></tfoot>" % "".join(cell("td", i, c) for i, c in enumerate(foot))
    return "<div class=\"tw\"><table><thead><tr>%s</tr></thead><tbody>%s</tbody>%s</table></div>" % (h, b, f)


def month_label_of(S, ym):
    try:
        return S["month_names"][S["months"].index(ym)]
    except (ValueError, KeyError, TypeError):
        return str(ym or "—")


# ----------------------------------------------------------------------------- findings
def load_findings(path):
    if not path or not os.path.isfile(path):
        return None
    try:
        d = read_json(path)
    except ValueError as e:
        log("WARNING: findings.json is not valid JSON: %s" % e)
        return None
    if not isinstance(d, dict):
        return None
    d.setdefault("findings", [])
    d.setdefault("actions", [])
    d.setdefault("data_gaps", [])
    return d


def bullets(items, placeholder):
    if items is None:
        return "<div class=\"placeholder\">%s</div>" % esc(placeholder)
    if not items:
        return "<p class=\"small\">— אין פריטים —</p>"
    out = []
    for it in items:
        if isinstance(it, dict):
            t, x = it.get("title", ""), it.get("text", "")
            out.append("<li>%s%s</li>" % ("<b>%s:</b> " % esc(t) if t else "", esc(x)))
        else:
            out.append("<li>%s</li>" % esc(it))
    return "<ul>%s</ul>" % "".join(out)


# ----------------------------------------------------------------------------- figures
class Figures(object):
    def __init__(self, fig_dir):
        self.dir = fig_dir
        self.embedded, self.missing = [], []
        self.n = 0

    def fig(self, key):
        self.n += 1
        caption = FIG_CAPTIONS.get(key, key)
        path = None
        if self.dir and os.path.isdir(self.dir):
            for f in sorted(os.listdir(self.dir)):
                if f.startswith(key + "_") and f.lower().endswith(".png"):
                    path = os.path.join(self.dir, f)
                    break
        if not path:
            self.missing.append(key)
            return "<div class=\"placeholder\">איור %s חסר (%s) — יש להריץ make_figures.py</div>" % (cnt(self.n), esc(caption))
        with open(path, "rb") as fh:
            data = base64.b64encode(fh.read()).decode("ascii")
        self.embedded.append(os.path.basename(path))
        return ("<figure><img src=\"data:image/png;base64,%s\" alt=\"%s\"><figcaption>איור %s: %s</figcaption></figure>"
                % (data, esc(caption), cnt(self.n), esc(caption)))


# ----------------------------------------------------------------------------- report
def build_report(cfg, S, F, figs, level):
    """Return the full HTML string."""
    H = []
    add = H.append
    N = S["n_months"]
    MN = S["month_names"]
    TE, TI = S["total_expenses"], S["total_income"]
    full = level != "overview"
    deep = level == "deep"
    win = S["window"]
    hh = S.get("household") or cfg["household"].get("label", "")
    generated = S.get("generated") or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    G, C = S["expenses_by_group"], S["expenses_by_cat"]
    IC = S["income_by_cat"]
    rc = S["recon"]
    sec = [0]

    def h2(title):
        sec[0] += 1
        return "<section class=\"page\"><h2>%d. %s</h2>" % (sec[0], esc(title))

    def share(v, tot):
        return pct(v / tot) if tot else pct(0)

    add("<!DOCTYPE html>\n<html dir=\"rtl\" lang=\"he\"><head><meta charset=\"utf-8\">")
    add("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    add("<title>דו\"ח ניתוח תזרים — %s</title><style>%s</style></head><body><div dir=\"rtl\" lang=\"he\" class=\"report\">" % (esc(hh), CSS))

    # ---------------- cover
    add("<section class=\"page cover\">")
    add("<h1>דו\"ח ניתוח הוצאות והכנסות — %s</h1>" % esc(hh))
    add("<div class=\"meta\">חלון הניתוח: %s–%s (%s חודשים) · רמת ניתוח: %s · הופק: <span class=\"num\">%s</span></div>"
        % (dmy(win["start"]), dmy(win["end"]), cnt(N), esc(level), esc(generated)))
    add("<div class=\"box\"><p><b>היקף.</b> הדו\"ח מנתח את כל התנועות הכספיות של משק הבית בחלון, על בסיס %s שורות במאגר "
        "(מהן %s שורות הוצאה ו-%s שורות הכנסה נסכמות). כל המספרים חושבו מחדש מהמאגר (<code>work/summary.json</code>) "
        "וזהים לאלה שבקובץ האקסל. הדו\"ח כולל תקציר מנהלים, מתודולוגיה, ניתוח הכנסות והוצאות"
        "%s.</p></div>" % (cnt(S["n_rows_db"]), cnt(S["n_expense_rows"]), cnt(S["n_income_rows"]),
                            ", מאזן וחיסכון, ממצאים מיוחדים, המלצות ונספח" if full else " (רמת סקירה)"))
    add("</section>")

    # ---------------- 1 executive summary
    add(h2(SECTION_TITLES[0]))
    add("<div class=\"kpis\">"
        "<div class=\"kpi\"><span class=\"v\">%s</span><span class=\"l\">ממוצע הכנסות חודשי (סה\"כ %s)</span></div>"
        "<div class=\"kpi\"><span class=\"v\">%s</span><span class=\"l\">ממוצע הוצאות חודשי (סה\"כ %s)</span></div>"
        "<div class=\"kpi\"><span class=\"v\">%s</span><span class=\"l\">מאזן חודשי ממוצע (הכנסות − הוצאות; סה\"כ %s)</span></div>"
        "<div class=\"kpi\"><span class=\"v\">%s</span><span class=\"l\">שיעור חיסכון = (הכנסות − הוצאות) ÷ הכנסות</span></div>"
        "<div class=\"kpi\"><span class=\"v\">%s</span><span class=\"l\">חיסכון והשקעות בפועל בחלון</span></div>"
        "<div class=\"kpi\"><span class=\"v\">%s</span><span class=\"l\">שארית התאמת התזרים (0 = כל תנועת בנק מוסברת)</span></div>"
        "</div>" % (ils0(S["avg_income"]), ils0(TI), ils0(S["avg_expenses"]), ils0(TE), ils0(S["avg_balance"]),
                    ils0(S["balance_total"]), pct(S["savings_rate"]), ils0(S["savings_total"]), ils0(rc["residual"])))
    add("<h3>ממצאים עיקריים</h3>")
    add(bullets(F["findings"] if F else None, FINDINGS_MISSING))
    add("<h3>פעולות מומלצות (תמצית)</h3>")
    if F is None:
        add("<div class=\"placeholder\">%s</div>" % esc(FINDINGS_MISSING))
    elif not F["actions"]:
        add("<p class=\"small\">— אין פעולות —</p>")
    else:
        add("<ol>%s</ol>" % "".join("<li>%s</li>" % esc(a.get("what", "")) for a in F["actions"]))
    add("<h3>המספרים בקצרה</h3>")
    g_rows = [[k, ils0(v["total"]), ils0(v["avg"]), share(v["total"], TE)] for k, v in G.items()]
    add(table(["קבוצה", "סה\"כ", "ממוצע חודשי", "חלק"], g_rows, colcls=["", "n", "n", "n"],
              foot=["סה\"כ הוצאות", ils0(TE), ils0(S["avg_expenses"]), pct(1 if TE else 0)]))
    add("</section>")

    # ---------------- 2 methodology
    add(h2(SECTION_TITLES[1]))
    add("<h3>מקורות הנתונים</h3>")
    src_rows = []
    for e in cfg.get("accounts", []):
        src_rows.append(["עו\"ש", esc(e.get("label", e["id"])), esc(owner_label(cfg, e.get("owner"))), esc(e.get("bank", ""))])
    for e in cfg.get("cards", []):
        src_rows.append(["כרטיס אשראי", esc(e.get("label", e["id"])), esc(owner_label(cfg, e.get("owner"))), esc(e.get("issuer", ""))])
    for e in cfg.get("benefit_programs", []):
        src_rows.append(["מועדון הטבות", esc(e.get("label", e["id"])), esc(owner_label(cfg, e.get("owner"))), esc(e.get("issuer", ""))])
    for e in cfg.get("p2p", []):
        src_rows.append(["אפליקציית תשלומים", esc(e.get("label", e["id"])), esc(owner_label(cfg, e.get("owner"))), esc(e.get("app", ""))])
    add(table(["סוג", "מקור", "בעלים", "מוסד / פורמט"], src_rows))
    pay_rows = [[esc(k), cnt(v["n_rows"]), ils0(v["total"]), share(v["total"], TE)] for k, v in S["expenses_by_pay"].items()]
    add(table(["אמצעי תשלום", "שורות הוצאה", "סה\"כ הוצאות בחלון", "חלק"], pay_rows, colcls=["", "n", "n", "n"],
              foot=["סה\"כ", cnt(S["n_expense_rows"]), ils0(TE), pct(1 if TE else 0)]))
    add("<h3>הגדרת חודש וחלון</h3>")
    add("<p>חודש הניתוח של כל תנועה הוא <b>חודש החיוב בבנק</b> (לכרטיסי אשראי: תאריך חיוב הכרטיס בעו\"ש, לא תאריך העסקה). "
        "החלון: %s–%s כולל (%s חודשים); תנועות שמחוץ לחלון נשמרו במאגר אך אינן נסכמות. "
        "סטטיסטיקות מחושבות תמיד על %s ערכים חודשיים (חודש ללא חיוב = 0). לצד הממוצע הרגיל מוצג <b>ממוצע ללא אפס</b> — "
        "ממוצע על החודשים שבהם הייתה תנועה בלבד — לסעיפים שאינם חודשיים.</p>"
        % (dmy(win["start"]), dmy(win["end"]), cnt(N), cnt(N)))
    add("<h3>מה אינו נספר כהוצאה</h3><ul>")
    add("<li><b>תשלומי כרטיסי אשראי בעו\"ש</b> — פירוט הכרטיס נושא את ההוצאות; חיוב הבנק מסומן כלא-נסכם (אין ספירה כפולה).</li>")
    add("<li><b>חיסכון והשקעות</b> — הפקדות בסך %s מוצגות בנפרד.</li>" % ils0(S["savings_total"]))
    add("<li><b>העברות פנימיות</b> — כסף שנשאר בחשבונות המשפחה (המרות מט\"ח, משיכות מזומן, טעינות, העברות בין חשבונות): נטו %s בעו\"ש.</li>"
        % ils0(S["internal_bank"]["net"]))
    if S["p2p"].get("present"):
        add("<li><b>תשלומי אפליקציה הממומנים בכרטיס</b> — נספרים פעם אחת, בתאריך התשלום באפליקציה; חיוב הכרטיס המקביל מסומן ככפילות (%s זוגות).</li>"
            % cnt(S["p2p"]["matched_pairs"]))
    if S["reimbursables"]["present"]:
        add("<li><b>הוצאות בהחזר</b> — %s הוצאות שיוחזרו ו-%s החזרים שהתקבלו הוצאו מהממוצעים (סעיף 5).</li>"
            % (ils0(S["reimbursables"]["paid_total"]), ils0(S["reimbursables"]["received_total"])))
    if S["nonstatement_expenses"]["n"]:
        add("<li><b>הוצאות לפי הערכה / שורות מחושבות</b> — %s נרשמו ממקור שאינו דף חשבון.</li>" % ils0(S["nonstatement_expenses"]["total"]))
    add("</ul>")
    add("<h3>התאמת חיובי כרטיסים לבנק</h3>")
    cd_rows = []
    for c in S["card_debits"]["table"]:
        cd_rows.append([esc(c["label"])] + [ils0(x) for x in c["bank_debits_by_month"]] +
                       [ils0(c["bank_debits"]), ils0(c["statement_ils"]), ils0(c["fx_expenses"]),
                        (ils0(c["diff"]) if c["diff"] is not None else "—")])
    add(table(["כרטיס"] + MN + ["חיובי בנק", "פירוט (₪, בחלון)", "מזה מט\"ח", "פער"], cd_rows,
              colcls=[""] + ["n"] * (N + 4)))
    add("<p class=\"small\">חיובי מט\"ח נסלקים מחשבון המט\"ח בתאריכי העסקה ולכן אינם בחיוב הבנק השקלי; \"פער\" = חיובי בנק פחות סכום שורות הפירוט בש\"ח. "
        "חיובי כרטיס בבנק שלא שויכו לאף כרטיס: %s.</p>" % ils0(S["card_debits"]["unmatched_total"]))
    add("<h3>פערי נתונים</h3>")
    gaps = list(F["data_gaps"]) if F else None
    auto = []
    if S["n_unknown"]:
        auto.append("\"%s\" / לבדיקה: %s פריטים בסך %s (רשימה בנספח%s)."
                    % ("אחר / לא מזוהה", cnt(S["n_unknown"]), ils0(S["unknown_total"]), " ובסעיף 6" if full else ""))
    if S["p2p"].get("unmatched_card_rows", {}).get("n"):
        auto.append("שורות תשלום-אפליקציה בכרטיס ללא צילום/קובץ תואם: %s שורות, %s." %
                    (cnt(S["p2p"]["unmatched_card_rows"]["n"]), ils0(S["p2p"]["unmatched_card_rows"]["total"])))
    if not rc["residual_ok"]:
        auto.append("התאמת התזרים אינה נסגרת: שארית %s (ראו סעיף 5)." % ils0(rc["residual"]))
    if S["card_debits"]["unmatched_total"]:
        auto.append("חיובי כרטיס בבנק שלא שויכו לכרטיס מוגדר: %s." % ils0(S["card_debits"]["unmatched_total"]))
    if gaps is None:
        add("<div class=\"placeholder\">%s</div>" % esc(FINDINGS_MISSING))
        add(bullets(auto, "") if auto else "")
    else:
        add(bullets(gaps + auto, ""))
    add("</section>")

    # ---------------- 3 income
    add(h2(SECTION_TITLES[2]))
    add("<p>סך ההכנסות בחלון %s (%s שורות), ממוצע %s לחודש.</p>" % (ils0(TI), cnt(S["n_income_rows"]), ils0(S["avg_income"])))
    add("<h3>לפי מקור</h3>")
    inc_rows = [[esc(v["cat"]), ils0(v["total"]), ils0(v["avg"]), nz_cell(v), share(v["total"], TI)] for k, v in IC.items()]
    add(table(["מקור הכנסה", "סה\"כ", "ממוצע חודשי", "ממוצע ללא אפס (חודשים)", "חלק"], inc_rows,
              colcls=["", "n", "n", "n", "n"], foot=["סה\"כ", ils0(TI), ils0(S["avg_income"]), "", pct(1 if TI else 0)]))
    add("<h3>לפי חודש ומקור</h3>")
    icm_rows = [[esc(v["cat"])] + [ils0(x) for x in v["by_month"]] + [ils0(v["total"])] for k, v in IC.items()]
    add(table(["מקור"] + MN + ["סה\"כ"], icm_rows, colcls=[""] + ["n"] * (N + 1),
              foot=["סה\"כ"] + [ils0(x) for x in S["income_by_month"]] + [ils0(TI)]))
    if full:
        add("<h3>לפי אדם</h3>")
        IP = S["income_by_person"]
        ip_rows = [[esc(p)] + [ils0(x) for x in v["by_month"]] + [ils0(v["total"]), ils0(v["avg"]), share(v["total"], TI)] for p, v in IP.items()]
        add(table(["אדם"] + MN + ["סה\"כ", "ממוצע", "חלק"], ip_rows, colcls=[""] + ["n"] * (N + 3)))
        add("<p class=\"small\">\"אדם\" = בעל החשבון/הכרטיס שבו נרשמה ההכנסה (חשבון משותף = %s).</p>" % esc(owner_label(cfg, None)))
    add("<h3>תנודתיות</h3>")
    ist = S["income_stats"]
    add("<p>ההכנסה החודשית נעה בין %s (%s) ל-%s (%s); חציון %s, סטיית תקן %s.</p>"
        % (ils0(ist["min"]), esc(month_label_of(S, ist["min_month"])), ils0(ist["max"]), esc(month_label_of(S, ist["max_month"])),
           ils0(ist["median"]), ils0(ist["stdev"])))
    add(figs.fig("05"))
    if full:
        add(figs.fig("06"))
    add("</section>")

    # ---------------- 4 expenses
    add(h2(SECTION_TITLES[3]))
    add("<p>סך ההוצאות בחלון %s (%s שורות), ממוצע %s לחודש.</p>" % (ils0(TE), cnt(S["n_expense_rows"]), ils0(S["avg_expenses"])))
    add("<h3>4.1 לפי חודש</h3>")
    m_rows = [[esc(MN[i]), ils0(S["expenses_by_month"][i]), ils0(S["fixed_by_month"][i]), ils0(S["variable_by_month"][i]),
               ils0(S["israel_by_month"][i]), ils0(S["abroad_by_month"][i])] for i in range(N)]
    add(table(["חודש", "סה\"כ הוצאות", "קבועות", "משתנות", "בישראל", "בחו\"ל"], m_rows, colcls=["", "n", "n", "n", "n", "n"],
              foot=["סה\"כ", ils0(TE), ils0(S["fixed_total"]), ils0(S["variable_total"]), ils0(S["israel_total"]), ils0(S["abroad_total"])]))
    est = S["expense_stats"]
    add("<p>החודש הנמוך: %s (%s); הגבוה: %s (%s).</p>" % (esc(month_label_of(S, est["min_month"])), ils0(est["min"]),
                                                            esc(month_label_of(S, est["max_month"])), ils0(est["max"])))
    add(figs.fig("01"))
    add("<h3>4.2 לפי קבוצה</h3>")
    add(table(["קבוצה", "סה\"כ", "ממוצע חודשי", "חלק"], g_rows, colcls=["", "n", "n", "n"],
              foot=["סה\"כ", ils0(TE), ils0(S["avg_expenses"]), pct(1 if TE else 0)]))
    gm_rows = [[esc(k)] + [ils0(x) for x in v["by_month"]] + [ils0(v["total"])] for k, v in G.items()]
    add(table(["קבוצה"] + MN + ["סה\"כ"], gm_rows, colcls=[""] + ["n"] * (N + 1),
              foot=["סה\"כ"] + [ils0(x) for x in S["expenses_by_month"]] + [ils0(TE)]))
    add(figs.fig("02"))
    add(figs.fig("03"))
    if full:
        add(figs.fig("04"))
    add("<h3>4.3 טבלת קטגוריות מלאה</h3>")
    c_rows = [[esc(v["group"]), esc(v["cat"]), ils0(v["total"]), ils0(v["avg"]), nz_cell(v),
               esc(month_label_of(S, v["max_month"])), share(v["total"], TE)] for k, v in C.items()]
    add(table(["קבוצה", "קטגוריה", "סה\"כ", "ממוצע חודשי", "ממוצע ללא אפס (חודשים)", "חודש שיא", "חלק"], c_rows,
              colcls=["", "", "n", "n", "n", "", "n"], foot=["סה\"כ", "", ils0(TE), ils0(S["avg_expenses"]), "", "", pct(1 if TE else 0)]))
    add("<p class=\"small\">\"ממוצע ללא אפס\" מראה את גודל החיוב כשהוא מגיע (סעיפים לא-חודשיים); \"ממוצע חודשי\" הוא הסכום שיש להפריש בכל חודש.</p>")
    add("<h3>4.4 הוצאות קבועות מול משתנות</h3>")
    add("<p>הוצאות קבועות (הקטגוריות: %s) הסתכמו ב-%s (%s, %s לחודש); משתנות %s (%s, %s לחודש).</p>"
        % (esc(", ".join(S["fixed_categories"]) or "—"), ils0(S["fixed_total"]), share(S["fixed_total"], TE), ils0(S["fixed_total"] / N),
           ils0(S["variable_total"]), share(S["variable_total"], TE), ils0(S["variable_total"] / N)))
    if full:
        add(figs.fig("10"))
    add("<h3>4.5 ישראל מול חו\"ל%s</h3>" % (" ונסיעות" if S["trips"] else ""))
    add("<p>הוצאות בחו\"ל (עסקאות במטבע זר): %s (%s); בישראל (בש\"ח): %s.</p>"
        % (ils0(S["abroad_total"]), share(S["abroad_total"], TE), ils0(S["israel_total"])))
    if S["abroad_by_currency"]:
        add(table(["מטבע", "עסקאות", "סה\"כ ₪"], [[esc(k), cnt(v["n"]), ils0(v["total"])] for k, v in S["abroad_by_currency"].items()],
                  colcls=["", "n", "n"]))
    if full:
        add(figs.fig("11"))
    if S["trips"] and full:
        add("<h4>פילוח לפי נסיעה</h4>")
        tsum = sum(v["total"] for v in S["trips"].values())
        add(table(["נסיעה", "שורות", "סה\"כ", "חלק"], [[esc(t), cnt(v["n"]), ils0(v["total"]), share(v["total"], tsum)] for t, v in S["trips"].items()],
                  colcls=["", "n", "n", "n"], foot=["סה\"כ", "", ils0(tsum), pct(1 if tsum else 0)]))
        if deep:
            for t, v in S["trips"].items():
                add("<h4>%s — הפריטים הגדולים</h4>" % esc(t))
                add(table(["תאריך עסקה", "פריט", "סכום"], [[dmy(i["date"]), esc(i["name"]), ils2(i["amount"])] for i in v["items"]],
                          colcls=["n", "", "n"]))
    add("<h3>4.6 לפי אמצעי תשלום</h3>")
    sm_rows = [[esc(k)] + [ils0(x) for x in v["by_month"]] + [ils0(v["total"]), share(v["total"], TE)] for k, v in S["expenses_by_pay"].items()]
    add(table(["אמצעי תשלום"] + MN + ["סה\"כ", "חלק"], sm_rows, colcls=[""] + ["n"] * (N + 2)))
    add(figs.fig("07"))
    if full and S["expenses_by_card"]:
        ctot = sum(v["total"] for v in S["expenses_by_card"].values())
        cr = [[esc(k), esc(v["owner"]), ils0(v["total"]), ils0(v["avg"]), share(v["total"], ctot)] for k, v in S["expenses_by_card"].items()]
        add(table(["כרטיס", "בעלים", "סה\"כ", "ממוצע חודשי", "חלק מהכרטיסים"], cr, colcls=["", "", "n", "n", "n"],
                  foot=["סה\"כ בכרטיסים", "", ils0(ctot), ils0(ctot / N), pct(1 if ctot else 0)]))
        add(figs.fig("13"))
    if full:
        add("<h3>4.7 בתי העסק הגדולים</h3>")
        tm_rows = [[cnt(i + 1), esc(t["name"]), esc(t["cat"]), cnt(t["count"]), ils0(t["total"]), ils0(t["avg"]), nz_cell(t)]
                   for i, t in enumerate(S["top_merchants"])]
        add(table(["#", "בית עסק", "קטגוריה", "עסקאות", "סה\"כ", "ממוצע חודשי", "ממוצע ללא אפס (חודשים)"], tm_rows,
                  colcls=["n", "", "", "n", "n", "n", "n"]))
        if deep:
            add(figs.fig("08"))
            add(figs.fig("09"))
        if S.get("has_secondary_scheme"):
            add("<h3>4.8 הסכימה המוצעת</h3>")
            NG = S["expenses_new_by_group"]
            add(table(["קבוצה (מוצעת)", "סה\"כ", "ממוצע חודשי", "חלק"],
                      [[esc(k), ils0(v["total"]), ils0(v["avg"]), share(v["total"], TE)] for k, v in NG.items()],
                      colcls=["", "n", "n", "n"], foot=["סה\"כ", ils0(TE), ils0(S["avg_expenses"]), pct(1 if TE else 0)]))
        if deep:
            add(figs.fig("14"))
    add("</section>")

    if not full:
        add("</div></body></html>")
        return "\n".join(H)

    # ---------------- 5 balance, savings, cash flow
    add(h2(SECTION_TITLES[4]))
    add("<h3>5.1 מאזן חודשי</h3>")
    b_rows = [[esc(MN[i]), ils0(S["income_by_month"][i]), ils0(S["expenses_by_month"][i]), ils0(S["balance_by_month"][i]),
               share(S["balance_by_month"][i], S["income_by_month"][i]), ils0(S["savings_by_month"][i]),
               ils0(S["balance_by_month"][i] - S["savings_by_month"][i])] for i in range(N)]
    add(table(["חודש", "הכנסות", "הוצאות", "מאזן", "שיעור חיסכון", "הפקדות לחיסכון/השקעה", "מאזן אחרי הפקדות"], b_rows,
              colcls=["", "n", "n", "n", "n", "n", "n"],
              foot=["סה\"כ", ils0(TI), ils0(TE), ils0(S["balance_total"]), pct(S["savings_rate"]), ils0(S["savings_total"]),
                    ils0(S["balance_total"] - S["savings_total"])]))
    add(figs.fig("12"))
    add("<h3>5.2 חיסכון והשקעות בחלון</h3>")
    add(table(["אפיק", "סכום"], [[esc(k), ils0(v)] for k, v in S["savings_by_cat"].items()], colcls=["", "n"],
              foot=["סה\"כ", ils0(S["savings_total"])]))
    add("<h3>5.3 התאמת תזרים מזומנים (בנק)</h3>")
    add("<p>הבדיקה: האם השינוי ביתרות העו\"ש בחלון מוסבר במלואו על ידי ההכנסות, ההוצאות, החיסכון וההעברות שזוהו. "
        "כל רכיב שאינו \"הכנסה\" או \"הוצאה\" אך הזיז כסף בעו\"ש מופיע כתיקון. השארית מוצגת תמיד.</p>")
    rc_rows = [[esc(c["sign"]), esc(RECON_LABELS.get(c["key"], c["key"])), ils0(c["value"])]
               for c in rc["components"] if c["present"] or abs(c["value"]) >= 0.005]
    add(table(["סימן", "רכיב", "סכום"], rc_rows, colcls=["n", "", "n"],
              foot=["=", "סה\"כ מוסבר %s | שינוי יתרות בנק בחלון %s (%s) | שארית" %
                    (ils0(rc["explained"]), ils0(rc["bank_change"]), "מיתרות דפי החשבון" if rc["bank_change_source"] == "balances" else "מסכום תנועות הבנק"),
                    ils0(rc["residual"])]))
    if rc["residual_ok"]:
        add("<p><b>ההתאמה נסגרת לשארית של %s</b> (עד לעיגול): כל תנועת בנק בחלון מוסברת, ולא נספרה הוצאה או הכנסה פעמיים.</p>" % ils0(rc["residual"]))
    else:
        add("<div class=\"note warn\"><b class=\"bad\">ההתאמה אינה נסגרת: שארית %s.</b> פער כרטיסים מול בנק (F3): %s; חיובי כרטיס לא משויכים: %s. "
            "יש לבדוק את הרכיב שסטה לאחר סבב הסיווג האחרון.</div>"
            % (ils0(rc["residual"]), ils0(rc["card_vs_bank_diff"]), ils0(rc["card_debits_unmatched"])))
    if rc.get("per_account"):
        pa = rc["per_account"]
        add(table(["חשבון", "יתרה לפני החלון", "יתרה בסוף החלון", "שינוי"],
                  [[esc(v["label"]), ils0(v["balance_before"]), ils0(v["balance_end"]), ils0(v["change"])] for v in pa.values()],
                  colcls=["", "n", "n", "n"], foot=["סה\"כ", "", "", ils0(rc["bank_change_from_balances"])]))
        add("<p class=\"small\">שינוי היתרות לפי סכום תנועות הבנק במאגר: %s.</p>" % ils0(rc["bank_change_from_rows"]))
    if S["internal_bank"]["by_cat"]:
        add("<h4>העברות פנימיות בעו\"ש לפי קטגוריה (חיובי = יוצא)</h4>")
        add(table(["קטגוריה", "סכום"], [[esc(k), ils0(v)] for k, v in S["internal_bank"]["by_cat"].items()], colcls=["", "n"],
                  foot=["נטו", ils0(S["internal_bank"]["net"])]))
    if S["reimbursables"]["present"]:
        R = S["reimbursables"]
        add("<h3>5.4 הוצאות בהחזר</h3>")
        add("<p>קבוצה לא-נסכמת: הוצאות שאמורות לחזור וההחזרים שהתקבלו. הן הוצאו מממוצעי ההוצאות וההכנסות.</p>")
        rb = [[dmy(r["date"]), esc(r["name"]), esc(r["type"]), esc(r["cat"]), ils2(r["amount"] if r["type"] == "הוצאה בהחזר" else -r["amount"])]
              for r in R["items"]]
        add(table(["תאריך", "פריט", "סוג", "קטגוריה", "סכום (+ הוצאה / − החזר)"], rb, colcls=["n", "", "", "", "n"],
                  foot=["", "הוצאות בהחזר %s − החזרים %s = יתרה פתוחה" % (ils0(R["paid_total"]), ils0(R["received_total"])), "", "", ils0(R["open_balance"])]))
    add("</section>")

    # ---------------- 6 special findings
    add(h2(SECTION_TITLES[5]))
    k = 0
    for club in S.get("benefit_programs", []):
        k += 1
        add("<h3>6.%d מועדון הטבות — %s</h3>" % (k, esc(club["label"])))
        add(table(["מדד", "ערך"], [["רכישות והטבות בערך נקוב", ils0(club["purchases_face"])], ["טעינות (העברה פנימית)", ils0(club["loads"])],
                                    ["הנחה מחושבת", ils0(club["discount_total"])], ["שיעור הנחה", pct(club["discount_rate"])],
                                    ["חיובי בנק", ils0(club["bank_debits"])], ["שינוי יתרת הכרטיס הנטען", ils0(club["balance_change"])]],
                  colcls=["", "n"]))
        add(table(["חודש"] + MN, [["הנחה"] + [ils0(x) for x in club["discount_by_month"]], ["חיוב בנק"] + [ils0(x) for x in club["bank_debits_by_month"]]],
                  colcls=[""] + ["n"] * N))
    if S["p2p"].get("configured") or S["p2p"].get("present"):
        k += 1
        P = S["p2p"]
        add("<h3>6.%d אפליקציות תשלומים</h3>" % k)
        add(table(["מדד", "ערך"], [["תשלומים (הוצאות) באפליקציה", ils0(P["expenses"])], ["תקבולים (הכנסות) באפליקציה", ils0(P["income"])],
                                    ["זוגות תואמים (שורת כרטיס ↔ תשלום)", cnt(P["matched_pairs"])],
                                    ["מומן בחיוב כרטיס בחלון", ils0(P["funded_by_card_in_window"])],
                                    ["שולם מיתרת האפליקציה / מומן מחוץ לחלון", ils0(P["paid_from_balance"])],
                                    ["שורות כרטיס ללא צילום תואם", "%s (%s)" % (cnt(P["unmatched_card_rows"]["n"]), ils0(P["unmatched_card_rows"]["total"]))]],
                  colcls=["", "n"]))
        if P["unmatched_card_rows"].get("by_person"):
            add(table(["אדם", "שורות כרטיס ללא צילום — סכום"], [[esc(p), ils0(v)] for p, v in P["unmatched_card_rows"]["by_person"].items()], colcls=["", "n"]))
    k += 1
    FX = S["fx"]
    add("<h3>6.%d מט\"ח</h3>" % k)
    add("<p>חיובי מט\"ח בפירוטי הכרטיסים: %s (%s שורות; %s בהמרה משוערת לפי שער בנק ישראל).</p>"
        % (ils0(FX["expenses"]), cnt(FX["n_rows"]), cnt(FX["estimated_conversions"])))
    if FX["by_currency"]:
        add(table(["מטבע", "סה\"כ ₪"], [[esc(c), ils0(v)] for c, v in FX["by_currency"].items()], colcls=["", "n"]))
    if S["nonstatement_expenses"]["n"]:
        k += 1
        add("<h3>6.%d הוצאות במזומן ולפי הערכה</h3>" % k)
        add(table(["מקור", "סכום"], [[esc(s), ils0(v)] for s, v in S["nonstatement_expenses"]["by_source"].items()], colcls=["", "n"],
                  foot=["סה\"כ", ils0(S["nonstatement_expenses"]["total"])]))
    k += 1
    add("<h3>6.%d חיובים חוזרים (מנויים אפשריים)</h3>" % k)
    add("<p class=\"small\">היוריסטיקה: בית עסק עם חיוב בכל חודש בחלון ובשונות נמוכה (סטיית תקן מתחת ל-10% מהממוצע).</p>")
    add(table(["בית עסק", "קטגוריה", "ממוצע חודשי", "סה\"כ"], [[esc(m["name"]), esc(m["cat"]), ils0(m["avg"]), ils0(m["total"])] for m in S["recurring_merchants"]],
              colcls=["", "", "n", "n"]))
    k += 1
    add("<h3>6.%d זיכויים והחזרים</h3>" % k)
    add(table(["תאריך", "פריט", "קטגוריה", "סכום"], [[dmy(r["date"]), esc(r["name"]), esc(r["cat"]), ils2(r["amount"])] for r in S["refunds"]],
              colcls=["n", "", "", "n"]))
    k += 1
    add("<h3>6.%d חיובים כפולים אפשריים</h3>" % k)
    add(table(["בית עסק", "תאריך", "סכום", "מס' פעמים"], [[esc(d["name"]), dmy(d["date"]), ils2(d["amount"]), cnt(d["n"])] for d in S["possible_double_charges"]],
              colcls=["", "n", "n", "n"]))
    k += 1
    add("<h3>6.%d פריטים לא מזוהים / לבדיקה</h3>" % k)
    add("<p class=\"small\">הסיווג נעשה בלשונית \"%s\" באקסל (עמודות \"%s\" ו-\"%s\").</p>" % (esc(SHEET_MANUAL), esc(USER_TEXT_COL), esc(USER_CAT_COL)))
    add(table(["מזהה", "תאריך", "שם בקובץ", "מקור", "סכום", "קטגוריה", "הערה"],
              [[esc(u["id"]), dmy(u["date"]), esc(u["name"]), esc(u["source"]), ils2(u["amount"]), esc(u["cat"]), esc(u["note"])] for u in S["unknown_items"]],
              colcls=["", "n", "", "", "n", "", ""], foot=["", "", "", "סה\"כ לא מזוהה", ils0(S["unknown_total"]), "", ""]))
    add("</section>")

    # ---------------- 7 recommendations
    add(h2(SECTION_TITLES[6]))
    if F is None:
        add("<div class=\"placeholder\">%s</div>" % esc(FINDINGS_MISSING))
    elif not F["actions"]:
        add("<p class=\"small\">— אין המלצות בקובץ findings.json —</p>")
    else:
        for i, a in enumerate(F["actions"], 1):
            add("<div class=\"action\"><h4>%s. %s</h4><p><b>מה:</b> %s</p><p><b>למה:</b> %s</p><p><b>איך:</b> %s</p></div>"
                % (cnt(i), esc(a.get("what", "")), esc(a.get("what", "")), esc(a.get("why", "")), esc(a.get("how", ""))))
    add("</section>")

    # ---------------- 8 appendix
    add(h2(SECTION_TITLES[7]))
    add("<h3>8.1 הרכב המאגר</h3>")
    add(table(["סוג תנועה", "שורות (כל המאגר)", "שורות בחלון"],
              [[esc(t), cnt(n), cnt(S["rows_by_type_in_window"].get(t, 0))] for t, n in S["rows_by_type"].items()],
              colcls=["", "n", "n"], foot=["סה\"כ", cnt(S["n_rows_db"]), cnt(S["n_rows_in_window"])]))
    add("<h3>8.2 הערות מיפוי ושאלות ותשובות</h3>")
    add("<p class=\"small\">סכימת הקטגוריות וכללי המיפוי מתועדים ב-<code>notes/category_scheme.md</code> ו-<code>rules/merchant_rules.csv</code>; "
        "ההחלטות המתודולוגיות ב-<code>notes/decisions.md</code> והשאלות והתשובות ב-<code>notes/questions_and_answers.md</code>.</p>")
    add("<h3>8.3 קבצי התוצר ואיך משתמשים בהם</h3><ul>")
    for f, d in [
        (os.path.basename(cfg["outputs"]["workbook"]), "קובץ האקסל: <b>%s</b> / <b>%s</b> (ממוצע, ממוצע ללא אפס, מקס/מין, חציון, סטיית תקן), <b>%s</b> (טבלאות וגרפים), "
         "<b>%s</b> (בחירה אינטראקטיבית), <b>%s</b>, <b>%s</b>, <b>%s</b>, <b>%s</b>, <b>%s</b> (שורות למילוי המשתמש), <b>%s</b> (כל התנועות)."
         % tuple(esc(x) for x in (SHEET_EXP, SHEET_INC, SHEET_ANALYSIS, SHEET_PIVOT, SHEET_BY_CAT, SHEET_TXNS, SHEET_VARIABLE, SHEET_TRANSFERS, SHEET_MANUAL, SHEET_DB))),
        (os.path.basename(cfg["outputs"]["dashboard"]), "דשבורד לדפדפן הנבנה מלשונית database: מסננים, KPI, פירוט קבוצה → קטגוריה → בית עסק → עסקאות, וטבלת כל השורות."),
        (os.path.basename(cfg["outputs"]["report_pdf"]), "דו\"ח זה."),
        ("work/summary.json", "כל המספרים שבדו\"ח, כפי שחושבו מהמאגר."),
        ("work/findings.json", "הממצאים, ההמלצות ופערי הנתונים שנכתבו לדו\"ח זה."),
        ("work/database.csv", "מאגר התנועות המנורמל (%s שורות)." % cnt(S["n_rows_db"])),
    ]:
        add("<li><code>%s</code> — %s</li>" % (esc(f), d))
    add("</ul></section>")
    add("</div></body></html>")
    return "\n".join(H)


def make_report(cfg, summary_path=None, findings_path=None, fig_dir=None, out_path=None, level=None):
    summary_path = summary_path or project_path(cfg, "work.summary")
    if not os.path.isfile(summary_path):
        raise FileNotFoundError(summary_path)
    S = read_json(summary_path)
    F = load_findings(findings_path or project_path(cfg, "work.findings"))
    figs = Figures(fig_dir or project_path(cfg, "work.figures"))
    level = level or cfg.level
    html = build_report(cfg, S, F, figs, level)
    out_path = out_path or project_path(cfg, "outputs.report_html")
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return {"ok": True, "html": out_path, "bytes": len(html.encode("utf-8")), "level": level,
            "sections": OVERVIEW_SECTIONS if level == "overview" else len(SECTION_TITLES),
            "figures_embedded": figs.embedded, "figures_missing": figs.missing, "findings_present": F is not None,
            "residual": S["recon"]["residual"], "residual_ok": S["recon"]["residual_ok"]}


def main(argv=None):
    def extra(p):
        p.add_argument("--summary", default=None, help="override work/summary.json")
        p.add_argument("--findings", default=None, help="override work/findings.json")
        p.add_argument("--figures", default=None, help="override work/figures directory")
        p.add_argument("--out", default=None, help="override the report HTML path")
        p.add_argument("--strict", action="store_true", help="exit 4 when |residual| >= 1")
    args, cfg = parse_args("Hebrew RTL report HTML from work/summary.json + work/findings.json (P16)", extra, argv)
    try:
        res = make_report(cfg, args.summary, args.findings, args.figures, args.out)
    except FileNotFoundError as e:
        fail("summary not found: %s" % e, "run make_summary.py first")
    if not res["findings_present"]:
        log("WARNING: work/findings.json missing — placeholders rendered; write it after reading summary.json")
    if res["figures_missing"]:
        log("WARNING: figures missing: %s — run make_figures.py" % ", ".join(res["figures_missing"]))
    if args.strict and not res["residual_ok"]:
        res.update({"ok": False, "error": "reconciliation residual %.2f >= 1" % res["residual"],
                    "hint": "fix the drifted component (F12) before handing the report over"})
        emit(res)
        sys.exit(4)
    emit(res)


if __name__ == "__main__":
    main()
