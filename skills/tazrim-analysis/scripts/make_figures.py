#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_figures.py — report figures from work/summary.json (P15).

Purpose : render the PNG figures the report embeds. Every number comes from summary.json
          (nothing is recomputed from the database). Charts of averages are framed as
          "ממוצע חודשי" (monthly averages), never N-month sums.
Inputs  : tazrim.config.json (level, people order, figures dir); work/summary.json.
Outputs : work/figures/NN_*.png (dpi 200, ~9 in wide, white background); ONE JSON object on
          stdout: {"ok": true, "figures": [...paths], "n": k, "level": ..., "font": ..., "bidi": ...}.
Exit    : 0 ok; 1 summary missing / matplotlib missing; 2 config error.

Figures per level (a chart with no data is still written, as an "אין נתונים" placeholder,
so the count per level is deterministic):
  overview (5) : 01 monthly expenses vs income, 02 expenses by group (avg), 03 group share donut,
                 05 income by source (avg), 07 expenses by payment method (avg)
  standard (11): + 04 group x month stacked, 06 income by person x month, 10 fixed vs variable,
                 11 Israel vs abroad, 12 savings by month, 13 per-card pie
  deep (14)    : + 08 top-25 merchants (avg), 09 top-25 merchants by no-zero average (sorted
                 independently), 14 expenses by person x month

Hebrew text: a font is picked from a search list (macOS: Arial Hebrew / Arial Unicode; Linux:
Noto Sans Hebrew / DejaVu Sans; else matplotlib's default with a warning) and strings pass
through python-bidi when importable, else a run-reversal fallback (heb()).
Python 3.8 compatible. Needs matplotlib.
"""
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import SHARED, emit, fail, log, owner_label, parse_args, project_path, read_json  # noqa: E402

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from matplotlib.ticker import FuncFormatter
except ImportError:  # pragma: no cover
    plt = None

FIGURES = {
    "overview": ["01", "02", "03", "05", "07"],
    "standard": ["01", "02", "03", "04", "05", "06", "07", "10", "11", "12", "13"],
    "deep": ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14"],
}
NAMES = {
    "01": "01_monthly_exp_inc.png", "02": "02_expenses_by_group.png", "03": "03_group_share_pie.png",
    "04": "04_stacked_group_month.png", "05": "05_income_by_source.png", "06": "06_income_by_person_month.png",
    "07": "07_expenses_by_pay.png", "08": "08_top_merchants_avg.png", "09": "09_top_merchants_avg_nz.png",
    "10": "10_fixed_vs_variable.png", "11": "11_israel_abroad.png", "12": "12_savings_by_month.png",
    "13": "13_by_card_pie.png", "14": "14_expenses_by_person_month.png",
}
FONT_SEARCH = [
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/System/Library/Fonts/ArialHB.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansHebrew-Regular.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansHebrew-Regular.ttf",
    "/usr/share/fonts/noto/NotoSansHebrew-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
]
FONT_FAMILIES = ["Arial Hebrew", "Arial Unicode MS", "Noto Sans Hebrew", "DejaVu Sans"]
PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
GRAY, GRAY_LIGHT, INK, INK2 = "#898781", "#c3c2b7", "#0b0b0b", "#52514e"
OTHER = "אחר"
AVG_LABEL = "ממוצע חודשי"
_HEB = re.compile(u"[֐-׿]")


# ----------------------------------------------------------------------------- bidi
try:
    from bidi.algorithm import get_display as _bidi  # type: ignore
    BIDI = "python-bidi"

    def heb(s):
        return _bidi(str(s))
except Exception:  # noqa: BLE001
    BIDI = "fallback"
    _MIRROR = {"(": ")", ")": "(", "[": "]", "]": "[", "{": "}", "}": "{", "<": ">", ">": "<"}
    _LTR = re.compile(r"[A-Za-z0-9][A-Za-z0-9.,%/\-]*[A-Za-z0-9%]|[A-Za-z0-9]")

    def heb(s):
        """Visual order for matplotlib: reverse RTL/neutral runs, keep latin/digit runs."""
        s = str(s)
        if not _HEB.search(s):
            return s
        runs, pos = [], 0
        for m in _LTR.finditer(s):
            if m.start() > pos:
                runs.append(("R", s[pos:m.start()]))
            runs.append(("L", m.group()))
            pos = m.end()
        if pos < len(s):
            runs.append(("R", s[pos:]))
        merged = []
        for kind, txt in runs:
            if kind == "R" and merged and merged[-1][0] == "L" and not _HEB.search(txt):
                merged.append(("N", txt))
            else:
                merged.append((kind, txt))
        out_runs, i = [], 0
        while i < len(merged):
            kind, txt = merged[i]
            if kind == "N":
                # UAX#9 N1: neutrals between two numbers take the paragraph (RTL) direction;
                # only join runs when at least one side contains latin letters.
                nxt = merged[i + 1][1] if i + 1 < len(merged) else ""
                has_letters = re.search(r"[A-Za-z]", out_runs[-1][1] if out_runs else "") or re.search(r"[A-Za-z]", nxt)
                if i + 1 < len(merged) and merged[i + 1][0] == "L" and has_letters:
                    out_runs[-1] = ("L", out_runs[-1][1] + txt + merged[i + 1][1])
                    i += 2
                    continue
                kind = "R"
            out_runs.append((kind, txt))
            i += 1
        out = []
        for kind, txt in reversed(out_runs):
            out.append(txt if kind == "L" else "".join(_MIRROR.get(c, c) for c in reversed(txt)))
        return "".join(out)


# ----------------------------------------------------------------------------- font
def setup_font():
    """Register the first Hebrew-capable font found; returns its name (or 'default')."""
    for p in FONT_SEARCH:
        if os.path.exists(p):
            try:
                font_manager.fontManager.addfont(p)
                name = font_manager.FontProperties(fname=p).get_name()
                plt.rcParams["font.family"] = name
                return name
            except Exception as e:  # noqa: BLE001
                log("font %s unusable: %s" % (p, e))
    for fam in FONT_FAMILIES:
        try:
            path = font_manager.findfont(fam, fallback_to_default=False)
            if path:
                plt.rcParams["font.family"] = fam
                return fam
        except Exception:  # noqa: BLE001
            continue
    log("WARNING: no Hebrew font found in the search list; Hebrew labels may render as boxes")
    return "default"


def setup_style():
    plt.rcParams.update({
        "font.size": 10, "axes.titlesize": 13, "axes.titleweight": "bold", "axes.labelsize": 10,
        "axes.edgecolor": GRAY_LIGHT, "axes.linewidth": 0.8, "axes.grid": True, "grid.color": "#e1e0d9",
        "grid.linewidth": 0.8, "axes.axisbelow": True, "xtick.color": INK2, "ytick.color": INK2,
        "text.color": INK, "axes.labelcolor": INK2, "legend.frameon": False,
        "figure.facecolor": "white", "axes.facecolor": "white", "savefig.facecolor": "white",
    })


# ----------------------------------------------------------------------------- drawing
class Painter(object):
    def __init__(self, S, cfg, out_dir):
        self.S, self.cfg, self.out = S, cfg, out_dir
        self.M = [heb(m) for m in S["month_names"]]
        self.N = max(1, int(S.get("n_months") or len(S["months"])))
        self.X = list(range(len(S["months"])))
        groups = sorted(S.get("expenses_by_group", {}).items(), key=lambda kv: -kv[1]["total"])
        self.groups = [(g, v["total"]) for g, v in groups]
        self.gcolor = {}
        for i, (g, _) in enumerate(self.groups):
            self.gcolor[g] = PALETTE[i] if i < len(PALETTE) else GRAY_LIGHT
        self.gcolor[OTHER] = GRAY
        people = [owner_label(cfg, p["id"]) for p in cfg["household"]["people"]] + [owner_label(cfg, SHARED)]
        self.pcolor = {p: (PALETTE[i] if i < len(PALETTE) else GRAY) for i, p in enumerate(people)}
        self.people = people
        self.written = []

    # -- helpers
    @staticmethod
    def fmt_nis(v, _pos=None):
        return "{:,.0f}".format(v)

    @staticmethod
    def fmt_val(v):
        return heb("{:,.0f} ₪".format(v))

    def new_fig(self, h=5.0):
        fig, ax = plt.subplots(figsize=(9, h))
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        return fig, ax

    def save(self, fig, key):
        path = os.path.join(self.out, NAMES[key])
        fig.tight_layout()
        fig.savefig(path, dpi=200)
        plt.close(fig)
        self.written.append(path)
        log("wrote %s" % os.path.basename(path))

    def empty(self, key, title):
        fig, ax = self.new_fig(3.2)
        ax.set_axis_off()
        ax.text(0.5, 0.55, heb(title), ha="center", va="center", fontsize=13, fontweight="bold")
        ax.text(0.5, 0.35, heb("אין נתונים"), ha="center", va="center", fontsize=11, color=INK2)
        self.save(fig, key)

    def finish(self, fig, ax, key, ylabel="₪"):
        ax.yaxis.set_major_formatter(FuncFormatter(self.fmt_nis))
        if ylabel:
            ax.set_ylabel(heb(ylabel))
        ax.tick_params(length=0)
        self.save(fig, key)

    def hbar(self, items, key, title, colors=None, h=None):
        items = [(k, v) for k, v in items if v is not None]
        if not items or max(abs(v) for _, v in items) == 0:
            return self.empty(key, title)
        items = sorted(items, key=lambda kv: -kv[1])
        labels = [heb(k) for k, _ in items]
        vals = [v for _, v in items]
        n = len(items)
        fig, ax = self.new_fig(h or max(3.2, 0.42 * n + 1.4))
        ys = list(range(n))[::-1]
        ax.barh(ys, vals, height=0.6, color=colors or [PALETTE[0]] * n, edgecolor="white", linewidth=1)
        ax.set_yticks(ys)
        ax.set_yticklabels(labels)
        vmax = max(vals) if max(vals) > 0 else 1.0
        for y, v in zip(ys, vals):
            ax.text(v + vmax * 0.012 if v >= 0 else 0, y, self.fmt_val(v), va="center", ha="left", fontsize=9, color=INK2)
        ax.set_xlim(min(0, min(vals) * 1.1), vmax * 1.18)
        ax.xaxis.set_major_formatter(FuncFormatter(self.fmt_nis))
        ax.grid(axis="y", visible=False)
        ax.set_title(heb(title), loc="right")
        ax.set_xlabel(heb("₪"))
        ax.tick_params(length=0)
        self.save(fig, key)

    def stacked(self, series, key, title, colors, h=5.2, legend_cols=None):
        if not series or all(abs(v) < 0.005 for _, vals in series for v in vals):
            return self.empty(key, title)
        fig, ax = self.new_fig(h)
        n = len(self.X)
        pos, neg = [0.0] * n, [0.0] * n
        for label, vals in series:
            bottoms, heights = [], []
            for i, v in enumerate(vals):
                if v >= 0:
                    bottoms.append(pos[i])
                    pos[i] += v
                else:
                    bottoms.append(neg[i])
                    neg[i] += v
                heights.append(v)
            ax.bar(self.X, heights, 0.62, bottom=bottoms, color=colors.get(label, GRAY), label=heb(label),
                   edgecolor="white", linewidth=1)
        tot = [sum(vals[i] for _, vals in series) for i in range(n)]
        top = max(pos) if max(pos) > 0 else 1.0
        for i, t in enumerate(tot):
            ax.text(i, pos[i] + top * 0.015, self.fmt_val(t), ha="center", va="bottom", fontsize=9, color=INK2)
        ax.set_ylim(min(0, min(neg)) * 1.15, top * 1.12)
        if min(neg) < 0:
            ax.axhline(0, color=GRAY_LIGHT, linewidth=0.8)
        ax.set_xticks(self.X)
        ax.set_xticklabels(self.M)
        ax.grid(axis="x", visible=False)
        ax.set_title(heb(title), loc="right")
        ax.legend(ncol=legend_cols or min(len(series), 5), loc="upper center", bbox_to_anchor=(0.5, -0.09), fontsize=9)
        self.finish(fig, ax, key)

    def donut(self, items, key, title, colors, center_label, center_value, fold_below=0.02):
        items = [(k, v) for k, v in items if v and v > 0]
        tot = sum(v for _, v in items)
        if not items or tot <= 0:
            return self.empty(key, title)
        big = [(k, v) for k, v in items if v / tot >= fold_below]
        small = [(k, v) for k, v in items if v / tot < fold_below]
        if small:
            big.append((OTHER, sum(v for _, v in small)))
        fig, ax = plt.subplots(figsize=(9, 5.8))
        vals = [v for _, v in big]
        labels = [heb("{} – {:,.0f} ₪ ({:.1f}%)".format(k, v, 100 * v / tot)) for k, v in big]
        cols = [colors.get(k, GRAY) if isinstance(colors, dict) else colors[i % len(colors)] for i, (k, _) in enumerate(big)]
        wedges, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                           wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
        ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=9.5)
        for w, (k, v) in zip(wedges, big):
            frac = v / tot
            if frac >= 0.08:
                ang = math.radians((w.theta1 + w.theta2) / 2)
                ax.text(0.79 * math.cos(ang), 0.79 * math.sin(ang), "{:.0f}%".format(100 * frac),
                        ha="center", va="center", fontsize=9.5, color="white", fontweight="bold")
        ax.text(0, 0.04, heb(center_label), ha="center", va="center", fontsize=10, color=INK2)
        ax.text(0, -0.08, self.fmt_val(center_value), ha="center", va="center", fontsize=13, color=INK, fontweight="bold")
        ax.set_title(heb(title), loc="right")
        ax.set_aspect("equal")
        self.save(fig, key)

    def bars_by_month(self, vals, key, title, color):
        if not vals or all(abs(v) < 0.005 for v in vals):
            return self.empty(key, title)
        fig, ax = self.new_fig(4.6)
        ax.bar(self.X, vals, 0.55, color=color, edgecolor="white")
        top = max(vals) if max(vals) > 0 else 1.0
        for i, v in enumerate(vals):
            ax.text(i, v + top * 0.015 if v >= 0 else v, self.fmt_val(v), ha="center", va="bottom", fontsize=9, color=INK2)
        ax.set_xticks(self.X)
        ax.set_xticklabels(self.M)
        ax.grid(axis="x", visible=False)
        ax.set_ylim(min(0, min(vals) * 1.15), top * 1.15)
        ax.set_title(heb(title), loc="right")
        self.finish(fig, ax, key)

    # -- the figures
    def fig01(self):
        S = self.S
        exp, inc, bal = S["expenses_by_month"], S["income_by_month"], S["balance_by_month"]
        if not any(exp) and not any(inc):
            return self.empty("01", "הוצאות מול הכנסות לפי חודש")
        fig, ax = self.new_fig(5.2)
        w = 0.36
        ax.bar([x - w / 2 for x in self.X], exp, w, color=PALETTE[1], label=heb("הוצאות"), edgecolor="white")
        ax.bar([x + w / 2 for x in self.X], inc, w, color=PALETTE[0], label=heb("הכנסות"), edgecolor="white")
        ax.plot(self.X, bal, color=PALETTE[2], linewidth=2, marker="o", markersize=7, markeredgecolor="white",
                markeredgewidth=1.5, label=heb("מאזן (הכנסות − הוצאות)"))
        for i, b in enumerate(bal):
            ax.annotate(self.fmt_val(b), (i, b), textcoords="offset points", xytext=(0, 10), ha="center", fontsize=8.5,
                        bbox=dict(boxstyle="round,pad=0.25", facecolor="white", edgecolor=GRAY_LIGHT, linewidth=0.6))
        ax.set_xticks(self.X)
        ax.set_xticklabels(self.M)
        ax.grid(axis="x", visible=False)
        top = max(max(inc), max(exp), 1.0)
        ax.set_ylim(min(0, min(bal) * 1.2), top * 1.15)
        ax.set_title(heb("הוצאות מול הכנסות לפי חודש"), loc="right")
        ax.legend(ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.09), fontsize=9)
        self.finish(fig, ax, "01")

    def fig02(self):
        self.hbar([(g, v / self.N) for g, v in self.groups], "02", "הוצאות לפי קבוצה – " + AVG_LABEL,
                  colors=[self.gcolor[g] for g, _ in sorted(self.groups, key=lambda kv: -kv[1])])

    def fig03(self):
        tot = sum(v for _, v in self.groups)
        self.donut(self.groups, "03", "התפלגות ההוצאות לפי קבוצה (קבוצות מתחת ל-2% אוחדו ל\"אחר\")",
                   self.gcolor, AVG_LABEL, tot / self.N)

    def fig04(self):
        gm = {g: v["by_month"] for g, v in self.S.get("expenses_by_group", {}).items()}
        order = [g for g, _ in self.groups]
        keep, fold = order[:len(PALETTE)], order[len(PALETTE):]
        series = [(g, gm[g]) for g in keep]
        if fold:
            series.append((OTHER, [sum(gm[g][i] for g in fold) for i in range(len(self.X))]))
        title = "הוצאות לפי קבוצה וחודש" + (" (\"אחר\" = " + ", ".join(fold) + ")" if fold else "")
        self.stacked(series, "04", title, self.gcolor, h=5.8, legend_cols=5)

    def fig05(self):
        inc = self.S.get("income_by_cat", {})
        self.hbar([(v.get("cat", k), v["total"] / self.N) for k, v in inc.items()], "05", "הכנסות לפי מקור – " + AVG_LABEL)

    def fig06(self):
        pm = self.S.get("income_by_person", {})
        self.stacked([(p, v["by_month"]) for p, v in pm.items()], "06", "הכנסות לפי אדם וחודש", self.pcolor)

    def fig07(self):
        pay = self.S.get("expenses_by_pay", {})
        self.hbar([(k, v["total"] / self.N) for k, v in pay.items()], "07", "הוצאות לפי אמצעי תשלום – " + AVG_LABEL)

    def fig08(self):
        tm = self.S.get("top_merchants", [])[:25]
        self.hbar([(m["name"], m["avg"]) for m in tm], "08", "בתי העסק הגדולים – " + AVG_LABEL, h=9.0)

    def fig09(self):
        tm = self.S.get("top_merchants_by_avg_nz", [])[:25]
        self.hbar([(m["name"], m["avg_nz"]) for m in tm], "09", "בתי העסק הגדולים – ממוצע חודשי ללא אפס", h=9.0)

    def fig10(self):
        self.stacked([("הוצאות קבועות", self.S["fixed_by_month"]), ("הוצאות משתנות", self.S["variable_by_month"])],
                     "10", "הוצאות קבועות מול משתנות לפי חודש", {"הוצאות קבועות": PALETTE[0], "הוצאות משתנות": PALETTE[1]})

    def fig11(self):
        self.stacked([("בישראל", self.S["israel_by_month"]), ("בחו\"ל", self.S["abroad_by_month"])],
                     "11", "הוצאות בישראל מול חו\"ל לפי חודש", {"בישראל": PALETTE[0], "בחו\"ל": PALETTE[1]})

    def fig12(self):
        self.bars_by_month(self.S.get("savings_by_month", []), "12", "חיסכון והשקעות לפי חודש", PALETTE[0])

    def fig13(self):
        cards = self.S.get("expenses_by_card", {})
        items = [(k, v["total"] / self.N) for k, v in cards.items()]
        tot = sum(v for _, v in items)
        self.donut(items, "13", "פילוח הוצאות לפי כרטיס – " + AVG_LABEL, PALETTE, "ממוצע חודשי בכרטיסים", tot, fold_below=0.0)

    def fig14(self):
        pm = self.S.get("expenses_by_person", {})
        self.stacked([(p, v["by_month"]) for p, v in pm.items()], "14", "הוצאות לפי אדם וחודש", self.pcolor)

    def run(self, keys):
        for k in keys:
            getattr(self, "fig" + k)()
        return self.written


def make_figures(cfg, summary_path=None, out_dir=None, level=None):
    if plt is None:
        raise RuntimeError("matplotlib is required (pip install -r requirements.txt)")
    summary_path = summary_path or project_path(cfg, "work.summary")
    if not os.path.isfile(summary_path):
        raise FileNotFoundError(summary_path)
    S = read_json(summary_path)
    out_dir = out_dir or project_path(cfg, "work.figures")
    os.makedirs(out_dir, exist_ok=True)
    font = setup_font()
    setup_style()
    level = level or cfg.level
    keys = FIGURES[level]
    written = Painter(S, cfg, out_dir).run(keys)
    return written, font


def main(argv=None):
    def extra(p):
        p.add_argument("--summary", default=None, help="override work/summary.json path")
        p.add_argument("--out", default=None, help="override work/figures directory")
    args, cfg = parse_args("render the report figures from work/summary.json (P15)", extra, argv)
    try:
        written, font = make_figures(cfg, args.summary, args.out)
    except FileNotFoundError as e:
        fail("summary not found: %s" % e, "run make_summary.py first")
    except RuntimeError as e:
        fail(str(e), "pip install -r requirements.txt")
    emit({"ok": True, "figures": written, "n": len(written), "level": cfg.level, "font": font, "bidi": BIDI})


if __name__ == "__main__":
    main()
