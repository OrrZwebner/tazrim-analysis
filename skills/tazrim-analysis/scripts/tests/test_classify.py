# -*- coding: utf-8 -*-
"""Tests for classify.py (P10) on the synthetic fixture: types, in_window/summed flags for a
2-month window, unknown handling per level, P2P pairing, discount row, template
canonicalisation, user labels, estimates/trips, the shipped scheme + starter rules, CLI contract.
Run: python3 -m unittest discover -s scripts/tests -v
"""
import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
SKILL = os.path.dirname(SCRIPTS)
sys.path.insert(0, SCRIPTS)

import common  # noqa: E402
import make_fixtures  # noqa: E402

try:
    import openpyxl  # noqa: F401
    import docx  # noqa: F401
    HAVE_DEPS = True
except ImportError:  # pragma: no cover
    HAVE_DEPS = False

T = common


def quiet(fn, *a, **kw):
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return fn(*a, **kw)


def build_project(root=None, level=None, edit=None):
    """Generate the fixture, optionally edit the config dict, run the four parsers, return
    (root, cfg)."""
    import parse_bank, parse_benefits, parse_card_cycle, parse_card_blocks  # noqa: E401
    root = root or tempfile.mkdtemp(prefix="tz_cls_")
    make_fixtures.generate(root)
    cfg_path = os.path.join(root, "tazrim.config.json")
    if edit:
        with open(cfg_path, encoding="utf-8") as fh:
            raw = json.load(fh)
        edit(raw)
        with open(cfg_path, "w", encoding="utf-8") as fh:
            json.dump(raw, fh, ensure_ascii=False)
    for mod, extra in ((parse_bank, []), (parse_card_cycle, ["--offline"]), (parse_card_blocks, []), (parse_benefits, [])):
        try:
            quiet(mod.main, ["--config", cfg_path] + extra)
        except SystemExit as e:
            assert (e.code or 0) == 0, mod.__name__
    cfg = common.load_config(cfg_path, level=level)
    return root, cfg


def db_rows(cfg):
    return common.read_csv(common.project_path(cfg, "work.database"))


def total(rows, type_=None, summed_only=True):
    return round(sum(float(r["amount"]) for r in rows
                     if (type_ is None or r["type"] == type_) and (not summed_only or r["summed"] == T.YES)), 2)


@unittest.skipUnless(HAVE_DEPS, "openpyxl / python-docx not installed")
def classify_fixture_template():
    """(group, cat) rows of the synthetic template, from make_fixtures (no literals here)."""
    import make_fixtures
    return list(make_fixtures.TEMPLATE_ROWS)


class TestClassifyStandard(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import classify
        cls.root, cls.cfg = build_project()
        with open(os.path.join(cls.root, "expected.json"), encoding="utf-8") as fh:
            cls.exp = json.load(fh)
        cls.summary = quiet(classify.run, cls.cfg)
        cls.rows = db_rows(cls.cfg)
        cls.by_id = {r["id"]: r for r in cls.rows}

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.root, ignore_errors=True)

    def find(self, name, **kw):
        hits = [r for r in self.rows if r["original_name"] == name and all(r[k] == v for k, v in kw.items())]
        self.assertTrue(hits, name)
        return hits[0]

    def test_columns_and_sort(self):
        self.assertEqual(list(self.rows[0].keys()), common.DB_COLUMNS)
        dates = [r["charge_date"] for r in self.rows]
        self.assertEqual(dates, sorted(dates))
        for r in self.rows:
            self.assertIn(r["type"], common.ROW_TYPES)
            self.assertEqual(r["month"], r["charge_date"][:7])
            self.assertEqual(r["month_name"], common.month_label(r["month"]))
            self.assertRegex(r["pay"], r"\(.+\)$")            # owner always suffixed
            self.assertTrue(r["group_tz"] and r["cat_tz"], r["id"])

    def test_types_and_window_totals(self):
        e = self.exp["classify"]
        self.assertAlmostEqual(total(self.rows, T.TYPE_INCOME), e["income_in_window"])
        self.assertAlmostEqual(total(self.rows, T.TYPE_EXPENSE), e["expense_in_window_if_all_unknowns_are_expenses"])
        self.assertEqual(sum(1 for r in self.rows if r["type"] == T.TYPE_CARD_DEBIT), e["card_debit_rows_non_summed"])
        self.assertEqual(sum(1 for r in self.rows if r["type"] == T.TYPE_DUPLICATE), e["duplicate_rows"])
        self.assertEqual(sum(1 for r in self.rows if r["type"] == T.TYPE_INTERNAL), e["internal_rows"])
        self.assertTrue(all(r["summed"] == T.NO for r in self.rows if r["type"] not in common.SUMMED_TYPES))
        self.assertTrue(all(r["summed"] == T.YES for r in self.rows if r["type"] in common.SUMMED_TYPES and r["in_window"] == T.YES))
        s = self.summary
        self.assertAlmostEqual(s["window_totals"]["income"], e["income_in_window"])
        self.assertAlmostEqual(s["window_totals"]["expense"], e["expense_in_window_if_all_unknowns_are_expenses"])
        self.assertEqual(s["months"], self.exp["window"]["months"])

    def test_amount_sign_rules(self):
        salary = self.find("משכורת לדוגמה")
        self.assertEqual(salary["type"], T.TYPE_INCOME)
        self.assertAlmostEqual(float(salary["amount"]), 8000.0)              # positive for income
        rent = self.find("שכר דירה לדוגמה")
        self.assertEqual((rent["type"], rent["group_tz"], rent["cat_tz"]), (T.TYPE_EXPENSE, "דיור", "שכר דירה"))
        self.assertAlmostEqual(float(rent["amount"]), 2000.0)
        disc = self.find("הנחת מועדון (זיכוי מחושב)")
        self.assertAlmostEqual(float(disc["amount"]), self.exp["club"]["discount_row_amount"])   # refund stays negative
        self.assertEqual(disc["type"], T.TYPE_EXPENSE)
        self.assertEqual(disc["summed"], T.YES)
        self.assertTrue(disc["linked_id"].startswith("bank_b-"))
        self.assertEqual(self.by_id[disc["linked_id"]]["type"], T.TYPE_CARD_DEBIT)

    def test_card_debits_generated_from_config(self):
        for name in ("חיוב לכרטיס ויזה 1234", "חיוב לכרטיס ויזה 5678", "9012 - כרטיס אשראי", "חיוב מועדון הטבות"):
            r = self.find(name)
            self.assertEqual(r["type"], T.TYPE_CARD_DEBIT)
            self.assertEqual(r["summed"], T.NO)
            self.assertEqual(r["cat_tz"], "תשלום כרטיס אשראי")

    def test_p2p_pair_and_incoming(self):
        e = self.exp["p2p"]
        self.assertEqual(self.summary["p2p_pairs"], e["pairs"])
        self.assertEqual(self.summary["p2p_unmatched_card_rows"], 0)
        card = self.find("BIT העברה לדוגמה")
        app = self.find("חבר לדוגמה")
        self.assertEqual(card["type"], T.TYPE_DUPLICATE)
        self.assertEqual(card["summed"], T.NO)
        self.assertEqual(card["linked_id"], app["id"])
        self.assertEqual(app["linked_id"], card["id"])
        self.assertEqual(app["type"], T.TYPE_EXPENSE)
        self.assertAlmostEqual(float(app["amount"]), e["outgoing"][0]["amount"])
        self.assertEqual(app["charge_date"], e["outgoing"][0]["date"])
        inc = self.find("חברה לדוגמה")
        self.assertEqual(inc["type"], T.TYPE_INCOME)
        self.assertAlmostEqual(float(inc["amount"]), 60.0)

    def test_unknowns_flagged_for_manual_sheet(self):
        unknown = [r for r in self.rows if r["cat_tz"] == common.UNKNOWN_CAT and r["type"] in common.SUMMED_TYPES]
        names = {r["original_name"] for r in unknown}
        for n in ("חנות בגדים לדוגמה", "הוראת קבע לדוגמה", "העברה מהחשבון", "עמלת ניהול חשבון"):
            self.assertIn(n, names)
        for r in unknown:
            self.assertIn("נדרש:", r["rule_note"])
            self.assertIn(common.NOTE_NO_RULE, r["rule_note"])
        fx = self.find("EXAMPLE STORE US")
        self.assertEqual((fx["group_tz"], fx["cat_tz"]), ('חופשות וחו"ל', 'הוצאות בחו"ל'))
        self.assertEqual(fx["orig_currency"], "USD")
        self.assertAlmostEqual(float(fx["amount"]), 350.0)
        self.assertEqual(len(unknown), 6)          # 4 bank/card rows + the P2P app row + the club benefit
        self.assertEqual(self.summary["unknown_in_window"], len(unknown))
        load = self.find("טעינת כסף - כרטיס נטען")
        self.assertEqual((load["type"], load["summed"]), (T.TYPE_INTERNAL, T.NO))

    def test_template_canonicalisation(self):
        """Rule/template categories are matched after stripping spaces/quotes; the group comes
        from the template."""
        import classify
        primary, secondary, src = classify.load_scheme(self.cfg)
        self.assertTrue(src.startswith("template:"))
        # template categories first, then the default scheme CSV (the workbook lists both) and every
        # generated category; UNKNOWN_CAT keeps the expense group
        from common import GENERATED_CATEGORIES, UNKNOWN_CAT
        n_tpl = self.exp["template"]["categories"]
        self.assertEqual(list(primary)[:n_tpl], [c for _, c in classify_fixture_template()])
        self.assertGreater(len(primary), n_tpl + len(GENERATED_CATEGORIES))
        for c, g in GENERATED_CATEGORIES.items():
            self.assertIn(c, primary)
        self.assertEqual(primary[UNKNOWN_CAT]["group"], "שונות")
        self.assertEqual(primary["ביגוד והנעלה"]["group"], "ביגוד וטיפוח")     # from the default CSV
        canon = classify.Canon(primary)
        self.assertEqual(canon.cat("סופר מרקט"), "סופרמרקט")
        self.assertEqual(canon.cat('"שכר דירה"'), "שכר דירה")
        self.assertEqual(canon.cat("שכר  דירה "), "שכר דירה")
        self.assertEqual(canon.group_of("שכר דירה"), "דיור")
        self.assertIsNone(canon.group_of("לא קיים"))
        self.assertEqual(canon.cat("לא קיים"), "לא קיים")
        # "משכורת" is absent from the fixture template but present in the default CSV -> not reported
        self.assertNotIn("משכורת", self.summary["categories_not_in_scheme"])
        self.assertEqual(self.summary["categories_not_in_scheme"], {})

    def test_window_flags_two_month_window(self):
        """Window Feb only: January rows keep their data but in_window/summed become לא."""
        import classify

        def edit(raw):
            raw["window"]["start"] = "2025-02-01"
        root, cfg = build_project(edit=edit)
        try:
            s = quiet(classify.run, cfg)
            rows = db_rows(cfg)
            self.assertEqual(s["months"], ["2025-02"])
            jan = [r for r in rows if r["month"] == "2025-01"]
            self.assertTrue(jan)
            self.assertTrue(all(r["in_window"] == T.NO and r["summed"] == T.NO for r in jan))
            feb = [r for r in rows if r["month"] == "2025-02" and r["type"] in common.SUMMED_TYPES]
            self.assertTrue(all(r["in_window"] == T.YES and r["summed"] == T.YES for r in feb))
            self.assertAlmostEqual(s["window_totals"]["income"], 60.0)
            # Feb expenses: 30 + 70 - 50 + 100 + 200 + 55 = 405
            self.assertAlmostEqual(s["window_totals"]["expense"], 405.0)
            self.assertEqual(len(rows), len(self.rows))
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_ids_stable_across_reruns(self):
        import classify
        quiet(classify.run, self.cfg)
        again = {r["id"] for r in db_rows(self.cfg)}
        self.assertEqual(again, set(self.by_id))


@unittest.skipUnless(HAVE_DEPS, "openpyxl / python-docx not installed")
class TestClassifyLevelsAndOverrides(unittest.TestCase):
    def test_overview_level(self):
        """overview: unknowns lumped silently (no נדרש marker), no P2P rows, no pairing."""
        import classify
        root, cfg = build_project(level="overview")
        try:
            s = quiet(classify.run, cfg)
            rows = db_rows(cfg)
            self.assertEqual(s["level"], "overview")
            self.assertFalse(any(r["source"].startswith("BIT") for r in rows))
            self.assertEqual(s["p2p_pairs"], 0)
            card = next(r for r in rows if r["original_name"] == "BIT העברה לדוגמה")
            self.assertEqual(card["type"], T.TYPE_EXPENSE)          # the card row IS the expense here
            self.assertEqual(card["summed"], T.YES)
            for r in rows:
                self.assertNotIn("נדרש:", r["rule_note"])
            unknown = [r for r in rows if r["cat_tz"] == common.UNKNOWN_CAT]
            self.assertTrue(unknown)
            self.assertAlmostEqual(s["window_totals"]["income"], 8000.0)
            self.assertAlmostEqual(s["window_totals"]["expense"], 4535.0)   # same cash total, counted on the card row
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_user_labels_then_interpreted(self):
        import classify
        root, cfg = build_project()
        try:
            rows = db_rows(cfg) if os.path.isfile(common.project_path(cfg, "work.database")) else None
            quiet(classify.run, cfg)
            rows = db_rows(cfg)
            shirt = next(r for r in rows if r["original_name"] == "חנות בגדים לדוגמה")
            transfer = next(r for r in rows if r["original_name"] == "העברה מהחשבון")
            with open(os.path.join(root, "rules", "user_labels.csv"), "w", encoding="utf-8") as fh:
                fh.write("id,user_text,user_cat\n%s,בגדים לילדים,סופר מרקט\n" % shirt["id"])
            with open(os.path.join(root, "rules", "user_labels_interpreted.csv"), "w", encoding="utf-8") as fh:
                fh.write("id,type,group_tz,cat_tz,group_new,cat_new,trip,name_clean,note\n"
                         "%s,חיסכון והשקעות,חיסכון והשקעות,קופת גמל,חיסכון והשקעות,קופת גמל,,הפקדה לקופת גמל,העברה לחיסכון\n" % transfer["id"])
            s = quiet(classify.run, cfg)
            rows = {r["id"]: r for r in db_rows(cfg)}
            sh = rows[shirt["id"]]
            self.assertEqual((sh["group_tz"], sh["cat_tz"]), ("מזון", "סופרמרקט"))   # canonicalised to the template
            self.assertEqual(sh["name_clean"], "חנות בגדים לדוגמה – בגדים לילדים")
            self.assertIn("סיווג ידני של המשתמש", sh["rule_note"])
            self.assertNotIn("נדרש:", sh["rule_note"])
            tr = rows[transfer["id"]]
            self.assertEqual((tr["type"], tr["cat_tz"], tr["summed"]), (T.TYPE_SAVINGS, "קופת גמל", T.NO))
            self.assertIn("סיווג לפי המשתמש: העברה לחיסכון", tr["rule_note"])
            self.assertAlmostEqual(s["window_totals"]["expense"], 4535.0 - 1000.0)
            self.assertEqual(s["user_labels_applied"]["user_labels"], 1)
            self.assertEqual(s["user_labels_applied"]["interpreted"], 1)
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_estimates_and_trips(self):
        import classify

        def edit(raw):
            raw["estimates"] = [{"id": "clean", "name": "ניקיון", "amount": 250, "every_n_months": 1, "day": 15,
                                 "group": "בית", "cat": "ניקיון", "note": "הערכה", "funded_by_cat": "עמלה"}]
            raw["trips"] = [{"label": "נסיעה לדוגמה", "countries": ["ארצות הברית"], "start": "2025-01-01", "end": "2025-01-31"}]
        root, cfg = build_project(edit=edit)
        try:
            s = quiet(classify.run, cfg)
            rows = db_rows(cfg)
            est = [r for r in rows if r["source"] == "ידני – הערכה"]
            self.assertEqual(len(est), 2)                       # one per window month
            self.assertEqual({r["month"] for r in est}, {"2025-01", "2025-02"})
            self.assertTrue(all(r["type"] == T.TYPE_EXPENSE and r["summed"] == T.YES for r in est))
            self.assertEqual(s["estimate_rows"], 2)
            self.assertAlmostEqual(s["window_totals"]["expense"], 4535.0 + 500.0)
            fx = next(r for r in rows if r["original_name"] == "EXAMPLE STORE US")
            self.assertEqual(fx["trip"], "נסיעה לדוגמה")
            self.assertEqual(s["trips"], {"נסיעה לדוגמה": 1})
            # overview skips both
            cfg_o = common.load_config(os.path.join(root, "tazrim.config.json"), level="overview")
            s = quiet(classify.run, cfg_o)
            self.assertEqual(s["estimate_rows"], 0)
            self.assertEqual(s["trips"], {})
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_default_scheme_when_no_template(self):
        import classify

        def edit(raw):
            raw["categories"]["from_template"] = None
        root, cfg = build_project(edit=edit)
        try:
            s = quiet(classify.run, cfg)
            self.assertEqual(s["scheme"], "default")
            self.assertGreaterEqual(s["scheme_categories"], 60)
            rows = db_rows(cfg)
            rent = next(r for r in rows if r["original_name"] == "שכר דירה לדוגמה")
            self.assertEqual(rent["group_tz"], "דיור")
            self.assertIn("סופרמרקט", s["categories_not_in_scheme"])   # fixture rule cat is not in the default scheme
            self.assertNotIn("משכורת", s["categories_not_in_scheme"])  # default scheme has income categories
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_starter_rules_fallback_and_longest_match(self):
        """Without rules/merchant_rules.csv the shipped starter is used; the longest pattern wins."""
        import classify
        root, cfg = build_project(edit=lambda raw: raw["categories"].update({"from_template": None}))
        try:
            os.remove(os.path.join(root, "rules", "merchant_rules.csv"))
            rules, path, warn = classify.load_rules(cfg)
            self.assertTrue(warn)
            self.assertTrue(path.endswith("merchant-rules-starter.csv"))
            lens = [len(r["pattern"]) for r in rules]
            self.assertEqual(lens, sorted(lens, reverse=True))
            row = {"original_name": "ביטוח לאומי - ילדים", "_section": "accounts", "_entry": {"id": "bank_a"}}
            self.assertEqual(classify.match_rule(row, rules)["cat"], "קצבת ילדים")
            row["original_name"] = "ביטוח לאומי"
            self.assertEqual(classify.match_rule(row, rules)["cat"], "קצבאות ביטוח לאומי")
            s = quiet(classify.run, cfg)
            self.assertTrue(any("starter" in w for w in s["warnings"]))
            self.assertEqual(s["rules"], len(rules))
        finally:
            shutil.rmtree(root, ignore_errors=True)


class TestShippedReferences(unittest.TestCase):
    """The default scheme and the starter rules are consistent and household-neutral."""

    def test_default_scheme(self):
        rows = common.read_csv(os.path.join(SKILL, "references", "category-scheme-default.csv"))
        self.assertEqual(list(rows[0].keys()), ["group", "cat", "fixed"])
        groups = []
        for r in rows:
            self.assertNotIn(",", r["cat"])
            self.assertNotIn(",", r["group"])
            self.assertIn(r["fixed"], ("yes", "no"))
            if r["group"] not in groups:
                groups.append(r["group"])
        expense_groups = [g for g in groups if g not in ("הכנסות", "חיסכון והשקעות", "העברות פנימיות", "תשלומי כרטיסי אשראי", "הוצאות בהחזר")]
        self.assertEqual(len(expense_groups), 14)
        cats = {(r["group"], r["cat"]) for r in rows}
        for must in (("שונות", "אחר / לא מזוהה"), ("הכנסות", "אחר / לא מזוהה"), ("שונות", "הנחת מועדון (זיכוי מחושב)"),
                     ("תשלומי כרטיסי אשראי", "תשלום כרטיס אשראי"), ("העברות פנימיות", "טעינת כרטיס הטבות"),
                     ("הוצאות בהחזר", "הוצאה בהחזר"), ("הוצאות בהחזר", "החזר הוצאה")):
            self.assertIn(must, cats)
        self.assertEqual(len(cats), len(rows))                          # no duplicates

    def test_starter_rules(self):
        scheme = {(r["group"], r["cat"]) for r in common.read_csv(os.path.join(SKILL, "references", "category-scheme-default.csv"))}
        path = os.path.join(SKILL, "references", "merchant-rules-starter.csv")
        rules = common.read_csv(path, skip_comments=True)
        self.assertEqual(list(rules[0].keys()), ["pattern", "name_clean", "type", "group", "cat", "note"])
        self.assertTrue(60 <= len(rules) <= 120, len(rules))
        for r in rules:
            self.assertTrue(r["pattern"].strip())
            self.assertIn((r["group"], r["cat"]), scheme, r["pattern"])
            self.assertIn(r["type"], common.ROW_TYPES)
            self.assertNotRegex(r["pattern"], r"\d{4}")                  # no card last-4 / account numbers
            self.assertNotRegex(r["pattern"], r'^(הו"ק ל|העברה ל|העברה מ|הע\. ל)')
        patterns = [r["pattern"] for r in rules]
        self.assertEqual(len(patterns), len(set(patterns)))


@unittest.skipUnless(HAVE_DEPS, "openpyxl / python-docx not installed")
class TestClassifyCLI(unittest.TestCase):
    def test_cli_contract(self):
        root, cfg = build_project()
        try:
            proc = subprocess.run([sys.executable, os.path.join(SCRIPTS, "classify.py"), "--config", cfg.path, "--level", "standard"],
                                  stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180)
            self.assertEqual(proc.returncode, 0, proc.stderr.decode("utf-8", "replace"))
            lines = [l for l in proc.stdout.decode("utf-8").splitlines() if l.strip()]
            self.assertEqual(len(lines), 1)
            obj = json.loads(lines[0])
            self.assertTrue(obj["ok"])
            self.assertTrue(os.path.isfile(obj["database"]))
            for k in ("by_type", "in_window_by_type", "window_totals", "unknown_in_window", "p2p_pairs"):
                self.assertIn(k, obj)
            shutil.rmtree(os.path.join(root, "work", "normalized"))
            proc = subprocess.run([sys.executable, os.path.join(SCRIPTS, "classify.py"), "--config", cfg.path],
                                  stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180)
            self.assertEqual(proc.returncode, 1)
            obj = json.loads(proc.stdout.decode("utf-8").strip())
            self.assertFalse(obj["ok"])
            self.assertIn("hint", obj)
        finally:
            shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
