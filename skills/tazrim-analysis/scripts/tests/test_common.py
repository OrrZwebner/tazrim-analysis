# -*- coding: utf-8 -*-
"""Tests for common.py: config loading + validation, months(), stable_id, labels, csv helpers.
Run: python3 -m unittest discover -s scripts/tests -v
"""
import atexit
import copy
import json
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

import common  # noqa: E402
from common import ConfigError, load_config, months  # noqa: E402

EXAMPLE = os.path.join(os.path.dirname(SCRIPTS), "assets", "config.example.json")


def _load_example():
    with open(EXAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


class TmpConfig(object):
    """Write a config dict to a temp project dir; returns the config path."""

    def __init__(self, data):
        self.dir = tempfile.mkdtemp(prefix="tz_cfg_")
        atexit.register(shutil.rmtree, self.dir, True)  # never leak temp dirs
        self.path = os.path.join(self.dir, "tazrim.config.json")
        with open(self.path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)


class TestConstants(unittest.TestCase):
    def test_columns(self):
        self.assertEqual(len(common.NORMALIZED_COLUMNS), 12)
        self.assertEqual(common.NORMALIZED_COLUMNS[0], "id")
        self.assertEqual(common.NORMALIZED_COLUMNS[-1], "row_ref")
        self.assertEqual(len(common.DB_COLUMNS), 27)
        self.assertEqual(common.DB_COLUMNS[:5], ["id", "source", "card", "pay", "person"])
        self.assertEqual(common.DB_COLUMNS[-1], "trip")
        self.assertTrue(set(common.NORMALIZED_COLUMNS) - {"amount_ils"} <= set(common.DB_COLUMNS))
        for c in common.DB_COLUMNS:
            self.assertIn(c, common.DB_HEADERS_HE)

    def test_row_types(self):
        self.assertEqual(len(common.ROW_TYPES), 8)
        self.assertEqual(common.SUMMED_TYPES, ("הוצאה", "הכנסה"))
        self.assertIn("תשלום כרטיס אשראי", common.NON_SUMMED_TYPES)

    def test_levels_and_sheets(self):
        self.assertEqual(common.LEVELS, ("overview", "standard", "deep"))
        ov, st, dp = (common.SHEETS_BY_LEVEL[l] for l in common.LEVELS)
        self.assertEqual(len(ov), 5)
        self.assertTrue(set(ov) < set(st) < set(dp))
        self.assertEqual(set(dp) - set(st), {"קטגוריות מוצעות"})
        self.assertIn("לסיווג ידני", st)
        self.assertNotIn("לסיווג ידני", ov)
        for s in dp:
            self.assertIn(s, common.SHEET_ORDER)

    def test_hebrew_months(self):
        self.assertEqual(common.HEBREW_MONTHS[1], "ינואר")
        self.assertEqual(common.HEBREW_MONTHS[12], "דצמבר")
        self.assertEqual(common.month_label("2025-02"), "פברואר 2025")
        self.assertEqual(common.month_label("garbage"), "garbage")


class TestLoadConfig(unittest.TestCase):
    def test_example_loads_with_defaults(self):
        cfg = load_config(EXAMPLE)
        self.assertEqual(cfg.level, "standard")
        self.assertEqual(cfg["window"]["start"], "2025-01-01")
        self.assertEqual(cfg["thresholds"]["highlight_expense_avg"], 3000)
        self.assertEqual(cfg["thresholds"]["highlight_income_avg"], 12000)
        self.assertEqual(cfg["verify"]["excel_recalc"], "auto")
        self.assertEqual(cfg["dashboard"]["chartjs"], "cdn")
        self.assertEqual(cfg["p2p"][0]["match_days"], 5)
        self.assertEqual(cfg["p2p"][0]["match_tolerance"], 0.01)
        self.assertEqual(cfg["work"]["database"], "work/database.csv")
        self.assertEqual(cfg["outputs"]["workbook"], "outputs/תזרים.xlsx")
        self.assertEqual(cfg.project_dir, os.path.dirname(EXAMPLE))
        # the example points at files that do not exist next to assets/ -> warnings, not errors
        self.assertTrue(any("no files match" in w for w in cfg.warnings))

    def test_minimal_config_gets_every_default(self):
        t = TmpConfig({"window": {"start": "2025-03-01", "end": "2025-03-31"},
                       "household": {"people": [{"id": "p1", "label": "א'"}]}})
        cfg = load_config(t.path)
        for key in ("accounts", "cards", "benefit_programs", "p2p", "trips", "estimates"):
            self.assertEqual(cfg[key], [])
        self.assertEqual(cfg["categories"]["scheme"], "default")
        self.assertEqual(cfg["fx"]["rates_dir"], "work/boi_rates")
        self.assertEqual(cfg["rules"]["merchant_rules"], "rules/merchant_rules.csv")
        self.assertEqual(cfg["household"]["shared_label"], 'משותף (עו"ש)')
        self.assertEqual(cfg.warnings, [])

    def test_level_override(self):
        cfg = load_config(EXAMPLE, level="deep")
        self.assertEqual(cfg.level, "deep")
        self.assertEqual(cfg["analysis_level"], "deep")

    def test_project_dir_override(self):
        d = tempfile.mkdtemp()
        cfg = load_config(EXAMPLE, project_dir=d)
        self.assertEqual(cfg.project_dir, os.path.abspath(d))
        self.assertEqual(common.project_path(cfg, "outputs.workbook"), os.path.join(d, "outputs", "תזרים.xlsx"))
        self.assertEqual(common.project_path(cfg, "foo/bar.csv"), os.path.join(d, "foo", "bar.csv"))
        self.assertEqual(common.project_path(cfg, "/abs/x.csv"), "/abs/x.csv")

    def _bad(self, mutate, needle):
        data = _load_example()
        mutate(data)
        t = TmpConfig(data)
        with self.assertRaises(ConfigError) as cm:
            load_config(t.path)
        self.assertIn(needle, str(cm.exception))
        self.assertIn("|", str(cm.exception))  # Hebrew | English

    def test_missing_file(self):
        with self.assertRaises(ConfigError):
            load_config("/nonexistent/tazrim.config.json")

    def test_invalid_json(self):
        d = tempfile.mkdtemp()
        p = os.path.join(d, "tazrim.config.json")
        with open(p, "w") as fh:
            fh.write("{not json")
        with self.assertRaises(ConfigError) as cm:
            load_config(p)
        self.assertIn("JSON", str(cm.exception))

    def test_missing_window_start(self):
        self._bad(lambda d: d["window"].pop("start"), "window.start")

    def test_missing_people(self):
        self._bad(lambda d: d["household"].pop("people"), "household.people")

    def test_person_without_label(self):
        self._bad(lambda d: d["household"]["people"].append({"id": "p3"}), "label")

    def test_bad_level(self):
        self._bad(lambda d: d.update(analysis_level="ultra"), "ultra")

    def test_bad_level_via_arg(self):
        with self.assertRaises(ConfigError):
            load_config(EXAMPLE, level="nope")

    def test_window_end_before_start(self):
        self._bad(lambda d: d["window"].update(end="2024-12-31"), "window.end")

    def test_bad_date_format(self):
        self._bad(lambda d: d["window"].update(start="01/01/2025"), "YYYY-MM-DD")

    def test_bad_month_rule(self):
        self._bad(lambda d: d["window"].update(month_rule="txn_date"), "month_rule")

    def test_unknown_owner(self):
        self._bad(lambda d: d["cards"][0].update(owner="p9"), "p9")

    def test_settles_from_not_account(self):
        self._bad(lambda d: d["cards"][0].update(settles_from="card_5678"), "settles_from")

    def test_benefit_settles_from_not_account(self):
        self._bad(lambda d: d["benefit_programs"][0].update(settles_from="nope"), "settles_from")

    def test_unknown_bank(self):
        self._bad(lambda d: d["accounts"][0].update(bank="leumi"), "leumi")

    def test_unknown_issuer(self):
        self._bad(lambda d: d["cards"][0].update(issuer="bank_xlsx_a"), "bank_xlsx_a")

    def test_duplicate_id(self):
        self._bad(lambda d: d["cards"][1].update(id="card_1234"), "card_1234")

    def test_bad_cycle_day(self):
        self._bad(lambda d: d["cards"][0].update(cycle_day=31), "cycle_day")

    def test_bad_p2p_thresholds(self):
        self._bad(lambda d: d["p2p"][0].update(match_days=-1), "match_days")
        self._bad(lambda d: d["p2p"][0].update(match_tolerance="x"), "match_tolerance")

    def test_bad_enums(self):
        self._bad(lambda d: d["verify"].update(excel_recalc="maybe"), "excel_recalc")
        self._bad(lambda d: d["dashboard"].update(chartjs="local"), "chartjs")
        self._bad(lambda d: d["fx"].update(source="ecb"), "fx.source")

    def test_glob_zero_files_is_a_warning(self):
        data = _load_example()
        t = TmpConfig(data)
        cfg = load_config(t.path)
        self.assertTrue(any("accounts[0].files" in w for w in cfg.warnings))


class TestMonths(unittest.TestCase):
    def _cfg(self, start, end):
        t = TmpConfig({"window": {"start": start, "end": end},
                       "household": {"people": [{"id": "p1", "label": "א'"}]}})
        return load_config(t.path)

    def test_one_month(self):
        self.assertEqual(months(self._cfg("2025-03-01", "2025-03-31")), ["2025-03"])

    def test_two_months(self):
        self.assertEqual(months(self._cfg("2025-01-01", "2025-02-28")), ["2025-01", "2025-02"])

    def test_seven_months_across_year_end(self):
        self.assertEqual(months(self._cfg("2024-10-01", "2025-04-30")),
                         ["2024-10", "2024-11", "2024-12", "2025-01", "2025-02", "2025-03", "2025-04"])

    def test_window_helpers(self):
        cfg = self._cfg("2025-01-01", "2025-02-28")
        self.assertTrue(common.in_window("2025-02-28", cfg))
        self.assertFalse(common.in_window("2025-03-01", cfg))
        self.assertFalse(common.in_window("2024-12-31", cfg))
        self.assertEqual(common.analysis_month("2025-02-10", cfg), "2025-02")
        self.assertEqual(common.flag(True), "כן")
        self.assertEqual(common.flag(False), "לא")


class TestIdsAndLabels(unittest.TestCase):
    def test_stable_id_deterministic(self):
        a = common.stable_id("C1", "/x/y/card_0225.xlsx", 12)
        b = common.stable_id("C1", "/other/dir/card_0225.xlsx", "12")
        self.assertEqual(a, b)                       # basename + row_ref only
        self.assertRegex(a, r"^C1-[0-9a-f]{6}$")
        self.assertNotEqual(a, common.stable_id("C1", "card_0225.xlsx", 13))
        self.assertNotEqual(a, common.stable_id("C2", "card_0225.xlsx", 12))

    def test_labels(self):
        cfg = load_config(EXAMPLE)
        self.assertEqual(common.owner_label(cfg, "shared"), 'משותף (עו"ש)')
        self.assertEqual(common.owner_label(cfg, "p1"), "בן/בת זוג א'")
        self.assertEqual(common.owner_label(cfg, "zz"), "zz")
        self.assertEqual(common.pay_label(cfg, "card_1234"), "ויזה 1234 (בן/בת זוג א')")
        self.assertEqual(common.pay_label(cfg, "bank_a"), "עו\"ש חשבון א' (משותף (עו\"ש))")
        self.assertEqual(common.pay_label(cfg, "ויזה 5678"), "ויזה 5678 (בן/בת זוג ב')")
        self.assertTrue(common.pay_label(cfg, "unknown").endswith(")"))
        self.assertEqual(common.person_of(cfg, "card_9012"), "בן/בת זוג ב'")

    def test_scheme_path_default(self):
        cfg = load_config(EXAMPLE)
        self.assertTrue(common.scheme_path(cfg).endswith(os.path.join("references", "category-scheme-default.csv")))


class TestCsvHelpers(unittest.TestCase):
    def test_roundtrip_and_comments(self):
        d = tempfile.mkdtemp()
        p = os.path.join(d, "x.csv")
        rows = [{"id": "A-1", "amount_ils": "1.00", "extra": "ignored"}, {"id": "B-2"}]
        n = common.write_csv(p, rows, ["id", "amount_ils"])
        self.assertEqual(n, 2)
        back = common.read_csv(p)
        self.assertEqual(back, [{"id": "A-1", "amount_ils": "1.00"}, {"id": "B-2", "amount_ils": ""}])
        with open(p, "a", encoding="utf-8") as fh:
            fh.write("# comment\nC-3,3\n")
        self.assertEqual(len(common.read_csv(p, skip_comments=True)), 3)

    def test_ensure_dirs(self):
        d = tempfile.mkdtemp()
        cfg = load_config(EXAMPLE, project_dir=d)
        made = common.ensure_dirs(cfg)
        for m in made:
            self.assertTrue(os.path.isdir(m))
        self.assertTrue(os.path.isdir(os.path.join(d, "work", "normalized")))
        self.assertTrue(os.path.isdir(os.path.join(d, "rules")))


class TestParseArgs(unittest.TestCase):
    def test_parse_args_with_example(self):
        args, cfg = common.parse_args("t", argv=["--config", EXAMPLE, "--level", "overview"])
        self.assertEqual(cfg.level, "overview")
        self.assertEqual(args.level, "overview")

    def test_parse_args_missing_config_exits_2(self):
        import io
        old = sys.stdout
        sys.stdout = io.StringIO()
        try:
            with self.assertRaises(SystemExit) as cm:
                common.parse_args("t", argv=["--config", "/nope/tazrim.config.json"])
            out = sys.stdout.getvalue()
        finally:
            sys.stdout = old
        self.assertEqual(cm.exception.code, 2)
        obj = json.loads(out.strip().splitlines()[-1])
        self.assertFalse(obj["ok"])
        self.assertIn("hint", obj)


class TestOptionalEntryKeys(unittest.TestCase):
    """Optional per-entry keys read by the parsers/classifier get defaults and are validated."""

    def _cfg(self, **over):
        ex = _load_example()
        ex.update(over)
        return TmpConfig(ex).path

    def test_example_carries_every_optional_key(self):
        cfg = load_config(EXAMPLE)
        self.assertEqual(cfg["cards"][0]["fx_settlement"]["transfer_pattern"], 'העברה לחשבון מט"ח')
        self.assertIsNone(cfg["cards"][1]["fx_settlement"]["transfer_pattern"])   # defaulted
        self.assertIs(cfg["cards"][2]["reconstruct_missing_cycles"], False)
        self.assertIs(cfg["cards"][0]["reconstruct_missing_cycles"], False)       # defaulted
        self.assertEqual(cfg["p2p"][0]["bank_withdrawal_pattern"], "משיכה מ-BIT")
        est = cfg["estimates"][0]
        self.assertEqual((est["every_n_months"], est["day"], est["funded_by_cat"]), (1, 15, None))
        self.assertEqual(cfg["trips"][0]["countries"], ["ארצות הברית", "USA"])
        for k in ("dynamic_merchant_rows", "dynamic_txn_rows", "top_merchants", "hyperlink_rows"):
            self.assertIsInstance(cfg["thresholds"][k], int)

    def test_estimate_defaults_and_validation(self):
        cfg = load_config(self._cfg(estimates=[{"name": "מזומן", "amount": 150}]))
        e = cfg["estimates"][0]
        self.assertEqual((e["every_n_months"], e["day"], e["cat"], e["group"], e["note"], e["funded_by_cat"], e["id"]),
                         (1, 15, common.UNKNOWN_CAT, None, "", None, None))
        with self.assertRaises(ConfigError):
            load_config(self._cfg(estimates=[{"amount": 100}]))
        with self.assertRaises(ConfigError):
            load_config(self._cfg(estimates=[{"name": "x", "amount": "100"}]))
        with self.assertRaises(ConfigError):
            load_config(self._cfg(estimates=[{"name": "x", "amount": 1, "day": 31}]))
        with self.assertRaises(ConfigError):
            load_config(self._cfg(estimates=[{"name": "x", "amount": 1, "every_n_months": 0}]))

    def test_card_optional_keys_validation(self):
        ex = _load_example()
        ex["cards"][0]["reconstruct_missing_cycles"] = "yes"
        with self.assertRaises(ConfigError):
            load_config(TmpConfig(ex).path)
        ex = _load_example()
        ex["cards"][0]["fx_settlement"] = {"transfer_pattern": "x"}          # currency missing
        with self.assertRaises(ConfigError):
            load_config(TmpConfig(ex).path)
        ex = _load_example()
        ex["cards"][0]["fx_settlement"] = {"currency": "usd"}
        cfg = load_config(TmpConfig(ex).path)
        self.assertEqual(cfg["cards"][0]["fx_settlement"], {"currency": "USD", "transfer_pattern": None})
        ex = _load_example()
        ex["trips"][0]["countries"] = "USA"
        with self.assertRaises(ConfigError):
            load_config(TmpConfig(ex).path)


class TestGeneratedCategories(unittest.TestCase):
    """Every category a script can assign without a rule is in the shipped scheme and the
    starter rules only use scheme categories; no category name carries a comma (FM5)."""

    def _scheme(self):
        import csv
        path = os.path.join(os.path.dirname(SCRIPTS), "references", "category-scheme-default.csv")
        with open(path, encoding="utf-8-sig") as fh:
            rows = [r for r in csv.DictReader(fh) if r.get("cat")]
        return rows

    def test_generated_categories_in_default_scheme(self):
        rows = self._scheme()
        pairs = {(r["group"], r["cat"]) for r in rows}
        for cat, group in common.GENERATED_CATEGORIES.items():
            self.assertIn((group, cat), pairs, "%s / %s missing from category-scheme-default.csv" % (group, cat))
        self.assertIn((common.UNKNOWN_GROUP_INCOME, common.UNKNOWN_CAT), pairs)

    def test_starter_rules_use_scheme_categories(self):
        import csv
        cats = {r["cat"] for r in self._scheme()}
        path = os.path.join(os.path.dirname(SCRIPTS), "references", "merchant-rules-starter.csv")
        with open(path, encoding="utf-8-sig") as fh:
            rules = [r for r in csv.DictReader(ln for ln in fh if not ln.startswith("#")) if r.get("cat")]
        self.assertGreater(len(rules), 20)
        self.assertEqual(sorted({r["cat"] for r in rules} - cats), [])

    def test_no_commas_in_category_names(self):
        for r in self._scheme():
            self.assertNotIn(",", r["cat"])
            self.assertNotIn(",", r["group"])
        for c in common.GENERATED_CATEGORIES:
            self.assertNotIn(",", c)


if __name__ == "__main__":
    unittest.main()
