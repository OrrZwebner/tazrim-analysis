# -*- coding: utf-8 -*-
"""Tests for make_summary.py on the synthetic project of synth_db_outputs.py.
Every expected value is a hand-derived literal from synth_db_outputs.EXPECTED.
Run: python3 -m unittest discover -s scripts/tests -v
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)
sys.path.insert(0, HERE)

import common  # noqa: E402
import synth_db_outputs as synth  # noqa: E402

try:
    import pandas  # noqa: F401
    import make_summary
except ImportError:  # pragma: no cover
    make_summary = None


@unittest.skipIf(make_summary is None, "pandas not installed")
class TestSummary(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="tz_sum_")
        cls.cfg_path, cls.E = synth.make_project(cls.tmp)
        cls.cfg = common.load_config(cls.cfg_path)
        cls.S = make_summary.build_summary(cls.cfg)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_window_and_counts(self):
        S, E = self.S, self.E
        self.assertEqual(S["n_months"], E["n_months"])
        self.assertEqual(S["months"], E["months"])
        self.assertEqual(S["n_rows_db"], E["n_rows_db"])
        self.assertEqual(S["n_expense_rows"], E["n_expense_rows"])
        self.assertEqual(S["n_income_rows"], E["n_income_rows"])

    def test_totals_and_series(self):
        S, E = self.S, self.E
        self.assertAlmostEqual(S["total_expenses"], E["total_expenses"], 2)
        self.assertAlmostEqual(S["total_income"], E["total_income"], 2)
        self.assertEqual(S["expenses_by_month"], E["expenses_by_month"])
        self.assertEqual(S["income_by_month"], E["income_by_month"])
        self.assertAlmostEqual(S["avg_expenses"], E["avg_expenses"], 2)
        self.assertAlmostEqual(S["avg_income"], E["avg_income"], 2)
        self.assertAlmostEqual(S["savings_total"], E["savings_total"], 2)

    def test_groups_and_category_stats(self):
        S, E = self.S, self.E
        self.assertEqual({k: v["total"] for k, v in S["expenses_by_group"].items()}, E["expenses_by_group"])
        # groups sorted by total desc
        totals = [v["total"] for v in S["expenses_by_group"].values()]
        self.assertEqual(totals, sorted(totals, reverse=True))
        st = S["expenses_by_cat"]["דיור / חשמל"]
        ex = E["cat_electricity"]
        for k in ("total", "avg", "avg_nz", "median", "stdev"):
            self.assertAlmostEqual(st[k], ex[k], 2, k)
        self.assertEqual(st["max_month"], ex["max_month"])
        self.assertEqual(st["min_month"], ex["min_month"])
        self.assertEqual(st["n_nonzero"], ex["n_nonzero"])
        self.assertAlmostEqual(sum(v["total"] for v in S["expenses_by_cat"].values()), E["total_expenses"], 2)

    def test_fixed_variable_israel_abroad(self):
        S, E = self.S, self.E
        self.assertAlmostEqual(S["fixed_total"], E["fixed_total"], 2)
        self.assertAlmostEqual(S["variable_total"], E["variable_total"], 2)
        self.assertAlmostEqual(S["abroad_total"], E["abroad_total"], 2)
        self.assertAlmostEqual(S["israel_total"], E["israel_total"], 2)
        self.assertEqual(S["trips"]["נסיעה לדוגמה"]["total"], E["trip_total"]["נסיעה לדוגמה"])

    def test_by_person_card_pay(self):
        S, E = self.S, self.E
        self.assertEqual({k: v["total"] for k, v in S["expenses_by_person"].items()}, E["expenses_by_person"])
        for p, v in E["income_by_person"].items():
            self.assertAlmostEqual(S["income_by_person"][p]["total"], v, 2)
        self.assertEqual({k: v["total"] for k, v in S["expenses_by_card"].items()}, E["expenses_by_card"])
        self.assertAlmostEqual(sum(v["total"] for v in S["expenses_by_pay"].values()), E["total_expenses"], 2)
        self.assertEqual(S["people"][0], synth.P1)  # config order

    def test_merchants_and_special_lists(self):
        S, E = self.S, self.E
        top = S["top_merchants"][0]
        self.assertEqual(top["name"], E["top_merchant"]["name"])
        self.assertAlmostEqual(top["total"], E["top_merchant"]["total"], 2)
        sup = [m for m in S["top_merchants"] if m["name"] == "סופר לדוגמה"][0]
        self.assertAlmostEqual(sup["total"], E["super_total"], 2)
        self.assertEqual(sup["count"], E["super_count"])
        self.assertTrue(len(S["top_merchants"]) <= 25 and len(S["top_merchants_by_avg_nz"]) <= 25)
        self.assertEqual(len(S["possible_double_charges"]), E["n_double_charges"])
        self.assertEqual(len(S["refunds"]), E["n_refunds"])
        self.assertEqual(S["n_unknown"], E["n_unknown"])
        self.assertAlmostEqual(S["unknown_total"], E["unknown_total"], 2)

    def test_components(self):
        S, E = self.S, self.E
        self.assertAlmostEqual(S["internal_bank"]["net"], E["internal_bank_net"], 2)
        self.assertAlmostEqual(S["card_debits"]["total"], E["card_debits_total"], 2)
        self.assertAlmostEqual(S["recon"]["card_vs_bank_diff"], E["card_vs_bank_diff"], 2)
        P = S["p2p"]
        self.assertAlmostEqual(P["income"], E["p2p"]["income"], 2)
        self.assertAlmostEqual(P["expenses"], E["p2p"]["expenses"], 2)
        self.assertAlmostEqual(P["paid_from_balance"], E["p2p"]["paid_from_balance"], 2)
        self.assertEqual(P["matched_pairs"], E["p2p"]["matched_pairs"])
        club = S["benefit_programs"][0]
        self.assertAlmostEqual(club["face_net"], E["club"]["face_net"], 2)
        self.assertAlmostEqual(club["bank_debits"], E["club"]["bank"], 2)
        self.assertAlmostEqual(club["balance_change"], E["club"]["diff"], 2)
        self.assertAlmostEqual(S["fx"]["expenses"], E["fx_expenses"], 2)
        R = S["reimbursables"]
        self.assertAlmostEqual(R["paid_total"], E["reimb_paid_bank"], 2)
        self.assertAlmostEqual(R["received_total"], E["reimb_recv_bank"], 2)
        self.assertAlmostEqual(R["open_balance"], E["reimb_open"], 2)
        self.assertAlmostEqual(S["nonstatement_expenses"]["total"], E["other_expenses"], 2)

    def test_reconciliation_identity_closes(self):
        rc, E = self.S["recon"], self.E
        self.assertEqual(rc["bank_change_source"], "balances")
        self.assertAlmostEqual(rc["bank_change_from_rows"], E["bank_change_rows"], 2)
        self.assertAlmostEqual(rc["bank_change_from_balances"], E["bank_change_balances"], 2)
        self.assertAlmostEqual(rc["explained"], E["bank_change_rows"], 2)
        self.assertAlmostEqual(rc["residual"], E["residual"], 2)
        self.assertTrue(rc["residual_ok"])
        self.assertIn("residual", rc["formula"])
        keys = [c["key"] for c in rc["components"]]
        for k in ("income", "expenses", "savings_bank", "internal_bank_net", "fx_expenses", "p2p_income",
                  "p2p_paid_from_balance", "club_balance_change", "nonstatement_expenses"):
            self.assertIn(k, keys)

    def test_residual_reported_when_identity_breaks(self):
        """Dropping a bank income row from the database (balances untouched) must surface a residual."""
        tmp = tempfile.mkdtemp(prefix="tz_sum_b_")
        try:
            cfg_path, _ = synth.make_project(tmp)
            db = os.path.join(tmp, "work", "database.csv")
            rows = common.read_csv(db)
            rows = [r for r in rows if r["id"] != "T-002"]  # 8,000 income in bank2
            common.write_csv(db, rows, common.DB_COLUMNS)
            S = make_summary.build_summary(common.load_config(cfg_path))
            self.assertFalse(S["recon"]["residual_ok"])
            self.assertAlmostEqual(S["recon"]["residual"], 8000.0, 2)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_rows_fallback_without_balances(self):
        tmp = tempfile.mkdtemp(prefix="tz_sum_c_")
        try:
            cfg_path, E = synth.make_project(tmp)
            shutil.rmtree(os.path.join(tmp, "work", "normalized"))
            S = make_summary.build_summary(common.load_config(cfg_path))
            self.assertEqual(S["recon"]["bank_change_source"], "rows")
            self.assertAlmostEqual(S["recon"]["bank_change"], E["bank_change_rows"], 2)
            self.assertAlmostEqual(S["recon"]["residual"], 0.0, 2)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_cli_writes_summary_and_digest(self):
        out = os.path.join(self.tmp, "work", "summary.json")
        p = subprocess.run([sys.executable, os.path.join(SCRIPTS, "make_summary.py"), "--config", self.cfg_path],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        self.assertEqual(p.returncode, 0, p.stderr)
        lines = [l for l in p.stdout.strip().splitlines() if l.strip()]
        self.assertEqual(len(lines), 1, "exactly one JSON object on stdout")
        d = json.loads(lines[0])
        self.assertTrue(d["ok"])
        self.assertEqual(d["n_months"], 2)
        self.assertAlmostEqual(d["recon"]["residual"], 0.0, 2)
        self.assertTrue(os.path.isfile(out))
        S = common.read_json(out)
        self.assertEqual(S["n_rows_db"], self.E["n_rows_db"])

    def test_cli_missing_database_fails_cleanly(self):
        tmp = tempfile.mkdtemp(prefix="tz_sum_d_")
        try:
            cfg_path, _ = synth.make_project(tmp)
            os.remove(os.path.join(tmp, "work", "database.csv"))
            p = subprocess.run([sys.executable, os.path.join(SCRIPTS, "make_summary.py"), "--config", cfg_path],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            self.assertNotEqual(p.returncode, 0)
            d = json.loads(p.stdout.strip().splitlines()[-1])
            self.assertFalse(d["ok"])
            self.assertIn("hint", d)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
