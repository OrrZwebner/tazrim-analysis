# -*- coding: utf-8 -*-
"""Tests for verify.py on the synthetic workbook: pandas-only fallback (Excel checks reported as
skipped), exact total-row lookup, failure detection, and — only when TAZRIM_TEST_EXCEL=1 on a
Mac with Microsoft Excel — the full AppleScript recalculation gate.
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
import synth_db  # noqa: E402
from synth_db import EXPECTED  # noqa: E402

try:
    import openpyxl
    import verify
except ImportError:  # pragma: no cover
    openpyxl = verify = None

BUILD = os.path.join(SCRIPTS, "build_excel.py")
TRANSFERS = os.path.join(SCRIPTS, "add_transfers_sheet.py")
VERIFY = os.path.join(SCRIPTS, "verify.py")


def run(script, *args):
    r = subprocess.run([sys.executable, script] + list(args), capture_output=True, text=True, timeout=900)
    lines = [ln for ln in r.stdout.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, "%s printed %d stdout lines" % (script, len(lines))
    return r.returncode, json.loads(lines[0]), r.stderr


def status_of(res, prefix):
    return {c["name"]: c["status"] for c in res["checks"] if c["name"].startswith(prefix)}


@unittest.skipIf(openpyxl is None, "openpyxl not installed")
class TestVerifyFallback(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = tempfile.mkdtemp(prefix="tz_verify_")
        cls.cfg_path, cls.db_path = synth_db.write_project(cls.root)
        rc, res, err = run(BUILD, "--config", cls.cfg_path, "--level", "standard")
        assert rc == 0, err[-800:]
        cls.xlsx = res["workbook"]
        rc, res, err = run(TRANSFERS, "--config", cls.cfg_path, "--level", "standard")
        assert rc == 0, err[-800:]

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.root, ignore_errors=True)

    def test_fallback_ok_with_excel_skipped(self):
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--level", "standard", "--excel-recalc", "never")
        self.assertEqual(rc, 0, err[-800:])
        self.assertTrue(res["ok"])
        self.assertFalse(res["excel_recalc"]["ran"])
        excel = status_of(res, "excel:")
        self.assertGreaterEqual(len(excel), 4)
        self.assertEqual(set(excel.values()), {"skipped"})
        for c in res["checks"]:
            if c["status"] == "skipped":
                self.assertIn("never", c["detail"])
        self.assertEqual(res["summary"]["fail"], 0)
        self.assertGreaterEqual(res["summary"]["pass"], 10)
        # every non-Excel check passed
        self.assertTrue(all(c["status"] == "pass" for c in res["checks"] if not c["name"].startswith("excel:")))

    def test_expected_totals_match_hand_constants(self):
        cfg = common.load_config(self.cfg_path)
        db = verify.load_db(self.db_path)
        t = verify.expected_totals(cfg, db)
        self.assertEqual(t["expense"], EXPECTED["expense_total"])
        self.assertEqual(t["income"], EXPECTED["income_total"])
        self.assertEqual(t["expense_by_month"], EXPECTED["expense_by_month"])
        self.assertEqual(t["income_by_month"], EXPECTED["income_by_month"])
        self.assertEqual(t["expense_by_category"], EXPECTED["cat_totals"])
        self.assertEqual(t["savings"], EXPECTED["savings_total"])
        self.assertEqual(t["internal"], EXPECTED["internal_total"])
        self.assertEqual(t["card_debits"], EXPECTED["card_debits_total"])
        self.assertEqual(t["reimbursable_open"], EXPECTED["reimbursable_open_balance"])

    def test_exact_total_row_lookup_ignores_subtotals(self):
        wb = openpyxl.load_workbook(self.xlsx)
        wi = wb[common.SHEET_INC]
        label = 'סה"כ הכנסות (נסכם ישירות מ-database)'
        tr = verify.find_total_row(wi, label)
        subtotal_rows = [r for r in range(3, tr) if str(wi.cell(r, 2).value).startswith('סה"כ הכנסות')]
        self.assertTrue(subtotal_rows, "a subtotal sharing the prefix must exist above the total")
        self.assertNotIn(tr, subtotal_rows)
        self.assertEqual(wi.cell(tr, 2).value, label)
        # synthetic sheet: a subtotal with the same prefix comes first; startswith would pick it
        ws = openpyxl.Workbook().active
        ws["B3"] = 'סה"כ הכנסות בן/בת זוג א׳'
        ws["B5"] = label
        self.assertEqual(verify.find_total_row(ws, label), 5)
        self.assertIsNone(verify.find_total_row(ws, "לא קיים"))

    def test_tampered_total_label_fails(self):
        bad = os.path.join(self.root, "outputs", "bad.xlsx")
        shutil.copy(self.xlsx, bad)
        wb = openpyxl.load_workbook(bad)
        we = wb[common.SHEET_EXP]
        tr = verify.find_total_row(we, verify.TOTAL_EXP_LABEL)
        we.cell(tr, 2, verify.TOTAL_EXP_LABEL + " (ישן)")
        wb.save(bad)
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--level", "standard", "--excel-recalc", "never", "--workbook", bad)
        self.assertEqual(rc, 1)
        self.assertFalse(res["ok"])
        failed = [c["name"] for c in res["checks"] if c["status"] == "fail"]
        self.assertTrue(any("total row" in n for n in failed), failed)

    def test_missing_named_range_fails(self):
        bad = os.path.join(self.root, "outputs", "noname.xlsx")
        shutil.copy(self.xlsx, bad)
        wb = openpyxl.load_workbook(bad)
        del wb.defined_names["DBcat"]
        wb.save(bad)
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--level", "standard", "--excel-recalc", "never", "--workbook", bad)
        self.assertEqual(rc, 1)
        named = [c for c in res["checks"] if c["name"].startswith("named ranges")][0]
        self.assertEqual(named["status"], "fail")
        self.assertIn("DBcat", named["detail"])

    def test_formula_error_scan(self):
        bad = os.path.join(self.root, "outputs", "ref.xlsx")
        shutil.copy(self.xlsx, bad)
        wb = openpyxl.load_workbook(bad)
        wb[common.SHEET_ANALYSIS]["B4"] = "=#REF!+1"
        wb.save(bad)
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--level", "standard", "--excel-recalc", "never", "--workbook", bad)
        self.assertEqual(rc, 1)
        scan = [c for c in res["checks"] if c["name"].startswith("no #REF!")][0]
        self.assertEqual(scan["status"], "fail")
        self.assertIn("B4", scan["detail"])

    def test_missing_workbook_fails_fast(self):
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--workbook", os.path.join(self.root, "nope.xlsx"))
        self.assertEqual(rc, 1)
        self.assertIn("workbook not found", res["error"])

    def test_auto_mode_off_mac_reports_reason(self):
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--level", "overview", "--excel-recalc", "never",
                           "--workbook", self.xlsx)
        self.assertEqual(rc, 0)
        self.assertEqual(res["excel_recalc"]["mode"], "never")
        self.assertTrue(res["excel_recalc"]["reason"])


@unittest.skipUnless(os.environ.get("TAZRIM_TEST_EXCEL") == "1" and verify is not None and verify.excel_available(),
                     "set TAZRIM_TEST_EXCEL=1 on a Mac with Microsoft Excel to run the recalculation gate")
class TestVerifyWithExcel(unittest.TestCase):
    """The project must live under the user's home: sandboxed Excel cannot save into /tmp."""

    @classmethod
    def setUpClass(cls):
        cls.root = tempfile.mkdtemp(prefix="tz_xl_", dir=os.path.join(os.path.expanduser("~"), "Library", "Caches"))
        cls.cfg_path, cls.db_path = synth_db.write_project(cls.root)
        rc, res, err = run(BUILD, "--config", cls.cfg_path, "--level", "deep")
        assert rc == 0, err[-800:]
        cls.xlsx = res["workbook"]
        rc, res, err = run(TRANSFERS, "--config", cls.cfg_path, "--level", "deep")
        assert rc == 0, err[-800:]

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.root, ignore_errors=True)

    def test_excel_recalculation_gate(self):
        rc, res, err = run(VERIFY, "--config", self.cfg_path, "--level", "deep", "--excel-recalc", "always")
        self.assertEqual(rc, 0, err[-1500:])
        self.assertTrue(res["excel_recalc"]["ran"])
        excel = status_of(res, "excel:")
        self.assertEqual(set(excel.values()), {"pass"}, excel)
        self.assertTrue(any("live filter" in n for n in excel))
        self.assertTrue(any("transfers-sheet" in n for n in excel))
        self.assertEqual(res["summary"]["skipped"], 0)
        for name in ("verify_tmp.xlsx", "verify_calc.xlsx"):
            self.assertFalse(os.path.exists(os.path.join(self.root, "work", name)))


if __name__ == "__main__":
    unittest.main()
