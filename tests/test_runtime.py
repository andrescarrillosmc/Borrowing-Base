"""test_runtime.py — Tests for the RuntimeData session model.

Covers three scenarios:
  1. Workbook import -> RuntimeData (file dependency ends here)
  2. Baseline probe from RuntimeData (no file I/O)
  3. Pro forma scenario from RuntimeData (no file I/O)

All integration tests skip automatically if the workbook is absent.

Usage:
    cd C:\\Users\\henry.yan\\borrowing-base
    PYTHONPATH=src python tests/test_runtime.py
"""

from __future__ import annotations

import sys
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from borrowing_base_workbench.engine import (
    import_workbook,
    probe_runtime,
    run_pro_forma_on_runtime,
)
from borrowing_base_workbench.loader import RuntimeData

WORKBOOK = Path(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")
_WB_SKIP = unittest.skipUnless(WORKBOOK.exists(), f"workbook not found: {WORKBOOK}")

# Scenario dict for a clean, eligible First Lien loan
_CLEAN_SCENARIO = {
    "company_name": "Runtime Test Co",
    "security_type": "First Lien",
    "ltm_revenue": "65000000",
    "ltm_adj_ebitda": "12000000",
    "drawn_revolver": "0",
    "first_out_balance": "5000000",
    "pari_passu": "0",
    "bdc_balance": "10000000",
    "total_sm_balance": "10000000",
    "cash_balance": "1000000",
    "interest_coverage": "2.5",
    "loan_denomination": "USD",
    "purchase_price": "1",
    "country": "United States",
    "investment_date": "03/01/2025",
    "maturity_date": "03/01/2030",
    "rate_type": "Floating",
    "industry_classification": "Software",
    "payment_frequency": "M",
    "pik_pct": "0",
    "spread": "0.0875",
    "sofr_floor": "0.02",
    "attach_point": "0.81",
}


# ---------------------------------------------------------------------------
# 1. Workbook import
# ---------------------------------------------------------------------------

class TestImportWorkbook(unittest.TestCase):
    """Workbook import produces RuntimeData with no further file dependency."""

    @_WB_SKIP
    def test_returns_runtime_data_instance(self):
        result = import_workbook(WORKBOOK)
        self.assertIsInstance(result, RuntimeData)

    @_WB_SKIP
    def test_source_path_set(self):
        rt = import_workbook(WORKBOOK)
        self.assertIsInstance(rt.source_path, Path)
        self.assertEqual(rt.source_path, WORKBOOK)

    @_WB_SKIP
    def test_imported_at_is_recent_datetime(self):
        before = datetime.now()
        rt = import_workbook(WORKBOOK)
        after = datetime.now()
        self.assertIsInstance(rt.imported_at, datetime)
        self.assertGreaterEqual(rt.imported_at, before)
        self.assertLessEqual(rt.imported_at, after)

    @_WB_SKIP
    def test_loans_loaded(self):
        rt = import_workbook(WORKBOOK)
        self.assertGreater(len(rt.loans), 0)

    @_WB_SKIP
    def test_obligors_loaded(self):
        rt = import_workbook(WORKBOOK)
        self.assertGreater(len(rt.obligors), 0)

    @_WB_SKIP
    def test_policy_loaded(self):
        rt = import_workbook(WORKBOOK)
        self.assertIsNotNone(rt.policy)
        self.assertGreater(rt.policy.rate_first_lien_large, 0)

    @_WB_SKIP
    def test_availability_meta_loaded(self):
        rt = import_workbook(WORKBOOK)
        self.assertIsNotNone(rt.availability_meta)
        self.assertGreater(rt.availability_meta.facility_amount, 0)

    @_WB_SKIP
    def test_portfolio_flags_loaded(self):
        rt = import_workbook(WORKBOOK)
        self.assertIsInstance(rt.portfolio_flags, dict)

    def test_missing_file_returns_error_dict(self):
        result = import_workbook(Path("/nonexistent/path.xlsm"))
        self.assertIsInstance(result, dict)
        self.assertEqual(result.get("status"), "error")
        self.assertIn("message", result)

    @_WB_SKIP
    def test_summary_contains_expected_keys(self):
        rt = import_workbook(WORKBOOK)
        s = rt.summary()
        for key in ("source_path", "imported_at", "loan_count", "obligor_count",
                    "facility_amount", "current_advances"):
            self.assertIn(key, s)

    @_WB_SKIP
    def test_to_workbook_data_roundtrip(self):
        """RuntimeData.to_workbook_data() must preserve all portfolio records."""
        rt = import_workbook(WORKBOOK)
        wd = rt.to_workbook_data()
        self.assertEqual(len(wd.loans), len(rt.loans))
        self.assertEqual(len(wd.obligors), len(rt.obligors))
        self.assertEqual(len(wd.vaes), len(rt.vaes))


# ---------------------------------------------------------------------------
# 2. Baseline probe from RuntimeData (no file I/O)
# ---------------------------------------------------------------------------

class TestProbeRuntime(unittest.TestCase):
    """Baseline probe runs on cached RuntimeData without touching the workbook."""

    @classmethod
    def setUpClass(cls):
        cls.runtime = import_workbook(WORKBOOK) if WORKBOOK.exists() else None

    @_WB_SKIP
    def test_probe_returns_ok(self):
        result = probe_runtime(self.runtime)
        self.assertEqual(result["status"], "ok")

    @_WB_SKIP
    def test_probe_availability(self):
        result = probe_runtime(self.runtime)
        avail = result["metrics"]["availability"]
        self.assertAlmostEqual(avail, 60_753_941, delta=1_000)

    @_WB_SKIP
    def test_probe_includes_concentration_limits(self):
        result = probe_runtime(self.runtime)
        self.assertIsInstance(result["concentration_limits"], list)
        self.assertGreater(len(result["concentration_limits"]), 0)

    @_WB_SKIP
    def test_probe_result_is_idempotent(self):
        """Calling probe_runtime twice returns identical results (no side effects)."""
        r1 = probe_runtime(self.runtime)
        r2 = probe_runtime(self.runtime)
        self.assertEqual(r1["metrics"]["availability"], r2["metrics"]["availability"])
        self.assertEqual(r1["metrics"]["net_adjusted_bv"], r2["metrics"]["net_adjusted_bv"])

    @_WB_SKIP
    def test_probe_does_not_mutate_runtime(self):
        """probe_runtime must leave the RuntimeData unchanged."""
        loan_count_before = len(self.runtime.loans)
        probe_runtime(self.runtime)
        self.assertEqual(len(self.runtime.loans), loan_count_before)

    @_WB_SKIP
    def test_probe_source_path_in_result(self):
        result = probe_runtime(self.runtime)
        self.assertIn("source_path", result)

    @_WB_SKIP
    def test_probe_imported_at_in_result(self):
        result = probe_runtime(self.runtime)
        self.assertIn("imported_at", result)


# ---------------------------------------------------------------------------
# 3. Scenario run from RuntimeData (no file I/O)
# ---------------------------------------------------------------------------

class TestRunProFormaOnRuntime(unittest.TestCase):
    """Pro forma scenarios execute against cached RuntimeData; no workbook re-read."""

    @classmethod
    def setUpClass(cls):
        cls.runtime = import_workbook(WORKBOOK) if WORKBOOK.exists() else None

    @_WB_SKIP
    def test_scenario_returns_ok(self):
        result = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        self.assertEqual(result["status"], "ok")

    @_WB_SKIP
    def test_scenario_has_before_and_after(self):
        result = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        self.assertIn("before", result)
        self.assertIn("after", result)

    @_WB_SKIP
    def test_scenario_has_eligibility(self):
        result = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        elig = result["eligibility"]
        self.assertIn("status", elig)
        self.assertIn("failed_tests", elig)

    @_WB_SKIP
    def test_clean_scenario_is_eligible(self):
        result = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        self.assertEqual(result["eligibility"]["status"], "Yes")

    @_WB_SKIP
    def test_clean_scenario_increases_availability(self):
        result = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        before = result["before"]["availability"]
        after = result["after"]["availability"]
        self.assertGreater(after, before)

    @_WB_SKIP
    def test_scenario_does_not_mutate_runtime(self):
        """Injecting a scenario must not change the base RuntimeData."""
        loan_count_before = len(self.runtime.loans)
        obligor_count_before = len(self.runtime.obligors)
        run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        self.assertEqual(len(self.runtime.loans), loan_count_before)
        self.assertEqual(len(self.runtime.obligors), obligor_count_before)

    @_WB_SKIP
    def test_multiple_scenarios_from_same_runtime(self):
        """Same RuntimeData baseline must produce consistent before metrics
        regardless of how many scenarios were previously run."""
        r1 = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        r2 = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        self.assertEqual(r1["before"]["availability"], r2["before"]["availability"])

    @_WB_SKIP
    def test_pik_scenario_ineligible(self):
        """A scenario with PIK >= 50% must be ineligible."""
        pik_scenario = {**_CLEAN_SCENARIO, "pik_pct": "0.55"}
        result = run_pro_forma_on_runtime(self.runtime, pik_scenario)
        self.assertEqual(result["eligibility"]["status"], "No")

    @_WB_SKIP
    def test_scenario_before_matches_standalone_probe(self):
        """The 'before' snapshot in a pro forma must equal a standalone probe."""
        probe = probe_runtime(self.runtime)
        scenario_result = run_pro_forma_on_runtime(self.runtime, _CLEAN_SCENARIO)
        probe_avail = probe["metrics"]["availability"]
        before_avail = scenario_result["before"]["availability"]
        self.assertAlmostEqual(probe_avail, before_avail, places=0)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    unittest.main(verbosity=2)
